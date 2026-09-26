"""Gegevensmodel: wat uit de bronbestanden komt (Dossier) en wat de gebruiker beslist (Beslissingen)."""
from __future__ import annotations

import datetime as dt
import json
from dataclasses import dataclass, field, asdict
from pathlib import Path


@dataclass
class Periode:
    """Eén boekjaar of tussentijdse periode met alle rekeningsaldi.

    Tekenconventie: activa en kosten volgen de boekhouding zoals Belgische detailbalansen ze tonen:
    activa positief (afschrijvingen 2x9 negatief), passiva/eigen vermogen positief,
    opbrengsten positief en kosten negatief.
    """
    begin: dt.date | None
    einde: dt.date
    tussentijds: bool
    rekeningen: dict[str, float]
    titels: dict[str, str] = field(default_factory=dict)
    bedrijfswinst: float | None = None
    winst_voor_belasting: float | None = None
    balanstotaal: float | None = None
    bron: str = ''

    @property
    def maanden(self) -> int:
        if not self.begin:
            return 12
        m = (self.einde.year - self.begin.year) * 12 + self.einde.month - self.begin.month + 1
        return max(1, m)

    def som(self, *prefixen: str, uitgezonderd: tuple[str, ...] = ()) -> float:
        return round(sum(v for k, v in self.rekeningen.items()
                         if k.startswith(prefixen) and not k.startswith(uitgezonderd)), 2)

    def items(self, *prefixen: str, uitgezonderd: tuple[str, ...] = ()) -> list[tuple[str, float]]:
        return sorted((k, v) for k, v in self.rekeningen.items()
                      if k.startswith(prefixen) and not k.startswith(uitgezonderd) and abs(v) > 0.004)


@dataclass
class Actief:
    rekening: str
    omschrijving: str
    datum: dt.date
    aanschafwaarde: float
    netto_boekwaarde: float


@dataclass
class Dossier:
    naam: str = ''
    straat: str = ''
    postcode_gemeente: str = ''
    ondernemingsnummer: str = ''
    activiteit: str = ''
    periodes: list[Periode] = field(default_factory=list)
    activa: list[Actief] = field(default_factory=list)
    afschrijvingstabel_totaal: float | None = None
    bestuurders: list[dict] = field(default_factory=list)
    aandeelhouders: list[dict] = field(default_factory=list)
    openstaande_klanten: float | None = None
    meldingen: list[str] = field(default_factory=list)   # controles en opvallende zaken uit het uitlezen

    def afgesloten(self) -> list[Periode]:
        return sorted([p for p in self.periodes if not p.tussentijds], key=lambda p: p.einde)

    def periode_op(self, datum: dt.date) -> Periode | None:
        return next((p for p in self.periodes if p.einde == datum), None)


@dataclass
class Beslissingen:
    """Alles wat per dossier door de gebruiker bepaald wordt (voorstellen staan in voorstellen.py)."""
    afsluitdatum: dt.date
    kolommen: dict[str, dt.date | None]          # H, K, N (afgesloten jaren) en R (prognose/tussentijds)
    weging: dict[str, float]                      # per kolom, som = 1
    bestuurdersvergoeding: float | None = None    # marktconforme totale kost per jaar; None = niet normaliseren
    marktconforme_huur: float | None = None       # per jaar; None = niet normaliseren
    multiple: float | None = None                 # EBITDA-multiple (Vlerick)
    multiple_bron: str = ''                       # bv. "VLERICK M&A Monitor 2026 (horeca)"
    vastgoed_marktwaarde: float | None = None     # schattingsverslag; None = boekwaarde
    vastgoed_schatting_datum: str = ''
    werkkapitaalcorrectie: float = 0.0
    rc_intrest_neutraliseren: bool = True
    goodwill_jaren: int = 7
    vereist_rendement: float | None = None      # rendement op EV voor goodwill; None = volgens samenstelling actief
    activiteit: str = ''

    def naar_json(self) -> dict:
        d = asdict(self)
        d['afsluitdatum'] = self.afsluitdatum.isoformat()
        d['kolommen'] = {k: (v.isoformat() if v else None) for k, v in self.kolommen.items()}
        return d

    @classmethod
    def van_json(cls, d: dict) -> 'Beslissingen':
        d = dict(d)
        d['afsluitdatum'] = dt.date.fromisoformat(d['afsluitdatum'])
        d['kolommen'] = {k: (dt.date.fromisoformat(v) if v else None) for k, v in d['kolommen'].items()}
        return cls(**d)


def dossier_naar_json(d: Dossier) -> dict:
    def conv(o):
        if isinstance(o, dt.date):
            return o.isoformat()
        raise TypeError(o)
    return json.loads(json.dumps(asdict(d), default=conv))


def dossier_van_json(j: dict) -> Dossier:
    d = Dossier(**{k: v for k, v in j.items() if k not in ('periodes', 'activa')})
    for p in j.get('periodes', []):
        d.periodes.append(Periode(
            begin=dt.date.fromisoformat(p['begin']) if p.get('begin') else None,
            einde=dt.date.fromisoformat(p['einde']), tussentijds=p['tussentijds'],
            rekeningen=p['rekeningen'], titels=p.get('titels', {}), bedrijfswinst=p.get('bedrijfswinst'),
            winst_voor_belasting=p.get('winst_voor_belasting'), balanstotaal=p.get('balanstotaal'),
            bron=p.get('bron', '')))
    for a in j.get('activa', []):
        d.activa.append(Actief(a['rekening'], a['omschrijving'], dt.date.fromisoformat(a['datum']),
                               a['aanschafwaarde'], a['netto_boekwaarde']))
    return d


def bewaar_json(pad: Path, data: dict) -> None:
    pad.write_text(json.dumps(data, indent=1, ensure_ascii=False), encoding='utf-8')
