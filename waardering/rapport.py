"""Kort verslag (Markdown) naast het Excel-bestand: bronnen, keuzes, controles en openstaande punten."""
from __future__ import annotations

import datetime as dt
from pathlib import Path

from .model import Beslissingen, Dossier
from .samenstellen import bedrijfswinst, eur


def schrijf(pad: Path, d: Dossier, b: Beslissingen, ok: list, fout: list, resultaat: dict, meldingen: list,
            signalen: list, kost: float) -> None:
    r = [f'# Waardering {d.naam or "?"} – toestand {b.afsluitdatum.strftime("%d/%m/%Y")}', '',
         f'Opgemaakt op {dt.date.today().strftime("%d/%m/%Y")}. Excel: `{pad.with_suffix(".xlsx").name}`', '']
    r += ['## Resultaat', '', '| | EUR |', '|---|---:|']
    for k, v in resultaat.items():
        r.append(f'| {k} | {eur(v) if isinstance(v, (int, float)) else (v or "–")} |')
    r += ['', '## Gebruikte periodes', '', '| Kolom | Periode | Maanden | Omzet | Bedrijfswinst | Weging | Bron |',
          '|---|---|---:|---:|---:|---:|---|']
    for k in 'HKNR':
        if not b.kolommen.get(k):
            continue
        p = d.periode_op(b.kolommen[k])
        r.append(f'| {k} | t.e.m. {p.einde.strftime("%d/%m/%Y")}{" (tussentijds)" if p.tussentijds else ""} | '
                 f'{p.maanden} | {eur(p.som("70"))} | {eur(bedrijfswinst(p))} | {b.weging.get(k, 0):.0%} | {p.bron} |')
    r += ['', '## Keuzes', '',
          f'- Bestuurdersvergoeding: ' + (f'{eur(b.bestuurdersvergoeding)} EUR per jaar (618 volledig vervangen)'
                                           if b.bestuurdersvergoeding else 'geen normalisatie'),
          f'- Vervangingshuur: ' + (f'{eur(b.marktconforme_huur)} EUR per jaar' if b.marktconforme_huur else 'geen'),
          f'- Multiple: {str(b.multiple).replace(".", ",") if b.multiple else "NIET INGEVULD"} '
          f'{("(" + b.multiple_bron + ")") if b.multiple_bron else ""}',
          f'- Onroerend goed: ' + (f'venale waarde {eur(b.vastgoed_marktwaarde)} EUR ({b.vastgoed_schatting_datum})'
                                   if b.vastgoed_marktwaarde else 'boekwaarde (geen schattingsverslag)'),
          f'- Werkkapitaalcorrectie: {eur(b.werkkapitaalcorrectie)} EUR',
          f'- R/C-intrest bestuurder geneutraliseerd: {"ja" if b.rc_intrest_neutraliseren else "neen"}',
          f'- Goodwill: {b.goodwill_jaren} jaar, vereist rendement ' +
          (f'{b.vereist_rendement:.1%}'.replace('.', ',') if b.vereist_rendement is not None
           else 'volgens samenstelling actief'),
          '- Vervangingsinvesteringen: af te schrijven waarde roerende goederen / gebruiksduur per sectie', '']
    r += ['## Controles', ''] + [f'- ✅ {x}' for x in ok] + [f'- ❌ {x}' for x in fout]
    if not ok and not fout:
        r.append('- Niet uitgevoerd (geen Excel of LibreOffice beschikbaar); open het bestand in Excel.')
    r += ['', '## Openstaande punten en meldingen', '']
    r += [f'- {m}' for m in meldingen] or ['- geen']
    r += ['- Motivering goodwillduur (tabblad Goodwill, B34:B39) en teksten met [ ] nakijken.',
          '- SWOT-analyse (verborgen tabblad Swot) invullen indien gewenst.', '']
    r += ['## Signalen voor normalisatie', ''] + ([f'- {s}' for s in signalen] or ['- geen'])
    r += ['', f'_Kost uitlezen via de Anthropic API: ± ${kost:.2f}_', '']
    pad.write_text('\n'.join(r), encoding='utf-8')
