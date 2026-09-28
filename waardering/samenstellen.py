"""Uitgelezen documenten samenvoegen tot één Dossier en voorstellen voor de beslissingen maken."""
from __future__ import annotations

import datetime as dt
import re

from .model import Actief, Beslissingen, Dossier, Periode

RC_PATROON = re.compile(r'\bR\s*/\s*C\b|rekening[- ]?courant', re.I)


def _datum(s: str) -> dt.date | None:
    s = (s or '').strip()
    for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%d.%m.%Y'):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def _datum_in_naam(naam: str) -> dt.date | None:
    """'voorlopige cijfers per 31052026.pdf' → 2026-05-31."""
    m = re.search(r'per[\s_-]*(\d{2})[.\-/]?(\d{2})[.\-/]?(\d{4})', naam, re.I)
    if not m:
        return None
    try:
        return dt.date(int(m[3]), int(m[2]), int(m[1]))
    except ValueError:
        return None


def samenstellen(resultaten: list[dict]) -> Dossier:
    d = Dossier()
    volgorde = sorted(resultaten, key=lambda r: r['documenttype'] != 'jaarrekening_of_fiscale_bundel')
    for r in volgorde:
        v = r.get('vennootschap', {})
        for veld in ('naam', 'straat', 'postcode_gemeente', 'ondernemingsnummer', 'activiteit'):
            if not getattr(d, veld) and v.get(veld):
                setattr(d, veld, v[veld].strip())
        for fout in r.get('_controle', []):
            d.meldingen.append(f"{r.get('_bestand', '?')}: rubriek sluit niet na correctie – {fout}")
        if r.get('opmerkingen'):
            d.meldingen.append(f"{r.get('_bestand', '?')}: {r['opmerkingen']}")

        for p in r.get('periodes', []):
            einde = _datum(p['periode_einde'])
            if not einde or not p['rubrieken']:
                continue
            reks, titels = {}, {}
            for rub in p['rubrieken']:
                for x in rub['rekeningen']:
                    nr = re.sub(r'\D', '', x['nummer'])
                    if not nr:
                        continue
                    reks[nr] = round(reks.get(nr, 0.0) + x['saldo'], 2)
                    titels.setdefault(nr, x['omschrijving'])
            begin = _datum(p['periode_begin'])
            if p['tussentijds']:
                # de einddatum ontbreekt soms op het document en wordt dan de afdrukdatum; de bestandsnaam
                # ('… per 31052026') is betrouwbaarder
                in_naam = _datum_in_naam(r.get('_bestand', ''))
                if in_naam and in_naam < einde and (not begin or in_naam > begin):
                    einde = in_naam
            nieuw = Periode(begin=begin, einde=einde, tussentijds=p['tussentijds'],
                            rekeningen=reks, titels=titels,
                            bedrijfswinst=p['bedrijfswinst'] or None,
                            winst_voor_belasting=p['winst_voor_belasting'] or None,
                            balanstotaal=p['balanstotaal'] or None, bron=r.get('_bestand', ''))
            bestaand = d.periode_op(einde)
            if bestaand is None:
                d.periodes.append(nieuw)
            elif len(nieuw.rekeningen) > len(bestaand.rekeningen):
                d.periodes[d.periodes.index(bestaand)] = nieuw

        if r['documenttype'] == 'afschrijvingstabel':
            d.afschrijvingstabel_totaal = r.get('afschrijvingstabel_totaal_aanschaf') or None
            for a in r.get('activa', []):
                datum = _datum(a['datum_aanschaf'])
                if datum and abs(a['aanschafwaarde']) > 0.004:
                    d.activa.append(Actief(re.sub(r'\D', '', a['rekening'])[:6], a['omschrijving'], datum,
                                           a['aanschafwaarde'], a['netto_boekwaarde']))
        if r['documenttype'] == 'openstaande_klanten':
            d.openstaande_klanten = r.get('openstaand_totaal')
        for pers in r.get('personen', []):
            rol = pers['rol'].lower()
            if 'bestuur' in rol or 'zaakvoerder' in rol:
                if pers['naam'] not in [b['naam'] for b in d.bestuurders]:
                    d.bestuurders.append(pers)
            if 'aandeelhouder' in rol or 'vennoot' in rol:
                if pers['naam'] not in [b['naam'] for b in d.aandeelhouders]:
                    d.aandeelhouders.append(pers)

    for p in d.periodes:
        if p.tussentijds and not p.begin:        # begint de dag na het laatste afgesloten boekjaar
            vorige = [q.einde for q in d.periodes if not q.tussentijds and q.einde < p.einde]
            if vorige:
                p.begin = max(vorige) + dt.timedelta(days=1)
        controleer_periode(p, d.meldingen)
    return d


def bedrijfswinst(p: Periode) -> float:
    return p.som('60', '61', '62', '63', '64', '70', '71', '72', '74')


def controleer_periode(p: Periode, meldingen: list[str]) -> None:
    bw = bedrijfswinst(p)
    if p.bedrijfswinst is not None and abs(bw - p.bedrijfswinst) > 0.02:
        meldingen.append(f'{p.einde}: bedrijfswinst uit de rekeningen {bw:,.2f} ≠ rapport {p.bedrijfswinst:,.2f}')
    wvb = bw + p.som('75', '65')
    if p.winst_voor_belasting is not None and abs(wvb - p.winst_voor_belasting) > 0.02:
        meldingen.append(f'{p.einde}: winst vóór belasting uit de rekeningen {wvb:,.2f} ≠ rapport '
                         f'{p.winst_voor_belasting:,.2f}')
    if p.tussentijds and abs(p.som('63')) < 0.01:
        meldingen.append(f'{p.einde}: tussentijdse cijfers zonder geboekte afschrijvingen (63).')


# ---------------------------------------------------------------------- voorstellen
def voorstel(d: Dossier) -> Beslissingen:
    afg = d.afgesloten()
    if not afg:
        raise ValueError('Geen afgesloten boekjaar met detail per rekening gevonden.')
    laatste3 = afg[-3:]
    kol = dict(zip(['H', 'K', 'N'][3 - len(laatste3):], [p.einde for p in laatste3]))
    for k in 'HKN':
        kol.setdefault(k, None)
    tussen = sorted([p for p in d.periodes if p.tussentijds and p.einde > afg[-1].einde], key=lambda p: p.einde)
    kol['R'] = tussen[-1].einde if tussen else None
    gevuld = [k for k in 'HKN' if kol[k]]
    weging = {k: (1 / len(gevuld) if k in gevuld else 0.0) for k in 'HKNR'}
    afsl = afg[-1]
    heeft_rc = any(RC_PATROON.search(afsl.titels.get(k, '')) and v > 0 for k, v in afsl.items('17', '48', '489'))
    return Beslissingen(afsluitdatum=afsl.einde, kolommen=kol, weging=weging, rc_intrest_neutraliseren=heeft_rc,
                        activiteit=d.activiteit)


EENMALIG = re.compile(r'terug|verzekering|schade|meerwaarde|minderwaarde|subsid|uitzonderlijk|eenmalig|'
                      r'niet[- ]recurrent|herstel', re.I)


def eur(x: float, dec: int = 0) -> str:
    t = f'{x:,.{dec}f}'
    return t.replace(',', 'X').replace('.', ',').replace('X', '.')


def signalen(d: Dossier, b: Beslissingen) -> list[str]:
    """Posten die de gebruiker best even bekijkt voor de normalisaties."""
    uit = []
    for p in sorted(d.periodes, key=lambda p: p.einde):
        if p.einde not in b.kolommen.values():
            continue
        j = p.einde.year
        bestuurder = p.som('618')
        if bestuurder:
            uit.append(f'{j}: geboekte bestuurdersvergoeding (618) {eur(-bestuurder)} EUR')
        huur = p.som('610')
        if huur < -3000:
            uit.append(f'{j}: huur (610) {eur(-huur)} EUR – aan een verbonden partij?')
        for k, v in p.items('76', '66'):
            uit.append(f'{j}: niet-recurrent – {k} {p.titels.get(k, "")} {eur(v, 2)}')
        for k, v in p.items('74', '64'):
            if EENMALIG.search(p.titels.get(k, '')):
                uit.append(f'{j}: mogelijk eenmalig – {k} {p.titels.get(k, "")} {eur(v, 2)}')
        for k, v in p.items('75'):
            if 'cadeau' in p.titels.get(k, '').lower():
                uit.append(f'{j}: {k} {p.titels.get(k, "")} {eur(v, 2)} (operationeel)')
    afsl = d.periode_op(b.afsluitdatum)
    if afsl:
        for k, v in afsl.items('493'):
            uit.append(f'Balans: {k} {afsl.titels.get(k, "")} {eur(v, 2)} (vooruitontvangen, bv. cadeaubonnen)')
        if afsl.som('22') > 0:
            uit.append(f'Balans: eigen onroerend goed (22), netto boekwaarde {eur(afsl.som("22"))} EUR – '
                       'schattingsverslag beschikbaar?')
    return uit
