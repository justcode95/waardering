"""Het blanco template invullen vanuit een Dossier en de Beslissingen van de gebruiker.

Alles wordt als traceerbare formule in de cel gezet (=12345.67+890), rijen worden alleen toegevoegd in de
MVA-secties wanneer er meer investeringsjaren zijn dan rijen. Celadressen hieronder volgen het template in
template/Waardering_template.xlsx.
"""
from __future__ import annotations

import datetime as dt
import re
from collections import defaultdict
from pathlib import Path

from .model import Beslissingen, Dossier, Periode
from .samenstellen import RC_PATROON, bedrijfswinst
from .xlsxedit import Book

FCF = 'Gecorrigeerde vrije cash flow'
PCT = {'H': 'I', 'K': 'L', 'N': 'O', 'R': 'S'}
MAANDEN = ['januari', 'februari', 'maart', 'april', 'mei', 'juni', 'juli', 'augustus', 'september', 'oktober',
           'november', 'december']
INTREST_75 = re.compile(r'intrest|interest|rente|belegging|beschikbare waarden|spaar', re.I)

# MVA-secties in het template: (sleutel, eerste itemrij, laatste itemrij, totaalrij, restwaardeformule)
MVA_SECTIES = [
    ('computers', 50, 54, 55, '0'),
    ('installaties', 60, 79, 80, 'IF(B{r}>2018,10%,0)'),
    ('meubilair', 84, 91, 92, 'IF(B{r}>2018,10%,0)'),
    ('rollend', 96, 98, 99, '10%'),
    ('inrichting', 103, 109, 110, '0'),
]
MVA_TITEL = {'computers': 'Computers/hardware', 'installaties': 'Installaties, machines en uitrusting',
             'meubilair': 'Meubilair', 'rollend': 'Rollend materieel', 'inrichting': 'Inrichting / huur gebouwen'}


def datum_tekst(d: dt.date) -> str:
    return f'{d.day} {MAANDEN[d.month - 1]} {d.year}'


def excel_datum(d: dt.date) -> int:
    return (d - dt.date(1899, 12, 30)).days


def eur(x: float, dec: int = 0) -> str:
    """Belgische notatie: 75.000 of 1.253,00."""
    t = f'{x:,.{dec}f}'
    return t.replace(',', 'X').replace('.', ',').replace('X', '.')


def breuk(w: float):
    """Weging als exacte formule (=1/3) wanneer dat kan."""
    from fractions import Fraction
    f = Fraction(w).limit_denominator(12)
    if abs(float(f) - w) < 1e-4 and f.denominator > 1:
        return f'={f.numerator}/{f.denominator}'
    return round(w, 6)


def getal(x: float) -> str:
    return f'{x:.2f}'.rstrip('0').rstrip('.') or '0'


def vastgoed_overtollig(afsl, b) -> bool:
    """Met een vervangingshuur wordt eigen vastgoed (22) als overtollig actief gewaardeerd."""
    return bool(b.marktconforme_huur) and afsl.som('22') > 0


def som_formule(waarden: list[float], teken: int = 1, factor: str | None = None) -> str | None:
    """=a+b-c; teken -1 geeft =-(a+b); factor voegt *R$11 toe."""
    waarden = [w for w in waarden if abs(w) > 0.004]
    if not waarden:
        return None
    body = '+'.join(getal(w * (1 if teken > 0 else -1)) for w in waarden).replace('+-', '-')
    if teken < 0:
        body = '-(' + '+'.join(getal(-w) for w in waarden).replace('+-', '-') + ')'
    if factor:
        return f'=({body})*{factor}'
    return '=' + body


def classificeer(rekening: str, omschrijving: str) -> str:
    o = omschrijving.lower()
    if rekening.startswith('21'):
        return 'immaterieel'
    if rekening.startswith('22'):
        return 'onroerend'
    if rekening.startswith('2401') or re.search(r'computer|laptop|\bpc\b|informatica|hardware|server|printer|'
                                                r'kantooruitrusting', o):
        return 'computers'
    if rekening.startswith(('232', '26')) or re.search(r'inrichting|verbouwing|renovatie', o):
        return 'inrichting'
    if rekening.startswith('241') or re.search(r'wagen|voertuig|\bauto\b|bestelwagen|fiets|rollend|truck', o):
        return 'rollend'
    if rekening.startswith('240'):
        return 'meubilair'
    return 'installaties'


def afschrijvingsrekening(rek: str, balans: dict[str, float]) -> str | None:
    """Zoek de bijhorende afschrijvingsrekening (zelfde klasse, eindigt op 9, langste gemeenschappelijk begin)."""
    kandidaten = [k for k in balans if k != rek and len(k) == len(rek) and k[:3] == rek[:3] and k.endswith('9')]

    def gemeen(k):
        n = 0
        while n < len(k) and k[n] == rek[n]:
            n += 1
        return n
    return max(kandidaten, key=gemeen) if kandidaten else None


class Invuller:
    def __init__(self, template: Path, dossier: Dossier, beslissing: Beslissingen, werkmap: Path):
        self.d, self.b = dossier, beslissing
        self.book = Book(str(template), str(werkmap))
        self.book.unshare_all()
        self.meldingen: list[str] = []
        self.ins: list[tuple[int, int]] = []   # (rij in origineel, aantal) voor MVA-rijnummering
        self.p = {k: (dossier.periode_op(v) if v else None) for k, v in beslissing.kolommen.items()}
        self.afsl = dossier.periode_op(beslissing.afsluitdatum)
        if self.afsl is None:
            raise ValueError(f'Geen balans gevonden op de afsluitdatum {beslissing.afsluitdatum}.')
        vorige = [p for p in dossier.afgesloten() if p.einde < beslissing.afsluitdatum]
        self.vorig = vorige[-1] if vorige else None

    # ------------------------------------------------------------------ hulp
    def nr(self, r: int) -> int:
        return r + sum(n for at, n in self.ins if r >= at)

    def factor(self, kol: str) -> str | None:
        p = self.p[kol]
        return f'{kol}$11' if p and p.maanden != 12 else None

    @staticmethod
    def rijen(sheet, rows, verborgen: bool):
        for r in rows:
            row = sheet._row(r)
            if verborgen:
                row.set('hidden', '1')
            else:
                row.attrib.pop('hidden', None)

    # ------------------------------------------------------------------ hoofd
    def vul(self, uit: Path, cached: dict | None = None) -> None:
        self.mva()            # eerst: kan rijen invoegen
        self.fcf()
        self.eigen_vermogen()
        self.goodwill()
        self.identificatie()
        self.marktwaarde()
        self.book.save(str(uit), cached=cached)

    # ------------------------------------------------------------------ identificatie + teksten
    def identificatie(self):
        w = self.book.sheet('weerhouden waarde')
        d, b = self.d, self.b
        w.set('A1', d.naam or '[NAAM VENNOOTSCHAP]')
        w.set('A2', d.straat or '[straat + nummer]')
        w.set('A3', d.postcode_gemeente or '[postcode + gemeente]')
        w.set('A4', f'ON. {d.ondernemingsnummer}' if d.ondernemingsnummer else 'ON. [ondernemingsnummer]')
        w.set('A6', f'Toestand op {datum_tekst(b.afsluitdatum)} (afsluitdatum)')
        act = b.activiteit or d.activiteit or '[omschrijving activiteit]'
        w.set('A10', w.get('A10').replace('[omschrijving activiteit]', act))
        jaren = [str(v.year) for k, v in b.kolommen.items() if v and k != 'R']
        tekst = ', '.join(jaren[:-1]) + ' en ' + jaren[-1] if len(jaren) > 1 else ''.join(jaren)
        w.set('A40', w.get('A40').replace('[jaren]', tekst))
        w.set('A102', f'3. Er wordt geen rekening gehouden met de eigen vermogensevolutie na '
                      f'{b.afsluitdatum.strftime("%d/%m/%Y")}. ')

    # ------------------------------------------------------------------ FCF
    def fcf(self):
        f = self.book.sheet(FCF)
        b = self.b
        for kol in 'HKNR':
            p = self.p[kol]
            pc = PCT[kol]
            if p is None:
                f.set(f'{kol}74', 0)
                continue
            fac = self.factor(kol)
            f.set(f'{kol}9', excel_datum(p.einde))
            f.set(f'{kol}10', f'({p.maanden} maanden x 12/{p.maanden})' if p.maanden != 12 else '(12 maanden)')
            f.set(f'{kol}11', f'=12/{p.maanden}' if p.maanden != 12 else 1)
            vals = lambda *pref, excl=(): [v for _, v in p.items(*pref, uitgezonderd=excl)]
            f.set(f'{kol}13', som_formule(vals('70'), 1, fac))
            aankopen = vals('600', '601', '602', '603', '604', '605', '606', '607', '608')
            f.set(f'{kol}21', som_formule([-v for v in aankopen], 1, fac))   # positieve kostbedragen
            f.set(f'{kol}24', som_formule([-v for v in vals('609')], 1, fac))
            for r, pref in ((33, ('610',)), (34, ('611',)), (35, ('612',)), (36, ('613',)), (37, ('614',)),
                            (38, ('615', '616', '617')), (39, ('618',))):
                f.set(f'{kol}{r}', som_formule(vals(*pref), 1, fac))
            f.set(f'{kol}41', som_formule(vals('62'), 1, fac))
            f.set(f'{kol}43', som_formule(vals('74'), 1, fac))
            f.set(f'{kol}45', som_formule(vals('64'), 1, fac))
            fin_op = [(k, v) for k, v in p.items('75')]
            f.set(f'{kol}47', som_formule([v for k, v in fin_op if not INTREST_75.search(p.titels.get(k, ''))], 1, fac))
            f.set(f'{kol}48', som_formule([v for k, v in fin_op if INTREST_75.search(p.titels.get(k, ''))], 1, fac))
            f.set(f'{kol}49', som_formule(vals('651', '652', '653', '654', '655', '656', '657', '658', '659'), 1, fac))
            f.set(f'{kol}50', som_formule(vals('650'), 1, fac))
            f.set(f'{kol}63', som_formule(vals('650'), 1, fac))
            # normalisatie financieel: fictieve intrest werkkapitaal + R/C-intrest bestuurder terugnemen
            rc = [v for k, v in p.items('650') if RC_PATROON.search(p.titels.get(k, '')) or
                  'bestuurder' in p.titels.get(k, '').lower()]
            fml = '=-Goodwill!$E$118*4%'
            if b.rc_intrest_neutraliseren and rc:
                fml += f'+({getal(-sum(rc))})*{fac}' if fac else '+' + getal(-sum(rc))
            f.set(f'{kol}65', fml)
            f.set(f'{kol}74', breuk(b.weging.get(kol, 0)))
            # a) bestuurder: rijen 85-87 per 618-rekening, rij 92 = genormaliseerde vergoeding
            r618 = sorted(p.items('618'), key=lambda kv: kv[1])
            groepen = [r618[:1], r618[1:2], r618[2:]]
            for r, g in zip((85, 86, 87), groepen):
                f.set(f'{kol}{r}', som_formule([v for _, v in g], 1, fac))
            f.set(f'{kol}92', -b.bestuurdersvergoeding if b.bestuurdersvergoeding else None)
            # b) huur: rij 101 = geboekte huur, rij 106 = marktconforme huur
            f.set(f'{kol}101', som_formule(vals('610'), 1, fac))
            f.set(f'{kol}106', -b.marktconforme_huur if b.marktconforme_huur else None)
            # controle EBITDA tegen bron (enkel H, K, N)
            if kol in 'HKN':
                vc = {'H': 'V', 'K': 'W', 'N': 'X'}[kol]
                f.set(f'{vc}55', p.bedrijfswinst if p.bedrijfswinst is not None else bedrijfswinst(p))
                f.set(f'{vc}56', -p.som('63'))
        # labels normalisatieblok (van de meest recente kolom met 618-rekeningen)
        ref = next((self.p[k] for k in 'NKHR' if self.p[k] and self.p[k].items('618')), None)
        if ref:
            r618 = sorted(ref.items('618'), key=lambda kv: kv[1])
            if r618:
                f.set('B85', f'{r618[0][0]} - {ref.titels.get(r618[0][0], "")}'[:60])
            if len(r618) > 1:
                f.set('B86', f'{r618[1][0]} - {ref.titels.get(r618[1][0], "")}'[:60])
            if len(r618) > 2:
                f.set('B87', ('/'.join(k for k, _ in r618[2:]) + ' - overige')[:60])
        huur_ref = next((self.p[k] for k in 'NKHR' if self.p[k] and self.p[k].items('610')), None)
        if huur_ref:
            f.set('B101', '610 - ' + ', '.join(huur_ref.titels.get(k, k) for k, _ in huur_ref.items('610'))[:55])
        # teksten
        if b.bestuurdersvergoeding:
            f.set('B82', f'Ter bepaling van de genormaliseerde vrije cashflow wordt een marktconforme '
                         f'bestuurdersvergoeding van {eur(b.bestuurdersvergoeding)} EUR (totale kost) weerhouden, ter '
                         f'vervanging van de huidige bezoldiging (618).')
        else:
            f.set('B82', 'Er wordt geen normalisatie van de bestuurdersvergoeding doorgevoerd; de werkelijk geboekte '
                         'vergoeding (618) wordt weerhouden.')
        if self.afsl.som('22') > 0 and not b.marktconforme_huur:
            f.set('B95', 'De vennootschap is eigenaar van het bedrijfspand. Er wordt geen huur aan een verbonden '
                         'partij betaald; bijgevolg is geen vervangingshuur van toepassing.')
        elif b.marktconforme_huur and self.afsl.som('22') > 0:
            f.set('B95', f'Er wordt een marktconforme vervangingshuur van {eur(b.marktconforme_huur)} EUR per jaar '
                         f'aangerekend. Het onroerend goed wordt bijgevolg als overtollig actief beschouwd (zie '
                         f'goodwill, overtollige activa).')
        elif b.marktconforme_huur:
            f.set('B95', f'De huur wordt genormaliseerd naar een marktconforme huur van '
                         f'{eur(b.marktconforme_huur)} EUR per jaar.')
        else:
            f.set('B95', 'Er wordt geen vervangingshuur genormaliseerd; de werkelijke huur wordt weerhouden.')
        f.set('B112', 'Er worden geen bijkomende kosten genormaliseerd.')
        delen = []
        delen.append(f'Correctie financiering werkkapitaal: {eur(b.werkkapitaalcorrectie)} EUR aan 4% (zie goodwill).'
                     if b.werkkapitaalcorrectie else
                     'Er wordt geen fictieve financiering van het werkkapitaal weerhouden.')
        if b.rc_intrest_neutraliseren:
            delen.append('De intrest op de rekening-courant van de bestuurder wordt geneutraliseerd, aangezien de '
                         'rekening-courant bij de overtollige activa/passiva in mindering wordt gebracht.')
        f.set('B138', ' '.join(delen))
        gew = [f'{self.p[k].einde.year}: {b.weging[k]:.0%}' for k in 'HKNR' if self.p[k] and b.weging.get(k)]
        f.set('B142', 'De genormaliseerde vrije cashflow wordt bepaald met volgende weging: ' + ', '.join(gew) + '.')
        # detailrijen tonen als ze gebruikt worden
        self.rijen(f, [21, 24, 33, 34, 35, 36, 37, 38, 39, 48, 50, 61, 62, 63, 64], False)

    # ------------------------------------------------------------------ MVA
    def mva(self):
        b, d = self.b, self.d
        activa = [a for a in d.activa if a.datum <= b.afsluitdatum]
        later = [a for a in d.activa if a.datum > b.afsluitdatum]
        groepen: dict[str, dict[int, list]] = defaultdict(lambda: defaultdict(list))
        for a in activa:
            groepen[classificeer(a.rekening, a.omschrijving)][a.datum.year].append(a)
        # rijen invoegen (van onder naar boven) waar een sectie te weinig itemrijen heeft
        for sleutel, first, last, tot, _ in reversed(MVA_SECTIES):
            nodig = len(groepen[sleutel]) - (last - first + 1)
            if nodig > 0:
                self.book.insert_rows('MVA', last, nodig, style_row=last - 1)
                self.ins.append((last, nodig))
        m = self.book.sheet('MVA')
        nr = self.nr
        m.set('H7', b.afsluitdatum.year + 1)
        bal = self.afsl.rekeningen
        # onroerend goed
        onr = sorted(groepen['onroerend'].items())
        rows = list(range(nr(16), nr(29) + 1))
        if len(onr) > len(rows):
            self.meldingen.append('Meer investeringsjaren onroerend goed dan rijen; laatste jaren samengevoegd.')
            onr = onr[:len(rows) - 1] + [(onr[-1][0], sum((g for _, g in onr[len(rows) - 1:]), []))]
        for i, r in enumerate(rows):
            if i < len(onr):
                jaar, items = onr[i]
                m.set(f'B{r}', jaar)
                m.set(f'C{r}', (items[0].omschrijving + (' e.a.' if len(items) > 1 else ''))[:60])
                m.set(f'D{r}', som_formule([a.aanschafwaarde for a in items]))
            else:
                for c in 'BCD':
                    m.set(f'{c}{r}', None)
        self.rijen(m, rows[len(onr):], True)
        m.set(f'L{nr(30)}', som_formule([v for _, v in self.afsl.items('22')]))
        if d.straat:
            m.set('C15', f'Bedrijfspand - {d.straat}, {d.postcode_gemeente}')
        if b.vastgoed_marktwaarde:
            m.set(f'D{nr(38)}', b.vastgoed_marktwaarde)
            m.set(f'C{nr(39)}', f'd.d. {b.vastgoed_schatting_datum}')
        # roerende secties
        for sleutel, first, last, tot, rest in MVA_SECTIES:
            jaren = sorted(groepen[sleutel].items())
            rows = list(range(nr(first), nr(last) + 1))
            for i, r in enumerate(rows):
                if i < len(jaren):
                    jaar, items = jaren[i]
                    m.set(f'B{r}', jaar, style_from=f'B{rows[0]}')
                    m.set(f'D{r}', som_formule([a.aanschafwaarde for a in items]), style_from=f'D{rows[0]}')
                else:
                    m.set(f'B{r}', None); m.set(f'D{r}', None)
                m.set(f'E{r}', f'=D{r}*' + rest.format(r=r), style_from=f'E{rows[0]}')
                m.set(f'F{r}', f'=D{r}-E{r}', style_from=f'F{rows[0]}')
                m.set(f'G{r}', m.get(f'G{rows[0]}'), style_from=f'G{rows[0]}')
                m.set(f'H{r}', f'=IF(B{r}="",G{r},IF($H$7-B{r}>G{r},G{r},$H$7-B{r}))', style_from=f'H{rows[0]}')
                m.set(f'I{r}', f'=F{r}/G{r}*H{r}', style_from=f'I{rows[0]}')
                m.set(f'J{r}', f'=D{r}-I{r}', style_from=f'J{rows[0]}')
                m.set(f'K{r}', f'=J{r}', style_from=f'K{rows[0]}')
            self.rijen(m, rows[:len(jaren)], False)
            self.rijen(m, rows[max(len(jaren), 1):], True)
            rekeningen = sorted({a.rekening for items in groepen[sleutel].values() for a in items})
            bw = []
            for rek in rekeningen:
                if rek in bal:
                    bw.append(bal[rek])
                    af = afschrijvingsrekening(rek, bal)
                    if af:
                        bw.append(bal[af])
            m.set(f'L{nr(tot)}', som_formule(bw) or 0)
            if jaren and rekeningen:
                m.set(f'C{nr(first) - 2}', f'{MVA_TITEL[sleutel]} (#{" / #".join(rekeningen)})'[:80])
        # controleblok
        imm = [a.aanschafwaarde for a in d.activa if classificeer(a.rekening, a.omschrijving) == 'immaterieel'
               and a.datum <= b.afsluitdatum]
        m.set(f'D{nr(118)}', som_formule(imm) or 0)
        m.set(f'D{nr(119)}', som_formule([a.aanschafwaarde for a in later]) or 0)
        m.set(f'D{nr(121)}', d.afschrijvingstabel_totaal)
        m.set(f'L{nr(119)}', som_formule([v for _, v in self.afsl.items('22', '23', '24', '25', '26', '27')]))
        if later:
            self.meldingen.append('Niet opgenomen (investering na afsluitdatum): ' +
                                  '; '.join(f'{a.rekening} {a.omschrijving} {a.datum}' for a in later))

    # ------------------------------------------------------------------ eigen vermogen
    def eigen_vermogen(self):
        e = self.book.sheet('gecorrigeerd eigen vermogen')
        a = self.afsl
        e.set('A11', '  Kapitaal en inbreng buiten kapitaal (10/11)')
        e.set('D11', som_formule([v for _, v in a.items('10', '11')]) or 0)
        e.set('A12', '  Herwaarderingsmeerwaarden en reserves (12/13)')
        e.set('D12', som_formule([v for _, v in a.items('12', '13')]) or 0)
        e.set('A13', '  Overgedragen resultaat en kapitaalsubsidies (14/15)')
        e.set('D13', som_formule([v for _, v in a.items('14', '15')]) or 0)
        aandelen = sum(x.get('aandelen') or 0 for x in self.d.aandeelhouders)
        e.set('B47', 'Inbreng / kapitaal')
        e.set('C47', '=D11')
        if aandelen:
            e.set('D47', aandelen)
            e.set('E47', 'aandelen')
        if self.d.openstaande_klanten is not None:
            e.set('A67', f'De openstaande handelsvorderingen per {self.b.afsluitdatum.strftime("%d/%m/%Y")} bedragen '
                         f'{eur(a.som("40"), 2)} EUR. Het meest recente overzicht van de openstaande klanten toont een '
                         f'totaal van {eur(self.d.openstaande_klanten, 2)} EUR. [Besluit over achterstallige saldi.]')

    # ------------------------------------------------------------------ goodwill
    def goodwill(self):
        g = self.book.sheet('Goodwill')
        a, b = self.afsl, self.b
        som = lambda *p, excl=(): som_formule([v for _, v in a.items(*p, uitgezonderd=excl)]) or 0
        g.set('C17', b.goodwill_jaren)
        g.set('C16', b.vereist_rendement if b.vereist_rendement is not None else '=E31')
        g.set('A42', f'(3) Overtollige activa en passiva worden als volgt berekend (toestand per '
                     f'{b.afsluitdatum.strftime("%d/%m/%Y")}):')
        if vastgoed_overtollig(a, b):
            # vervangingshuur aangerekend: het eigen vastgoed hoort niet meer bij de bedrijfsmiddelen
            g.set('B43', 'Onroerend goed (vervangingshuur aangerekend)')
            g.set('D43', f'=MVA!K{self.nr(33)}')
            if not b.vastgoed_marktwaarde:
                self.meldingen.append('Vervangingshuur aangerekend: het onroerend goed staat als overtollig actief '
                                      'aan boekwaarde; vul de venale waarde in (schattingsverslag).')
        g.set('D46', som('412'))
        g.set('D47', som_formule([-v for _, v in a.items('450')]) or 0)
        rc = [(k, v) for k, v in a.items('17', '41', '48') if RC_PATROON.search(a.titels.get(k, ''))]
        rc_waarde = [(-v if k.startswith(('17', '48')) else v) for k, v in rc]
        g.set('B48', ('R/C bestuurder (' + ', '.join(k for k, _ in rc) + ')') if rc else 'R/C bestuurder')
        g.set('D48', som_formule(rc_waarde) or 0)
        g.set('D51', b.afsluitdatum.strftime('%d.%m.%Y'))
        balans = {53: ('20',), 54: ('21',), 55: ('22', '23', '24', '25', '26', '27'), 56: ('28',), 57: ('29',),
                  60: ('30', '31'), 61: ('32',), 62: ('33',), 63: ('34',), 64: ('35',), 65: ('36',), 66: ('37',),
                  68: ('40',), 69: ('41',), 70: ('50', '51', '52', '53'), 71: ('54', '55', '56', '57', '58'),
                  72: ('490', '491')}
        for r, pref in balans.items():
            g.set(f'D{r}', som(*pref))
        # leningen (3a): 17 (excl. R/C) + 42 + 43, per lening op één lijn (zelfde omschrijving)
        rc_rek = {k for k, _ in rc}
        leningen: dict[str, list] = defaultdict(list)
        for k, v in a.items('17', '42', '43'):
            if k in rc_rek:
                continue
            titel = re.sub(r'\s+', ' ', a.titels.get(k, k)).strip().lower()
            leningen[titel].append((k, v))
        lijnen = list(leningen.items())
        if len(lijnen) > 6:
            lijnen = lijnen[:5] + [('overige leningen', sum((x for _, x in lijnen[5:]), []))]
        for i, r in enumerate(range(77, 83)):
            if i < len(lijnen):
                titel, items = lijnen[i]
                g.set(f'B{r}', ('/'.join(k for k, _ in items) + ' - ' + a.titels.get(items[0][0], titel))[:70])
                g.set(f'C{r}', som_formule([v for _, v in items]))
            else:
                g.set(f'B{r}', None); g.set(f'C{r}', None)
        kt = a.som('42', '43', '44', '45', '46', '47', '48') + a.som('492', '493')
        uitsluiten = a.som('42', '43') + a.som('450') + sum(v for k, v in rc if k.startswith('48'))
        g.set('C87', '=-(' + getal(kt) + ')+' + getal(uitsluiten))
        # werkkapitaaltabel: D = vorig boekjaar, E = afsluitdatum
        for col, p in (('D', self.vorig), ('E', a)):
            if p is None:
                continue
            g.set(f'{col}98', excel_datum(p.einde))
            ov = [v for k, v in p.items('41') if not k.startswith('412') and not RC_PATROON.search(p.titels.get(k, ''))]
            wk = {99: [v for _, v in p.items('3')], 100: [v for _, v in p.items('40')], 101: ov,
                  102: [v for _, v in p.items('490', '491')], 105: [-v for _, v in p.items('44')],
                  106: [-v for _, v in p.items('45', uitgezonderd=('450', '454', '455', '456', '457', '458',
                                                                    '459'))],
                  107: [-v for _, v in p.items('454', '455', '456', '457', '458', '459')],
                  108: [-v for _, v in p.items('492', '493')]}
            for r, vals in wk.items():
                g.set(f'{col}{r}', som_formule(vals) or 0)
            g.set(f'{col}113', som_formule([v for _, v in p.items('10', '11', '12', '13', '14', '15')]) or 0)
            totaal = p.balanstotaal or p.som('20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '3', '4',
                                              '5', '490', '491')
            g.set(f'{col}114', round(totaal, 2))
        g.set('B106', 'Te betalen BTW en andere belastingen')
        g.set('E118', b.werkkapitaalcorrectie)
        ev = sum(v for _, v in a.items('10', '11', '12', '13', '14', '15'))
        totaal = a.balanstotaal or 1
        g.set('B91', f'Op basis van onderstaand overzicht bedraagt de solvabiliteit van de onderneming '
                     f'{ev / totaal:.0%}. ' + ('Er wordt een theoretische financiering van het werkkapitaal van '
                                                f'{eur(b.werkkapitaalcorrectie)} EUR weerhouden.'
                                                if b.werkkapitaalcorrectie else
                                                'Er wordt geen theoretische financiering van het werkkapitaal '
                                                'weerhouden.'))
        g.set('J133', f'=ROUND(({getal(a.som("10", "11"))})/1000,0)')
        g.set('J134', f'=ROUND(({getal(a.som("12", "13", "14", "15"))})/1000,0)')
        g.set('J137', f'=ROUND(({getal(a.som("44", "45", "46", "47") + a.som("48") - sum(v for k, v in rc if k.startswith("48")))})/1000,0)')
        g.set('J138', f'=ROUND(({getal(a.som("492", "493"))})/1000,0)')

    # ------------------------------------------------------------------ marktwaarde
    def marktwaarde(self):
        mw = self.book.sheet('Marktwaarde ')
        if self.b.multiple:
            mw.set('D19', self.b.multiple)
        if self.b.multiple_bron:
            mw.set('A19', f'Weerhouden multiplicator obv {self.b.multiple_bron}')
