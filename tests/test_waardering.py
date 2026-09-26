"""Tests met een volledig verzonnen testbedrijf (geen klantgegevens in de repository)."""
import datetime as dt
import shutil
from pathlib import Path

import pytest

from waardering import herberekenen
from waardering.invoer import getal
from waardering.invullen import Invuller, afschrijvingsrekening, breuk, classificeer, som_formule
from waardering.model import Actief, Beslissingen, Dossier, Periode, dossier_naar_json, dossier_van_json
from waardering.samenstellen import controleer_periode, samenstellen, signalen, voorstel
from waardering.uitlezen import controleer

TEMPLATE = Path(__file__).resolve().parent.parent / 'template' / 'Waardering_template.xlsx'

BALANS = {  # identiek op elke jaareinde, balanstotaal 297.000
    '220000': 50000, '221000': 200000, '221009': -100000, '230000': 70000, '230009': -40000,
    '240000': 10000, '240009': -6000, '241000': 30000, '241009': -15000, '340000': 5000, '400000': 12000,
    '550000': 80000, '490000': 1000,
    '100000': 20000, '140000': 150000, '173000': 60000, '423000': 10000, '440000': 15000, '450000': 5000,
    '451000': 7000, '179000': 30000}
TITELS = {'179000': 'R/C bestuurder Peeters', '750000': 'Intresten spaarrekening', '650100': 'Intrest R/C bestuurder',
          '173000': 'Investeringskrediet 1234', '423000': 'Investeringskrediet 1234', '618000': 'Bezoldiging bestuurder',
          '618100': 'VAA bestuurder'}


def resultaat(jaar: int, omzet: float, tussentijds=False) -> dict:
    pl = {'700000': omzet, '604000': -0.4 * omzet, '611000': -10000, '612000': -8000, '613000': -12000,
          '614000': -6000, '618000': -36000, '618100': -4000, '620000': -120000, '630000': -20000,
          '640000': -3000, '740000': 2000, '750000': 500, '650000': -3000, '650100': -1000, '670000': -20000}
    if tussentijds:
        pl = {k: round(v * 5 / 12, 2) for k, v in pl.items() if not k.startswith('63')}
    rek = {**BALANS, **pl}
    bw = sum(v for k, v in pl.items() if k[:2] in ('60', '61', '62', '63', '64', '70', '74'))
    wvb = bw + pl['750000'] + pl['650000'] + pl['650100']
    return {'documenttype': 'tussentijdse_cijfers' if tussentijds else 'jaarrekening_of_fiscale_bundel',
            'vennootschap': {'naam': 'Testbedrijf BV', 'straat': 'Teststraat 1', 'postcode_gemeente': '9000 Gent',
                             'ondernemingsnummer': '0123.456.789', 'activiteit': 'testactiviteiten'},
            'periodes': [{'periode_begin': f'{jaar}-01-01',
                          'periode_einde': f'{jaar}-05-31' if tussentijds else f'{jaar}-12-31',
                          'tussentijds': tussentijds, 'bedrijfswinst': round(bw, 2),
                          'winst_voor_belasting': round(wvb, 2), 'balanstotaal': 297000,
                          'rubrieken': [{'naam': 'alles', 'totaal': round(sum(rek.values()), 2),
                                         'rekeningen': [{'nummer': k, 'omschrijving': TITELS.get(k, k), 'saldo': v}
                                                        for k, v in rek.items()]}]}],
            'activa': [], 'afschrijvingstabel_totaal_aanschaf': 0, 'openstaand_totaal': 0, 'opmerkingen': '',
            'personen': [{'naam': 'Jan Peeters', 'rol': 'bestuurder en aandeelhouder', 'aandelen': 100,
                          'bezoldiging': 40000}], '_controle': [], '_bestand': f'{jaar}.pdf'}


def afschrijvingstabel() -> dict:
    items = [('220000', 'Grond', '2015-06-01', 50000), ('221000', 'Gebouw', '2015-06-01', 200000),
             ('230000', 'Machine A', '2021-03-01', 50000), ('230000', 'Machine B', '2024-09-01', 20000),
             ('240000', 'Bureaus', '2022-02-01', 10000), ('241000', 'Bestelwagen', '2023-05-01', 30000),
             ('230000', 'Machine C', '2026-02-01', 5000)]
    return {'documenttype': 'afschrijvingstabel',
            'vennootschap': {'naam': '', 'straat': '', 'postcode_gemeente': '', 'ondernemingsnummer': '',
                             'activiteit': ''},
            'periodes': [], 'personen': [], 'openstaand_totaal': 0, 'opmerkingen': '', '_controle': [],
            'activa': [{'rekening': r, 'omschrijving': o, 'datum_aanschaf': d, 'aanschafwaarde': a,
                        'netto_boekwaarde': a / 2} for r, o, d, a in items],
            'afschrijvingstabel_totaal_aanschaf': 365000, '_bestand': 'afschrijvingstabel.pdf'}


@pytest.fixture
def dossier():
    return samenstellen([resultaat(2023, 400000), resultaat(2024, 420000), resultaat(2025, 450000),
                         resultaat(2026, 200000, tussentijds=True), afschrijvingstabel()])


# ------------------------------------------------------------------ eenheden
def test_getal_invoer():
    assert getal('75.000') == 75000
    assert getal('4,5') == 4.5
    assert getal('1.253,20') == 1253.2
    assert getal('0.12') == 0.12
    assert getal('') is None


def test_som_formule_en_breuk():
    assert som_formule([100, -20.5]) == '=100-20.5'
    assert som_formule([-100, -20.5], -1) == '=-(100+20.5)'
    assert som_formule([10], 1, 'R$11') == '=(10)*R$11'
    assert som_formule([0]) is None
    assert breuk(1 / 3) == '=1/3'
    assert breuk(0.5) == '=1/2'
    assert breuk(1) == 1


def test_classificatie_en_afschrijvingsrekening():
    assert classificeer('240100', 'Laptop HP') == 'computers'
    assert classificeer('232000', 'Inrichting') == 'inrichting'
    assert classificeer('241200', 'Fiets') == 'rollend'
    assert classificeer('240000', 'Stoelen') == 'meubilair'
    assert classificeer('230000', 'Pomp') == 'installaties'
    bal = {'240000': 1, '240009': -1, '240100': 1, '240109': -1}
    assert afschrijvingsrekening('240000', bal) == '240009'
    assert afschrijvingsrekening('240100', bal) == '240109'


def test_controle_rubrieken():
    r = {'periodes': [{'periode_einde': '2025-12-31', 'rubrieken': [
        {'naam': 'A', 'totaal': 30, 'rekeningen': [{'saldo': 10}, {'saldo': 20}]},
        {'naam': 'B', 'totaal': 30, 'rekeningen': [{'saldo': 10}, {'saldo': 21}]}]}]}
    fouten = controleer(r)
    assert len(fouten) == 1 and 'B' in fouten[0]


# ------------------------------------------------------------------ samenstellen en voorstel
def test_samenstellen_en_voorstel(dossier):
    assert dossier.naam == 'Testbedrijf BV'
    assert len(dossier.afgesloten()) == 3
    b = voorstel(dossier)
    assert b.afsluitdatum == dt.date(2025, 12, 31)
    assert b.kolommen == {'H': dt.date(2023, 12, 31), 'K': dt.date(2024, 12, 31), 'N': dt.date(2025, 12, 31),
                          'R': dt.date(2026, 5, 31)}
    assert abs(sum(b.weging.values()) - 1) < 1e-6 and b.weging['R'] == 0
    assert b.rc_intrest_neutraliseren
    assert not [m for m in dossier.meldingen if '≠' in m]            # bedrijfswinst sluit aan
    assert any('zonder geboekte afschrijvingen' in m for m in dossier.meldingen)
    assert any('618' in s for s in signalen(dossier, b))


def test_json_heen_en_terug(dossier):
    terug = dossier_van_json(dossier_naar_json(dossier))
    assert terug.periodes[0].rekeningen == dossier.periodes[0].rekeningen
    b = voorstel(dossier)
    assert Beslissingen.van_json(b.naar_json()) == b


def test_bedrijfswinst_verschil_gemeld():
    p = Periode(None, dt.date(2025, 12, 31), False, {'700000': 100, '600000': -40}, bedrijfswinst=70)
    meld = []
    controleer_periode(p, meld)
    assert meld and '≠' in meld[0]


# ------------------------------------------------------------------ volledige waardering
@pytest.mark.skipif(not shutil.which('soffice'), reason='LibreOffice nodig om te herberekenen')
def test_volledige_waardering(dossier, tmp_path):
    b = voorstel(dossier)
    b.bestuurdersvergoeding = 60000
    b.multiple = 4.5
    uit = tmp_path / 'test.xlsx'
    inv = Invuller(TEMPLATE, dossier, b, tmp_path / 'werk')
    inv.vul(uit)
    assert any('Machine C' in m for m in inv.meldingen)                 # investering na afsluitdatum
    w = herberekenen.met_libreoffice(uit)
    assert w, 'LibreOffice kon niet herberekenen'
    ok, fout = herberekenen.controles(w, b.kolommen, True, 297000)
    assert not fout, fout
    assert len(ok) == 6
    f = w['Gecorrigeerde vrije cash flow']
    # EBITDA 2025 = bedrijfswinst + afschrijvingen + fin. operationeel - normalisatie bestuurder
    bw = 450000 * 0.6 - 10000 - 8000 - 12000 - 6000 - 40000 - 120000 - 20000 - 3000 + 2000
    assert f['N54'] == pytest.approx(bw + 20000 - (60000 - 40000))
    assert f['N65'] == pytest.approx(1000)                               # R/C-intrest teruggenomen
    res = herberekenen.resultaten(w)
    assert res['weerhouden waarde'] and res['weerhouden waarde'] > 0
    assert w['weerhouden waarde']['A1'] == 'Testbedrijf BV'


@pytest.mark.skipif(not shutil.which('soffice'), reason='LibreOffice nodig om te herberekenen')
def test_keten_zoals_het_venster(tmp_path, monkeypatch):
    """Map met PDF's → lees_dossier (API nagebootst) → maak_waardering → Excel + verslag + bewaarde keuzes."""
    from waardering import verwerk
    from waardering.uitlezen import Uitlezer

    nep = {'2023.pdf': resultaat(2023, 400000), '2024.pdf': resultaat(2024, 420000),
           '2025.pdf': resultaat(2025, 450000), 'afschrijvingstabel.pdf': afschrijvingstabel()}
    for naam in nep:
        (tmp_path / naam).write_bytes(b'%PDF-1.4 test')
    monkeypatch.setattr(Uitlezer, '__init__', lambda self, *a, **k: setattr(self, 'kost', 0.0) or
                        setattr(self, 'log', print))
    monkeypatch.setattr(Uitlezer, 'lees', lambda self, pdf, cache: nep[pdf.name])
    cfg = {'template': str(TEMPLATE), 'api_key': 'x', 'model': 'claude-opus-5'}
    d, b, _ = verwerk.lees_dossier(tmp_path, cfg, print)
    assert b.kolommen['R'] is None
    b.multiple = 4
    uit, res, ok, fout = verwerk.maak_waardering(tmp_path, d, b, cfg, print)
    assert uit.name == 'Testbedrijf BV - Waardebepaling - 31-12-2025.xlsx'
    assert not fout and res['weerhouden waarde'] > 0
    verslag = uit.with_suffix('.md').read_text(encoding='utf-8')
    assert 'Controles' in verslag and '❌' not in verslag
    assert (tmp_path / verwerk.KEUZES).exists()
    import openpyxl
    wb = openpyxl.load_workbook(uit, data_only=True)                     # berekende waarden staan in het bestand
    assert wb['weerhouden waarde']['D82'].value == res['weerhouden waarde']
