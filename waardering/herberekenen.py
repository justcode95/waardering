"""Werkboek herberekenen en de controles uitlezen.

Op Windows gebeurt dit met Excel zelf (via pywin32); Excel slaat meteen de berekende waarden op. Zonder Excel wordt
LibreOffice gebruikt als dat geïnstalleerd is. Lukt geen van beide, dan herberekent Excel bij het openen
(fullCalcOnLoad) en worden de controles overgeslagen.
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

from .samenstellen import eur

ZICHTBAAR = ['weerhouden waarde', 'substantiële waarde', 'gecorrigeerd eigen vermogen', 'MVA', 'Goodwill',
             'Gecorrigeerde vrije cash flow', 'OLO', 'Rendementswaarde', 'rendementswaarde 1', 'rendementswaarde 2',
             'Marktwaarde ']


def met_excel(pad: Path) -> bool:
    if sys.platform != 'win32':
        return False
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except ImportError:
        return False
    pythoncom.CoInitialize()                 # we draaien in een achtergrondthread
    try:
        try:
            excel = win32com.client.DispatchEx('Excel.Application')
        except Exception:                    # noqa: BLE001 - Excel niet geïnstalleerd
            return False
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            wb = excel.Workbooks.Open(str(pad.resolve()), UpdateLinks=0)
            excel.CalculateFull()
            wb.Save()
            wb.Close(SaveChanges=False)
        finally:
            excel.Quit()
        return True
    finally:
        pythoncom.CoUninitialize()


def waarden_openpyxl(pad: Path) -> dict:
    import openpyxl
    wb = openpyxl.load_workbook(pad, data_only=True)
    uit = {}
    for ws in wb.worksheets:
        if ws.title in ZICHTBAAR:
            uit[ws.title] = {c.coordinate: c.value for row in ws.iter_rows() for c in row if c.value is not None}
    return uit


def met_libreoffice(pad: Path) -> dict | None:
    """Herberekent via LibreOffice (UNO) en geeft de waarden terug; None als LibreOffice ontbreekt."""
    if getattr(sys, 'frozen', False) or not shutil.which('soffice'):
        return None                          # in de .exe kan het hulpscript niet los draaien
    script = Path(__file__).with_name('_lo_recalc.py')
    uit = pad.with_suffix('.waarden.json')
    r = subprocess.run([sys.executable, str(script), str(pad), str(uit), *ZICHTBAAR], capture_output=True,
                       text=True, timeout=600)
    if r.returncode != 0 or not uit.exists():
        return None
    import json
    data = json.loads(uit.read_text(encoding='utf-8'))
    uit.unlink()
    return data


def controles(w: dict, kolommen: dict, afschrijvingstabel: bool, balanstotaal: float | None) -> tuple[list, list]:
    """Geeft (ok-lijst, fouten-lijst) terug op basis van de berekende waarden."""
    ok, fout = [], []
    f = w.get('Gecorrigeerde vrije cash flow', {})
    for kol, vc in (('H', 'V'), ('K', 'W'), ('N', 'X')):
        if kolommen.get(kol):
            v = f.get(f'{vc}58')
            (ok if v == 'OK' else fout).append(f'EBITDA {kolommen[kol]}: aansluiting met bedrijfswinst + '
                                              f'afschrijvingen {"OK" if v == "OK" else "FOUT"}')
    m = w.get('MVA', {})
    tot_row = next((int(k[1:]) for k, v in m.items() if k.startswith('C') and v == 'TOTAAL MVA'), None)
    if tot_row:
        nbw = m.get(f'L{tot_row + 8}')
        if isinstance(nbw, (int, float)):
            (ok if abs(nbw) < 0.02 else fout).append(f'MVA netto boekwaarde vs balans: verschil {eur(nbw, 2)}')
        if afschrijvingstabel:
            d = m.get(f'D{tot_row + 10}')
            if isinstance(d, (int, float)):
                (ok if abs(d) < 0.02 else fout).append(f'MVA aanschafwaarde vs afschrijvingstabel: verschil {eur(d, 2)}')
    g = w.get('Goodwill', {})
    if balanstotaal and isinstance(g.get('D74'), (int, float)):
        verschil = g['D74'] - balanstotaal
        (ok if abs(verschil) < 0.02 else fout).append(f'Goodwill balanstotaal vs jaarrekening: verschil {eur(verschil, 2)}')
    for blad, vals in w.items():
        err = [k for k, v in vals.items() if isinstance(v, str) and v.startswith('#') and v != '# jaar']
        if err:
            fout.append(f'{blad}: foutwaarden in {", ".join(err[:8])}')
    return ok, fout


def resultaten(w: dict) -> dict:
    ww = w.get('weerhouden waarde', {})
    f = w.get('Gecorrigeerde vrije cash flow', {})
    g = w.get('Goodwill', {})
    return {'substantiële waarde': ww.get('B75'), 'rendementswaarde': ww.get('B76'),
            'marktwaarde (EBITDA)': ww.get('B77'), 'weerhouden waarde': ww.get('D82'),
            'gecorrigeerd eigen vermogen': g.get('D7'), 'goodwill': g.get('D22'), 'overtollige activa': g.get('E42'),
            'weerhouden EBITDA': f.get('K55'), 'weerhouden FCF na belastingen': f.get('K76')}
