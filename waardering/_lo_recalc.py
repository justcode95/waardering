"""(Interne hulp voor herberekenen.py) Herberekent een xlsx met LibreOffice (UNO) en schrijft alle celwaarden van de gevraagde bladen naar JSON.
Gebruik dit om na bewerking met xlsxedit.py de aansluitingscontroles te doen en om cached waarden terug te
schrijven (Book.save(..., cached=json)).

Gebruik:
    python recalc.py werkmap.xlsx waarden.json ["blad 1" "blad 2" ...] [--iterate "rendementswaarde 2!E17=E16"]

--iterate zet de doelcel herhaald gelijk aan (de afgeronde waarde van) de broncel tot die stabiel is
(bv. te ontlenen kapitaal in rendementswaarde 2). De eindwaarde wordt afgedrukt; zet ze daarna vast in het bestand.

Vereist: libreoffice-calc (niet enkel libreoffice-core) en python3-uno. Kill een oude 'soffice.bin' als het
laden faalt met 'type detection failed'.
Foutcodes worden vertaald: 524 → #REF!, 532 → #DIV/0!, 519 → #VALUE!, 525 → #NAME?.
"""
import argparse, json, os, subprocess, time
import uno
from com.sun.star.beans import PropertyValue

ERR = {524: '#REF!', 532: '#DIV/0!', 519: '#VALUE!', 525: '#NAME?', 502: '#NUM!', 503: '#NUM!', 32767: '#N/A'}


def colname(c):
    s, n = '', c + 1
    while n:
        n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('xlsx'); ap.add_argument('out'); ap.add_argument('sheets', nargs='*')
    ap.add_argument('--iterate', action='append', default=[])
    ap.add_argument('--port', type=int, default=2002)
    a = ap.parse_args()
    proc = subprocess.Popen(['soffice', '--headless', '--invisible', '--norestore', '--nologo',
                             f'--accept=socket,host=localhost,port={a.port};urp;'],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
    ctx = None
    for _ in range(60):
        try:
            ctx = resolver.resolve(f'uno:socket,host=localhost,port={a.port};urp;StarOffice.ComponentContext'); break
        except Exception:
            time.sleep(0.5)
    desktop = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
    p = PropertyValue(); p.Name = 'Hidden'; p.Value = True
    doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(a.xlsx)), '_blank', 0, (p,))
    doc.calculateAll()
    for it in a.iterate:
        sheet, rest = it.split('!'); tgt, src = rest.split('=')
        sh = doc.Sheets.getByName(sheet)
        for i in range(100):
            v = round(sh.getCellRangeByName(src).Value, 2)
            if abs(v - sh.getCellRangeByName(tgt).Value) < 0.005:
                break
            sh.getCellRangeByName(tgt).Value = v
            doc.calculateAll()
        print(f'{it}: {sh.getCellRangeByName(tgt).Value} na {i} iteraties')
    res = {}
    for sh in doc.Sheets:
        if a.sheets and sh.Name not in a.sheets:
            continue
        cur = sh.createCursor(); cur.gotoEndOfUsedArea(False)
        maxr, maxc = cur.RangeAddress.EndRow, min(cur.RangeAddress.EndColumn, 60)
        vals = {}
        for r in range(maxr + 1):
            for c in range(maxc + 1):
                cell = sh.getCellByPosition(c, r)
                t = cell.Type.value
                if t == 'EMPTY':
                    continue
                ref = f'{colname(c)}{r + 1}'
                if t == 'FORMULA':
                    if cell.getError():
                        vals[ref] = ERR.get(cell.getError(), '#VALUE!')
                    elif cell.FormulaResultType2 == 2:
                        vals[ref] = cell.String
                    else:
                        vals[ref] = cell.Value
                elif t == 'TEXT':
                    vals[ref] = cell.String
                else:
                    vals[ref] = cell.Value
        res[sh.Name] = vals
    doc.close(True)
    try:
        desktop.terminate()
    except Exception:
        pass
    proc.wait(timeout=30)
    json.dump(res, open(a.out, 'w'), indent=0, ensure_ascii=False)


if __name__ == '__main__':
    main()
