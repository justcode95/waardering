"""Minimal in-place xlsx editor that preserves everything it does not touch.

Edits cell values/formulas directly in the sheet XML (lxml keeps prefixes),
can expand shared formulas and insert rows with workbook-wide reference shifting.
"""
import os, re, shutil, zipfile, copy
from lxml import etree

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
N = '{%s}' % NS


def col2n(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n


def n2col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def split(ref):
    m = re.fullmatch(r'\$?([A-Z]{1,3})\$?(\d+)', ref)
    return m[1], int(m[2])


# token regex: string literal | optional sheet prefix + cell or range ref
REF_RE = re.compile(
    r'"(?:[^"]|"")*"'
    r'|(?<![A-Za-z0-9_.$\]\'])'
    r'((?:\'(?:[^\']|\'\')+\'|[A-Za-z_][A-Za-z0-9_.]*)!)?'
    r'(\$?)([A-Z]{1,3})(\$?)(\d+)'
    r'(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?'
    r'(?![A-Za-z0-9_(!])')


def _sheetname(prefix):
    if not prefix:
        return None
    p = prefix[:-1]
    if p.startswith("'"):
        p = p[1:-1].replace("''", "'")
    return p


def transform_formula(f, fn):
    """fn(sheet_or_None, cabs, col, rabs, row) -> (col, row). Applied to each ref endpoint."""
    def rep(m):
        if m.group(0).startswith('"'):
            return m.group(0)
        pre, ca, c, ra, r = m.group(1), m.group(2), m.group(3), m.group(4), m.group(5)
        sh = _sheetname(pre)
        c1, r1 = fn(sh, ca, c, ra, int(r))
        out = (pre or '') + ca + c1 + ra + str(r1)
        if m.group(7):
            c2, r2 = fn(sh, m.group(6), m.group(7), m.group(8), int(m.group(9)))
            out += ':' + m.group(6) + c2 + m.group(8) + str(r2)
        return out
    return REF_RE.sub(rep, f)


class Book:
    def __init__(self, src, workdir):
        if os.path.exists(workdir):
            shutil.rmtree(workdir)
        with zipfile.ZipFile(src) as z:
            z.extractall(workdir)
            self.order = [i.filename for i in z.infolist()]
        self.d = workdir
        self.wb = etree.parse(os.path.join(workdir, 'xl/workbook.xml'))
        rels = etree.parse(os.path.join(workdir, 'xl/_rels/workbook.xml.rels'))
        rmap = {r.get('Id'): r.get('Target') for r in rels.getroot()}
        self.paths = {}
        for s in self.wb.getroot().iter(N + 'sheet'):
            self.paths[s.get('name')] = 'xl/' + rmap[s.get('{%s}id' % RNS)]
        self.trees = {}
        self.sst_path = os.path.join(workdir, 'xl/sharedStrings.xml')
        self.sst = etree.parse(self.sst_path)
        self.strings = []
        for si in self.sst.getroot().iter(N + 'si'):
            self.strings.append(''.join(si.itertext()))
        self.str_index = {}
        for i, s in enumerate(self.strings):
            self.str_index.setdefault(s, i)

    def tree(self, name):
        if name not in self.trees:
            self.trees[name] = etree.parse(os.path.join(self.d, self.paths[name]))
        return self.trees[name]

    def sheet(self, name):
        return Sheet(self, name)

    def add_string(self, s):
        if s in self.str_index:
            return self.str_index[s]
        si = etree.SubElement(self.sst.getroot(), N + 'si')
        t = etree.SubElement(si, N + 't')
        t.text = s
        if s != s.strip() or '\n' in s:
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        self.strings.append(s)
        self.str_index[s] = len(self.strings) - 1
        return len(self.strings) - 1

    # --- shared formulas ------------------------------------------------
    def unshare_all(self):
        for name in self.paths:
            self._unshare(name)

    def _unshare(self, name):
        root = self.tree(name).getroot()
        masters = {}
        cells = list(root.iter(N + 'c'))
        for c in cells:
            f = c.find(N + 'f')
            if f is not None and f.get('t') == 'shared' and f.text:
                masters[f.get('si')] = (c.get('r'), f.text)
        for c in cells:
            f = c.find(N + 'f')
            if f is None or f.get('t') != 'shared':
                continue
            mref, mtext = masters[f.get('si')]
            mc, mr = split(mref)
            cc, cr = split(c.get('r'))
            dr, dc = cr - mr, col2n(cc) - col2n(mc)

            def fn(sh, ca, col, ra, row):
                if not ca:
                    col = n2col(col2n(col) + dc)
                if not ra:
                    row = row + dr
                return col, row
            f.text = transform_formula(mtext, fn)
            for a in ('t', 'si', 'ref'):
                if a in f.attrib:
                    del f.attrib[a]

    # --- row insertion --------------------------------------------------
    def insert_rows(self, sheetname, at, n, style_row):
        """Insert n rows before row `at` in sheetname; copy row/cell styles from style_row
        (numbered before insertion)."""
        def fn_for(current_sheet):
            def fn(sh, ca, col, ra, row):
                target = sh if sh is not None else current_sheet
                if target == sheetname and row >= at:
                    row += n
                return col, row
            return fn
        # formulas everywhere
        for name in self.paths:
            root = self.tree(name).getroot()
            fn = fn_for(name)
            for f in root.iter(N + 'f'):
                if f.text:
                    f.text = transform_formula(f.text, fn)
        for dn in self.wb.getroot().iter(N + 'definedName'):
            if dn.text:
                dn.text = transform_formula(dn.text, fn_for(None))
        root = self.tree(sheetname).getroot()
        sd = root.find(N + 'sheetData')
        tmpl = None
        for row in sd.findall(N + 'row'):
            if int(row.get('r')) == style_row:
                tmpl = copy.deepcopy(row)
        for row in sorted(sd.findall(N + 'row'), key=lambda r: -int(r.get('r'))):
            r = int(row.get('r'))
            if r >= at:
                row.set('r', str(r + n))
                for c in row.findall(N + 'c'):
                    cc, cr = split(c.get('r'))
                    c.set('r', f'{cc}{cr + n}')
        # insert copies of template row (styles only, no values)
        rows = sd.findall(N + 'row')
        idx = next((i for i, r in enumerate(rows) if int(r.get('r')) > at + n - 1), len(rows))
        for k in range(n):
            nr = copy.deepcopy(tmpl)
            rn = at + k
            nr.set('r', str(rn))
            for c in nr.findall(N + 'c'):
                cc, _ = split(c.get('r'))
                c.set('r', f'{cc}{rn}')
                for ch in list(c):
                    c.remove(ch)
                if 't' in c.attrib:
                    del c.attrib['t']
            sd.insert(list(sd).index(rows[idx]) if idx < len(rows) else len(sd), nr)
        # merged cells / sqref / dimension
        for tag in ('mergeCell', 'hyperlink'):
            for e in root.iter(N + tag):
                e.set('ref', self._shift_area(e.get('ref'), at, n))
        for e in root.iter():
            if e.get('sqref'):
                e.set('sqref', ' '.join(self._shift_area(a, at, n) for a in e.get('sqref').split()))
        dim = root.find(N + 'dimension')
        if dim is not None:
            dim.set('ref', self._shift_area(dim.get('ref'), at, n))

    @staticmethod
    def _shift_area(a, at, n):
        def s(ref):
            m = re.fullmatch(r'(\$?[A-Z]{1,3}\$?)(\d+)', ref)
            if not m:
                return ref
            r = int(m[2])
            return m[1] + str(r + n if r >= at else r)
        return ':'.join(s(p) for p in a.split(':'))


    # --- sheet deletion -----------------------------------------------
    def delete_sheet(self, name):
        """Remove a worksheet and the parts only it uses; fix defined names and indices."""
        wbroot = self.wb.getroot()
        sheets_el = wbroot.find(N + 'sheets')
        sheet_els = list(sheets_el)
        idx = next(i for i, s in enumerate(sheet_els) if s.get('name') == name)
        el = sheet_els[idx]
        rid = el.get('{%s}id' % RNS)
        sheets_el.remove(el)
        # defined names
        dns = wbroot.find(N + 'definedNames')
        if dns is not None:
            quoted = "'" + name.replace("'", "''") + "'!"
            for dn in list(dns):
                lsid = dn.get('localSheetId')
                txt = dn.text or ''
                if (lsid is not None and int(lsid) == idx) or quoted in txt or re.search(
                        r'(?<![\w\'\]])' + re.escape(name) + '!', txt):
                    dns.remove(dn)
                elif lsid is not None and int(lsid) > idx:
                    dn.set('localSheetId', str(int(lsid) - 1))
            if len(dns) == 0:
                wbroot.remove(dns)
        for bv in wbroot.iter(N + 'workbookView'):
            for a in ('activeTab', 'firstSheet'):
                if a in bv.attrib:
                    bv.set(a, '0')
        # relationships + parts
        relp = os.path.join(self.d, 'xl/_rels/workbook.xml.rels')
        rels = etree.parse(relp)
        for r in list(rels.getroot()):
            if r.get('Id') == rid:
                rels.getroot().remove(r)
        rels.write(relp, xml_declaration=True, encoding='UTF-8', standalone=True)
        part = self.paths.pop(name)
        self.trees.pop(name, None)
        self._remove_part(part)
        self.removed_sheets = getattr(self, 'removed_sheets', []) + [name]

    def _remove_part(self, part):
        """Delete a part, its rels file and (recursively) internal targets used by nobody else."""
        p = os.path.join(self.d, part)
        folder, fn = os.path.split(part)
        relf = os.path.join(self.d, folder, '_rels', fn + '.rels')
        targets = []
        if os.path.exists(relf):
            for r in etree.parse(relf).getroot():
                if r.get('TargetMode') == 'External':
                    continue
                t = os.path.normpath(os.path.join(folder, r.get('Target'))).replace(os.sep, '/')
                targets.append(t)
            os.remove(relf)
        if os.path.exists(p):
            os.remove(p)
        ct = os.path.join(self.d, '[Content_Types].xml')
        s = open(ct, encoding='utf-8').read()
        s = re.sub(r'<Override [^>]*PartName="/' + re.escape(part) + r'"[^>]*/>', '', s)
        open(ct, 'w', encoding='utf-8').write(s)
        for t in targets:
            if not self._part_referenced(t):
                self._remove_part(t)

    def _part_referenced(self, part):
        for dirpath, _, files in os.walk(self.d):
            for f in files:
                if not f.endswith('.rels'):
                    continue
                relf = os.path.join(dirpath, f)
                owner_folder = os.path.relpath(os.path.dirname(dirpath), self.d)
                for r in etree.parse(relf).getroot():
                    if r.get('TargetMode') == 'External':
                        continue
                    tgt = r.get('Target')
                    full = (tgt.lstrip('/') if tgt.startswith('/') else
                            os.path.normpath(os.path.join(owner_folder, tgt)).replace(os.sep, '/'))
                    if full == part:
                        return True
        return False

    def remove_unused_external_links(self):
        """Drop externalReferences that no formula or defined name uses any more."""
        wbroot = self.wb.getroot()
        er = wbroot.find(N + 'externalReferences')
        if er is None:
            return []
        used = set()
        texts = [dn.text or '' for dn in wbroot.iter(N + 'definedName')]
        for name in self.paths:
            texts += [f.text or '' for f in self.tree(name).getroot().iter(N + 'f')]
        for t in texts:
            used.update(int(m) for m in re.findall(r'\[(\d+)\]', t))
        relp = os.path.join(self.d, 'xl/_rels/workbook.xml.rels')
        rels = etree.parse(relp)
        rmap = {r.get('Id'): r for r in rels.getroot()}
        removed = []
        for i, ref in enumerate(list(er), start=1):
            if i in used:
                continue
            rid = ref.get('{%s}id' % RNS)
            er.remove(ref)
            r = rmap.get(rid)
            if r is not None:
                rels.getroot().remove(r)
                self._remove_part('xl/' + r.get('Target'))
            removed.append(i)
        if len(er) == 0:
            wbroot.remove(er)
        rels.write(relp, xml_declaration=True, encoding='UTF-8', standalone=True)
        return removed

    def fix_app_props(self):
        """Rewrite docProps/app.xml sheet list after deletions."""
        p = os.path.join(self.d, 'docProps/app.xml')
        if not os.path.exists(p):
            return
        t = etree.parse(p)
        root = t.getroot()
        ns = {'ep': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
              'vt': 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'}
        names = [s.get('name') for s in self.wb.getroot().find(N + 'sheets')]
        hp = root.find('ep:HeadingPairs', ns)
        tp = root.find('ep:TitlesOfParts', ns)
        VT = '{%s}' % ns['vt']
        if hp is not None:
            vec = hp.find('vt:vector', ns)
            for ch in list(vec):
                vec.remove(ch)
            v1 = etree.SubElement(vec, VT + 'variant'); etree.SubElement(v1, VT + 'lpstr').text = 'Worksheets'
            v2 = etree.SubElement(vec, VT + 'variant'); etree.SubElement(v2, VT + 'i4').text = str(len(names))
            vec.set('size', '2')
        if tp is not None:
            vec = tp.find('vt:vector', ns)
            for ch in list(vec):
                vec.remove(ch)
            for n_ in names:
                etree.SubElement(vec, VT + 'lpstr').text = n_
            vec.set('size', str(len(names)))
        t.write(p, xml_declaration=True, encoding='UTF-8', standalone=True)


    # --- cleanup helpers ----------------------------------------------
    def remove_defined_names(self, keep):
        """Remove every defined name for which keep(name, text) is False."""
        dns = self.wb.getroot().find(N + 'definedNames')
        removed = []
        if dns is None:
            return removed
        for dn in list(dns):
            if not keep(dn.get('name'), dn.text or ''):
                dns.remove(dn)
                removed.append(dn.get('name'))
        if len(dns) == 0:
            self.wb.getroot().remove(dns)
        return removed

    def remove_workbook_part(self, target_suffix):
        """Remove a workbook-level related part (e.g. 'customXml/item2.xml') and its relationship."""
        relp = os.path.join(self.d, 'xl/_rels/workbook.xml.rels')
        rels = etree.parse(relp)
        for r in list(rels.getroot()):
            if r.get('Target', '').endswith(target_suffix):
                rels.getroot().remove(r)
                full = os.path.normpath(os.path.join('xl', r.get('Target'))).replace(os.sep, '/')
                rels.write(relp, xml_declaration=True, encoding='UTF-8', standalone=True)
                self._remove_part(full)
                return True
        return False

    def compact_shared_strings(self):
        """Drop shared strings no cell uses any more (old dossier texts) and renumber."""
        root = self.sst.getroot()
        old = list(root.iter(N + 'si'))
        used, order = {}, []
        cells = []
        for name in self.paths:
            for c in self.tree(name).getroot().iter(N + 'c'):
                if c.get('t') == 's':
                    v = c.find(N + 'v')
                    if v is not None:
                        i = int(v.text)
                        if i not in used:
                            used[i] = len(order); order.append(i)
                        cells.append((v, i))
        for v, i in cells:
            v.text = str(used[i])
        for si in old:
            root.remove(si)
        for i in order:
            root.append(old[i])
        self.strings = [self.strings[i] for i in order]
        self.str_index = {}
        for k, t in enumerate(self.strings):
            self.str_index.setdefault(t, k)
        return len(old) - len(order)

    # --- save -----------------------------------------------------------
    def save(self, out, cached=None):
        """cached: {sheetname: {ref: value}} written as cached <v> for formula cells."""
        if cached:
            for name, vals in cached.items():
                root = self.tree(name).getroot()
                for c in root.iter(N + 'c'):
                    f = c.find(N + 'f')
                    if f is None or c.get('r') not in vals:
                        continue
                    v = c.find(N + 'v')
                    if v is None:
                        v = etree.SubElement(c, N + 'v')
                    val = vals[c.get('r')]
                    if isinstance(val, str) and val.startswith('#'):
                        c.set('t', 'e'); v.text = val
                    elif isinstance(val, str):
                        c.set('t', 'str'); v.text = val
                    elif isinstance(val, bool):
                        c.set('t', 'b'); v.text = '1' if val else '0'
                    else:
                        if 't' in c.attrib:
                            del c.attrib['t']
                        v.text = repr(float(val)) if val is not None else '0'
        for name, t in self.trees.items():
            t.write(os.path.join(self.d, self.paths[name]), xml_declaration=True, encoding='UTF-8', standalone=True)
        root = self.sst.getroot()
        root.set('uniqueCount', str(len(self.strings)))
        root.attrib.pop('count', None)
        self.sst.write(self.sst_path, xml_declaration=True, encoding='UTF-8', standalone=True)
        # full recalc on open, drop calcChain
        wr = self.wb.getroot()
        cp = wr.find(N + 'calcPr')
        cp.set('fullCalcOnLoad', '1')
        self.wb.write(os.path.join(self.d, 'xl/workbook.xml'), xml_declaration=True, encoding='UTF-8', standalone=True)
        cc = os.path.join(self.d, 'xl/calcChain.xml')
        if os.path.exists(cc):
            os.remove(cc)
            for p, pat in (('xl/_rels/workbook.xml.rels', r'<Relationship [^>]*calcChain[^>]*/>'),
                           ('[Content_Types].xml', r'<Override [^>]*calcChain[^>]*/>')):
                fp = os.path.join(self.d, p)
                s = open(fp, encoding='utf-8').read()
                open(fp, 'w', encoding='utf-8').write(re.sub(pat, '', s))
        if os.path.exists(out):
            os.remove(out)
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
            for fn in self.order:
                p = os.path.join(self.d, fn)
                if os.path.exists(p):
                    z.write(p, fn)


class Sheet:
    def __init__(self, book, name):
        self.b = book
        self.name = name
        self.root = book.tree(name).getroot()
        self.sd = self.root.find(N + 'sheetData')

    def _row(self, r, create=True):
        rows = self.sd.findall(N + 'row')
        for row in rows:
            rr = int(row.get('r'))
            if rr == r:
                return row
            if rr > r:
                if not create:
                    return None
                new = etree.Element(N + 'row', r=str(r))
                row.addprevious(new)
                return new
        if not create:
            return None
        return etree.SubElement(self.sd, N + 'row', r=str(r))

    def cell(self, ref, create=True):
        c, r = split(ref)
        row = self._row(r, create)
        if row is None:
            return None
        for cell in row.findall(N + 'c'):
            cc, _ = split(cell.get('r'))
            if cc == c:
                return cell
            if col2n(cc) > col2n(c):
                if not create:
                    return None
                new = etree.Element(N + 'c', r=ref)
                cell.addprevious(new)
                return new
        if not create:
            return None
        return etree.SubElement(row, N + 'c', r=ref)

    def style(self, ref):
        c = self.cell(ref, create=False)
        return c.get('s') if c is not None else None

    def get(self, ref):
        c = self.cell(ref, create=False)
        if c is None:
            return None
        f = c.find(N + 'f')
        if f is not None:
            return '=' + (f.text or '')
        v = c.find(N + 'v')
        if v is None:
            is_ = c.find(N + 'is')
            return ''.join(is_.itertext()) if is_ is not None else None
        if c.get('t') == 's':
            return self.b.strings[int(v.text)]
        if c.get('t') in ('str', 'e', 'inlineStr'):
            return v.text
        return float(v.text)

    def set(self, ref, value, style_from=None):
        c = self.cell(ref)
        if style_from:
            s = self.style(style_from)
            if s is not None:
                c.set('s', s)
        for ch in list(c):
            c.remove(ch)
        if 't' in c.attrib:
            del c.attrib['t']
        if value is None or value == '':
            return
        if isinstance(value, str) and value.startswith('='):
            f = etree.SubElement(c, N + 'f')
            f.text = value[1:]
        elif isinstance(value, str):
            c.set('t', 's')
            v = etree.SubElement(c, N + 'v')
            v.text = str(self.b.add_string(value))
        else:
            v = etree.SubElement(c, N + 'v')
            v.text = repr(float(value)) if not float(value).is_integer() else str(int(value))

    def clear(self, ref):
        c = self.cell(ref, create=False)
        if c is not None:
            for ch in list(c):
                c.remove(ch)
            if 't' in c.attrib:
                del c.attrib['t']

    def used_refs(self):
        return [c.get('r') for c in self.sd.iter(N + 'c') if len(c)]
