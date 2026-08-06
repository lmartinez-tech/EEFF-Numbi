import io
import json
import os
import re
import zipfile
from collections import deque
from pathlib import Path
from xml.etree import ElementTree as ET

DATA_DIR = Path(os.environ.get('EEFF_DATA_DIR', 'data'))
GUIDE = Path(os.environ.get('EEFF_GUIDE_PATH', DATA_DIR / 'guia_eeff.xlsx'))
SOURCES = [GUIDE] + [
    DATA_DIR / name for name in (
        '1. ENERO Balance de prueba por tercero.xlsx',
        '2. FEBRERO Balance de prueba por tercero.xlsx',
        '3. MARZO Balance de prueba por tercero.xlsx',
        '4. ABRIL Balance de prueba por tercero.xlsx',
        '5. MAYO Balance de prueba por tercero.xlsx',
        'Balance de prueba por tercero 2025.xlsx',
    )
]
NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'

def local(tag): return tag.rsplit('}', 1)[-1]

def parse_shared(z):
    if 'xl/sharedStrings.xml' not in z.namelist(): return []
    out = []
    for event, elem in ET.iterparse(z.open('xl/sharedStrings.xml'), events=('end',)):
        if local(elem.tag) == 'si':
            out.append(''.join(t.text or '' for t in elem.iter() if local(t.tag) == 't'))
            elem.clear()
    return out

def col_num(ref):
    m = re.match(r'([A-Z]+)', ref or '')
    n = 0
    for ch in (m.group(1) if m else ''): n = n * 26 + ord(ch) - 64
    return n

def cell_value(cell, shared):
    typ = cell.attrib.get('t')
    v = next((x.text for x in cell if local(x.tag) == 'v'), None)
    if typ == 'inlineStr':
        return ''.join(x.text or '' for x in cell.iter() if local(x.tag) == 't')
    if typ == 's' and v is not None:
        try: return shared[int(v)]
        except: return v
    if typ == 'b': return v == '1'
    if v is None: return None
    try: return float(v) if any(x in v for x in '.eE') else int(v)
    except: return v

def inspect_sheet(z, target, shared):
    first, tail, formulas, styles = [], deque(maxlen=10), [], {}
    dimension = None
    merges, drawing_refs, validations, hyperlinks = [], [], [], []
    formula_count = 0
    max_seen_row = 0
    stream = z.open(target)
    for event, elem in ET.iterparse(stream, events=('start','end')):
        tag = local(elem.tag)
        if event == 'start' and tag == 'dimension': dimension = elem.attrib.get('ref')
        if event == 'end' and tag == 'row':
            r = int(elem.attrib.get('r', 0)); max_seen_row = max(max_seen_row, r)
            vals = {}
            for cell in elem:
                if local(cell.tag) != 'c': continue
                ref = cell.attrib.get('r'); cnum = col_num(ref)
                f = next((x.text for x in cell if local(x.tag) == 'f'), None)
                value = cell_value(cell, shared)
                if value is not None: vals[cnum] = value
                style = cell.attrib.get('s')
                if style is not None: styles[style] = styles.get(style, 0) + 1
                if f is not None:
                    formula_count += 1
                    if len(formulas) < 100: formulas.append({'cell': ref, 'formula': '=' + f, 'cached': value})
            if vals:
                max_col = min(max(vals), 30)
                row_out = [vals.get(c) for c in range(1, max_col + 1)]
                rec = {'row': r, 'values': row_out}
                if len(first) < 20: first.append(rec)
                tail.append(rec)
            elem.clear()
        elif event == 'end' and tag == 'mergeCell':
            if len(merges) < 100: merges.append(elem.attrib.get('ref'))
            elem.clear()
        elif event == 'end' and tag == 'drawing':
            drawing_refs.append(elem.attrib.get(f'{{{NS_REL}}}id')); elem.clear()
        elif event == 'end' and tag == 'dataValidation':
            if len(validations) < 20: validations.append(dict(elem.attrib)); elem.clear()
        elif event == 'end' and tag == 'hyperlink':
            if len(hyperlinks) < 20: hyperlinks.append(dict(elem.attrib)); elem.clear()
    return {'target': target, 'dimension': dimension, 'max_seen_row': max_seen_row, 'first_nonempty_rows': first,
            'tail_nonempty_rows': list(tail), 'formula_count': formula_count, 'sample_formulas': formulas,
            'style_counts': styles, 'merged_ranges': merges, 'drawing_refs': drawing_refs,
            'data_validations': validations, 'hyperlinks': hyperlinks}

def inspect_book(path):
    with zipfile.ZipFile(path) as z:
        shared = parse_shared(z)
        wbroot = ET.fromstring(z.read('xl/workbook.xml'))
        relroot = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        rels = {r.attrib['Id']: r.attrib['Target'] for r in relroot if local(r.tag) == 'Relationship'}
        sheets = []
        for s in wbroot.iter():
            if local(s.tag) != 'sheet': continue
            rid = s.attrib.get(f'{{{NS_REL}}}id')
            target = rels.get(rid, '')
            if not target.startswith('xl/'): target = 'xl/' + target.lstrip('/')
            info = inspect_sheet(z, target, shared)
            info.update({'name': s.attrib.get('name'), 'state': s.attrib.get('state', 'visible'), 'sheetId': s.attrib.get('sheetId')})
            sheets.append(info)
        return {
            'path': str(path), 'zip_entries': z.namelist(), 'shared_string_count': len(shared), 'sheets': sheets,
            'has_vba': 'xl/vbaProject.bin' in z.namelist(),
            'external_links': [n for n in z.namelist() if n.startswith('xl/externalLinks/') and n.endswith('.xml')],
            'connections': 'xl/connections.xml' in z.namelist(),
        }

out = Path('analysis_artifacts'); out.mkdir(exist_ok=True)
results = []
for src in SOURCES:
    print('Reading', src.name, flush=True)
    results.append(inspect_book(src))
(out / 'structure_ooxml.json').write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding='utf-8')
print(out / 'structure_ooxml.json')
