import json
import os
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

DATA_DIR = Path(os.environ.get('EEFF_DATA_DIR', 'data'))
GUIDE = Path(os.environ.get('EEFF_GUIDE_PATH', DATA_DIR / 'guia_eeff.xlsx'))
SOURCES = [
    ('2025', '2025-12-31', DATA_DIR / 'Balance de prueba por tercero 2025.xlsx'),
    ('Enero 2026', '2026-01-31', DATA_DIR / '1. ENERO Balance de prueba por tercero.xlsx'),
    ('Febrero 2026', '2026-02-28', DATA_DIR / '2. FEBRERO Balance de prueba por tercero.xlsx'),
    ('Marzo 2026', '2026-03-31', DATA_DIR / '3. MARZO Balance de prueba por tercero.xlsx'),
    ('Abril 2026', '2026-04-30', DATA_DIR / '4. ABRIL Balance de prueba por tercero.xlsx'),
    ('Mayo 2026', '2026-05-31', DATA_DIR / '5. MAYO Balance de prueba por tercero.xlsx'),
]
NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

def local(tag): return tag.rsplit('}', 1)[-1]
def col_num(ref):
    m = re.match(r'([A-Z]+)', ref or '')
    n = 0
    for ch in (m.group(1) if m else ''): n = n * 26 + ord(ch) - 64
    return n
def shared_strings(z):
    if 'xl/sharedStrings.xml' not in z.namelist(): return []
    out=[]
    for event,e in ET.iterparse(z.open('xl/sharedStrings.xml'),events=('end',)):
        if local(e.tag)=='si': out.append(''.join(t.text or '' for t in e.iter() if local(t.tag)=='t')); e.clear()
    return out
def value(c, shared):
    typ=c.attrib.get('t'); v=next((x.text for x in c if local(x.tag)=='v'),None)
    if typ=='inlineStr': return ''.join(x.text or '' for x in c.iter() if local(x.tag)=='t')
    if typ=='s' and v is not None:
        try:return shared[int(v)]
        except:return v
    if v is None:return None
    try:return float(v) if any(x in v for x in '.eE') else int(v)
    except:return v
def sheet_target(z, sheet_name):
    wb=ET.fromstring(z.read('xl/workbook.xml')); rel=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rels={x.attrib['Id']:x.attrib['Target'] for x in rel if local(x.tag)=='Relationship'}
    for s in wb.iter():
        if local(s.tag)=='sheet' and s.attrib.get('name')==sheet_name:
            target=rels[s.attrib[f'{{{NS_REL}}}id']]
            return target if target.startswith('xl/') else 'xl/'+target.lstrip('/')
    raise KeyError(sheet_name)
def read_rows(path, sheet_name):
    with zipfile.ZipFile(path) as z:
        shared=shared_strings(z); target=sheet_target(z,sheet_name); rows=[]
        for event,e in ET.iterparse(z.open(target),events=('end',)):
            if local(e.tag)!='row': continue
            vals={}
            for c in e:
                if local(c.tag)=='c': vals[col_num(c.attrib.get('r'))]=value(c,shared)
            if vals: rows.append([vals.get(i) for i in range(1,max(vals)+1)])
            e.clear()
        return rows
def code_text(v):
    if v is None:return ''
    if isinstance(v,float) and v.is_integer():return str(int(v))
    return str(v).strip()

def mapping_for(code):
    p2=code[:2]; p1=code[:1]
    if p1=='1':
        lines={'11':'Efectivo y equivalentes','12':'Inversiones','13':'Cuentas por cobrar','14':'Inventarios','15':'Propiedad, planta y equipo','16':'Intangibles','17':'Diferidos','18':'Otros activos','19':'Otros activos'}
        return 'ESF','Activos',lines.get(p2,'Otros activos')
    if p1=='2':
        lines={'21':'Obligaciones financieras','22':'Proveedores','23':'Cuentas por pagar','24':'Impuestos por pagar','25':'Beneficios a empleados','26':'Provisiones','27':'Ingresos diferidos','28':'Otros pasivos','29':'Otros pasivos'}
        return 'ESF','Pasivos',lines.get(p2,'Otros pasivos')
    if p1=='3':
        lines={'31':'Capital social','32':'Prima en colocación','33':'Reservas','34':'Revalorización del patrimonio','35':'Dividendos decretados','36':'Resultado del ejercicio','37':'Resultados acumulados','38':'Otro resultado integral'}
        return 'ESF','Patrimonio',lines.get(p2,'Otros componentes del patrimonio')
    if p1=='4':
        return 'ER','Ingresos','Ingresos operacionales' if p2=='41' else 'Otros ingresos'
    if p1=='5':
        lines={'51':'Gastos de administración','52':'Gastos de ventas','53':'Otros gastos','54':'Impuesto a las ganancias','59':'Ganancias y pérdidas'}
        return 'ER','Gastos',lines.get(p2,'Otros gastos')
    if p1 in ('6','7'):
        return 'ER','Costos','Costo de ventas y servicios'
    return 'REVISAR','Sin clasificar','Cuenta por revisar'

records=[]; source_meta=[]
for period, period_end, path in SOURCES:
    rows=read_rows(path,'Sheet1')
    source_meta.append({'periodo':period,'fecha_cierre':period_end,'archivo':path.name,'empresa':rows[1][0] if len(rows)>1 else '', 'nit':rows[2][0] if len(rows)>2 else ''})
    for row in rows:
        if len(row)<11 or str(row[0]).strip().lower()!='auxiliar' or str(row[1]).strip().lower() not in ('sí','si'): continue
        code=code_text(row[2]); report,group,line=mapping_for(code)
        initial=float(row[7] or 0); debit=float(row[8] or 0); credit=float(row[9] or 0); final=float(row[10] or 0)
        records.append({
            'periodo':period,'fecha_cierre':period_end,'archivo':path.name,'codigo':code,'cuenta_base':code[:4],
            'nombre_cuenta':row[3] or '','identificacion':code_text(row[4]),'sucursal':row[5] or '',
            'tercero':row[6] or '','saldo_inicial':initial,'debito':debit,'credito':credit,'saldo_final':final,
            'movimiento_neto':debit-credit,'estado':report,'grupo_eeff':group,'linea_eeff':line,
        })

unique={}
for r in records:
    key=r['cuenta_base']
    if key not in unique:
        unique[key]={'cuenta_base':key,'nombre_referencia':r['nombre_cuenta'],'estado':r['estado'],'grupo_eeff':r['grupo_eeff'],'linea_eeff':r['linea_eeff']}

guide_setup=[]
for row in read_rows(GUIDE,'Setup'):
    if len(row)>=15 and row[10] is not None:
        guide_setup.append({'cuenta':code_text(row[10]),'estado':row[11] or '','grupo':row[12] or '','subgrupo':row[13] or '','item':row[14] or ''})

checks=[]
by_period={}
for r in records:
    p=by_period.setdefault(r['periodo'],{'saldo_inicial':0,'debito':0,'credito':0,'saldo_final':0,'rows':0})
    for k in ('saldo_inicial','debito','credito','saldo_final'): p[k]+=r[k]
    p['rows']+=1
for period,vals in by_period.items():
    vals['ecuacion_movimiento']=vals['saldo_inicial']+vals['debito']-vals['credito']-vals['saldo_final']

out=Path('analysis_artifacts'); out.mkdir(exist_ok=True)
(out/'normalized_data.json').write_text(json.dumps({'sources':source_meta,'records':records,'mapping':sorted(unique.values(),key=lambda x:x['cuenta_base']),'guide_setup':guide_setup,'period_checks':by_period},ensure_ascii=False,indent=2),encoding='utf-8')
print('records',len(records),'mapping',len(unique),'guide_setup',len(guide_setup))
print(json.dumps(by_period,ensure_ascii=False,indent=2))
