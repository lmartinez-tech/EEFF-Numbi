import fs from 'node:fs/promises';
import path from 'node:path';
import { SpreadsheetFile, Workbook } from '@oai/artifact-tool';

const input = JSON.parse(await fs.readFile(path.resolve('analysis_artifacts/normalized_data.json'),'utf8'));
const companyName = input.sources?.[0]?.empresa || 'Empresa';
const companyNit = input.sources?.[0]?.nit || '';
const outDir = path.resolve('outputs/019fd42d-e520-79e1-bbd2-4365facc785b');
await fs.mkdir(outDir,{recursive:true});
const outputPath = path.join(outDir,'EEFF_balance_por_tercero.xlsx');
const statusPath = path.join(outDir,'build_status.txt');
const status = async (message) => { await fs.writeFile(statusPath,message); };
await status('iniciando');

const wb = Workbook.create();
const inicio = wb.worksheets.add('Inicio');
const carga = wb.worksheets.add('Carga BP');
const mapeo = wb.worksheets.add('Mapeo Cuentas');
const er = wb.worksheets.add('Estado de Resultados');
const esf = wb.worksheets.add('Estado de Situación Financiera');
const terceros = wb.worksheets.add('Saldos por Tercero');
const checks = wb.worksheets.add('Checks');
const fuentes = wb.worksheets.add('Fuentes');

const C = {
  navy:'#17365D', blue:'#1F4E78', teal:'#0F6B78', sky:'#D9EAF7', pale:'#EAF2F8',
  green:'#E2F0D9', greenText:'#006100', yellow:'#FFF2CC', red:'#FCE4D6',
  redText:'#9C0006', gray:'#E7E6E6', darkGray:'#666666', white:'#FFFFFF', black:'#000000'
};
const moneyFmt = '#,##0;[Red](#,##0);-';
const decimalFmt = '#,##0.00;[Red](#,##0.00);-';
const dateFmt = 'mmm yyyy';
const isoDate = (s) => new Date(`${s}T00:00:00Z`);
const monthNames = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
const monthEnds = Array.from({length:12},(_,i)=>new Date(Date.UTC(2026,i+1,0)));
const periods = [{label:'2025',date:isoDate('2025-12-31')},...monthEnds.map((d,i)=>({label:`${monthNames[i]} 2026`,date:d}))];

function title(sheet, range, text) {
  sheet.getRange(range).merge();
  sheet.getRange(range).values=[[text]];
  sheet.getRange(range).format={fill:C.navy,font:{bold:true,color:C.white,size:16},verticalAlignment:'center',horizontalAlignment:'left'};
  sheet.getRange(range).format.rowHeight=30;
}
function header(range) {
  range.format={fill:C.blue,font:{bold:true,color:C.white},horizontalAlignment:'center',verticalAlignment:'center',wrapText:true,borders:{preset:'outside',style:'thin',color:C.blue}};
}
function section(range) {
  range.format={fill:C.sky,font:{bold:true,color:C.navy},borders:{bottom:{style:'thin',color:C.blue}}};
}
function total(range, strong=false) {
  range.format={font:{bold:true,color:C.black},borders:{top:{style:strong?'double':'thin',color:C.navy}}};
}
function setWidths(sheet, defs) {
  for (const [range,width] of Object.entries(defs)) sheet.getRange(range).format.columnWidth=width;
}

// Inicio
inicio.showGridLines=false;
title(inicio,'A1:H1','Estados financieros desde balance de prueba por tercero');
inicio.getRange('A3:B7').values=[
  ['Empresa',companyName],['NIT',companyNit],['Fecha de corte',null],
  ['Unidad','COP - pesos colombianos'],['Versión','1.0 | 2026-08-05']
];
inicio.getRange('B5').formulas=[["=MAX('Carga BP'!$A$4:$A$10000)"]];
inicio.getRange('B5').format.numberFormat='dd mmm yyyy';
inicio.getRange('A3:A7').format={fill:C.pale,font:{bold:true,color:C.navy}};
inicio.getRange('A9:B9').merge(); inicio.getRange('A9:B9').values=[['Estado del modelo']]; section(inicio.getRange('A9:B9'));
inicio.getRange('A10:B10').values=[['Resultado',null]];
inicio.getRange('B10').formulas=[["='Checks'!B2"]];
inicio.getRange('B10').format={font:{bold:true,size:14},horizontalAlignment:'center'};
inicio.getRange('B10').conditionalFormats.add('containsText',{text:'OK',format:{fill:C.green,font:{bold:true,color:C.greenText}}});
inicio.getRange('B10').conditionalFormats.add('containsText',{text:'REVISAR',format:{fill:C.red,font:{bold:true,color:C.redText}}});

inicio.getRange('D3:H3').merge(); inicio.getRange('D3:H3').values=[['Resumen al último corte cargado']]; section(inicio.getRange('D3:H3'));
inicio.getRange('D4:E8').values=[
  ['Total activos',null],['Total pasivos',null],['Total patrimonio',null],['Ingresos acumulados 2026',null],['Resultado neto acumulado 2026',null]
];
inicio.getRange('D4:D8').format={fill:C.pale,font:{bold:true,color:C.navy}};
inicio.getRange('E4').formulas=[["=INDEX('Estado de Situación Financiera'!$B$15:$N$15,1,MATCH($B$5,'Estado de Situación Financiera'!$B$5:$N$5,0))"]];
inicio.getRange('E5').formulas=[["=INDEX('Estado de Situación Financiera'!$B$26:$N$26,1,MATCH($B$5,'Estado de Situación Financiera'!$B$5:$N$5,0))"]];
inicio.getRange('E6').formulas=[["=INDEX('Estado de Situación Financiera'!$B$34:$N$34,1,MATCH($B$5,'Estado de Situación Financiera'!$B$5:$N$5,0))"]];
inicio.getRange('E7').formulas=[["='Estado de Resultados'!O10"]];
inicio.getRange('E8').formulas=[["='Estado de Resultados'!O24"]];
inicio.getRange('E4:E8').format={numberFormat:moneyFmt,font:{color:'#008000'}};

inicio.getRange('A12:H12').merge(); inicio.getRange('A12:H12').values=[['Cómo actualizar el archivo']]; section(inicio.getRange('A12:H12'));
inicio.getRange('A13:H19').merge(true);
inicio.getRange('A13:H19').values=[
  ['1. Abra el nuevo “Balance de prueba por tercero” y conserve únicamente las filas con Nivel = Auxiliar y Transaccional = Sí.'],
  ['2. En “Carga BP”, pegue al final de la tabla las columnas A:M: fecha de cierre, periodo, archivo y las 10 columnas contables indicadas.'],
  ['3. La tabla tiene 1.000 filas de carga reservadas con fórmulas en N:R. Si aparece REVISAR, clasifique la cuenta base en “Mapeo Cuentas”.'],
  ['4. Revise “Checks”: los periodos cargados deben quedar en OK y la diferencia del estado de situación financiera debe ser menor a $1.'],
  ['5. Los reportes y el resumen se actualizan por fecha. No pegue subtotales de Clase, Grupo, Cuenta o Subcuenta porque duplicarían los saldos.'],
  ['Nota: 2025 corresponde al archivo anual suministrado; enero–mayo de 2026 corresponden a los cinco archivos mensuales suministrados.'],
  ['Convención: celdas azules = entradas editables; fórmulas = negro/verde; amarillo = requiere revisión.']
];
inicio.getRange('A13:H19').format={wrapText:true,verticalAlignment:'center'};
inicio.getRange('A13:H19').format.rowHeight=28;
setWidths(inicio,{'A:A':24,'B:B':28,'C:C':3,'D:D':33,'E:E':20,'F:H':12});
inicio.freezePanes.freezeRows(1);

// Mapeo
mapeo.showGridLines=false;
title(mapeo,'A1:F1','Mapeo de cuentas a estados financieros');
mapeo.getRange('A3:F3').values=[['Cuenta base','Nombre de referencia','Estado','Grupo EEFF','Línea EEFF','Referencia de la guía']]; header(mapeo.getRange('A3:F3'));
const guideMap = new Map(input.guide_setup.map(x=>[x.cuenta,x]));
const mapRows = input.mapping.map(x=>{
  const g=guideMap.get(x.cuenta_base);
  let line=x.linea_eeff;
  if(x.cuenta_base==='4175') line='Devoluciones y descuentos';
  if(x.cuenta_base==='3610') line='Resultados acumulados';
  return [x.cuenta_base,x.nombre_referencia,x.estado,x.grupo_eeff,line,g?`${g.subgrupo} | ${g.item}`:'Clasificación sugerida por prefijo PUC'];
});
mapeo.getRangeByIndexes(3,0,mapRows.length,6).values=mapRows;
mapeo.getRange(`A4:A${mapRows.length+3}`).format.numberFormat='@';
mapeo.getRange(`C4:E${mapRows.length+3}`).format={font:{color:'#0000FF'},fill:'#F7FBFF'};
mapeo.getRange(`C4:C${mapRows.length+3}`).dataValidation={rule:{type:'list',values:['ESF','ER','REVISAR']}};
mapeo.getRange(`A3:F${mapRows.length+3}`).format.borders={insideHorizontal:{style:'thin',color:'#E7E6E6'},bottom:{style:'thin',color:'#A6A6A6'}};
mapeo.getRange('A2:F2').merge(); mapeo.getRange('A2:F2').values=[['Edite Estado, Grupo y Línea cuando una cuenta nueva quede marcada como REVISAR. Las fórmulas de los reportes usan este mapeo.']];
mapeo.getRange('A2:F2').format={fill:C.yellow,font:{italic:true,color:C.darkGray},wrapText:true};
setWidths(mapeo,{'A:A':14,'B:B':34,'C:C':12,'D:D':18,'E:E':34,'F:F':42});
mapeo.freezePanes.freezeRows(3);

// Carga BP
carga.showGridLines=true;
title(carga,'A1:R1','Base consolidada de balances de prueba por tercero');
const rawHeaders=['Fecha cierre','Periodo','Archivo fuente','Código cuenta','Cuenta base','Nombre cuenta','Identificación','Sucursal','Tercero','Saldo inicial','Débito','Crédito','Saldo final','Control movimiento','Estado','Grupo EEFF','Línea EEFF','Valor EEFF'];
carga.getRange('A3:R3').values=[rawHeaders]; header(carga.getRange('A3:R3'));
const normalizedMap = new Map(mapRows.map(r=>[r[0],{estado:r[2],grupo:r[3],linea:r[4]}]));
const rawRows=input.records.map(r=>{
  const m=normalizedMap.get(r.cuenta_base) || {estado:'REVISAR',grupo:'Sin clasificar',linea:'Cuenta por revisar'};
  const control=r.saldo_inicial+r.debito-r.credito-r.saldo_final;
  const reportValue=m.estado==='ESF' ? (m.grupo==='Activos'?r.saldo_final:-r.saldo_final) :
    (m.estado==='ER' ? (m.grupo==='Ingresos'?r.credito-r.debito:r.debito-r.credito) : 0);
  return [isoDate(r.fecha_cierre),r.periodo,r.archivo,r.codigo,r.cuenta_base,r.nombre_cuenta,r.identificacion,r.sucursal,r.tercero,r.saldo_inicial,r.debito,r.credito,r.saldo_final,control,m.estado,m.grupo,m.linea,reportValue];
});
const firstDataRow=4, lastDataRow=firstDataRow+rawRows.length-1;
const reserveEnd=lastDataRow+1000;
carga.getRangeByIndexes(firstDataRow-1,0,rawRows.length,18).values=rawRows;
carga.getRange(`N${lastDataRow+1}`).formulas=[[`=IF(A${lastDataRow+1}="","",J${lastDataRow+1}+K${lastDataRow+1}-L${lastDataRow+1}-M${lastDataRow+1})`]];
carga.getRange(`N${lastDataRow+1}:N${reserveEnd}`).fillDown();
carga.getRange(`O${lastDataRow+1}`).formulas=[[`=IF(E${lastDataRow+1}="","",IFERROR(VLOOKUP(E${lastDataRow+1},'Mapeo Cuentas'!$A$4:$E$500,3,FALSE),"REVISAR"))`]];
carga.getRange(`O${lastDataRow+1}:O${reserveEnd}`).fillDown();
carga.getRange(`P${lastDataRow+1}`).formulas=[[`=IF(E${lastDataRow+1}="","",IFERROR(VLOOKUP(E${lastDataRow+1},'Mapeo Cuentas'!$A$4:$E$500,4,FALSE),"Sin clasificar"))`]];
carga.getRange(`P${lastDataRow+1}:P${reserveEnd}`).fillDown();
carga.getRange(`Q${lastDataRow+1}`).formulas=[[`=IF(E${lastDataRow+1}="","",IFERROR(VLOOKUP(E${lastDataRow+1},'Mapeo Cuentas'!$A$4:$E$500,5,FALSE),"Cuenta por revisar"))`]];
carga.getRange(`Q${lastDataRow+1}:Q${reserveEnd}`).fillDown();
carga.getRange(`R${lastDataRow+1}`).formulas=[[`=IF(A${lastDataRow+1}="","",IF(O${lastDataRow+1}="ESF",IF(P${lastDataRow+1}="Activos",M${lastDataRow+1},-M${lastDataRow+1}),IF(O${lastDataRow+1}="ER",IF(P${lastDataRow+1}="Ingresos",L${lastDataRow+1}-K${lastDataRow+1},K${lastDataRow+1}-L${lastDataRow+1}),0)))`]];
carga.getRange(`R${lastDataRow+1}:R${reserveEnd}`).fillDown();
carga.getRange(`A${firstDataRow}:A${reserveEnd}`).format.numberFormat='dd-mmm-yyyy';
carga.getRange(`D${firstDataRow}:G${reserveEnd}`).format.numberFormat='@';
carga.getRange(`J${firstDataRow}:N${reserveEnd}`).format.numberFormat=decimalFmt;
carga.getRange(`R${firstDataRow}:R${reserveEnd}`).format.numberFormat=moneyFmt;
carga.getRange(`A${firstDataRow}:M${lastDataRow}`).format.font={color:C.black};
carga.getRange(`N${firstDataRow}:R${lastDataRow}`).format.font={color:C.black};
carga.getRange(`N${lastDataRow+1}:R${reserveEnd}`).format.font={color:'#008000'};
carga.getRange(`O${firstDataRow}:O${reserveEnd}`).conditionalFormats.add('containsText',{text:'REVISAR',format:{fill:C.yellow,font:{bold:true,color:C.redText}}});
const rawTable=carga.tables.add(`A3:R${reserveEnd}`,true,'tblCargaBP'); rawTable.style='TableStyleMedium2';
setWidths(carga,{'A:A':14,'B:B':16,'C:C':38,'D:E':14,'F:F':34,'G:G':18,'H:H':14,'I:I':38,'J:N':16,'O:P':16,'Q:Q':34,'R:R':16});
carga.freezePanes.freezeRows(3); carga.freezePanes.freezeColumns(2);
await status('carga construida');

// Report helpers
function reportBase(sheet,titleText,cols=15){
  sheet.showGridLines=false;
  const end=String.fromCharCode(64+cols);
  title(sheet,`A1:${end}1`,titleText);
  sheet.getRange(`A2:${end}2`).merge(); sheet.getRange(`A2:${end}2`).values=[[`${companyName}${companyNit ? ` | NIT ${companyNit}` : ''} | COP - pesos colombianos`]];
  sheet.getRange(`A2:${end}2`).format={font:{italic:true,color:C.darkGray}};
  sheet.getRange('A4').values=[['Concepto']];
  sheet.getRange('B4').values=[['Comparativo']];
  sheet.getRange('C4:N4').values=[monthNames];
  if(cols===15) sheet.getRange('O4').values=[['Acumulado 2026']];
  header(sheet.getRange(`A4:${end}4`));
  sheet.getRange('B5').values=[[isoDate('2025-12-31')]];
  sheet.getRange('C5:N5').values=[monthEnds];
  if(cols===15) sheet.getRange('O5').values=[['YTD']];
  sheet.getRange(`B5:N5`).format={numberFormat:dateFmt,font:{bold:true,color:C.navy},horizontalAlignment:'center'};
  sheet.getRange('O5').format={font:{bold:true,color:C.navy},horizontalAlignment:'center'};
  setWidths(sheet,{'A:A':42,'B:O':15});
  sheet.freezePanes.freezeRows(5); sheet.freezePanes.freezeColumns(1);
}

// Estado de Resultados
reportBase(er,'Estado de Resultados',15);
const erLabels={7:'INGRESOS',8:'Ingresos operacionales',9:'Devoluciones y descuentos',10:'Ingresos netos',12:'Costo de ventas y servicios',13:'Utilidad bruta',15:'Gastos de administración',16:'Gastos de ventas',17:'Total gastos operacionales',18:'Resultado operacional',20:'Otros ingresos',21:'Otros gastos',22:'Resultado antes de impuestos',23:'Impuesto a las ganancias',24:'RESULTADO NETO'};
for(const [r,label] of Object.entries(erLabels)) er.getRange(`A${r}`).values=[[label]];
section(er.getRange('A7:O7'));
const erDirect={8:'Ingresos operacionales',9:'Devoluciones y descuentos',12:'Costo de ventas y servicios',15:'Gastos de administración',16:'Gastos de ventas',20:'Otros ingresos',21:'Otros gastos',23:'Impuesto a las ganancias'};
for(const [row,line] of Object.entries(erDirect)){
  er.getRange(`B${row}`).formulas=[[`=SUMIFS('Carga BP'!$R$4:$R$10000,'Carga BP'!$A$4:$A$10000,B$5,'Carga BP'!$Q$4:$Q$10000,$A${row})`]];
  er.getRange(`B${row}:N${row}`).fillRight();
  er.getRange(`O${row}`).formulas=[[`=SUM(C${row}:N${row})`]];
}
for(const col of ['B','C','D','E','F','G','H','I','J','K','L','M','N','O']){
  er.getRange(`${col}10`).formulas=[[`=SUM(${col}8:${col}9)`]];
  er.getRange(`${col}13`).formulas=[[`=${col}10-${col}12`]];
  er.getRange(`${col}17`).formulas=[[`=SUM(${col}15:${col}16)`]];
  er.getRange(`${col}18`).formulas=[[`=${col}13-${col}17`]];
  er.getRange(`${col}22`).formulas=[[`=${col}18+${col}20-${col}21`]];
  er.getRange(`${col}24`).formulas=[[`=${col}22-${col}23`]];
}
er.getRange('B8:O24').format.numberFormat=moneyFmt;
er.getRange('B8:O24').format.font={color:'#008000'};
for(const r of [10,13,17,18,22]) total(er.getRange(`A${r}:O${r}`));
total(er.getRange('A24:O24'),true); er.getRange('A24:O24').format.fill=C.green;

// Estado de Situación Financiera
reportBase(esf,'Estado de Situación Financiera',14);
const esfLabels={7:'ACTIVOS',8:'Efectivo y equivalentes',9:'Inversiones',10:'Cuentas por cobrar',11:'Inventarios',12:'Propiedad, planta y equipo',13:'Intangibles',14:'Otros activos',15:'TOTAL ACTIVOS',17:'PASIVOS',18:'Obligaciones financieras',19:'Proveedores',20:'Cuentas por pagar',21:'Impuestos por pagar',22:'Beneficios a empleados',23:'Provisiones',24:'Ingresos diferidos',25:'Otros pasivos',26:'TOTAL PASIVOS',28:'PATRIMONIO',29:'Capital social',30:'Prima en colocación',31:'Reservas',32:'Resultados acumulados',33:'Resultado del periodo',34:'TOTAL PATRIMONIO',36:'TOTAL PASIVO Y PATRIMONIO',37:'DIFERENCIA DE BALANCE'};
for(const [r,label] of Object.entries(esfLabels)) esf.getRange(`A${r}`).values=[[label]];
for(const r of [7,17,28]) section(esf.getRange(`A${r}:N${r}`));
const esfDirect={8:'Efectivo y equivalentes',9:'Inversiones',10:'Cuentas por cobrar',11:'Inventarios',12:'Propiedad, planta y equipo',13:'Intangibles',14:'Otros activos',18:'Obligaciones financieras',19:'Proveedores',20:'Cuentas por pagar',21:'Impuestos por pagar',22:'Beneficios a empleados',23:'Provisiones',24:'Ingresos diferidos',25:'Otros pasivos',29:'Capital social',30:'Prima en colocación',31:'Reservas',32:'Resultados acumulados'};
for(const [row,line] of Object.entries(esfDirect)){
  esf.getRange(`B${row}`).formulas=[[`=SUMIFS('Carga BP'!$R$4:$R$10000,'Carga BP'!$A$4:$A$10000,B$5,'Carga BP'!$Q$4:$Q$10000,$A${row})`]];
  esf.getRange(`B${row}:N${row}`).fillRight();
}
for(let c=2;c<=14;c++){
  const col=String.fromCharCode(64+c);
  esf.getRange(`${col}15`).formulas=[[`=SUM(${col}8:${col}14)`]];
  esf.getRange(`${col}26`).formulas=[[`=SUM(${col}18:${col}25)`]];
  esf.getRange(`${col}33`).formulas=[[c===2 ? `=IF(COUNTIF('Carga BP'!$A$4:$A$10000,B$5)=0,0,'Estado de Resultados'!B24)` : `=IF(COUNTIF('Carga BP'!$A$4:$A$10000,${col}$5)=0,0,SUM('Estado de Resultados'!$B$24:${col}$24))`]];
  esf.getRange(`${col}34`).formulas=[[`=SUM(${col}29:${col}33)`]];
  esf.getRange(`${col}36`).formulas=[[`=${col}26+${col}34`]];
  esf.getRange(`${col}37`).formulas=[[`=${col}15-${col}36`]];
}
esf.getRange('B8:N37').format.numberFormat=moneyFmt;
esf.getRange('B8:N37').format.font={color:'#008000'};
for(const r of [15,26,34,36]) total(esf.getRange(`A${r}:N${r}`),r===36);
esf.getRange('A37:N37').format={fill:C.yellow,font:{bold:true,color:C.redText},borders:{top:{style:'double',color:C.navy}}};
esf.getRange('B37:N37').conditionalFormats.add('cellIs',{operator:'between',formula:[-1,1],format:{fill:C.green,font:{bold:true,color:C.greenText}}});

// Saldos por tercero
terceros.showGridLines=false;
title(terceros,'A1:I1','Saldos por tercero al último corte');
terceros.getRange('A2:B3').values=[['Fecha de corte',null],['Nota','Resumen del último corte suministrado. Para nuevos periodos, filtre Carga BP por tercero o actualice esta tabla.']];
terceros.getRange('B2').formulas=[["=MAX('Carga BP'!$A$4:$A$10000)"]]; terceros.getRange('B2').format.numberFormat='dd mmm yyyy';
terceros.getRange('A5:I5').values=[['Identificación','Tercero','Cuentas por cobrar','Proveedores','Cuentas por pagar','Impuestos','Beneficios empleados','Otros pasivos','Saldo neto seleccionado']]; header(terceros.getRange('A5:I5'));
const partyMap=new Map();
for(const r of input.records){if(r.identificacion && !partyMap.has(r.identificacion)) partyMap.set(r.identificacion,r.tercero);}
const latestDate=input.sources.map(s=>s.fecha_cierre).sort().at(-1);
const partyLines=['Cuentas por cobrar','Proveedores','Cuentas por pagar','Impuestos por pagar','Beneficios a empleados','Otros pasivos'];
const partyBalances=new Map();
for(const r of input.records){
  if(r.fecha_cierre!==latestDate || !r.identificacion) continue;
  const m=normalizedMap.get(r.cuenta_base) || {estado:'REVISAR',grupo:'Sin clasificar',linea:'Cuenta por revisar'};
  const val=m.estado==='ESF' ? (m.grupo==='Activos'?r.saldo_final:-r.saldo_final) : 0;
  const key=`${r.identificacion}|||${m.linea}`;
  partyBalances.set(key,(partyBalances.get(key)||0)+val);
}
const partyRows=[...partyMap.entries()].sort((a,b)=>a[1].localeCompare(b[1],'es')).map(([id,name])=>{
  const vals=partyLines.map(line=>partyBalances.get(`${id}|||${line}`)||0);
  return [id,name,...vals,vals[0]-vals.slice(1).reduce((a,b)=>a+b,0)];
}).sort((a,b)=>Math.abs(b[8])-Math.abs(a[8]));
const partyStart=6, partyEnd=partyStart+partyRows.length-1;
terceros.getRangeByIndexes(partyStart-1,0,partyRows.length,9).values=partyRows;
terceros.getRange(`C${partyStart}:I${partyEnd}`).format={numberFormat:moneyFmt,font:{color:C.black}};
terceros.getRange(`A${partyStart}:A${partyEnd}`).format.numberFormat='@';
const partyTable=terceros.tables.add(`A5:I${partyEnd}`,true,'tblSaldosTercero'); partyTable.style='TableStyleMedium2';
setWidths(terceros,{'A:A':18,'B:B':40,'C:I':18}); terceros.freezePanes.freezeRows(5); terceros.freezePanes.freezeColumns(2);

// Checks
checks.showGridLines=false;
title(checks,'A1:H1','Controles y conciliaciones');
checks.getRange('A2:B2').values=[['ESTADO DEL MODELO',null]];
checks.getRange('B2').formulas=[[`=IF(COUNTIF(H6:H18,"REVISAR")=0,"OK","REVISAR")`]];
checks.getRange('A2:B2').format={fill:C.navy,font:{bold:true,color:C.white,size:13}};
checks.getRange('B2').conditionalFormats.add('containsText',{text:'OK',format:{fill:C.green,font:{bold:true,color:C.greenText}}});
checks.getRange('B2').conditionalFormats.add('containsText',{text:'REVISAR',format:{fill:C.red,font:{bold:true,color:C.redText}}});
checks.getRange('A5:H5').values=[['Fecha','Periodo','Filas auxiliares','Débitos - créditos','Control movimiento','Cuentas sin mapear','Diferencia ESF','Estado']]; header(checks.getRange('A5:H5'));
const checkRows=periods.map(p=>[p.date,p.label,null,null,null,null,null,null]);
checks.getRangeByIndexes(5,0,checkRows.length,8).values=checkRows;
for(let r=6;r<=18;r++){
  checks.getRange(`C${r}`).formulas=[[`=COUNTIF('Carga BP'!$A$4:$A$10000,A${r})`]];
  checks.getRange(`D${r}`).formulas=[[`=SUMIFS('Carga BP'!$K$4:$K$10000,'Carga BP'!$A$4:$A$10000,A${r})-SUMIFS('Carga BP'!$L$4:$L$10000,'Carga BP'!$A$4:$A$10000,A${r})`]];
  checks.getRange(`E${r}`).formulas=[[`=SUMIFS('Carga BP'!$N$4:$N$10000,'Carga BP'!$A$4:$A$10000,A${r})`]];
  checks.getRange(`F${r}`).formulas=[[`=COUNTIFS('Carga BP'!$A$4:$A$10000,A${r},'Carga BP'!$O$4:$O$10000,"REVISAR")`]];
  checks.getRange(`G${r}`).formulas=[[`=IFERROR(INDEX('Estado de Situación Financiera'!$B$37:$N$37,1,MATCH(A${r},'Estado de Situación Financiera'!$B$5:$N$5,0)),0)`]];
  checks.getRange(`H${r}`).formulas=[[`=IF(C${r}=0,"PENDIENTE",IF(AND(ABS(D${r})<1,ABS(E${r})<1,F${r}=0,ABS(G${r})<1),"OK","REVISAR"))`]];
}
checks.getRange('A6:A18').format.numberFormat='dd-mmm-yyyy'; checks.getRange('D6:G18').format.numberFormat=decimalFmt;
checks.getRange('C6:H18').format.font={color:'#008000'};
checks.getRange('H6:H18').conditionalFormats.add('containsText',{text:'OK',format:{fill:C.green,font:{bold:true,color:C.greenText}}});
checks.getRange('H6:H18').conditionalFormats.add('containsText',{text:'REVISAR',format:{fill:C.red,font:{bold:true,color:C.redText}}});
checks.getRange('H6:H18').conditionalFormats.add('containsText',{text:'PENDIENTE',format:{fill:C.gray,font:{color:C.darkGray}}});
checks.getRange('A20:H22').merge(true); checks.getRange('A20:H22').values=[
  ['Tolerancia: $1 por diferencias de redondeo.'],['El control de movimiento valida Saldo inicial + Débito - Crédito = Saldo final.'],['La diferencia ESF incorpora el resultado del periodo para conciliar Activos = Pasivos + Patrimonio.']
]; checks.getRange('A20:H22').format={fill:C.pale,font:{italic:true,color:C.darkGray},wrapText:true};
setWidths(checks,{'A:A':24,'B:B':18,'C:C':16,'D:G':20,'H:H':16}); checks.freezePanes.freezeRows(5);

// Fuentes
fuentes.showGridLines=false;
title(fuentes,'A1:F1','Fuentes y trazabilidad');
fuentes.getRange('A3:F3').values=[['Tipo','Periodo / corte','Archivo','Empresa','NIT','Uso']]; header(fuentes.getRange('A3:F3'));
const sourceRows=input.sources.map(s=>['Balance de prueba por tercero',s.periodo,s.archivo,s.empresa || companyName,s.nit || companyNit,'Filas Auxiliar / Sí importadas a Carga BP']);
sourceRows.push(['Referencia de diseño','Plantilla','guia_eeff.xlsx','','','Guía de estructura, presentación y mapeo; no se mezclaron sus saldos con la compañía objetivo']);
fuentes.getRangeByIndexes(3,0,sourceRows.length,6).values=sourceRows;
fuentes.getRange(`A3:F${sourceRows.length+3}`).format.borders={insideHorizontal:{style:'thin',color:'#E7E6E6'},bottom:{style:'thin',color:'#A6A6A6'}};
fuentes.getRange(`A${sourceRows.length+6}:F${sourceRows.length+8}`).merge(true);
fuentes.getRange(`A${sourceRows.length+6}:F${sourceRows.length+8}`).values=[
  ['Criterio de importación: solo Nivel = Auxiliar y Transaccional = Sí para evitar doble conteo.'],
  ['Criterio de signo: activos por saldo final; pasivos y patrimonio con signo invertido; ingresos = crédito - débito; costos y gastos = débito - crédito.'],
  ['Las clasificaciones son editables en Mapeo Cuentas y deben revisarse cuando se incorporen cuentas nuevas.']
];
fuentes.getRange(`A${sourceRows.length+6}:F${sourceRows.length+8}`).format={fill:C.pale,wrapText:true,font:{italic:true,color:C.darkGray}};
setWidths(fuentes,{'A:A':28,'B:B':18,'C:C':48,'D:D':34,'E:E':18,'F:F':60}); fuentes.freezePanes.freezeRows(3);
await status('hojas construidas');

// Global sheet polish
for(const sheet of [inicio,carga,mapeo,er,esf,terceros,checks,fuentes]){
  const used=sheet.getUsedRange();
  if(used) used.format.font.name='Aptos';
}

// Compact verification before export
await fs.writeFile(path.join(outDir,'inspect_checks.ndjson'),(await wb.inspect({kind:'table',range:'Checks!A1:H18',include:'values,formulas',tableMaxRows:20,tableMaxCols:8,maxChars:12000})).ndjson);
await fs.writeFile(path.join(outDir,'inspect_errors.ndjson'),(await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A',options:{useRegex:true,maxResults:300},summary:'final formula error scan',maxChars:6000})).ndjson);
await status('verificación lógica completa; renderizando');

for(const [sheetName,range,file] of (process.env.SKIP_RENDER==='1' ? [] : [
  ['Inicio','A1:H19','preview_inicio.png'],['Estado de Resultados','A1:O24','preview_er.png'],
  ['Estado de Situación Financiera','A1:N37','preview_esf.png'],['Checks','A1:H22','preview_checks.png'],
  ['Mapeo Cuentas','A1:F30','preview_mapeo.png'],['Carga BP','A1:R20','preview_carga.png'],
  ['Saldos por Tercero','A1:I25','preview_terceros.png'],['Fuentes','A1:F18','preview_fuentes.png']
])){
  const png=await wb.render({sheetName,range,scale:1,format:'png'});
  await fs.writeFile(path.join(outDir,file),new Uint8Array(await png.arrayBuffer()));
  await status(`renderizado ${sheetName}`);
}

const out=await SpreadsheetFile.exportXlsx(wb);
await out.save(outputPath);
await status('done');
