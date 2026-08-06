import fs from 'node:fs/promises';
import path from 'node:path';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';
const dir=path.resolve('outputs/019fd42d-e520-79e1-bbd2-4365facc785b');
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(dir,'EEFF_balance_por_tercero.xlsx')));
for(const [sheetName,range,file] of [['Saldos por Tercero','A1:I25','preview_terceros_final.png'],['Fuentes','A1:F18','preview_fuentes_final.png'],['Checks','A1:H22','preview_checks_final.png'],['Estado de Situación Financiera','A1:N37','preview_esf_final.png']]){
  const png=await wb.render({sheetName,range,scale:1,format:'png'});
  await fs.writeFile(path.join(dir,file),new Uint8Array(await png.arrayBuffer()));
}
