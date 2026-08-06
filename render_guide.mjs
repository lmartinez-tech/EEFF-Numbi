import fs from 'node:fs/promises';
import path from 'node:path';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const out = path.resolve('analysis_artifacts');
await fs.mkdir(out,{recursive:true});
await fs.writeFile(path.join(out,'render_status.txt'),'importing');
try {
  const dataDir=path.resolve(process.env.EEFF_DATA_DIR || 'data');
  const source=path.resolve(process.env.EEFF_GUIDE_PATH || path.join(dataDir,'guia_eeff.xlsx'));
  const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(source));
  await fs.writeFile(path.join(out,'render_status.txt'),'rendering');
  for (const [name,range,file] of [
    ['1. Estado de Resultados','A6:Q56','guide_er.png'],
    ['2. Estado de Situacion Financie','A6:Q90','guide_esf.png'],
  ]) {
    const png=await wb.render({sheetName:name,range,scale:1,format:'png'});
    await fs.writeFile(path.join(out,file),new Uint8Array(await png.arrayBuffer()));
  }
  await fs.writeFile(path.join(out,'render_status.txt'),'done');
} catch (e) {
  await fs.writeFile(path.join(out,'render_status.txt'),'error\n'+String(e?.stack||e));
  throw e;
}
