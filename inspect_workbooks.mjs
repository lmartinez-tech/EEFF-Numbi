import fs from 'node:fs/promises';
import path from 'node:path';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const dataDir = path.resolve(process.env.EEFF_DATA_DIR || 'data');
const guide = path.resolve(process.env.EEFF_GUIDE_PATH || path.join(dataDir, 'guia_eeff.xlsx'));
const sources = [guide, ...[
  '1. ENERO Balance de prueba por tercero.xlsx',
  '2. FEBRERO Balance de prueba por tercero.xlsx',
  '3. MARZO Balance de prueba por tercero.xlsx',
  '4. ABRIL Balance de prueba por tercero.xlsx',
  '5. MAYO Balance de prueba por tercero.xlsx',
  'Balance de prueba por tercero 2025.xlsx',
].map((name) => path.join(dataDir, name))];

const outDir = path.resolve('analysis_artifacts');
await fs.mkdir(outDir, { recursive: true });
const summaries = [];

for (let fileIndex = 0; fileIndex < sources.length; fileIndex++) {
  const source = sources[fileIndex];
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(source));
  const sheetNames = workbook.worksheets.items.map(s => s.name);
  const entry = { source, sheets: [] };
  for (const sheetName of sheetNames) {
    const sheet = workbook.worksheets.getItem(sheetName);
    let usedAddress = null;
    try { usedAddress = sheet.getUsedRange()?.address ?? null; } catch { usedAddress = null; }
    const region = await workbook.inspect({
      kind: 'region', sheetId: sheetName, range: usedAddress || 'A1:Z50',
      maxChars: 5000, tableMaxRows: 12, tableMaxCols: 16, tableMaxCellChars: 100,
    });
    const formulas = await workbook.inspect({
      kind: 'formula', sheetId: sheetName, range: usedAddress || 'A1:Z50',
      maxChars: 2500, options: {maxResults: 30},
    });
    entry.sheets.push({name: sheetName, usedAddress, region: region.ndjson, formulas: formulas.ndjson});
    if (fileIndex === 0) {
      try {
        const preview = await workbook.render({sheetName, autoCrop:'all', scale:0.8, format:'png'});
        const safe = sheetName.replace(/[<>:"/\\|?*]/g,'_');
        await fs.writeFile(path.join(outDir, `guide_${String(entry.sheets.length).padStart(2,'0')}_${safe}.png`), new Uint8Array(await preview.arrayBuffer()));
      } catch (error) {
        entry.sheets.at(-1).renderError = String(error);
      }
    }
  }
  summaries.push(entry);
  console.log(`Inspected ${path.basename(source)}: ${sheetNames.length} sheets`);
}

await fs.writeFile(path.join(outDir, 'workbook_summary.json'), JSON.stringify(summaries, null, 2));
console.log(path.join(outDir, 'workbook_summary.json'));
