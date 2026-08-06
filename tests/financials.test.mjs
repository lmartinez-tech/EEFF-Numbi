import assert from "node:assert/strict";
import { mkdir, readFile, unlink, writeFile } from "node:fs/promises";
import path from "node:path";
import test, { after } from "node:test";
import { fileURLToPath, pathToFileURL } from "node:url";
import ts from "typescript";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const runtimeDir = path.join(root, "tmp", "test-runtime");
const runtimeFile = path.join(runtimeDir, `financials-${process.pid}.mjs`);

async function loadFinancials() {
  await mkdir(runtimeDir, { recursive: true });
  const macroSource = (await readFile(path.join(root, "app", "lib", "macro-mapping.ts"), "utf8"))
    .replace("export const MACRO_MAPPING", "const MACRO_MAPPING");
  const financialSource = (await readFile(path.join(root, "app", "lib", "financials.ts"), "utf8"))
    .replace('import { MACRO_MAPPING } from "./macro-mapping";\n', "");
  const output = ts.transpileModule(`${macroSource}\n${financialSource}`, {
    compilerOptions: { module: ts.ModuleKind.ES2022, target: ts.ScriptTarget.ES2022 },
  }).outputText;
  await writeFile(runtimeFile, output, "utf8");
  return import(`${pathToFileURL(runtimeFile).href}?${Date.now()}`);
}

after(async () => {
  try { await unlink(runtimeFile); } catch { /* test runtime may not have been created */ }
});

test("calculates an interim balance without forcing a residual equity plug", async () => {
  const { calculate } = await loadFinancials();
  const common = {
    fileName: "synthetic.xlsx",
    company: "Empresa de prueba",
    nit: "900000000-1",
    periodLabel: "Enero 2026",
    periodEnd: "2026-01-31",
    annual: false,
    identification: "",
    branch: "",
    thirdParty: "",
    initial: 0,
  };
  const records = [
    { ...common, code: "110505", base: "1105", accountName: "Caja", debit: 100, credit: 0, final: 100 },
    { ...common, code: "130505", base: "1305", accountName: "Clientes", debit: 100, credit: 0, final: 100 },
    { ...common, code: "220505", base: "2205", accountName: "Proveedores", debit: 0, credit: 50, final: -50 },
    { ...common, code: "310505", base: "3105", accountName: "Capital", debit: 0, credit: 100, final: -100 },
    { ...common, code: "413501", base: "4135", accountName: "Ingresos", debit: 0, credit: 100, final: -100 },
    { ...common, code: "510505", base: "5105", accountName: "Gasto", debit: 50, credit: 0, final: 50 },
  ];
  const period = calculate(records, {}).latest;
  assert.ok(period);
  assert.equal(period.er["Ingresos netos"], 100);
  assert.equal(period.er["Resultado neto"], 50);
  assert.equal(period.esf["Total activos"], 200);
  assert.equal(period.esf["Total pasivos"], 50);
  assert.equal(period.esf["Total patrimonio"], 150);
  assert.ok(Math.abs(period.check.balance) < 1);
  assert.equal(period.check.status, "OK");
});

test("marks the material 1625 suggestion for manual confirmation", async () => {
  const { defaultMapping } = await loadFinancials();
  assert.equal(defaultMapping("1625").needsReview, true);
  assert.equal(defaultMapping("4135").line, "Ingresos operacionales");
  assert.equal(defaultMapping("6180").line, "Costo de ventas y servicios");
});
