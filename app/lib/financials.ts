import { strFromU8, unzipSync } from "fflate";

export type Mapping = {
  statement: "ESF" | "ER" | "REVISAR";
  group: string;
  line: string;
};

export type TrialRecord = {
  fileName: string;
  company: string;
  nit: string;
  periodLabel: string;
  periodEnd: string;
  annual: boolean;
  code: string;
  base: string;
  accountName: string;
  identification: string;
  branch: string;
  thirdParty: string;
  initial: number;
  debit: number;
  credit: number;
  final: number;
};

export type ParsedFile = {
  name: string;
  company: string;
  nit: string;
  periodLabel: string;
  periodEnd: string;
  annual: boolean;
  records: TrialRecord[];
};

export const ER_LINES = [
  "Ingresos operacionales",
  "Devoluciones y descuentos",
  "Costo de ventas y servicios",
  "Gastos de administración",
  "Gastos de ventas",
  "Otros ingresos",
  "Otros gastos",
  "Impuesto a las ganancias",
] as const;

export const ESF_LINES = [
  "Efectivo y equivalentes",
  "Inversiones",
  "Cuentas por cobrar",
  "Inventarios",
  "Propiedad, planta y equipo",
  "Intangibles",
  "Otros activos",
  "Obligaciones financieras",
  "Proveedores",
  "Cuentas por pagar",
  "Impuestos por pagar",
  "Beneficios a empleados",
  "Provisiones",
  "Ingresos diferidos",
  "Otros pasivos",
  "Capital social",
  "Prima en colocación",
  "Reservas",
  "Resultados acumulados",
] as const;

export const MAPPING_OPTIONS = [...ESF_LINES, ...ER_LINES];

const monthNumbers: Record<string, number> = {
  enero: 1,
  febrero: 2,
  marzo: 3,
  abril: 4,
  mayo: 5,
  junio: 6,
  julio: 7,
  agosto: 8,
  septiembre: 9,
  octubre: 10,
  noviembre: 11,
  diciembre: 12,
};

const normalize = (value: string) =>
  value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();

const excelColumn = (reference: string) => {
  const letters = reference.match(/[A-Z]+/)?.[0] ?? "A";
  return [...letters].reduce((sum, letter) => sum * 26 + letter.charCodeAt(0) - 64, 0) - 1;
};

function parseSheet(xml: string, shared: string[]) {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  if (doc.querySelector("parsererror")) throw new Error("El archivo contiene XML inválido.");
  const output = new Map<number, unknown[]>();
  for (const row of Array.from(doc.getElementsByTagNameNS("*", "row"))) {
    const rowNumber = Number(row.getAttribute("r") || 0);
    const values: unknown[] = [];
    for (const cell of Array.from(row.getElementsByTagNameNS("*", "c"))) {
      const reference = cell.getAttribute("r") || "A1";
      const column = excelColumn(reference);
      const type = cell.getAttribute("t");
      let result: unknown = null;
      if (type === "inlineStr") {
        result = Array.from(cell.getElementsByTagNameNS("*", "t"))
          .map((node) => node.textContent || "")
          .join("");
      } else {
        const raw = cell.getElementsByTagNameNS("*", "v")[0]?.textContent ?? "";
        if (type === "s") result = shared[Number(raw)] ?? raw;
        else if (type === "b") result = raw === "1";
        else if (raw !== "") result = Number.isFinite(Number(raw)) ? Number(raw) : raw;
      }
      values[column] = result;
    }
    output.set(rowNumber, values);
  }
  return output;
}

function readSharedStrings(files: Record<string, Uint8Array>) {
  const source = files["xl/sharedStrings.xml"];
  if (!source) return [];
  const doc = new DOMParser().parseFromString(strFromU8(source), "application/xml");
  return Array.from(doc.getElementsByTagNameNS("*", "si")).map((item) =>
    Array.from(item.getElementsByTagNameNS("*", "t"))
      .map((node) => node.textContent || "")
      .join(""),
  );
}

function periodFromDescription(description: string) {
  const match = description.match(/De\s+(.+?)\s+(\d{4})\s+a\s+(.+?)\s+(\d{4})/i);
  if (!match) throw new Error("No fue posible identificar el periodo del balance.");
  const startMonth = monthNumbers[normalize(match[1])];
  const endMonth = monthNumbers[normalize(match[3])];
  const startYear = Number(match[2]);
  const endYear = Number(match[4]);
  if (!startMonth || !endMonth) throw new Error("El nombre del mes no es reconocido.");
  const lastDay = new Date(Date.UTC(endYear, endMonth, 0)).getUTCDate();
  const end = `${endYear}-${String(endMonth).padStart(2, "0")}-${String(lastDay).padStart(2, "0")}`;
  const annual = startMonth === 1 && endMonth === 12 && startYear === endYear;
  const monthName = Object.entries(monthNumbers).find(([, value]) => value === endMonth)?.[0] || "periodo";
  const label = annual ? String(endYear) : `${monthName[0].toUpperCase()}${monthName.slice(1)} ${endYear}`;
  return { label, end, annual };
}

const asText = (value: unknown) => (value == null ? "" : String(value).trim());
const asNumber = (value: unknown) => {
  if (typeof value === "number") return value;
  const parsed = Number(String(value ?? "").replace(/,/g, ""));
  return Number.isFinite(parsed) ? parsed : 0;
};

export async function parseWorkbook(file: File): Promise<ParsedFile> {
  const files = unzipSync(new Uint8Array(await file.arrayBuffer()));
  const worksheet = files["xl/worksheets/sheet1.xml"];
  if (!worksheet) throw new Error("No se encontró la primera hoja del libro.");
  const rows = parseSheet(strFromU8(worksheet), readSharedStrings(files));
  const company = asText(rows.get(3)?.[0]);
  const nit = asText(rows.get(4)?.[0]);
  const period = periodFromDescription(asText(rows.get(5)?.[0]));
  const headerEntry = [...rows.entries()].find(([, cells]) =>
    cells.some((cell) => normalize(asText(cell)) === "codigo cuenta contable"),
  );
  if (!headerEntry) throw new Error("No se encontró el encabezado de cuentas contables.");
  const records: TrialRecord[] = [];
  for (const [rowNumber, row] of rows.entries()) {
    if (rowNumber <= headerEntry[0]) continue;
    if (normalize(asText(row[0])) !== "auxiliar" || normalize(asText(row[1])) !== "si") continue;
    const code = asText(row[2]).replace(/\.0$/, "");
    if (!code) continue;
    records.push({
      fileName: file.name,
      company,
      nit,
      periodLabel: period.label,
      periodEnd: period.end,
      annual: period.annual,
      code,
      base: code.slice(0, 4),
      accountName: asText(row[3]),
      identification: asText(row[4]).replace(/\.0$/, ""),
      branch: asText(row[5]),
      thirdParty: asText(row[6]),
      initial: asNumber(row[7]),
      debit: asNumber(row[8]),
      credit: asNumber(row[9]),
      final: asNumber(row[10]),
    });
  }
  if (!records.length) throw new Error("El archivo no contiene filas Auxiliar / Sí para procesar.");
  return { name: file.name, company, nit, periodLabel: period.label, periodEnd: period.end, annual: period.annual, records };
}

export function defaultMapping(base: string): Mapping {
  const first = base.slice(0, 1);
  const two = base.slice(0, 2);
  if (base === "4175") return { statement: "ER", group: "Ingresos", line: "Devoluciones y descuentos" };
  if (base === "3610") return { statement: "ESF", group: "Patrimonio", line: "Resultados acumulados" };
  if (first === "1") {
    const lines: Record<string, string> = { 11: "Efectivo y equivalentes", 12: "Inversiones", 13: "Cuentas por cobrar", 14: "Inventarios", 15: "Propiedad, planta y equipo", 16: "Intangibles", 17: "Otros activos", 18: "Otros activos", 19: "Otros activos" };
    return { statement: "ESF", group: "Activos", line: lines[two] || "Otros activos" };
  }
  if (first === "2") {
    const lines: Record<string, string> = { 21: "Obligaciones financieras", 22: "Proveedores", 23: "Cuentas por pagar", 24: "Impuestos por pagar", 25: "Beneficios a empleados", 26: "Provisiones", 27: "Ingresos diferidos", 28: "Otros pasivos", 29: "Otros pasivos" };
    return { statement: "ESF", group: "Pasivos", line: lines[two] || "Otros pasivos" };
  }
  if (first === "3") {
    const lines: Record<string, string> = { 31: "Capital social", 32: "Prima en colocación", 33: "Reservas", 34: "Resultados acumulados", 35: "Resultados acumulados", 36: "Resultados acumulados", 37: "Resultados acumulados", 38: "Resultados acumulados" };
    return { statement: "ESF", group: "Patrimonio", line: lines[two] || "Resultados acumulados" };
  }
  if (first === "4") return { statement: "ER", group: "Ingresos", line: two === "41" ? "Ingresos operacionales" : "Otros ingresos" };
  if (first === "5") {
    const lines: Record<string, string> = { 51: "Gastos de administración", 52: "Gastos de ventas", 53: "Otros gastos", 54: "Impuesto a las ganancias", 59: "Otros gastos" };
    return { statement: "ER", group: "Gastos", line: lines[two] || "Otros gastos" };
  }
  if (first === "6" || first === "7") return { statement: "ER", group: "Costos", line: "Costo de ventas y servicios" };
  return { statement: "REVISAR", group: "Sin clasificar", line: "Cuenta por revisar" };
}

export function mappingFromLine(line: string): Mapping {
  if ((ER_LINES as readonly string[]).includes(line)) {
    const group = line.includes("Ingreso") || line.includes("Devoluciones") ? "Ingresos" : line.includes("Costo") ? "Costos" : "Gastos";
    return { statement: "ER", group, line };
  }
  if ((ESF_LINES as readonly string[]).includes(line)) {
    const assets = ESF_LINES.slice(0, 7) as readonly string[];
    const liabilities = ESF_LINES.slice(7, 15) as readonly string[];
    return { statement: "ESF", group: assets.includes(line) ? "Activos" : liabilities.includes(line) ? "Pasivos" : "Patrimonio", line };
  }
  return { statement: "REVISAR", group: "Sin clasificar", line: "Cuenta por revisar" };
}

const add = (target: Record<string, number>, key: string, value: number) => {
  target[key] = (target[key] || 0) + value;
};

export type PeriodResult = {
  label: string;
  date: string;
  annual: boolean;
  rows: number;
  er: Record<string, number>;
  esf: Record<string, number>;
  check: { debitCredit: number; movement: number; unmapped: number; balance: number; status: "OK" | "REVISAR" };
};

export function calculate(records: TrialRecord[], overrides: Record<string, string>) {
  const grouped = new Map<string, TrialRecord[]>();
  for (const record of records) {
    const key = record.periodEnd;
    grouped.set(key, [...(grouped.get(key) || []), record]);
  }
  const periods: PeriodResult[] = [...grouped.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([date, periodRecords]) => {
      const er: Record<string, number> = {};
      const esf: Record<string, number> = {};
      let debitCredit = 0;
      let movement = 0;
      let unmapped = 0;
      for (const record of periodRecords) {
        const mapping = overrides[record.base] ? mappingFromLine(overrides[record.base]) : defaultMapping(record.base);
        debitCredit += record.debit - record.credit;
        movement += record.initial + record.debit - record.credit - record.final;
        if (mapping.statement === "REVISAR") {
          unmapped += 1;
          continue;
        }
        if (mapping.statement === "ER") add(er, mapping.line, mapping.group === "Ingresos" ? record.credit - record.debit : record.debit - record.credit);
        if (mapping.statement === "ESF") add(esf, mapping.line, mapping.group === "Activos" ? record.final : -record.final);
      }
      er["Ingresos netos"] = (er["Ingresos operacionales"] || 0) + (er["Devoluciones y descuentos"] || 0);
      er["Utilidad bruta"] = er["Ingresos netos"] - (er["Costo de ventas y servicios"] || 0);
      er["Total gastos operacionales"] = (er["Gastos de administración"] || 0) + (er["Gastos de ventas"] || 0);
      er["Resultado operacional"] = er["Utilidad bruta"] - er["Total gastos operacionales"];
      er["Resultado antes de impuestos"] = er["Resultado operacional"] + (er["Otros ingresos"] || 0) - (er["Otros gastos"] || 0);
      er["Resultado neto"] = er["Resultado antes de impuestos"] - (er["Impuesto a las ganancias"] || 0);
      const totalAssets = ESF_LINES.slice(0, 7).reduce((sum, line) => sum + (esf[line] || 0), 0);
      const totalLiabilities = ESF_LINES.slice(7, 15).reduce((sum, line) => sum + (esf[line] || 0), 0);
      const baseEquity = ESF_LINES.slice(15).reduce((sum, line) => sum + (esf[line] || 0), 0);
      esf["Total activos"] = totalAssets;
      esf["Total pasivos"] = totalLiabilities;
      esf["Resultado del periodo"] = totalAssets - totalLiabilities - baseEquity;
      esf["Total patrimonio"] = baseEquity + esf["Resultado del periodo"];
      esf["Total pasivo y patrimonio"] = totalLiabilities + esf["Total patrimonio"];
      const balance = totalAssets - esf["Total pasivo y patrimonio"];
      const ok = Math.abs(debitCredit) < 1 && Math.abs(movement) < 1 && unmapped === 0 && Math.abs(balance) < 1;
      const first = periodRecords[0];
      return { label: first.periodLabel, date, annual: first.annual, rows: periodRecords.length, er, esf, check: { debitCredit, movement, unmapped, balance, status: ok ? "OK" : "REVISAR" } };
    });

  const latest = periods.at(-1);
  const latestRecords = latest ? grouped.get(latest.date) || [] : [];
  const thirdParties = new Map<string, { identification: string; name: string; receivable: number; payable: number; taxes: number; other: number; net: number }>();
  for (const record of latestRecords) {
    if (!record.identification) continue;
    const mapping = overrides[record.base] ? mappingFromLine(overrides[record.base]) : defaultMapping(record.base);
    if (mapping.statement !== "ESF") continue;
    const current = thirdParties.get(record.identification) || { identification: record.identification, name: record.thirdParty || "Sin nombre", receivable: 0, payable: 0, taxes: 0, other: 0, net: 0 };
    const value = mapping.group === "Activos" ? record.final : -record.final;
    if (mapping.line === "Cuentas por cobrar") current.receivable += value;
    else if (mapping.line === "Proveedores" || mapping.line === "Cuentas por pagar") current.payable += value;
    else if (mapping.line === "Impuestos por pagar") current.taxes += value;
    else if (mapping.group === "Pasivos") current.other += value;
    current.net = current.receivable - current.payable - current.taxes - current.other;
    thirdParties.set(record.identification, current);
  }

  const accounts = new Map<string, { base: string; name: string; mapping: Mapping; records: number }>();
  for (const record of records) {
    const current = accounts.get(record.base);
    const mapping = overrides[record.base] ? mappingFromLine(overrides[record.base]) : defaultMapping(record.base);
    accounts.set(record.base, { base: record.base, name: current?.name || record.accountName, mapping, records: (current?.records || 0) + 1 });
  }
  return {
    periods,
    latest,
    thirdParties: [...thirdParties.values()].sort((a, b) => Math.abs(b.net) - Math.abs(a.net)),
    accounts: [...accounts.values()].sort((a, b) => a.base.localeCompare(b.base)),
  };
}

export function csvDownload(filename: string, rows: (string | number)[][]) {
  const csv = rows.map((row) => row.map((cell) => `"${String(cell ?? "").replace(/"/g, '""')}"`).join(",")).join("\r\n");
  const blob = new Blob(["\ufeff", csv], { type: "text/csv;charset=utf-8" });
  const anchor = document.createElement("a");
  anchor.href = URL.createObjectURL(blob);
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(anchor.href);
}
