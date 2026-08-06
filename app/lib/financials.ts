import { strFromU8, unzipSync } from "fflate";
import { MACRO_MAPPING } from "./macro-mapping";

export type Mapping = {
  statement: "ESF" | "ER" | "REVISAR";
  group: "Activos" | "Pasivos" | "Patrimonio" | "Ingresos" | "Gastos" | "Sin clasificar";
  subgroup: string;
  item: string;
  line: string;
  source: "Macro" | "PUC" | "Manual";
  needsReview?: boolean;
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

export const ER_DETAIL_LINES = [
  "Ingresos operacionales",
  "Devoluciones y descuentos",
  "Costo de ventas y servicios",
  "Gastos de administración",
  "Gastos de ventas",
  "Otros gastos operacionales",
  "Depreciación",
  "Amortización",
  "Otros ingresos",
  "Gastos financieros",
  "Impuesto a las ganancias",
] as const;

export const CURRENT_ASSETS = [
  "Efectivo y equivalentes",
  "Inversiones corrientes",
  "Cuentas comerciales por cobrar",
  "Otras cuentas por cobrar",
  "Inventarios",
  "Activos diferidos",
  "Otros activos corrientes",
] as const;

export const NON_CURRENT_ASSETS = [
  "Propiedad, planta y equipo",
  "Intangibles",
  "Impuesto diferido",
  "Cuentas comerciales no corrientes",
  "Otras cuentas por cobrar no corrientes",
  "Activos financieros no corrientes",
  "Otros activos no corrientes",
] as const;

export const CURRENT_LIABILITIES = [
  "Obligaciones financieras",
  "Cuentas por pagar a socios",
  "Proveedores",
  "Cuentas por pagar",
  "Impuestos por pagar",
  "Beneficios a empleados",
  "Provisiones",
  "Ingresos diferidos",
  "Otros pasivos corrientes",
] as const;

export const NON_CURRENT_LIABILITIES = [
  "Obligaciones financieras no corrientes",
  "Acreedores comerciales no corrientes",
  "Beneficios a empleados no corrientes",
  "Vinculados económicos",
  "Otras cuentas por pagar no corrientes",
  "Provisiones no corrientes",
  "Impuestos no corrientes",
  "Otros pasivos no corrientes",
] as const;

export const EQUITY_LINES = [
  "Capital social",
  "Prima en colocación",
  "Reservas",
  "Superávit y revaluación",
  "Resultados acumulados",
  "Resultado del periodo",
  "Otros componentes del patrimonio",
] as const;

export const ESF_DETAIL_LINES = [
  ...CURRENT_ASSETS,
  ...NON_CURRENT_ASSETS,
  ...CURRENT_LIABILITIES,
  ...NON_CURRENT_LIABILITIES,
  ...EQUITY_LINES,
] as const;

export const MAPPING_OPTIONS = [...ESF_DETAIL_LINES, ...ER_DETAIL_LINES];

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

const includes = (value: string, fragment: string) => normalize(value).includes(normalize(fragment));

const excelColumn = (reference: string) => {
  const letters = reference.match(/[A-Z]+/)?.[0] ?? "A";
  return [...letters].reduce((sum, letter) => sum * 26 + letter.charCodeAt(0) - 64, 0) - 1;
};

function parseSheet(xml: string, shared: string[]) {
  const output = new Map<number, unknown[]>();
  const rowPattern = /<row\b[^>]*\br="(\d+)"[^>]*>([\s\S]*?)<\/row>/g;
  let rowMatch: RegExpExecArray | null;
  while ((rowMatch = rowPattern.exec(xml))) {
    const rowNumber = Number(rowMatch[1]);
    const values: unknown[] = [];
    const cellPattern = /<c\b([^>]*)>([\s\S]*?)<\/c>/g;
    let cellMatch: RegExpExecArray | null;
    while ((cellMatch = cellPattern.exec(rowMatch[2]))) {
      const attributes = cellMatch[1];
      const body = cellMatch[2];
      const reference = attributes.match(/\br="([^"]+)"/)?.[1] || "A1";
      const column = excelColumn(reference);
      const type = attributes.match(/\bt="([^"]+)"/)?.[1];
      let result: unknown = null;
      if (type === "inlineStr") {
        result = [...body.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((match) => decodeXml(match[1])).join("");
      } else {
        const raw = body.match(/<v>([\s\S]*?)<\/v>/)?.[1] ?? "";
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
  const xml = strFromU8(source);
  return [...xml.matchAll(/<si>([\s\S]*?)<\/si>/g)].map((item) =>
    [...item[1].matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((text) => decodeXml(text[1])).join(""),
  );
}

function decodeXml(value: string) {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&")
    .replace(/&#(\d+);/g, (_, code) => String.fromCodePoint(Number(code)))
    .replace(/&#x([0-9a-f]+);/gi, (_, code) => String.fromCodePoint(Number.parseInt(code, 16)));
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
  if (!headerEntry) throw new Error("No se encontró el encabezado de cuentas contables de Siigo.");
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
  if (!company || !nit) throw new Error("Faltan la empresa o el NIT en las filas iniciales del archivo.");
  if (!records.length) throw new Error("El archivo no contiene filas Auxiliar / Sí para procesar.");
  return { name: file.name, company, nit, periodLabel: period.label, periodEnd: period.end, annual: period.annual, records };
}

type MacroEntry = { statement: "BG" | "P&G"; group: string; subgroup: string; item: string };

function esfLine(base: string, group: string, subgroup: string, item: string) {
  const code2 = base.slice(0, 2);
  if (group === "Activos") {
    if (includes(subgroup, "efectivo")) return "Efectivo y equivalentes";
    if (includes(subgroup, "inversion")) return code2 === "18" ? "Activos financieros no corrientes" : "Inversiones corrientes";
    if (includes(subgroup, "deudores comerciales")) return "Cuentas comerciales por cobrar";
    if (includes(subgroup, "otros deudores")) return "Otras cuentas por cobrar";
    if (includes(subgroup, "inventario") || includes(subgroup, "materiales")) return "Inventarios";
    if (includes(subgroup, "propiedad") || code2 === "15") return "Propiedad, planta y equipo";
    if (includes(item, "impuesto diferido") || includes(subgroup, "impuestos diferidos")) return "Impuesto diferido";
    if (includes(subgroup, "cuentas comerciales") && code2 === "18") return "Cuentas comerciales no corrientes";
    if (includes(subgroup, "otras cuentas por cobrar") && code2 === "18") return "Otras cuentas por cobrar no corrientes";
    if (includes(subgroup, "financieros") && code2 === "18") return "Activos financieros no corrientes";
    if (includes(subgroup, "intangible") || code2 === "16" || code2 === "19") return "Intangibles";
    if (includes(subgroup, "diferido") && code2 !== "18") return "Activos diferidos";
    return code2 === "18" || code2 === "19" ? "Otros activos no corrientes" : "Otros activos corrientes";
  }
  if (group === "Pasivos") {
    const nonCurrent = code2 === "29";
    if (includes(subgroup, "obligaciones financi")) return nonCurrent ? "Obligaciones financieras no corrientes" : "Obligaciones financieras";
    if (includes(subgroup, "socios") || includes(subgroup, "accionistas") || includes(subgroup, "deudas con accionistas")) return "Cuentas por pagar a socios";
    if (includes(subgroup, "proveedor")) return "Proveedores";
    if (includes(subgroup, "acreedores comerciales")) return nonCurrent ? "Acreedores comerciales no corrientes" : "Cuentas por pagar";
    if (includes(subgroup, "cuentas por pagar")) return nonCurrent ? "Otras cuentas por pagar no corrientes" : "Cuentas por pagar";
    if (includes(subgroup, "impuesto") || includes(subgroup, "gravamen")) return nonCurrent ? "Impuestos no corrientes" : "Impuestos por pagar";
    if (includes(subgroup, "salarios") || includes(subgroup, "beneficios")) return nonCurrent ? "Beneficios a empleados no corrientes" : "Beneficios a empleados";
    if (includes(subgroup, "provision")) return nonCurrent ? "Provisiones no corrientes" : "Provisiones";
    if (includes(subgroup, "vinculados")) return "Vinculados económicos";
    if (includes(subgroup, "diferido") || includes(item, "anticipado")) return "Ingresos diferidos";
    return nonCurrent ? "Otros pasivos no corrientes" : "Otros pasivos corrientes";
  }
  if (includes(subgroup, "prima")) return "Prima en colocación";
  if (includes(subgroup, "reserva")) return "Reservas";
  if (includes(subgroup, "superavit") || includes(subgroup, "revalu")) return "Superávit y revaluación";
  if (includes(subgroup, "utilidad acumulada")) return "Resultados acumulados";
  if (includes(item, "resultado ejercicio") || base === "3605" || base === "3610") return "Resultado del periodo";
  if (includes(subgroup, "capital")) return "Capital social";
  return "Otros componentes del patrimonio";
}

function fromMacro(base: string, entry: MacroEntry): Mapping {
  if (entry.statement === "BG") {
    const group = entry.group as "Activos" | "Pasivos" | "Patrimonio";
    return {
      statement: "ESF",
      group,
      subgroup: entry.subgroup,
      item: entry.item,
      line: esfLine(base, entry.group, entry.subgroup, entry.item),
      source: "Macro",
    };
  }
  const group = normalize(entry.group);
  const subgroup = normalize(entry.subgroup);
  let line: (typeof ER_DETAIL_LINES)[number] = "Otros gastos operacionales";
  let signGroup: "Ingresos" | "Gastos" = "Gastos";
  if (group === "ingresos") {
    signGroup = "Ingresos";
    line = subgroup.includes("devolucion") ? "Devoluciones y descuentos" : "Ingresos operacionales";
  } else if (group.includes("ingresos no") || group.includes("intereses ingreso")) {
    signGroup = "Ingresos";
    line = "Otros ingresos";
  } else if (group.includes("costo de ventas")) line = "Costo de ventas y servicios";
  else if (group.includes("depreciacion") || group.includes("amortizacion")) line = subgroup.includes("amort") ? "Amortización" : "Depreciación";
  else if (group.includes("intereses egreso")) line = "Gastos financieros";
  else if (group.includes("impuestos")) line = "Impuesto a las ganancias";
  else if (subgroup.includes("ventas")) line = "Gastos de ventas";
  else if (subgroup.includes("otros gastos")) line = "Otros gastos operacionales";
  else line = "Gastos de administración";
  return { statement: "ER", group: signGroup, subgroup: entry.subgroup, item: entry.item, line, source: "Macro" };
}

const explicitExtensions: Record<string, Mapping> = {
  "1625": { statement: "ESF", group: "Activos", subgroup: "Aportes en capital de trabajo", item: "Confirmar naturaleza", line: "Otros activos no corrientes", source: "PUC", needsReview: true },
  "2420": { statement: "ESF", group: "Pasivos", subgroup: "Impuestos por pagar", item: "Industria y comercio", line: "Impuestos por pagar", source: "PUC" },
  "4135": { statement: "ER", group: "Ingresos", subgroup: "Operacionales", item: "Comercio al por mayor y al detal", line: "Ingresos operacionales", source: "PUC" },
  "4180": { statement: "ER", group: "Ingresos", subgroup: "Operacionales", item: "Servicios", line: "Ingresos operacionales", source: "PUC" },
  "6180": { statement: "ER", group: "Gastos", subgroup: "Costo de servicios", item: "Servicios de transporte", line: "Costo de ventas y servicios", source: "PUC" },
};

export function defaultMapping(base: string): Mapping {
  const macro = (MACRO_MAPPING as Record<string, MacroEntry>)[base];
  if (macro) return fromMacro(base, macro);
  if (explicitExtensions[base]) return explicitExtensions[base];
  const first = base.slice(0, 1);
  const two = base.slice(0, 2);
  if (first === "1") {
    const lines: Record<string, string> = { "11": "Efectivo y equivalentes", "12": "Inversiones corrientes", "13": "Otras cuentas por cobrar", "14": "Inventarios", "15": "Propiedad, planta y equipo", "16": "Intangibles", "17": "Activos diferidos", "18": "Otros activos no corrientes", "19": "Otros activos no corrientes" };
    return { statement: "ESF", group: "Activos", subgroup: "Sugerencia PUC", item: "", line: lines[two] || "Otros activos corrientes", source: "PUC" };
  }
  if (first === "2") {
    const lines: Record<string, string> = { "21": "Obligaciones financieras", "22": "Proveedores", "23": "Cuentas por pagar", "24": "Impuestos por pagar", "25": "Beneficios a empleados", "26": "Provisiones", "27": "Ingresos diferidos", "28": "Otros pasivos corrientes", "29": "Otros pasivos no corrientes" };
    return { statement: "ESF", group: "Pasivos", subgroup: "Sugerencia PUC", item: "", line: lines[two] || "Otros pasivos corrientes", source: "PUC" };
  }
  if (first === "3") {
    const lines: Record<string, string> = { "31": "Capital social", "32": "Prima en colocación", "33": "Reservas", "34": "Otros componentes del patrimonio", "35": "Superávit y revaluación", "36": "Resultado del periodo", "37": "Resultados acumulados", "38": "Otros componentes del patrimonio" };
    return { statement: "ESF", group: "Patrimonio", subgroup: "Sugerencia PUC", item: "", line: lines[two] || "Otros componentes del patrimonio", source: "PUC" };
  }
  if (first === "4") {
    const returns = two === "41" && (base === "4170" || base === "4175");
    const other = two === "42";
    return { statement: "ER", group: "Ingresos", subgroup: "Sugerencia PUC", item: "", line: returns ? "Devoluciones y descuentos" : other ? "Otros ingresos" : "Ingresos operacionales", source: "PUC" };
  }
  if (first === "5") {
    const lines: Record<string, string> = { "51": "Gastos de administración", "52": "Gastos de ventas", "53": "Otros gastos operacionales", "54": "Impuesto a las ganancias" };
    return { statement: "ER", group: "Gastos", subgroup: "Sugerencia PUC", item: "", line: lines[two] || "Otros gastos operacionales", source: "PUC" };
  }
  if (first === "6" || first === "7") return { statement: "ER", group: "Gastos", subgroup: "Costo de ventas", item: "", line: "Costo de ventas y servicios", source: "PUC" };
  return { statement: "REVISAR", group: "Sin clasificar", subgroup: "Sin clasificar", item: "", line: "Cuenta por revisar", source: "PUC", needsReview: true };
}

export function mappingFromLine(line: string): Mapping {
  if ((ER_DETAIL_LINES as readonly string[]).includes(line)) {
    const income = line === "Ingresos operacionales" || line === "Devoluciones y descuentos" || line === "Otros ingresos";
    return { statement: "ER", group: income ? "Ingresos" : "Gastos", subgroup: "Clasificación manual", item: "", line, source: "Manual" };
  }
  if ((ESF_DETAIL_LINES as readonly string[]).includes(line)) {
    const assets = [...CURRENT_ASSETS, ...NON_CURRENT_ASSETS] as readonly string[];
    const liabilities = [...CURRENT_LIABILITIES, ...NON_CURRENT_LIABILITIES] as readonly string[];
    return { statement: "ESF", group: assets.includes(line) ? "Activos" : liabilities.includes(line) ? "Pasivos" : "Patrimonio", subgroup: "Clasificación manual", item: "", line, source: "Manual" };
  }
  return { statement: "REVISAR", group: "Sin clasificar", subgroup: "Sin clasificar", item: "", line: "Cuenta por revisar", source: "Manual", needsReview: true };
}

const add = (target: Record<string, number>, key: string, value: number) => {
  target[key] = (target[key] || 0) + value;
};

const sum = (source: Record<string, number>, lines: readonly string[]) => lines.reduce((total, line) => total + (source[line] || 0), 0);

export type PeriodResult = {
  label: string;
  date: string;
  annual: boolean;
  rows: number;
  er: Record<string, number>;
  esf: Record<string, number>;
  check: {
    debitCredit: number;
    movement: number;
    unmapped: number;
    review: number;
    unmappedAmount: number;
    balance: number;
    status: "OK" | "REVISAR";
  };
};

export type CashFlowResult = { label: string; date: string; previousLabel: string; lines: Record<string, number>; residual: number };
export type EquityChangeResult = { label: string; date: string; previousLabel: string; lines: Record<string, number> };

export function calculate(records: TrialRecord[], overrides: Record<string, string>) {
  const grouped = new Map<string, TrialRecord[]>();
  for (const record of records) grouped.set(record.periodEnd, [...(grouped.get(record.periodEnd) || []), record]);

  const mapRecord = (record: TrialRecord) => overrides[record.base] ? mappingFromLine(overrides[record.base]) : defaultMapping(record.base);

  const periods: PeriodResult[] = [...grouped.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([date, periodRecords]) => {
      const er: Record<string, number> = {};
      const esf: Record<string, number> = {};
      const unmapped = new Set<string>();
      const review = new Set<string>();
      let debitCredit = 0;
      let movement = 0;
      let unmappedAmount = 0;
      let openProfitAndLossBalance = 0;
      for (const record of periodRecords) {
        const mapping = mapRecord(record);
        debitCredit += record.debit - record.credit;
        movement += record.initial + record.debit - record.credit - record.final;
        if (mapping.statement === "REVISAR") {
          unmapped.add(record.base);
          unmappedAmount += Math.abs(record.final || record.debit - record.credit);
          continue;
        }
        if (mapping.needsReview) {
          review.add(record.base);
          unmappedAmount += Math.abs(record.final || record.debit - record.credit);
        }
        if (mapping.statement === "ER") {
          add(er, mapping.line, mapping.group === "Ingresos" ? record.credit - record.debit : record.debit - record.credit);
          openProfitAndLossBalance += record.final;
        }
        if (mapping.statement === "ESF") add(esf, mapping.line, mapping.group === "Activos" ? record.final : -record.final);
      }

      er["Ingresos netos"] = (er["Ingresos operacionales"] || 0) + (er["Devoluciones y descuentos"] || 0);
      er["Utilidad bruta"] = er["Ingresos netos"] - (er["Costo de ventas y servicios"] || 0);
      er["Total gastos operacionales"] = (er["Gastos de administración"] || 0) + (er["Gastos de ventas"] || 0) + (er["Otros gastos operacionales"] || 0);
      er["Utilidad operativa"] = er["Utilidad bruta"] - er["Total gastos operacionales"];
      er["EBITDA"] = er["Utilidad operativa"] + (er["Depreciación"] || 0) + (er["Amortización"] || 0);
      er["EBIT"] = er["EBITDA"] - (er["Depreciación"] || 0) - (er["Amortización"] || 0);
      er["Resultado antes de impuestos"] = er["EBIT"] + (er["Otros ingresos"] || 0) - (er["Gastos financieros"] || 0);
      er["Resultado neto"] = er["Resultado antes de impuestos"] - (er["Impuesto a las ganancias"] || 0);

      esf["Total activos corrientes"] = sum(esf, CURRENT_ASSETS);
      esf["Total activos no corrientes"] = sum(esf, NON_CURRENT_ASSETS);
      esf["Total activos"] = esf["Total activos corrientes"] + esf["Total activos no corrientes"];
      esf["Total pasivos corrientes"] = sum(esf, CURRENT_LIABILITIES);
      esf["Total pasivos no corrientes"] = sum(esf, NON_CURRENT_LIABILITIES);
      esf["Total pasivos"] = esf["Total pasivos corrientes"] + esf["Total pasivos no corrientes"];
      // Siigo mantiene las cuentas 4–7 abiertas durante el año. Su saldo de cierre
      // se presenta en patrimonio para conciliar el balance, sin usar una partida residual.
      esf["Resultado del periodo"] = (esf["Resultado del periodo"] || 0) - openProfitAndLossBalance;
      esf["Total patrimonio"] = sum(esf, EQUITY_LINES);
      esf["Total pasivo y patrimonio"] = esf["Total pasivos"] + esf["Total patrimonio"];
      const balance = esf["Total activos"] - esf["Total pasivo y patrimonio"];
      const ok = Math.abs(debitCredit) < 1 && Math.abs(movement) < 1 && unmapped.size === 0 && review.size === 0 && Math.abs(balance) < 1;
      const first = periodRecords[0];
      return {
        label: first.periodLabel,
        date,
        annual: first.annual,
        rows: periodRecords.length,
        er,
        esf,
        check: { debitCredit, movement, unmapped: unmapped.size, review: review.size, unmappedAmount, balance, status: ok ? "OK" : "REVISAR" },
      };
    });

  const cashFlows: CashFlowResult[] = periods.slice(1).map((period, index) => {
    const previous = periods[index];
    const receivableLines = ["Cuentas comerciales por cobrar", "Otras cuentas por cobrar", "Cuentas comerciales no corrientes", "Otras cuentas por cobrar no corrientes"];
    const operatingAssetLines = ["Activos diferidos", "Otros activos corrientes", "Impuesto diferido", "Otros activos no corrientes"];
    const operatingLiabilityLines = [...CURRENT_LIABILITIES, ...NON_CURRENT_LIABILITIES].filter((line) => !line.startsWith("Obligaciones financieras"));
    const lines: Record<string, number> = {
      "Resultado neto": period.er["Resultado neto"] || 0,
      "Depreciación y amortización": (period.er["Depreciación"] || 0) + (period.er["Amortización"] || 0),
      "Variación de cuentas por cobrar": sum(previous.esf, receivableLines) - sum(period.esf, receivableLines),
      "Variación de inventarios": (previous.esf["Inventarios"] || 0) - (period.esf["Inventarios"] || 0),
      "Variación de otros activos operativos": sum(previous.esf, operatingAssetLines) - sum(period.esf, operatingAssetLines),
      "Variación de pasivos operativos": sum(period.esf, operatingLiabilityLines) - sum(previous.esf, operatingLiabilityLines),
    };
    lines["Flujo de operación"] = sum(lines, ["Resultado neto", "Depreciación y amortización", "Variación de cuentas por cobrar", "Variación de inventarios", "Variación de otros activos operativos", "Variación de pasivos operativos"]);
    lines["Flujo de inversión"] = -(
      (period.esf["Inversiones corrientes"] || 0) - (previous.esf["Inversiones corrientes"] || 0) +
      (period.esf["Propiedad, planta y equipo"] || 0) - (previous.esf["Propiedad, planta y equipo"] || 0) +
      (period.esf["Intangibles"] || 0) - (previous.esf["Intangibles"] || 0) +
      (period.esf["Activos financieros no corrientes"] || 0) - (previous.esf["Activos financieros no corrientes"] || 0)
    );
    const debtLines = ["Obligaciones financieras", "Obligaciones financieras no corrientes"];
    lines["Flujo de financiación"] = sum(period.esf, debtLines) - sum(previous.esf, debtLines) + (period.esf["Total patrimonio"] || 0) - (previous.esf["Total patrimonio"] || 0) - (period.er["Resultado neto"] || 0);
    const classified = lines["Flujo de operación"] + lines["Flujo de inversión"] + lines["Flujo de financiación"];
    const cashChange = (period.esf["Efectivo y equivalentes"] || 0) - (previous.esf["Efectivo y equivalentes"] || 0);
    lines["Otras variaciones por conciliar"] = cashChange - classified;
    lines["Variación neta del efectivo"] = cashChange;
    lines["Efectivo inicial"] = previous.esf["Efectivo y equivalentes"] || 0;
    lines["Efectivo final"] = period.esf["Efectivo y equivalentes"] || 0;
    return { label: period.label, date: period.date, previousLabel: previous.label, lines, residual: lines["Otras variaciones por conciliar"] };
  });

  const equityChanges: EquityChangeResult[] = periods.slice(1).map((period, index) => {
    const previous = periods[index];
    const directLines = ["Capital social", "Prima en colocación", "Reservas", "Superávit y revaluación", "Otros componentes del patrimonio"];
    const lines: Record<string, number> = { "Patrimonio inicial": previous.esf["Total patrimonio"] || 0 };
    for (const line of directLines) lines[`Variación · ${line}`] = (period.esf[line] || 0) - (previous.esf[line] || 0);
    lines["Resultado del periodo"] = period.er["Resultado neto"] || 0;
    lines["Otros movimientos en resultados acumulados"] = (period.esf["Total patrimonio"] || 0) - lines["Patrimonio inicial"] - sum(lines, directLines.map((line) => `Variación · ${line}`)) - lines["Resultado del periodo"];
    lines["Patrimonio final"] = period.esf["Total patrimonio"] || 0;
    return { label: period.label, date: period.date, previousLabel: previous.label, lines };
  });

  const latest = periods.at(-1);
  const latestRecords = latest ? grouped.get(latest.date) || [] : [];
  const thirdParties = new Map<string, { identification: string; name: string; receivable: number; payable: number; taxes: number; other: number; net: number }>();
  for (const record of latestRecords) {
    if (!record.identification) continue;
    const mapping = mapRecord(record);
    if (mapping.statement !== "ESF") continue;
    const current = thirdParties.get(record.identification) || { identification: record.identification, name: record.thirdParty || "Sin nombre", receivable: 0, payable: 0, taxes: 0, other: 0, net: 0 };
    const value = mapping.group === "Activos" ? record.final : -record.final;
    if (mapping.line.includes("por cobrar")) current.receivable += value;
    else if (mapping.line === "Proveedores" || mapping.line.includes("Cuentas por pagar") || mapping.line.includes("Acreedores")) current.payable += value;
    else if (mapping.line.includes("Impuestos")) current.taxes += value;
    else if (mapping.group === "Pasivos") current.other += value;
    current.net = current.receivable - current.payable - current.taxes - current.other;
    thirdParties.set(record.identification, current);
  }

  const accounts = new Map<string, { base: string; name: string; mapping: Mapping; records: number; amount: number }>();
  const expenseMap = new Map<string, { group: string; subgroup: string; item: string; values: Record<string, number>; total: number }>();
  for (const record of records) {
    const current = accounts.get(record.base);
    const mapping = mapRecord(record);
    accounts.set(record.base, { base: record.base, name: current?.name || record.accountName, mapping, records: (current?.records || 0) + 1, amount: (current?.amount || 0) + Math.abs(record.final || record.debit - record.credit) });
    if (mapping.statement === "ER" && mapping.group === "Gastos") {
      const key = `${mapping.line}::${mapping.subgroup}::${mapping.item || record.accountName}`;
      const currentExpense = expenseMap.get(key) || { group: mapping.line, subgroup: mapping.subgroup, item: mapping.item || record.accountName, values: {}, total: 0 };
      const value = record.debit - record.credit;
      add(currentExpense.values, record.periodEnd, value);
      currentExpense.total += value;
      expenseMap.set(key, currentExpense);
    }
  }

  return {
    periods,
    latest,
    cashFlows,
    equityChanges,
    latestCashFlow: cashFlows.at(-1),
    latestEquityChange: equityChanges.at(-1),
    thirdParties: [...thirdParties.values()].sort((a, b) => Math.abs(b.net) - Math.abs(a.net)),
    accounts: [...accounts.values()].sort((a, b) => Number(b.mapping.needsReview) - Number(a.mapping.needsReview) || a.base.localeCompare(b.base)),
    expenses: [...expenseMap.values()].sort((a, b) => Math.abs(b.total) - Math.abs(a.total)),
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
