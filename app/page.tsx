"use client";

import { ChangeEvent, DragEvent, ReactNode, useMemo, useRef, useState } from "react";
import {
  calculate,
  csvDownload,
  CURRENT_ASSETS,
  CURRENT_LIABILITIES,
  EQUITY_LINES,
  MAPPING_OPTIONS,
  NON_CURRENT_ASSETS,
  NON_CURRENT_LIABILITIES,
  parseWorkbook,
  type ParsedFile,
  type PeriodResult,
} from "./lib/financials";

type Tab = "resumen" | "resultados" | "situacion" | "patrimonio" | "flujo" | "terceros" | "analisis" | "mapeo" | "controles";

const tabs: { id: Tab; label: string; short: string }[] = [
  { id: "resumen", label: "Resumen ejecutivo", short: "Resumen" },
  { id: "resultados", label: "Estado de resultados", short: "Resultados" },
  { id: "situacion", label: "Situación financiera", short: "Situación" },
  { id: "patrimonio", label: "Cambios en patrimonio", short: "Patrimonio" },
  { id: "flujo", label: "Flujo de efectivo", short: "Flujo" },
  { id: "terceros", label: "Cartera y proveedores", short: "Terceros" },
  { id: "analisis", label: "Gastos y costos", short: "Gastos" },
  { id: "mapeo", label: "Mapeo de cuentas", short: "Mapeo" },
  { id: "controles", label: "Controles", short: "Controles" },
];

const money = new Intl.NumberFormat("es-CO", { maximumFractionDigits: 0 });
const compact = new Intl.NumberFormat("es-CO", { notation: "compact", maximumFractionDigits: 1 });
const percentage = new Intl.NumberFormat("es-CO", { maximumFractionDigits: 1 });
const formatMoney = (value = 0) => `${value < 0 ? "(" : ""}$ ${money.format(Math.abs(value))}${value < 0 ? ")" : ""}`;
const formatCompact = (value = 0) => `$ ${compact.format(value)}`;
const safeRatio = (numerator: number, denominator: number) => denominator ? numerator / denominator : 0;

type DisplayRow = { label: string; section?: boolean; total?: boolean; highlight?: boolean };

const resultRows: DisplayRow[] = [
  { label: "Ingresos", section: true },
  { label: "Ingresos operacionales" },
  { label: "Devoluciones y descuentos" },
  { label: "Ingresos netos", total: true },
  { label: "Costo de ventas y servicios" },
  { label: "Utilidad bruta", total: true },
  { label: "Operación", section: true },
  { label: "Gastos de administración" },
  { label: "Gastos de ventas" },
  { label: "Otros gastos operacionales" },
  { label: "Total gastos operacionales", total: true },
  { label: "Utilidad operativa", total: true },
  { label: "EBITDA", total: true, highlight: true },
  { label: "Depreciación" },
  { label: "Amortización" },
  { label: "EBIT", total: true },
  { label: "Otros ingresos" },
  { label: "Gastos financieros" },
  { label: "Resultado antes de impuestos", total: true },
  { label: "Impuesto a las ganancias" },
  { label: "Resultado neto", total: true, highlight: true },
];

const balanceRows: DisplayRow[] = [
  { label: "Activos corrientes", section: true },
  ...CURRENT_ASSETS.map((label) => ({ label })),
  { label: "Total activos corrientes", total: true },
  { label: "Activos no corrientes", section: true },
  ...NON_CURRENT_ASSETS.map((label) => ({ label })),
  { label: "Total activos no corrientes", total: true },
  { label: "Total activos", total: true, highlight: true },
  { label: "Pasivos corrientes", section: true },
  ...CURRENT_LIABILITIES.map((label) => ({ label })),
  { label: "Total pasivos corrientes", total: true },
  { label: "Pasivos no corrientes", section: true },
  ...NON_CURRENT_LIABILITIES.map((label) => ({ label })),
  { label: "Total pasivos no corrientes", total: true },
  { label: "Total pasivos", total: true },
  { label: "Patrimonio", section: true },
  ...EQUITY_LINES.map((label) => ({ label })),
  { label: "Total patrimonio", total: true },
  { label: "Total pasivo y patrimonio", total: true, highlight: true },
];

const cashFlowRows: DisplayRow[] = [
  { label: "Actividades de operación", section: true },
  { label: "Resultado neto" },
  { label: "Depreciación y amortización" },
  { label: "Variación de cuentas por cobrar" },
  { label: "Variación de inventarios" },
  { label: "Variación de otros activos operativos" },
  { label: "Variación de pasivos operativos" },
  { label: "Flujo de operación", total: true },
  { label: "Actividades de inversión y financiación", section: true },
  { label: "Flujo de inversión", total: true },
  { label: "Flujo de financiación", total: true },
  { label: "Otras variaciones por conciliar" },
  { label: "Variación neta del efectivo", total: true, highlight: true },
  { label: "Efectivo inicial" },
  { label: "Efectivo final", total: true },
];

function SectionHeading({ eyebrow, title, description, action }: { eyebrow: string; title: string; description?: string; action?: ReactNode }) {
  return (
    <div className="section-heading">
      <div><span className="eyebrow">{eyebrow}</span><h3>{title}</h3>{description && <p>{description}</p>}</div>
      {action}
    </div>
  );
}

function FinancialTable({ periods, type }: { periods: PeriodResult[]; type: "er" | "esf" }) {
  const rows = type === "er" ? resultRows : balanceRows;
  const field = type === "er" ? "er" : "esf";
  return (
    <div className="table-scroll">
      <table className="financial-table">
        <thead><tr><th>Concepto</th>{periods.map((period) => <th key={period.date}>{period.label}<small>{period.annual ? "Anual" : "Movimiento mensual"}</small></th>)}</tr></thead>
        <tbody>
          {rows.map((row) => row.section ? (
            <tr className="section-row" key={row.label}><td colSpan={periods.length + 1}>{row.label}</td></tr>
          ) : (
            <tr className={`${row.total ? "total-row" : ""} ${row.highlight ? "highlight-row" : ""}`} key={row.label}>
              <td>{row.label}</td>
              {periods.map((period) => {
                const value = period[field][row.label] || 0;
                return <td className={value < 0 ? "negative" : ""} key={period.date}>{formatMoney(value)}</td>;
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

export default function Home() {
  const [uploaded, setUploaded] = useState<ParsedFile[]>([]);
  const [overrides, setOverrides] = useState<Record<string, string>>({});
  const [tab, setTab] = useState<Tab>("resumen");
  const [busy, setBusy] = useState(false);
  const [dragging, setDragging] = useState(false);
  const [message, setMessage] = useState("");
  const [partySearch, setPartySearch] = useState("");
  const [mappingSearch, setMappingSearch] = useState("");
  const inputRef = useRef<HTMLInputElement>(null);

  const records = useMemo(() => uploaded.flatMap((file) => file.records), [uploaded]);
  const analysis = useMemo(() => calculate(records, overrides), [records, overrides]);
  const latest = analysis.latest;
  const company = uploaded[0]?.company || "Tu empresa";
  const nit = uploaded[0]?.nit || "";
  const allOk = analysis.periods.length > 0 && analysis.periods.every((period) => period.check.status === "OK");
  const currentYearPeriods = latest ? analysis.periods.filter((period) => period.date.slice(0, 4) === latest.date.slice(0, 4) && !period.annual) : [];
  const ytdRevenue = currentYearPeriods.reduce((total, period) => total + (period.er["Ingresos netos"] || 0), 0);
  const ytdResult = currentYearPeriods.reduce((total, period) => total + (period.er["Resultado neto"] || 0), 0);
  const currentAssets = latest?.esf["Total activos corrientes"] || 0;
  const currentLiabilities = latest?.esf["Total pasivos corrientes"] || 0;
  const workingCapital = currentAssets - currentLiabilities;
  const currentRatio = safeRatio(currentAssets, currentLiabilities);
  const leverage = safeRatio(latest?.esf["Total pasivos"] || 0, latest?.esf["Total activos"] || 0);
  const roe = safeRatio(ytdResult, latest?.esf["Total patrimonio"] || 0);
  const netMargin = safeRatio(ytdResult, ytdRevenue);
  const maxTrend = Math.max(1, ...currentYearPeriods.flatMap((period) => [Math.abs(period.er["Ingresos netos"] || 0), Math.abs(period.er["Resultado neto"] || 0)]));
  const filteredParties = analysis.thirdParties.filter((party) => `${party.identification} ${party.name}`.toLowerCase().includes(partySearch.toLowerCase()));
  const filteredAccounts = analysis.accounts.filter((account) => `${account.base} ${account.name} ${account.mapping.line}`.toLowerCase().includes(mappingSearch.toLowerCase()));

  async function loadFiles(fileList: FileList | File[]) {
    const candidates = Array.from(fileList).filter((file) => file.name.toLowerCase().endsWith(".xlsx"));
    if (!candidates.length) {
      setMessage("Selecciona uno o varios balances de Siigo en formato .xlsx.");
      return;
    }
    setBusy(true);
    setMessage("");
    const settled = await Promise.allSettled(candidates.map(parseWorkbook));
    const parsed = settled.filter((item): item is PromiseFulfilledResult<ParsedFile> => item.status === "fulfilled").map((item) => item.value);
    const expectedNit = uploaded[0]?.nit || parsed[0]?.nit;
    const accepted = parsed.filter((file) => file.nit === expectedNit);
    const companyMismatch = parsed.length - accepted.length;
    const failures = settled.filter((item): item is PromiseRejectedResult => item.status === "rejected");
    setUploaded((current) => {
      const next = new Map(current.map((file) => [file.periodEnd, file]));
      accepted.forEach((file) => next.set(file.periodEnd, file));
      return [...next.values()].sort((a, b) => a.periodEnd.localeCompare(b.periodEnd));
    });
    const notes = [`${accepted.length} periodo(s) procesado(s)`];
    if (failures.length) notes.push(`${failures.length} archivo(s) no válido(s): ${failures.map((item) => item.reason?.message || "error desconocido").join(" · ")}`);
    if (companyMismatch) notes.push(`${companyMismatch} archivo(s) corresponde(n) a otro NIT`);
    setMessage(`${notes.join(". ")}.`);
    setBusy(false);
    if (accepted.length) setTab("resumen");
  }

  function onDrop(event: DragEvent<HTMLDivElement>) {
    event.preventDefault();
    setDragging(false);
    loadFiles(event.dataTransfer.files);
  }

  function downloadStatement(type: "er" | "esf") {
    const rows = (type === "er" ? resultRows : balanceRows).filter((row) => !row.section);
    const field = type === "er" ? "er" : "esf";
    csvDownload(type === "er" ? "estado-de-resultados.csv" : "situacion-financiera.csv", [
      ["Concepto", ...analysis.periods.map((period) => period.label)],
      ...rows.map((row) => [row.label, ...analysis.periods.map((period) => period[field][row.label] || 0)]),
    ]);
  }

  const uploader = (
    <div
      className={`dropzone ${dragging ? "is-dragging" : ""}`}
      onDragOver={(event) => { event.preventDefault(); setDragging(true); }}
      onDragLeave={() => setDragging(false)}
      onDrop={onDrop}
    >
      <img src="/numbi-upload.png" alt="" />
      <div><strong>{busy ? "Procesando balances…" : "Arrastra aquí los balances de Siigo"}</strong><span>Uno o varios periodos · formato .xlsx</span></div>
      <button onClick={() => inputRef.current?.click()} disabled={busy}>{busy ? "Leyendo…" : "Seleccionar archivos"}</button>
    </div>
  );

  return (
    <main>
      <input ref={inputRef} type="file" accept=".xlsx" multiple hidden onChange={(event: ChangeEvent<HTMLInputElement>) => event.target.files && loadFiles(event.target.files)} />
      <header className="topbar">
        <div className="brand-lockup"><span className="wordmark">numbi</span><span className="product-name">EEFF</span></div>
        <div className="privacy-pill"><span className="privacy-dot" /> Tus archivos no salen del navegador</div>
      </header>

      {!uploaded.length ? (
        <>
          <section className="hero">
            <div className="hero-copy">
              <span className="eyebrow">Estados financieros sin reprocesos</span>
              <h1>Tus estados financieros, <span>listos para decidir.</span></h1>
              <p>Sube el balance de prueba por tercero descargado de Siigo. Numbi valida el archivo, aplica la estructura de tu macro y prepara una lectura financiera completa.</p>
              {uploader}
              {message && <p className="upload-message">{message}</p>}
              <div className="trust-row"><span>✓ Procesamiento local</span><span>✓ Mapeo auditable</span><span>✓ Controles sin ajustes ocultos</span></div>
            </div>
            <div className="hero-panel" aria-label="Proceso de la herramienta">
              <div className="panel-head"><span>Del balance a la decisión</span><span className="live-dot">3 pasos</span></div>
              <div className="flow-step"><img src="/numbi-upload.png" alt="" /><div><small>01</small><strong>Carga</strong><span>Uno o varios balances por tercero</span></div></div>
              <div className="flow-line" />
              <div className="flow-step"><img src="/numbi-calculator.png" alt="" /><div><small>02</small><strong>Valida</strong><span>Formato, movimientos, mapeo y ecuación</span></div></div>
              <div className="flow-line" />
              <div className="flow-step"><img src="/numbi-growth.png" alt="" /><div><small>03</small><strong>Presenta</strong><span>EEFF, indicadores y análisis por tercero</span></div></div>
              <img className="robot" src="/numbi-assistant.png" alt="Asistente Numbi" />
            </div>
          </section>
          <section className="promise-grid">
            <article><img src="/numbi-report.png" alt="" /><div><span>Alcance de la macro</span><strong>Cuatro estados financieros</strong><p>Resultados, situación financiera, cambios en patrimonio y flujo de efectivo.</p></div></article>
            <article><img src="/numbi-approved.png" alt="" /><div><span>Control visible</span><strong>Sin cuadrar por diferencia</strong><p>La herramienta muestra la diferencia contable real y las cuentas por confirmar.</p></div></article>
            <article><img src="/numbi-calculator.png" alt="" /><div><span>Formato esperado</span><strong>Balance de Siigo</strong><p>Lee las filas Auxiliar / Sí para evitar dobles conteos entre niveles.</p></div></article>
          </section>
        </>
      ) : (
        <section className="workspace">
          <div className="workspace-header">
            <div><span className="eyebrow">Presentación de estados financieros</span><h1>{company}</h1><p>{nit ? `NIT ${nit} · ` : ""}{records.length.toLocaleString("es-CO")} registros auxiliares · corte {latest?.label}</p></div>
            <div className={`model-status ${allOk ? "ok" : "review"}`}><span>{allOk ? "✓" : "!"}</span><div><small>Estado del modelo</small><strong>{allOk ? "Todo conciliado" : "Requiere confirmación"}</strong></div></div>
          </div>

          <div className="import-bar">
            <div className="file-chips">
              {uploaded.map((file) => <button key={file.periodEnd} title="Quitar periodo" onClick={() => setUploaded((current) => current.filter((item) => item.periodEnd !== file.periodEnd))}><span>{file.periodLabel}</span><b>×</b></button>)}
            </div>
            <button className="secondary-button" onClick={() => inputRef.current?.click()}>＋ Agregar balances</button>
          </div>
          {message && <p className="workspace-message">{message}</p>}

          <div className="dashboard-grid">
            <nav className="side-nav" aria-label="Secciones del reporte">
              {tabs.map((item, index) => <button className={tab === item.id ? "active" : ""} onClick={() => setTab(item.id)} key={item.id}><span>{String(index + 1).padStart(2, "0")}</span>{item.label}</button>)}
            </nav>
            <select className="mobile-nav" value={tab} onChange={(event) => setTab(event.target.value as Tab)}>{tabs.map((item) => <option value={item.id} key={item.id}>{item.label}</option>)}</select>

            <div className="dashboard-content">
              {tab === "resumen" && latest && (
                <section className="report-section">
                  <SectionHeading eyebrow={`Último corte · ${latest.label}`} title="Resumen ejecutivo" description="Indicadores calculados directamente desde los balances cargados." />
                  <div className="kpi-grid">
                    <article><span>Capital de trabajo</span><strong>{formatCompact(workingCapital)}</strong><small>Activo corriente − pasivo corriente</small></article>
                    <article className={currentRatio < 1 ? "alert" : "positive"}><span>Razón corriente</span><strong>{percentage.format(currentRatio)}x</strong><small>Capacidad de pago de corto plazo</small></article>
                    <article><span>Endeudamiento</span><strong>{percentage.format(leverage * 100)}%</strong><small>Pasivos sobre activos</small></article>
                    <article className={ytdResult < 0 ? "alert" : "positive"}><span>Resultado acumulado</span><strong>{formatCompact(ytdResult)}</strong><small>Margen neto {percentage.format(netMargin * 100)}%</small></article>
                    <article><span>ROE</span><strong>{percentage.format(roe * 100)}%</strong><small>Resultado sobre patrimonio</small></article>
                  </div>
                  <div className="insight-grid">
                    <article className="trend-card">
                      <div className="card-title"><div><span>Tendencia mensual</span><strong>Ingresos y resultado neto</strong></div><div className="legend"><i />Ingresos <i />Resultado</div></div>
                      <div className="bar-chart">
                        {currentYearPeriods.map((period) => {
                          const revenue = period.er["Ingresos netos"] || 0;
                          const result = period.er["Resultado neto"] || 0;
                          return <div className="bar-group" key={period.date}><div className="bars"><i style={{ height: `${Math.max(3, Math.abs(revenue) / maxTrend * 100)}%` }} title={formatMoney(revenue)} /><i className={result < 0 ? "down" : "up"} style={{ height: `${Math.max(3, Math.abs(result) / maxTrend * 100)}%` }} title={formatMoney(result)} /></div><span>{period.label.split(" ")[0].slice(0, 3)}</span></div>;
                        })}
                      </div>
                    </article>
                    <article className="decision-card">
                      <div className="card-title"><div><span>Lectura del corte</span><strong>Señales para revisar</strong></div></div>
                      <ul>
                        <li className={latest.check.review ? "warning" : "good"}><b>{latest.check.review ? `${latest.check.review} cuenta(s) por confirmar` : "Mapeo confirmado"}</b><span>{latest.check.review ? `Impacto visible de ${formatMoney(latest.check.unmappedAmount)}.` : "No hay clasificaciones pendientes."}</span></li>
                        <li className={Math.abs(latest.check.balance) >= 1 ? "warning" : "good"}><b>{Math.abs(latest.check.balance) >= 1 ? "Diferencia en la ecuación" : "Ecuación contable conciliada"}</b><span>{formatMoney(latest.check.balance)}</span></li>
                        <li className={workingCapital < 0 ? "warning" : "good"}><b>{workingCapital < 0 ? "Capital de trabajo negativo" : "Capital de trabajo positivo"}</b><span>{formatMoney(workingCapital)}</span></li>
                      </ul>
                    </article>
                  </div>
                </section>
              )}

              {tab === "resultados" && <section className="report-section"><SectionHeading eyebrow="COP · pesos colombianos" title="Estado de resultados" description="Movimientos del periodo, conservando el nivel de detalle de la macro." action={<button onClick={() => downloadStatement("er")}>Descargar CSV</button>} /><FinancialTable periods={analysis.periods} type="er" /></section>}
              {tab === "situacion" && <section className="report-section"><SectionHeading eyebrow="COP · saldos al cierre" title="Estado de situación financiera" description="La diferencia se calcula sin incorporar cuentas de ajuste ni resultados residuales." action={<button onClick={() => downloadStatement("esf")}>Descargar CSV</button>} /><FinancialTable periods={analysis.periods} type="esf" /></section>}

              {tab === "patrimonio" && (
                <section className="report-section">
                  <SectionHeading eyebrow={analysis.latestEquityChange ? `${analysis.latestEquityChange.previousLabel} → ${analysis.latestEquityChange.label}` : "Comparativo requerido"} title="Estado de cambios en el patrimonio" description="Conciliación entre los dos cortes más recientes." />
                  {analysis.latestEquityChange ? <div className="table-scroll compact-table"><table><thead><tr><th>Movimiento</th><th>Valor</th></tr></thead><tbody>{Object.entries(analysis.latestEquityChange.lines).map(([line, value]) => <tr className={line.includes("final") ? "highlight-row" : ""} key={line}><td>{line}</td><td className={value < 0 ? "negative" : ""}>{formatMoney(value)}</td></tr>)}</tbody></table></div> : <EmptyState icon="/numbi-report.png" title="Carga al menos dos periodos" text="El cambio patrimonial se calcula comparando dos saldos de cierre." />}
                </section>
              )}

              {tab === "flujo" && (
                <section className="report-section">
                  <SectionHeading eyebrow={analysis.latestCashFlow ? `${analysis.latestCashFlow.previousLabel} → ${analysis.latestCashFlow.label}` : "Método indirecto"} title="Estado de flujo de efectivo" description="Estimado desde variaciones de balance. Las partidas no identificadas quedan expuestas para conciliación." />
                  {analysis.latestCashFlow ? <div className="table-scroll compact-table"><table><thead><tr><th>Concepto</th><th>Valor</th></tr></thead><tbody>{cashFlowRows.map((row) => row.section ? <tr className="section-row" key={row.label}><td colSpan={2}>{row.label}</td></tr> : <tr className={`${row.total ? "total-row" : ""} ${row.highlight ? "highlight-row" : ""}`} key={row.label}><td>{row.label}</td><td className={(analysis.latestCashFlow?.lines[row.label] || 0) < 0 ? "negative" : ""}>{formatMoney(analysis.latestCashFlow?.lines[row.label] || 0)}</td></tr>)}</tbody></table></div> : <EmptyState icon="/numbi-growth.png" title="Carga al menos dos periodos" text="El flujo indirecto necesita un saldo inicial y uno final." />}
                </section>
              )}

              {tab === "terceros" && (
                <section className="report-section"><SectionHeading eyebrow={`Último corte · ${latest?.label}`} title="Cartera y proveedores por tercero" action={<input aria-label="Buscar tercero" placeholder="Buscar NIT o tercero…" value={partySearch} onChange={(event) => setPartySearch(event.target.value)} />} /><div className="table-scroll"><table><thead><tr><th>Identificación</th><th>Tercero</th><th>Cuentas por cobrar</th><th>Proveedores y CxP</th><th>Impuestos</th><th>Otros pasivos</th><th>Saldo neto</th></tr></thead><tbody>{filteredParties.slice(0, 150).map((party) => <tr key={party.identification}><td>{party.identification}</td><td>{party.name}</td><td>{formatMoney(party.receivable)}</td><td>{formatMoney(party.payable)}</td><td>{formatMoney(party.taxes)}</td><td>{formatMoney(party.other)}</td><td className={party.net < 0 ? "negative" : ""}>{formatMoney(party.net)}</td></tr>)}</tbody></table></div></section>
              )}

              {tab === "analisis" && (
                <section className="report-section"><SectionHeading eyebrow="Apertura por naturaleza" title="Análisis de gastos y costos" description="Ordenado por impacto acumulado en los periodos cargados." /><div className="table-scroll"><table><thead><tr><th>Línea</th><th>Subgrupo</th><th>Detalle</th>{analysis.periods.map((period) => <th key={period.date}>{period.label}</th>)}<th>Total</th></tr></thead><tbody>{analysis.expenses.slice(0, 120).map((expense) => <tr key={`${expense.group}-${expense.subgroup}-${expense.item}`}><td>{expense.group}</td><td>{expense.subgroup}</td><td>{expense.item}</td>{analysis.periods.map((period) => <td key={period.date}>{formatMoney(expense.values[period.date] || 0)}</td>)}<td><b>{formatMoney(expense.total)}</b></td></tr>)}</tbody></table></div></section>
              )}

              {tab === "mapeo" && (
                <section className="report-section"><SectionHeading eyebrow="Clasificación auditable" title="Mapeo de cuentas" description="La base proviene de la hoja Setup de la macro. Las sugerencias PUC pueden confirmarse o cambiarse." action={<input aria-label="Buscar cuenta" placeholder="Buscar código o cuenta…" value={mappingSearch} onChange={(event) => setMappingSearch(event.target.value)} />} /><div className="table-scroll"><table><thead><tr><th>Cuenta</th><th>Nombre</th><th>Origen</th><th>Grupo de macro</th><th>Línea EEFF</th><th>Impacto</th></tr></thead><tbody>{filteredAccounts.map((account) => <tr key={account.base} className={account.mapping.needsReview ? "needs-review" : ""}><td><b>{account.base}</b></td><td>{account.name}</td><td><span className={`source-tag ${account.mapping.source.toLowerCase()}`}>{account.mapping.source}</span></td><td>{account.mapping.subgroup}</td><td><select value={overrides[account.base] || account.mapping.line} onChange={(event) => setOverrides((current) => ({ ...current, [account.base]: event.target.value }))}><option value="Cuenta por revisar">Cuenta por revisar</option>{MAPPING_OPTIONS.map((line) => <option key={line} value={line}>{line}</option>)}</select></td><td>{formatCompact(account.amount)}</td></tr>)}</tbody></table></div></section>
              )}

              {tab === "controles" && (
                <section className="report-section"><SectionHeading eyebrow="Tolerancia · $1" title="Controles y conciliaciones" description="Cada diferencia se mantiene visible; el sistema no crea partidas automáticas para cuadrar." /><div className="table-scroll"><table><thead><tr><th>Periodo</th><th>Filas</th><th>Débitos − créditos</th><th>Control de movimiento</th><th>Sin mapear</th><th>Por confirmar</th><th>Diferencia ESF</th><th>Estado</th></tr></thead><tbody>{analysis.periods.map((period) => <tr key={period.date}><td>{period.label}</td><td>{period.rows.toLocaleString("es-CO")}</td><td>{formatMoney(period.check.debitCredit)}</td><td>{formatMoney(period.check.movement)}</td><td>{period.check.unmapped}</td><td>{period.check.review}</td><td className={Math.abs(period.check.balance) >= 1 ? "negative" : ""}>{formatMoney(period.check.balance)}</td><td><span className={`status-tag ${period.check.status === "OK" ? "ok" : "review"}`}>{period.check.status}</span></td></tr>)}</tbody></table></div><div className="check-notes"><p><b>Débitos − créditos:</b> valida el equilibrio del balance exportado.</p><p><b>Control de movimiento:</b> saldo inicial + débito − crédito = saldo final.</p><p><b>Diferencia ESF:</b> activos − pasivos − patrimonio, sin partidas de ajuste.</p><p><b>Por confirmar:</b> cuentas con una sugerencia material que requiere criterio contable.</p></div></section>
              )}
            </div>
          </div>
        </section>
      )}

      <footer><span className="wordmark">numbi</span><p>Decisiones inteligentes en tiempo real.</p><span>Procesamiento local · Los archivos no se almacenan</span></footer>
    </main>
  );
}

function EmptyState({ icon, title, text }: { icon: string; title: string; text: string }) {
  return <div className="empty-state"><img src={icon} alt="" /><strong>{title}</strong><p>{text}</p></div>;
}
