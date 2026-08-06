"use client";

import { ChangeEvent, DragEvent, useMemo, useRef, useState } from "react";
import {
  calculate,
  csvDownload,
  ER_LINES,
  ESF_LINES,
  MAPPING_OPTIONS,
  parseWorkbook,
  type ParsedFile,
  type PeriodResult,
} from "./lib/financials";

type Tab = "resumen" | "resultados" | "situacion" | "terceros" | "mapeo" | "checks";

const tabs: { id: Tab; label: string; short: string }[] = [
  { id: "resumen", label: "Resumen ejecutivo", short: "Resumen" },
  { id: "resultados", label: "Estado de resultados", short: "Resultados" },
  { id: "situacion", label: "Situación financiera", short: "Situación" },
  { id: "terceros", label: "Saldos por tercero", short: "Terceros" },
  { id: "mapeo", label: "Mapeo de cuentas", short: "Mapeo" },
  { id: "checks", label: "Controles", short: "Controles" },
];

const money = new Intl.NumberFormat("es-CO", { maximumFractionDigits: 0 });
const compact = new Intl.NumberFormat("es-CO", { notation: "compact", maximumFractionDigits: 1 });
const formatMoney = (value = 0) => `${value < 0 ? "(" : ""}$ ${money.format(Math.abs(value))}${value < 0 ? ")" : ""}`;

const resultRows = [
  { label: "INGRESOS", section: true },
  { label: "Ingresos operacionales" },
  { label: "Devoluciones y descuentos" },
  { label: "Ingresos netos", total: true },
  { label: "Costo de ventas y servicios" },
  { label: "Utilidad bruta", total: true },
  { label: "Gastos de administración" },
  { label: "Gastos de ventas" },
  { label: "Total gastos operacionales", total: true },
  { label: "Resultado operacional", total: true },
  { label: "Otros ingresos" },
  { label: "Otros gastos" },
  { label: "Resultado antes de impuestos", total: true },
  { label: "Impuesto a las ganancias" },
  { label: "Resultado neto", total: true, highlight: true },
];

const balanceRows = [
  { label: "ACTIVOS", section: true },
  ...ESF_LINES.slice(0, 7).map((label) => ({ label })),
  { label: "Total activos", total: true },
  { label: "PASIVOS", section: true },
  ...ESF_LINES.slice(7, 15).map((label) => ({ label })),
  { label: "Total pasivos", total: true },
  { label: "PATRIMONIO", section: true },
  ...ESF_LINES.slice(15).map((label) => ({ label })),
  { label: "Resultado del periodo" },
  { label: "Total patrimonio", total: true },
  { label: "Total pasivo y patrimonio", total: true, highlight: true },
];

function FinancialTable({ periods, type }: { periods: PeriodResult[]; type: "er" | "esf" }) {
  const rows = type === "er" ? resultRows : balanceRows;
  const field = type === "er" ? "er" : "esf";
  return (
    <div className="table-scroll">
      <table className="financial-table">
        <thead>
          <tr>
            <th>Concepto</th>
            {periods.map((period) => <th key={period.date}>{period.label}</th>)}
          </tr>
        </thead>
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
  const inputRef = useRef<HTMLInputElement>(null);
  const records = useMemo(() => uploaded.flatMap((file) => file.records), [uploaded]);
  const analysis = useMemo(() => calculate(records, overrides), [records, overrides]);
  const latest = analysis.latest;
  const company = uploaded[0]?.company || "Tu empresa";
  const nit = uploaded[0]?.nit || "";
  const allOk = analysis.periods.length > 0 && analysis.periods.every((period) => period.check.status === "OK");
  const ytdPeriods = latest ? analysis.periods.filter((period) => period.date.slice(0, 4) === latest.date.slice(0, 4) && !period.annual) : [];
  const ytdRevenue = ytdPeriods.reduce((sum, period) => sum + (period.er["Ingresos netos"] || 0), 0);
  const ytdResult = ytdPeriods.reduce((sum, period) => sum + (period.er["Resultado neto"] || 0), 0);
  const maxTrend = Math.max(1, ...ytdPeriods.flatMap((period) => [Math.abs(period.er["Ingresos netos"] || 0), Math.abs(period.er["Resultado neto"] || 0)]));
  const filteredParties = analysis.thirdParties.filter((party) => `${party.identification} ${party.name}`.toLowerCase().includes(partySearch.toLowerCase()));

  async function loadFiles(fileList: FileList | File[]) {
    const candidates = Array.from(fileList).filter((file) => file.name.toLowerCase().endsWith(".xlsx"));
    if (!candidates.length) {
      setMessage("Selecciona archivos de Excel con extensión .xlsx.");
      return;
    }
    setBusy(true);
    setMessage("");
    const settled = await Promise.allSettled(candidates.map(parseWorkbook));
    const successes = settled.filter((item): item is PromiseFulfilledResult<ParsedFile> => item.status === "fulfilled").map((item) => item.value);
    const failures = settled.filter((item): item is PromiseRejectedResult => item.status === "rejected");
    setUploaded((current) => {
      const next = new Map(current.map((file) => [file.name, file]));
      successes.forEach((file) => next.set(file.name, file));
      return [...next.values()].sort((a, b) => a.periodEnd.localeCompare(b.periodEnd));
    });
    if (failures.length) setMessage(`${successes.length} archivo(s) procesados. ${failures.length} no se pudieron leer: ${failures.map((item) => item.reason?.message || "error desconocido").join(" · ")}`);
    else setMessage(`${successes.length} archivo(s) procesados correctamente.`);
    setBusy(false);
    setTab("resumen");
  }

  function onDrop(event: DragEvent<HTMLDivElement>) {
    event.preventDefault();
    setDragging(false);
    loadFiles(event.dataTransfer.files);
  }

  function downloadStatement(type: "er" | "esf") {
    const rows = type === "er" ? resultRows.filter((row) => !row.section) : balanceRows.filter((row) => !row.section);
    const field = type === "er" ? "er" : "esf";
    csvDownload(type === "er" ? "estado-de-resultados.csv" : "situacion-financiera.csv", [
      ["Concepto", ...analysis.periods.map((period) => period.label)],
      ...rows.map((row) => [row.label, ...analysis.periods.map((period) => period[field][row.label] || 0)]),
    ]);
  }

  return (
    <main>
      <header className="topbar">
        <div className="brand-lockup"><span className="wordmark">numbi</span><span className="product-name">EEFF</span></div>
        <div className="privacy-pill"><span className="privacy-dot" /> Procesamiento privado en tu navegador</div>
      </header>

      <section className="hero">
        <div className="hero-copy">
          <span className="eyebrow">Estados financieros sin reprocesos</span>
          <h1>Tus números,<br /><span>siempre al día.</span></h1>
          <p>Sube tus balances de prueba por tercero. Numbi los valida, clasifica y convierte en reportes financieros listos para analizar.</p>
          <div
            className={`dropzone ${dragging ? "is-dragging" : ""}`}
            onDragOver={(event) => { event.preventDefault(); setDragging(true); }}
            onDragLeave={() => setDragging(false)}
            onDrop={onDrop}
          >
            <img src="/icon-upload.png" alt="" />
            <div>
              <strong>{busy ? "Procesando balances…" : "Arrastra aquí tus archivos .xlsx"}</strong>
              <span>Puedes cargar uno o varios periodos a la vez</span>
            </div>
            <button onClick={() => inputRef.current?.click()} disabled={busy}>{busy ? "Leyendo…" : "Seleccionar archivos"}</button>
            <input ref={inputRef} type="file" accept=".xlsx" multiple hidden onChange={(event: ChangeEvent<HTMLInputElement>) => event.target.files && loadFiles(event.target.files)} />
          </div>
          {message && <p className="upload-message">{message}</p>}
          <div className="trust-row"><span>✓ Sin guardar archivos</span><span>✓ Solo filas Auxiliar / Sí</span><span>✓ Control contable automático</span></div>
        </div>
        <div className="hero-panel" aria-label="Proceso de la herramienta">
          <div className="panel-head"><span>Flujo inteligente</span><span className="live-dot">local</span></div>
          <div className="flow-step"><img src="/icon-upload.png" alt="" /><div><small>01</small><strong>Carga</strong><span>Balance por tercero</span></div></div>
          <div className="flow-line" />
          <div className="flow-step"><img src="/icon-calculator.png" alt="" /><div><small>02</small><strong>Valida</strong><span>Movimientos y mapeo</span></div></div>
          <div className="flow-line" />
          <div className="flow-step"><img src="/icon-growth.png" alt="" /><div><small>03</small><strong>Analiza</strong><span>Estados y terceros</span></div></div>
          <img className="robot" src="/numbi-robot.png" alt="Mascota Numbi" />
        </div>
      </section>

      {!uploaded.length ? (
        <section className="empty-preview">
          <div><span>01</span><strong>Detectamos el periodo</strong><p>Leemos automáticamente empresa, NIT y fecha de corte.</p></div>
          <div><span>02</span><strong>Evitamos doble conteo</strong><p>Procesamos únicamente el nivel Auxiliar marcado como transaccional.</p></div>
          <div><span>03</span><strong>Conciliamos el resultado</strong><p>Revisamos débitos, créditos, saldos y ecuación financiera.</p></div>
        </section>
      ) : (
        <section className="workspace">
          <div className="workspace-header">
            <div><span className="eyebrow">Panel financiero</span><h2>{company}</h2><p>{nit ? `NIT ${nit} · ` : ""}{records.length.toLocaleString("es-CO")} registros auxiliares · {uploaded.length} archivo(s)</p></div>
            <div className={`model-status ${allOk ? "ok" : "review"}`}><span>{allOk ? "✓" : "!"}</span><div><small>Estado del modelo</small><strong>{allOk ? "Todo conciliado" : "Requiere revisión"}</strong></div></div>
          </div>

          <div className="file-chips">
            {uploaded.map((file) => <button key={file.name} title="Quitar archivo" onClick={() => setUploaded((current) => current.filter((item) => item.name !== file.name))}><span>{file.periodLabel}</span>{file.name}<b>×</b></button>)}
            <button className="add-file" onClick={() => inputRef.current?.click()}>＋ Agregar periodo</button>
          </div>

          <div className="dashboard-grid">
            <nav className="side-nav" aria-label="Secciones del reporte">
              {tabs.map((item, index) => <button className={tab === item.id ? "active" : ""} onClick={() => setTab(item.id)} key={item.id}><span>{String(index + 1).padStart(2, "0")}</span>{item.label}</button>)}
            </nav>
            <select className="mobile-nav" value={tab} onChange={(event) => setTab(event.target.value as Tab)}>{tabs.map((item) => <option value={item.id} key={item.id}>{item.label}</option>)}</select>

            <div className="dashboard-content">
              {tab === "resumen" && latest && (
                <>
                  <div className="section-heading"><div><span className="eyebrow">Último corte · {latest.label}</span><h3>Resumen ejecutivo</h3></div></div>
                  <div className="kpi-grid">
                    <article><span>Activos</span><strong>{formatMoney(latest.esf["Total activos"])}</strong><small>Saldo al corte</small></article>
                    <article><span>Pasivos</span><strong>{formatMoney(latest.esf["Total pasivos"])}</strong><small>{latest.esf["Total activos"] ? `${((latest.esf["Total pasivos"] / latest.esf["Total activos"]) * 100).toFixed(1)}% de los activos` : "-"}</small></article>
                    <article><span>Ingresos acumulados</span><strong>{formatMoney(ytdRevenue)}</strong><small>Año {latest.date.slice(0, 4)}</small></article>
                    <article className={ytdResult < 0 ? "alert" : "positive"}><span>Resultado neto</span><strong>{formatMoney(ytdResult)}</strong><small>Acumulado del año</small></article>
                  </div>
                  <div className="insight-grid">
                    <article className="trend-card">
                      <div className="card-title"><div><span>Tendencia mensual</span><strong>Ingresos vs. resultado neto</strong></div><div className="legend"><i />Ingresos <i />Resultado</div></div>
                      <div className="bar-chart">
                        {ytdPeriods.map((period) => {
                          const revenue = period.er["Ingresos netos"] || 0;
                          const result = period.er["Resultado neto"] || 0;
                          return <div className="bar-group" key={period.date}><div className="bars"><i style={{ height: `${Math.max(3, Math.abs(revenue) / maxTrend * 100)}%` }} title={formatMoney(revenue)} /><i className={result < 0 ? "down" : "up"} style={{ height: `${Math.max(3, Math.abs(result) / maxTrend * 100)}%` }} title={formatMoney(result)} /></div><span>{period.label.split(" ")[0].slice(0, 3)}</span></div>;
                        })}
                      </div>
                    </article>
                    <article className="composition-card">
                      <div className="card-title"><div><span>Composición</span><strong>Activos al último corte</strong></div></div>
                      {["Efectivo y equivalentes", "Cuentas por cobrar", "Inventarios", "Intangibles", "Otros activos"].map((line) => {
                        const value = latest.esf[line] || 0;
                        const share = latest.esf["Total activos"] ? Math.max(0, value / latest.esf["Total activos"] * 100) : 0;
                        return <div className="composition-row" key={line}><div><span>{line}</span><b>{share.toFixed(1)}%</b></div><div className="progress"><i style={{ width: `${Math.min(100, share)}%` }} /></div></div>;
                      })}
                    </article>
                  </div>
                </>
              )}

              {tab === "resultados" && (
                <section className="report-section"><div className="section-heading"><div><span className="eyebrow">COP · pesos colombianos</span><h3>Estado de resultados</h3></div><button onClick={() => downloadStatement("er")}>Descargar CSV</button></div><FinancialTable periods={analysis.periods} type="er" /></section>
              )}
              {tab === "situacion" && (
                <section className="report-section"><div className="section-heading"><div><span className="eyebrow">COP · pesos colombianos</span><h3>Estado de situación financiera</h3></div><button onClick={() => downloadStatement("esf")}>Descargar CSV</button></div><FinancialTable periods={analysis.periods} type="esf" /></section>
              )}
              {tab === "terceros" && (
                <section className="report-section"><div className="section-heading"><div><span className="eyebrow">Último corte · {latest?.label}</span><h3>Saldos por tercero</h3></div><input aria-label="Buscar tercero" placeholder="Buscar NIT o tercero…" value={partySearch} onChange={(event) => setPartySearch(event.target.value)} /></div><div className="table-scroll"><table><thead><tr><th>Identificación</th><th>Tercero</th><th>Cuentas por cobrar</th><th>Proveedores y CxP</th><th>Impuestos</th><th>Otros pasivos</th><th>Saldo neto</th></tr></thead><tbody>{filteredParties.slice(0, 100).map((party) => <tr key={party.identification}><td>{party.identification}</td><td>{party.name}</td><td>{formatMoney(party.receivable)}</td><td>{formatMoney(party.payable)}</td><td>{formatMoney(party.taxes)}</td><td>{formatMoney(party.other)}</td><td className={party.net < 0 ? "negative" : ""}>{formatMoney(party.net)}</td></tr>)}</tbody></table></div></section>
              )}
              {tab === "mapeo" && (
                <section className="report-section"><div className="section-heading"><div><span className="eyebrow">Clasificación auditable</span><h3>Mapeo de cuentas</h3><p>Las sugerencias siguen el PUC colombiano. Puedes cambiar cualquier línea y los reportes se recalculan.</p></div></div><div className="table-scroll"><table><thead><tr><th>Cuenta base</th><th>Nombre de referencia</th><th>Estado</th><th>Grupo</th><th>Línea EEFF</th><th>Registros</th></tr></thead><tbody>{analysis.accounts.map((account) => <tr key={account.base} className={account.mapping.statement === "REVISAR" ? "needs-review" : ""}><td><b>{account.base}</b></td><td>{account.name}</td><td>{account.mapping.statement}</td><td>{account.mapping.group}</td><td><select value={overrides[account.base] || account.mapping.line} onChange={(event) => setOverrides((current) => ({ ...current, [account.base]: event.target.value }))}><option value="Cuenta por revisar">Cuenta por revisar</option>{MAPPING_OPTIONS.map((line) => <option key={line} value={line}>{line}</option>)}</select></td><td>{account.records.toLocaleString("es-CO")}</td></tr>)}</tbody></table></div></section>
              )}
              {tab === "checks" && (
                <section className="report-section"><div className="section-heading"><div><span className="eyebrow">Tolerancia de redondeo · $1</span><h3>Controles y conciliaciones</h3></div></div><div className="table-scroll"><table><thead><tr><th>Periodo</th><th>Filas</th><th>Débitos - créditos</th><th>Control movimiento</th><th>Sin mapear</th><th>Diferencia ESF</th><th>Estado</th></tr></thead><tbody>{analysis.periods.map((period) => <tr key={period.date}><td>{period.label}</td><td>{period.rows.toLocaleString("es-CO")}</td><td>{formatMoney(period.check.debitCredit)}</td><td>{formatMoney(period.check.movement)}</td><td>{period.check.unmapped}</td><td>{formatMoney(period.check.balance)}</td><td><span className={`status-tag ${period.check.status === "OK" ? "ok" : "review"}`}>{period.check.status}</span></td></tr>)}</tbody></table></div><div className="check-notes"><p><b>Débitos - créditos:</b> valida que el archivo esté equilibrado.</p><p><b>Control movimiento:</b> Saldo inicial + Débito - Crédito = Saldo final.</p><p><b>Diferencia ESF:</b> Activos = Pasivos + Patrimonio, incorporando el resultado del periodo.</p></div></section>
              )}
            </div>
          </div>
        </section>
      )}

      <footer><span className="wordmark">numbi</span><p>Decisiones inteligentes en tiempo real.</p><span>Los datos se procesan localmente y no salen de tu navegador.</span></footer>
    </main>
  );
}
