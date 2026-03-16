import * as XLSX from "xlsx";

// === CONFIGURACIÓN DE COLUMNAS (0 = A, 1 = B, ...) ===
const CONFIG = {
  COL_FECHA: 1,
  COL_HG1_LANEC_INI: 2,
  COL_HG1_LANEC_FIN: 3,
  COL_HG2_LANEC_INI: 8,
  COL_HG2_LANEC_FIN: 9,
  COL_AUX_KWH: 24,
  COL_LANEC_PARCIAL_KWH: 25,
  COL_GRACA_PARCIAL_KWH: 30,
  COL_TOTAL_LG_KWH: 35,
  COL_GEN1_KWH: 36,
  COL_GEN2_KWH: 37,
  COL_HFO_GAL: 43,
  COL_DO_GAL: 46,
  COL_STOCK_HFO_TOTAL: 49,
  COL_STOCK_DO_TOTAL: 51,
};

const HORO_BASE_U1 = 0;
const HORO_BASE_U2 = 21041;
const OBJ_MTO_HORAS_U1 = 8000;
const OBJ_MTO_HORAS_U2 = 6000;

const COSTOS_VARIABLES: Record<string, number> = {
  combustible_transporte: 0.1153,
  lubricantes_quimicos:  0.0182,
  agua_insumos:          0.0070,
  repuestos_predictivo:  0.0299,
  impacto_ambiental:     0.0029,
  servicios_auxiliares:  0.0034,
  margen_variable:       0.0138,
};

const COSTO_VARIABLE_TOTAL = Object.values(COSTOS_VARIABLES).reduce((a, b) => a + b, 0);
const COSTO_FIJO_MENSUAL_POR_UNIDAD = 30720;

const P_INST_TOTAL = 5100;
const P_INST_EFECTIVA = 0.85 * P_INST_TOTAL;
const P_CONTR_LANEC = 3800;
const P_CONTR_GRACA = 1000;
const P_CONTR_TOT = P_CONTR_LANEC + P_CONTR_GRACA;

// ========= UTILIDADES =========

function num(v: unknown): number {
  if (v === null || v === undefined) return 0;
  if (typeof v === "number") return v;
  if (typeof v === "string") {
    const s = v.replace(/\./g, "").replace(",", ".");
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  }
  return 0;
}

function posNum(v: unknown): number {
  const n = num(v);
  return n < 0 ? 0 : n;
}

function fmt(v: number | unknown, dec = 2): string {
  return Number(v).toLocaleString("es-EC", {
    minimumFractionDigits: dec,
    maximumFractionDigits: dec,
  });
}

function pad2(n: number): string {
  return n < 10 ? "0" + n : "" + n;
}

function jsDateKey(d: Date): string {
  return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function excelDateKey(v: any): string | null {
  if (v == null) return null;
  if (typeof v === "number") {
    const dc = XLSX.SSF.parse_date_code(v);
    if (!dc) return null;
    return `${dc.y}-${pad2(dc.m)}-${pad2(dc.d)}`;
  }
  if (v instanceof Date) return jsDateKey(v);
  if (typeof v === "string") {
    const s = v.trim();
    const m = s.match(/^(\d{1,2})[/\-](\d{1,2})[/\-](\d{2,4})$/);
    if (m) {
      const d = parseInt(m[1], 10), mo = parseInt(m[2], 10);
      let y = parseInt(m[3], 10);
      if (y < 100) y += 2000;
      return `${y}-${pad2(mo)}-${pad2(d)}`;
    }
    const d2 = new Date(s);
    if (!isNaN(d2.getTime())) return jsDateKey(d2);
  }
  return null;
}

function parseFechaRobusta(v: unknown): Date | null {
  if (v === null || v === undefined || v === "") return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if (typeof v === "number" && isFinite(v)) {
    // Usar XLSX.SSF.parse_date_code para evitar desfase de zona horaria.
    // La fórmula (v-25569)*86400*1000 crea un timestamp UTC que JavaScript
    // convierte a hora local, desplazando el día -1 en zonas UTC-N (p. ej. Ecuador UTC-5).
    const dc = XLSX.SSF.parse_date_code(v);
    if (!dc) return null;
    const dt = new Date(dc.y, dc.m - 1, dc.d, 0, 0, 0, 0);
    return isNaN(dt.getTime()) ? null : dt;
  }
  if (typeof v === "string") {
    const s = v.trim();
    const iso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (iso) {
      const dt = new Date(+iso[1], +iso[2] - 1, +iso[3]);
      return isNaN(dt.getTime()) ? null : dt;
    }
    const m = s.match(/^(\d{1,2})[/\-](\d{1,2})[/\-](\d{2}|\d{4})$/);
    if (m) {
      let a = +m[1], b = +m[2], y = +m[3];
      if (y < 100) y += 2000;
      let day, mon;
      if (a > 12 && b <= 12) { day = a; mon = b; }
      else if (b > 12 && a <= 12) { day = b; mon = a; }
      else { day = a; mon = b; }
      const dt = new Date(y, mon - 1, day);
      return isNaN(dt.getTime()) ? null : dt;
    }
    const dt2 = new Date(s);
    return isNaN(dt2.getTime()) ? null : dt2;
  }
  return null;
}

function getMesNombreES(monthIndex: number): string {
  const meses = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
  return meses[monthIndex] || "";
}

function getSheetNameFromDate(fechaJS: Date): string {
  return `${getMesNombreES(fechaJS.getMonth())} ${fechaJS.getFullYear()}`;
}

function getDaysInMonth(year: number, monthIndex: number): number {
  return new Date(year, monthIndex + 1, 0).getDate();
}

function formatFechaLarga(str: string): string {
  const d = new Date(str + "T00:00:00");
  if (isNaN(d.getTime())) return "";
  const meses = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];
  return `${d.getDate()} de ${meses[d.getMonth()]} de ${d.getFullYear()}`;
}

function monthStrFromDate(d: Date): string {
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}`;
}

function prevMonthStr(yyyyMM: string): string | null {
  const [yS, mS] = (yyyyMM || "").split("-");
  let y = parseInt(yS, 10), m = parseInt(mS, 10);
  if (!y || !m) return null;
  m -= 1;
  if (m === 0) { m = 12; y -= 1; }
  return `${y}-${pad2(m)}`;
}

// ========= SEMÁFORO =========

function semaforo(v: number): string {
  if (!Number.isFinite(v)) return "SIN DATO";
  if (v >= 0.85) return "VERDE";
  if (v >= 0.75) return "AMARILLO";
  if (v >= 0.65) return "NARANJA";
  return "ROJO";
}

function badgeEstado(estado: string): string {
  const cls: Record<string, string> = {
    VERDE: "badge badge-verde",
    AMARILLO: "badge badge-amarillo",
    NARANJA: "badge badge-naranja",
    ROJO: "badge badge-rojo",
    "SIN DATO": "badge badge-nd",
    NORMAL: "badge badge-verde",
    "ALERTA OPERATIVA": "badge badge-naranja",
    "CRÍTICO": "badge badge-rojo",
  };
  return `<span class="${cls[estado] ?? "badge badge-nd"}">${estado}</span>`;
}

// ========= HELPERS HOJA =========

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function getProdSheetAndRows(wbProd: XLSX.WorkBook, fechaJS: Date): { rows: any[][] } {
  const name = getSheetNameFromDate(fechaJS);
  const ws = wbProd.Sheets[name];
  if (ws) {
    return { rows: XLSX.utils.sheet_to_json(ws, { header: 1, raw: true }) as unknown[][] as any[][] };
  }
  const keyTarget = jsDateKey(fechaJS);
  for (const n of wbProd.SheetNames) {
    const wsTry = wbProd.Sheets[n];
    const rows = XLSX.utils.sheet_to_json(wsTry, { header: 1, raw: true }) as any[][];
    for (let i = 0; i < rows.length; i++) {
      if (excelDateKey(rows[i][CONFIG.COL_FECHA]) === keyTarget) {
        return { rows };
      }
    }
  }
  const ws0 = wbProd.Sheets[wbProd.SheetNames[0]];
  return { rows: XLSX.utils.sheet_to_json(ws0, { header: 1, raw: true }) as any[][] };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function findRowByDate(rows: any[][], fechaJS: Date): any[] | null {
  const k = jsDateKey(fechaJS);
  for (let i = 0; i < rows.length; i++) {
    if (excelDateKey(rows[i][CONFIG.COL_FECHA]) === k) return rows[i];
  }
  return null;
}

// ========= ENCABEZADO =========

function rptHeader(tipo: string, subtitulo: string): string {
  // El bloque <style> se inyecta dentro del DOM del informe para garantizar que
  // html2canvas aplique vertical-align:middle en los estilos computados de cada
  // celda, independientemente del soporte de hojas externas durante la captura.
  const inlineStyles = `<style>
.data-table td,.data-table th{vertical-align:middle!important}
</style>`;
  return `${inlineStyles}<div class="rpt-header">
  <div class="rpt-header-body">
    <div class="rpt-header-left">
      <div class="rpt-logo-circle">
        <svg viewBox="0 0 32 32" fill="none" xmlns="http://www.w3.org/2000/svg" width="28" height="28">
          <circle cx="16" cy="16" r="14" fill="rgba(255,255,255,0.15)" stroke="rgba(255,255,255,0.5)" stroke-width="1.5"/>
          <path d="M16 7v4M16 21v4M7 16h4M21 16h4M10.1 10.1l2.8 2.8M19.1 19.1l2.8 2.8M19.1 10.1l-2.8 2.8M10.1 19.1l2.8 2.8" stroke="rgba(255,255,255,0.9)" stroke-width="1.5" stroke-linecap="round"/>
          <circle cx="16" cy="16" r="3.5" fill="rgba(255,255,255,0.9)"/>
        </svg>
      </div>
      <div>
        <div class="rpt-empresa">Central El Morro &mdash; Morro Energy S.A.</div>
        <div class="rpt-tipo">${tipo}</div>
      </div>
    </div>
    <div class="rpt-header-right">
      <div class="rpt-subtitulo-label">Período</div>
      <div class="rpt-subtitulo">${subtitulo}</div>
    </div>
  </div>
  <div class="rpt-header-stripe"></div>
</div>`;
}

function seccion(n: string | number, titulo: string): string {
  return `<div class="rpt-section-title"><span class="rpt-section-num">${n}</span><span class="rpt-section-label">${titulo}</span></div>`;
}

// ========= ANÁLISIS EJECUTIVO DE COMBUSTIBLE =========

interface FuelMetric {
  date: Date;
  kWh: number;
  gen1: number;
  gen2: number;
  hfo: number;
  dsl: number;
  fuel: number;
  // por unidad (prorrateo por energía)
  hfo_u1: number;  dsl_u1: number;  fuel_u1: number;
  hfo_u2: number;  dsl_u2: number;  fuel_u2: number;
  pctDO: number;
  gal_h: number;
  gal_h_u1: number;      gal_h_hfo_u1: number;  gal_h_do_u1: number;
  gal_h_u2: number;      gal_h_hfo_u2: number;  gal_h_do_u2: number;
  horasOp: number;
  h1: number;
  h2: number;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function buildFuelMetricFromRow(row: any[]): FuelMetric | null {
  const d = parseFechaRobusta(row[CONFIG.COL_FECHA]);
  if (!d) return null;
  const aux = posNum(row[CONFIG.COL_AUX_KWH]);
  const lan = posNum(row[CONFIG.COL_LANEC_PARCIAL_KWH]);
  const gra = posNum(row[CONFIG.COL_GRACA_PARCIAL_KWH]);
  const gen1 = posNum(row[CONFIG.COL_GEN1_KWH]);
  const gen2 = posNum(row[CONFIG.COL_GEN2_KWH]);
  const kWh = aux + lan + gra;
  const hfo = posNum(row[CONFIG.COL_HFO_GAL]);
  const dsl = posNum(row[CONFIG.COL_DO_GAL]);
  const fuel = hfo + dsl;
  const h1 = Math.max(0, posNum(row[CONFIG.COL_HG1_LANEC_FIN]) - posNum(row[CONFIG.COL_HG1_LANEC_INI]));
  const h2 = Math.max(0, posNum(row[CONFIG.COL_HG2_LANEC_FIN]) - posNum(row[CONFIG.COL_HG2_LANEC_INI]));
  const horasOp = Math.max(h1, h2);
  if (!(kWh > 0 && fuel > 0 && horasOp > 0)) return null;

  // Prorrateo del combustible total por energía generada por cada unidad
  const genTot = gen1 + gen2;
  const w1 = genTot > 0 ? gen1 / genTot : (h1 > 0 ? h1 / (h1 + h2 || 1) : 0);
  const w2 = genTot > 0 ? gen2 / genTot : (h2 > 0 ? h2 / (h1 + h2 || 1) : 0);
  // Prorratear HFO y diésel por separado con el mismo peso
  const hfo_u1 = hfo * w1;  const dsl_u1 = dsl * w1;  const fuel_u1 = fuel * w1;
  const hfo_u2 = hfo * w2;  const dsl_u2 = dsl * w2;  const fuel_u2 = fuel * w2;

  return {
    date: new Date(d.getFullYear(), d.getMonth(), d.getDate()),
    kWh, gen1, gen2, hfo, dsl, fuel,
    hfo_u1, dsl_u1, fuel_u1,
    hfo_u2, dsl_u2, fuel_u2,
    pctDO: dsl / fuel,
    gal_h: fuel / horasOp,
    gal_h_u1:     h1 > 0 ? fuel_u1 / h1 : NaN,
    gal_h_hfo_u1: h1 > 0 ? hfo_u1  / h1 : NaN,
    gal_h_do_u1:  h1 > 0 ? dsl_u1  / h1 : NaN,
    gal_h_u2:     h2 > 0 ? fuel_u2 / h2 : NaN,
    gal_h_hfo_u2: h2 > 0 ? hfo_u2  / h2 : NaN,
    gal_h_do_u2:  h2 > 0 ? dsl_u2  / h2 : NaN,
    horasOp, h1, h2,
  };
}

function meanSafe(arr: number[]): number {
  const a = arr.filter(v => Number.isFinite(v));
  if (!a.length) return NaN;
  return a.reduce((s, x) => s + x, 0) / a.length;
}

function getAllFuelMetrics(wbProd: XLSX.WorkBook): FuelMetric[] {
  const out: FuelMetric[] = [];
  for (const name of wbProd.SheetNames) {
    const ws = wbProd.Sheets[name];
    if (!ws) continue;
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true }) as any[][];
    for (const row of rows) {
      const m = buildFuelMetricFromRow(row);
      if (m) out.push(m);
    }
  }
  out.sort((a, b) => a.date.getTime() - b.date.getTime());
  return out;
}

function lastNDaysWithData(metrics: FuelMetric[], endDate: Date, n: number): FuelMetric[] {
  const end = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
  const out: FuelMetric[] = [];
  for (let i = metrics.length - 1; i >= 0 && out.length < n; i--) {
    if (metrics[i].date <= end) out.push(metrics[i]);
  }
  return out.reverse();
}

function buildFuelExecutiveHTML(wbProd: XLSX.WorkBook, fechaJS: Date, mode = "daily"): string {
  try {
    const metrics = getAllFuelMetrics(wbProd);
    const win90 = lastNDaysWithData(metrics, fechaJS, 90);
    if (!win90 || win90.length < 20) {
      return `<div class="rpt-notice">Análisis Ejecutivo de Combustible: Información insuficiente para referencia 90D (días válidos: ${(win90||[]).length}).</div>`;
    }
    const galh_ref         = meanSafe(win90.map(d => d.gal_h));
    const win90u1          = win90.filter(d => d.h1 > 0);
    const win90u2          = win90.filter(d => d.h2 > 0);
    const galh_ref_u1      = meanSafe(win90u1.map(d => d.gal_h_u1));
    const galh_hfo_ref_u1  = meanSafe(win90u1.map(d => d.gal_h_hfo_u1));
    const galh_do_ref_u1   = meanSafe(win90u1.map(d => d.gal_h_do_u1));
    const galh_ref_u2      = meanSafe(win90u2.map(d => d.gal_h_u2));
    const galh_hfo_ref_u2  = meanSafe(win90u2.map(d => d.gal_h_hfo_u2));
    const galh_do_ref_u2   = meanSafe(win90u2.map(d => d.gal_h_do_u2));
    const pctDO_ref        = meanSafe(win90.map(d => d.pctDO));
    const fmtPct = (x: number) => Number.isFinite(x) ? (x * 100).toFixed(1) + "%" : "—";
    const fmt1   = (x: number) => Number.isFinite(x) ? x.toFixed(1) : "—";
    const fmt0   = (x: number) => Number.isFinite(x) ? x.toFixed(0) : "—";

    // Helpers de impacto y etiqueta — lenguaje institucional
    const impactoLabel = (gal: number) =>
      !Number.isFinite(gal) ? "—"
      : Math.abs(gal) < 5 ? "En línea con referencia"
      : gal > 0 ? `+${fmt0(gal)} gal (por encima de referencia)`
      : `${fmt0(gal)} gal (por debajo de referencia)`;
    const impactoClass = (gal: number) =>
      !Number.isFinite(gal) ? "" : gal > 0 ? "warn" : "fuel-ahorro";
    const desvLabel = (d: number) =>
      !Number.isFinite(d) ? "—" : (d > 0 ? `+${fmt1(d)}` : fmt1(d));
    const desvClass = (d: number) =>
      !Number.isFinite(d) ? "" : d > 0 ? "warn" : "fuel-ahorro";

    if (String(mode).toLowerCase() !== "monthly") {
      const key = jsDateKey(fechaJS);
      let today: FuelMetric | null = null;
      for (let i = metrics.length - 1; i >= 0; i--) {
        if (jsDateKey(metrics[i].date) === key) { today = metrics[i]; break; }
        if (metrics[i].date <= fechaJS && !today) today = metrics[i];
      }
      if (!today) return `<div class="rpt-notice">Análisis Ejecutivo de Combustible: No se encontró un día válido para análisis.</div>`;

      // Impacto total por unidad en el día (para la fila Total U1/U2)
      const dU1 = today.h1 > 0 && Number.isFinite(today.gal_h_u1) && Number.isFinite(galh_ref_u1) ? today.gal_h_u1 - galh_ref_u1 : NaN;
      const dU2 = today.h2 > 0 && Number.isFinite(today.gal_h_u2) && Number.isFinite(galh_ref_u2) ? today.gal_h_u2 - galh_ref_u2 : NaN;
      const impU1_dia = Number.isFinite(dU1) ? dU1 * today.h1 : NaN;
      const impU2_dia = Number.isFinite(dU2) ? dU2 * today.h2 : NaN;

      // Impacto HFO y DO por separado en el día (U1+U2)
      const dHfoU1 = today.h1 > 0 && Number.isFinite(today.gal_h_hfo_u1) && Number.isFinite(galh_hfo_ref_u1) ? today.gal_h_hfo_u1 - galh_hfo_ref_u1 : NaN;
      const dHfoU2 = today.h2 > 0 && Number.isFinite(today.gal_h_hfo_u2) && Number.isFinite(galh_hfo_ref_u2) ? today.gal_h_hfo_u2 - galh_hfo_ref_u2 : NaN;
      const dDoU1  = today.h1 > 0 && Number.isFinite(today.gal_h_do_u1)  && Number.isFinite(galh_do_ref_u1)  ? today.gal_h_do_u1  - galh_do_ref_u1  : NaN;
      const dDoU2  = today.h2 > 0 && Number.isFinite(today.gal_h_do_u2)  && Number.isFinite(galh_do_ref_u2)  ? today.gal_h_do_u2  - galh_do_ref_u2  : NaN;
      const impHfoU1_dia = Number.isFinite(dHfoU1) ? dHfoU1 * today.h1 : NaN;
      const impHfoU2_dia = Number.isFinite(dHfoU2) ? dHfoU2 * today.h2 : NaN;
      const impDoU1_dia  = Number.isFinite(dDoU1)  ? dDoU1  * today.h1 : NaN;
      const impDoU2_dia  = Number.isFinite(dDoU2)  ? dDoU2  * today.h2 : NaN;
      const impHfoTotal_dia = (Number.isFinite(impHfoU1_dia) ? impHfoU1_dia : 0) + (Number.isFinite(impHfoU2_dia) ? impHfoU2_dia : 0);
      const impDoTotal_dia  = (Number.isFinite(impDoU1_dia)  ? impDoU1_dia  : 0) + (Number.isFinite(impDoU2_dia)  ? impDoU2_dia  : 0);

      // Acumulado mes HFO y DO por separado
      const dm = today.date.getMonth(), dy = today.date.getFullYear();
      const endD = new Date(fechaJS.getFullYear(), fechaJS.getMonth(), fechaJS.getDate());
      let impAcumHfoU1 = 0, impAcumDoU1 = 0, impAcumHfoU2 = 0, impAcumDoU2 = 0;
      for (const d of metrics) {
        if (d.date.getFullYear() === dy && d.date.getMonth() === dm && d.date <= endD) {
          if (d.h1 > 0 && Number.isFinite(d.gal_h_hfo_u1) && Number.isFinite(galh_hfo_ref_u1))
            impAcumHfoU1 += (d.gal_h_hfo_u1 - galh_hfo_ref_u1) * d.h1;
          if (d.h1 > 0 && Number.isFinite(d.gal_h_do_u1) && Number.isFinite(galh_do_ref_u1))
            impAcumDoU1  += (d.gal_h_do_u1  - galh_do_ref_u1)  * d.h1;
          if (d.h2 > 0 && Number.isFinite(d.gal_h_hfo_u2) && Number.isFinite(galh_hfo_ref_u2))
            impAcumHfoU2 += (d.gal_h_hfo_u2 - galh_hfo_ref_u2) * d.h2;
          if (d.h2 > 0 && Number.isFinite(d.gal_h_do_u2) && Number.isFinite(galh_do_ref_u2))
            impAcumDoU2  += (d.gal_h_do_u2  - galh_do_ref_u2)  * d.h2;
        }
      }
      const impAcumHfoTotal = impAcumHfoU1 + impAcumHfoU2;
      const impAcumDoTotal  = impAcumDoU1  + impAcumDoU2;

      // Diagnóstico institucional
      const msgs: string[] = [];
      if (today.h1 > 0 && Number.isFinite(dU1)) {
        if (dU1 > 1.0) msgs.push(`U1 por encima de referencia en consumo (+${fmt1(dU1)} gal/h)`);
        else if (dU1 < -1.0) msgs.push(`U1 por debajo de referencia en consumo (${fmt1(dU1)} gal/h)`);
      }
      if (today.h2 > 0 && Number.isFinite(dU2)) {
        if (dU2 > 1.0) msgs.push(`U2 por encima de referencia en consumo (+${fmt1(dU2)} gal/h)`);
        else if (dU2 < -1.0) msgs.push(`U2 por debajo de referencia en consumo (${fmt1(dU2)} gal/h)`);
      }
      const causa = msgs.length ? msgs.join(". ") + "." : "Operación dentro de los parámetros de referencia 90D.";

      // ── Helpers desviación para HFO/DO ─────────────────────────────────────
      const dHfoU1r = Number.isFinite(today.gal_h_hfo_u1)&&Number.isFinite(galh_hfo_ref_u1)?today.gal_h_hfo_u1-galh_hfo_ref_u1:NaN;
      const dDoU1r  = Number.isFinite(today.gal_h_do_u1) &&Number.isFinite(galh_do_ref_u1) ?today.gal_h_do_u1-galh_do_ref_u1  :NaN;
      const dHfoU2r = Number.isFinite(today.gal_h_hfo_u2)&&Number.isFinite(galh_hfo_ref_u2)?today.gal_h_hfo_u2-galh_hfo_ref_u2:NaN;
      const dDoU2r  = Number.isFinite(today.gal_h_do_u2) &&Number.isFinite(galh_do_ref_u2) ?today.gal_h_do_u2-galh_do_ref_u2  :NaN;

      const rowsU1 = today.h1 > 0 ? `
        <tr class="rpt-row-grupo"><td colspan="6" class="label">Unidad 1 &nbsp;<span class="rpt-muted">(${fmt1(today.h1)} h operadas)</span></td></tr>
        <tr>
          <td class="label" style="padding-left:16px">HFO</td>
          <td class="num">${fmt1(today.gal_h_hfo_u1)}</td>
          <td class="num">${fmt1(galh_hfo_ref_u1)}</td>
          <td class="num ${desvClass(dHfoU1r)}">${desvLabel(dHfoU1r)}</td>
          <td class="num ${impactoClass(impHfoU1_dia)}">${impactoLabel(impHfoU1_dia)}</td>
          <td class="num ${impactoClass(impAcumHfoU1)}">${impactoLabel(impAcumHfoU1)}</td>
        </tr>
        <tr>
          <td class="label" style="padding-left:16px">Diésel (DO)</td>
          <td class="num">${fmt1(today.gal_h_do_u1)}</td>
          <td class="num">${fmt1(galh_do_ref_u1)}</td>
          <td class="num ${desvClass(dDoU1r)}">${desvLabel(dDoU1r)}</td>
          <td class="num ${impactoClass(impDoU1_dia)}">${impactoLabel(impDoU1_dia)}</td>
          <td class="num ${impactoClass(impAcumDoU1)}">${impactoLabel(impAcumDoU1)}</td>
        </tr>`
        : `<tr><td class="label">Unidad 1</td><td colspan="5" style="text-align:left;color:#9ca3af;font-style:italic">No operó este día</td></tr>`;

      const rowsU2 = today.h2 > 0 ? `
        <tr class="rpt-row-grupo"><td colspan="6" class="label">Unidad 2 &nbsp;<span class="rpt-muted">(${fmt1(today.h2)} h operadas)</span></td></tr>
        <tr>
          <td class="label" style="padding-left:16px">HFO</td>
          <td class="num">${fmt1(today.gal_h_hfo_u2)}</td>
          <td class="num">${fmt1(galh_hfo_ref_u2)}</td>
          <td class="num ${desvClass(dHfoU2r)}">${desvLabel(dHfoU2r)}</td>
          <td class="num ${impactoClass(impHfoU2_dia)}">${impactoLabel(impHfoU2_dia)}</td>
          <td class="num ${impactoClass(impAcumHfoU2)}">${impactoLabel(impAcumHfoU2)}</td>
        </tr>
        <tr>
          <td class="label" style="padding-left:16px">Diésel (DO)</td>
          <td class="num">${fmt1(today.gal_h_do_u2)}</td>
          <td class="num">${fmt1(galh_do_ref_u2)}</td>
          <td class="num ${desvClass(dDoU2r)}">${desvLabel(dDoU2r)}</td>
          <td class="num ${impactoClass(impDoU2_dia)}">${impactoLabel(impDoU2_dia)}</td>
          <td class="num ${impactoClass(impAcumDoU2)}">${impactoLabel(impAcumDoU2)}</td>
        </tr>`
        : `<tr><td class="label">Unidad 2</td><td colspan="5" style="text-align:left;color:#9ca3af;font-style:italic">No operó este día</td></tr>`;

      return `<div class="rpt-fuel-box">
  <div class="rpt-fuel-header">
    <span class="rpt-fuel-title">Análisis Ejecutivo de Combustible</span>
  </div>
  <p class="rpt-fuel-causa">${causa}</p>
  <table class="data-table">
    <thead><tr>
      <th>Combustible</th>
      <th>gal/h del día</th>
      <th>Referencia 90D</th>
      <th>Desviación (gal/h)</th>
      <th>Balance del día (gal)</th>
      <th>Acumulado mes (gal)</th>
    </tr></thead>
    <tbody>
      ${rowsU1}
      ${rowsU2}
      <tr class="rpt-row-total">
        <td class="label" colspan="4"><strong>Balance total HFO del día (U1+U2)</strong></td>
        <td class="num ${impactoClass(impHfoTotal_dia)}" colspan="2"><strong>${impactoLabel(impHfoTotal_dia)}</strong></td>
      </tr>
      <tr class="rpt-row-total">
        <td class="label" colspan="4"><strong>Balance total DO del día (U1+U2)</strong></td>
        <td class="num ${impactoClass(impDoTotal_dia)}" colspan="2"><strong>${impactoLabel(impDoTotal_dia)}</strong></td>
      </tr>
    </tbody>
  </table>
  <table class="data-table" style="margin-top:6px">
    <colgroup>
      <col style="width:34%"><col style="width:16%"><col style="width:16%"><col style="width:16%"><col style="width:18%">
    </colgroup>
    <thead><tr>
      <th>Acumulado mes por producto y unidad</th>
      <th>HFO – U1</th>
      <th>DO – U1</th>
      <th>HFO – U2</th>
      <th>DO – U2</th>
    </tr></thead>
    <tbody>
      <tr class="rpt-row-grand">
        <td class="label"><strong>Balance acumulado del mes</strong></td>
        <td class="num ${impactoClass(impAcumHfoU1)}">${impactoLabel(impAcumHfoU1)}</td>
        <td class="num ${impactoClass(impAcumDoU1)}">${impactoLabel(impAcumDoU1)}</td>
        <td class="num ${impactoClass(impAcumHfoU2)}">${impactoLabel(impAcumHfoU2)}</td>
        <td class="num ${impactoClass(impAcumDoU2)}">${impactoLabel(impAcumDoU2)}</td>
      </tr>
    </tbody>
  </table>
  <p class="rpt-muted" style="margin-top:5px;font-size:10.5px">* Balance positivo = por encima de referencia (mayor consumo que histórico). Referencia 90D calculada por unidad solo con días operados.</p>
</div>`;
    }

    // ── MENSUAL ──────────────────────────────────────────────────────────────
    const endM = new Date(fechaJS.getFullYear(), fechaJS.getMonth(), fechaJS.getDate());
    const yM = endM.getFullYear(), mM = endM.getMonth();
    let sumHfoU1=0, sumDslU1=0, sumFuelU1=0, sumHfoU2=0, sumDslU2=0, sumFuelU2=0;
    let sumH1=0, sumH2=0, sumDsl=0, sumFuel=0;
    for (const d of metrics) {
      if (d.date.getFullYear() === yM && d.date.getMonth() === mM && d.date <= endM) {
        sumHfoU1 += d.hfo_u1; sumDslU1 += d.dsl_u1; sumFuelU1 += d.fuel_u1;
        sumHfoU2 += d.hfo_u2; sumDslU2 += d.dsl_u2; sumFuelU2 += d.fuel_u2;
        sumH1 += d.h1; sumH2 += d.h2;
        sumDsl += d.dsl; sumFuel += d.fuel;
      }
    }
    if (!(sumFuel > 0 && (sumH1 + sumH2) > 0)) {
      return `<div class="rpt-notice">Análisis Ejecutivo de Combustible: Sin datos suficientes del mes.</div>`;
    }
    const galh_mes_u1      = sumH1 > 0 ? sumFuelU1 / sumH1 : NaN;
    const galh_hfo_mes_u1  = sumH1 > 0 ? sumHfoU1  / sumH1 : NaN;
    const galh_do_mes_u1   = sumH1 > 0 ? sumDslU1  / sumH1 : NaN;
    const galh_mes_u2      = sumH2 > 0 ? sumFuelU2 / sumH2 : NaN;
    const galh_hfo_mes_u2  = sumH2 > 0 ? sumHfoU2  / sumH2 : NaN;
    const galh_do_mes_u2   = sumH2 > 0 ? sumDslU2  / sumH2 : NaN;
    const dU1m = Number.isFinite(galh_mes_u1) && Number.isFinite(galh_ref_u1) ? galh_mes_u1 - galh_ref_u1 : NaN;
    const dU2m = Number.isFinite(galh_mes_u2) && Number.isFinite(galh_ref_u2) ? galh_mes_u2 - galh_ref_u2 : NaN;
    const impU1m = Number.isFinite(dU1m) ? dU1m * sumH1 : NaN;
    const impU2m = Number.isFinite(dU2m) ? dU2m * sumH2 : NaN;
    const impTotalM = (Number.isFinite(impU1m) ? impU1m : 0) + (Number.isFinite(impU2m) ? impU2m : 0);
    const pctDO_mes = sumFuel > 0 ? sumDsl / sumFuel : NaN;

    // Impactos HFO y DO separados (mensual)
    const d_hfo_u1m_total = Number.isFinite(galh_hfo_mes_u1) && Number.isFinite(galh_hfo_ref_u1) ? galh_hfo_mes_u1 - galh_hfo_ref_u1 : NaN;
    const d_do_u1m_total  = Number.isFinite(galh_do_mes_u1)  && Number.isFinite(galh_do_ref_u1)  ? galh_do_mes_u1  - galh_do_ref_u1  : NaN;
    const d_hfo_u2m_total = Number.isFinite(galh_hfo_mes_u2) && Number.isFinite(galh_hfo_ref_u2) ? galh_hfo_mes_u2 - galh_hfo_ref_u2 : NaN;
    const d_do_u2m_total  = Number.isFinite(galh_do_mes_u2)  && Number.isFinite(galh_do_ref_u2)  ? galh_do_mes_u2  - galh_do_ref_u2  : NaN;
    const impHfoU1m = Number.isFinite(d_hfo_u1m_total) && sumH1 > 0 ? d_hfo_u1m_total * sumH1 : NaN;
    const impDoU1m  = Number.isFinite(d_do_u1m_total)  && sumH1 > 0 ? d_do_u1m_total  * sumH1 : NaN;
    const impHfoU2m = Number.isFinite(d_hfo_u2m_total) && sumH2 > 0 ? d_hfo_u2m_total * sumH2 : NaN;
    const impDoU2m  = Number.isFinite(d_do_u2m_total)  && sumH2 > 0 ? d_do_u2m_total  * sumH2 : NaN;
    const impHfoTotalM = (Number.isFinite(impHfoU1m) ? impHfoU1m : 0) + (Number.isFinite(impHfoU2m) ? impHfoU2m : 0);
    const impDoTotalM  = (Number.isFinite(impDoU1m)  ? impDoU1m  : 0) + (Number.isFinite(impDoU2m)  ? impDoU2m  : 0);

    const msgsM: string[] = [];
    if (sumH1 > 0 && Number.isFinite(dU1m)) {
      if (dU1m > 1.0) msgsM.push(`U1 por encima de referencia en el período (+${fmt1(dU1m)} gal/h)`);
      else if (dU1m < -1.0) msgsM.push(`U1 por debajo de referencia en el período (${fmt1(dU1m)} gal/h)`);
    }
    if (sumH2 > 0 && Number.isFinite(dU2m)) {
      if (dU2m > 1.0) msgsM.push(`U2 por encima de referencia en el período (+${fmt1(dU2m)} gal/h)`);
      else if (dU2m < -1.0) msgsM.push(`U2 por debajo de referencia en el período (${fmt1(dU2m)} gal/h)`);
    }
    const causaM = msgsM.length ? msgsM.join(". ") + "." : "Operación dentro de los parámetros de referencia 90D.";

    const d_hfo_u1m = Number.isFinite(galh_hfo_mes_u1)&&Number.isFinite(galh_hfo_ref_u1) ? galh_hfo_mes_u1-galh_hfo_ref_u1 : NaN;
    const d_do_u1m  = Number.isFinite(galh_do_mes_u1)&&Number.isFinite(galh_do_ref_u1)   ? galh_do_mes_u1-galh_do_ref_u1   : NaN;
    const d_hfo_u2m = Number.isFinite(galh_hfo_mes_u2)&&Number.isFinite(galh_hfo_ref_u2) ? galh_hfo_mes_u2-galh_hfo_ref_u2 : NaN;
    const d_do_u2m  = Number.isFinite(galh_do_mes_u2)&&Number.isFinite(galh_do_ref_u2)   ? galh_do_mes_u2-galh_do_ref_u2   : NaN;

    const rowsU1m = sumH1 > 0 ? `
      <tr class="rpt-row-grupo"><td colspan="5" class="label">Unidad 1 &nbsp;<span class="rpt-muted">(${fmt0(sumH1)} h en el período)</span></td></tr>
      <tr>
        <td class="label" style="padding-left:16px">HFO</td>
        <td class="num">${fmt1(galh_hfo_mes_u1)}</td>
        <td class="num">${fmt1(galh_hfo_ref_u1)}</td>
        <td class="num ${desvClass(d_hfo_u1m)}">${desvLabel(d_hfo_u1m)}</td>
        <td class="num ${impactoClass(impHfoU1m)}">${impactoLabel(impHfoU1m)}</td>
      </tr>
      <tr>
        <td class="label" style="padding-left:16px">Diésel (DO)</td>
        <td class="num">${fmt1(galh_do_mes_u1)}</td>
        <td class="num">${fmt1(galh_do_ref_u1)}</td>
        <td class="num ${desvClass(d_do_u1m)}">${desvLabel(d_do_u1m)}</td>
        <td class="num ${impactoClass(impDoU1m)}">${impactoLabel(impDoU1m)}</td>
      </tr>`
      : `<tr><td class="label">Unidad 1</td><td colspan="4" style="text-align:left;color:#9ca3af;font-style:italic">Sin horas en el período</td></tr>`;

    const rowsU2m = sumH2 > 0 ? `
      <tr class="rpt-row-grupo"><td colspan="5" class="label">Unidad 2 &nbsp;<span class="rpt-muted">(${fmt0(sumH2)} h en el período)</span></td></tr>
      <tr>
        <td class="label" style="padding-left:16px">HFO</td>
        <td class="num">${fmt1(galh_hfo_mes_u2)}</td>
        <td class="num">${fmt1(galh_hfo_ref_u2)}</td>
        <td class="num ${desvClass(d_hfo_u2m)}">${desvLabel(d_hfo_u2m)}</td>
        <td class="num ${impactoClass(impHfoU2m)}">${impactoLabel(impHfoU2m)}</td>
      </tr>
      <tr>
        <td class="label" style="padding-left:16px">Diésel (DO)</td>
        <td class="num">${fmt1(galh_do_mes_u2)}</td>
        <td class="num">${fmt1(galh_do_ref_u2)}</td>
        <td class="num ${desvClass(d_do_u2m)}">${desvLabel(d_do_u2m)}</td>
        <td class="num ${impactoClass(impDoU2m)}">${impactoLabel(impDoU2m)}</td>
      </tr>`
      : `<tr><td class="label">Unidad 2</td><td colspan="4" style="text-align:left;color:#9ca3af;font-style:italic">Sin horas en el período</td></tr>`;

    return `<div class="rpt-fuel-box">
  <div class="rpt-fuel-header">
    <span class="rpt-fuel-title">Análisis Ejecutivo de Combustible (Mensual)</span>
  </div>
  <p class="rpt-fuel-causa">${causaM}</p>
  <table class="data-table">
    <thead><tr>
      <th>Combustible</th>
      <th>gal/h período</th>
      <th>Referencia 90D</th>
      <th>Desviación (gal/h)</th>
      <th>Balance período (gal)</th>
    </tr></thead>
    <tbody>
      ${rowsU1m}
      ${rowsU2m}
      <tr class="rpt-row-grand">
        <td class="label" colspan="4"><strong>Balance total HFO del período (U1+U2)</strong></td>
        <td class="num ${impactoClass(impHfoTotalM)}"><strong>${impactoLabel(impHfoTotalM)}</strong></td>
      </tr>
      <tr class="rpt-row-grand">
        <td class="label" colspan="4"><strong>Balance total DO del período (U1+U2)</strong></td>
        <td class="num ${impactoClass(impDoTotalM)}"><strong>${impactoLabel(impDoTotalM)}</strong></td>
      </tr>
    </tbody>
  </table>
  <p class="rpt-muted" style="margin-top:5px;font-size:10.5px">% Diésel del período: ${fmtPct(pctDO_mes)} &nbsp;|&nbsp; Referencia 90D: ${fmtPct(pctDO_ref)} &nbsp;|&nbsp; Balance positivo = por encima de referencia.</p>
</div>`;
  } catch (e) {
    console.error("FuelExecutive error:", e);
    return `<div class="rpt-notice rpt-notice-error">Análisis Ejecutivo de Combustible: No disponible por error de cálculo.</div>`;
  }
}

// ========= IDOM =========

interface RefPrev {
  mesRef: string;
  mesesUsados: string[];
  diasConDato: number;
  R_ref: number;
  IC_ref: number;
}

function calcularReferenciaPromedio3Meses(wbProd: XLSX.WorkBook, fechaJS: Date): RefPrev {
  const mesActual = monthStrFromDate(fechaJS);
  const m1 = prevMonthStr(mesActual);
  const m2 = m1 ? prevMonthStr(m1) : null;
  const m3 = m2 ? prevMonthStr(m2) : null;
  const meses = [m1, m2, m3].filter(Boolean) as string[];
  if (!meses.length) return { mesRef: "", mesesUsados: [], diasConDato: 0, R_ref: 0, IC_ref: 0 };

  let sumKWh = 0, sumGal = 0, sumHoras = 0;
  const diasSet = new Set<string>();

  for (const mesStr of meses) {
    const [yS, mS] = mesStr.split("-");
    const year = parseInt(yS, 10);
    const monthIndex = parseInt(mS, 10) - 1;
    const fechaCorte = new Date(year, monthIndex + 1, 0);
    fechaCorte.setHours(0, 0, 0, 0);
    const { rows } = getProdSheetAndRows(wbProd, fechaCorte);
    for (const r of rows) {
      const key = excelDateKey(r[CONFIG.COL_FECHA]);
      if (!key) continue;
      const d = new Date(key + "T00:00:00");
      if (d.getFullYear() !== year || d.getMonth() !== monthIndex) continue;

      const aux_kwh = posNum(r[CONFIG.COL_AUX_KWH]);
      const lan = posNum(r[CONFIG.COL_LANEC_PARCIAL_KWH]);
      const gra = posNum(r[CONFIG.COL_GRACA_PARCIAL_KWH]);
      const total_gen_kwh = lan + gra + aux_kwh;
      const hfo = posNum(r[CONFIG.COL_HFO_GAL]);
      const dsl = posNum(r[CONFIG.COL_DO_GAL]);
      const fuelTot = hfo + dsl;
      const h1i = posNum(r[CONFIG.COL_HG1_LANEC_INI]);
      const h1f = posNum(r[CONFIG.COL_HG1_LANEC_FIN]);
      const h2i = posNum(r[CONFIG.COL_HG2_LANEC_INI]);
      const h2f = posNum(r[CONFIG.COL_HG2_LANEC_FIN]);
      const u1_dia = Math.max(0, h1f - h1i);
      const u2_dia = Math.max(0, h2f - h2i);
      const horasOperDia = Math.max(u1_dia, u2_dia);

      if (total_gen_kwh > 0 && fuelTot > 0 && horasOperDia > 0) {
        sumKWh += total_gen_kwh;
        sumGal += fuelTot;
        sumHoras += horasOperDia;
        diasSet.add(jsDateKey(d));
      }
    }
  }

  const diasConDato = diasSet.size;
  const R_ref = sumGal > 0 ? sumKWh / sumGal : 0;
  const IC_ref = sumHoras > 0 ? (sumKWh / sumHoras) / P_INST_EFECTIVA : 0;
  return { mesRef: meses.join(", "), mesesUsados: meses, diasConDato, R_ref, IC_ref };
}

interface IDOMResult {
  R_dia: number;
  IE: number;
  ID: number;
  IC_dia: number;
  IDOM: number;
  estado: string;
  driver: string;
  Loss_kWh: number;
  ref: RefPrev;
}

function calcularIDOMDia(
  total_gen_kwh: number, fuelTot: number, horasOperDia: number,
  u1_dia: number, u2_dia: number, ref: RefPrev
): IDOMResult | null {
  if (!(total_gen_kwh > 0 && fuelTot > 0 && horasOperDia > 0 && ref && ref.R_ref > 0 && ref.IC_ref > 0)) return null;
  const R_dia = total_gen_kwh / fuelTot;
  const IE = R_dia / ref.R_ref;
  const Pavg = total_gen_kwh / horasOperDia;
  const IC_dia = Pavg / P_INST_EFECTIVA;
  const ID_ok = (((u1_dia > 0) ? 1 : 0) + ((u2_dia > 0) ? 1 : 0)) / 2;
  const IDOM = 0.4 * IE + 0.3 * ID_ok + 0.3 * (IC_dia / ref.IC_ref);
  const estado = semaforo(IDOM);
  const Loss_kWh = Math.max(0, fuelTot * ref.R_ref - total_gen_kwh);
  const penR = Math.max(0, ref.R_ref - R_dia);
  const penC = Math.max(0, ref.IC_ref - IC_dia);
  const driver = (ID_ok === 0.5) ? "DISPONIBILIDAD" : (penC > penR ? "CARGA" : "RENDIMIENTO");
  return { R_dia, IE, ID: ID_ok, IC_dia, IDOM, estado, driver, Loss_kWh, ref };
}

function notaPieIDOM(ref: RefPrev): string {
  if (!ref || !ref.R_ref || !ref.IC_ref) return "";
  return `<div class="rpt-nota-tecnica">
  <strong>Nota técnica — KPIs IDOM</strong><br>
  <strong>R_ref</strong>: Σ(kWh)/Σ(gal) del promedio móvil de 3 meses (${ref.mesRef}, días con dato: ${ref.diasConDato}).<br>
  <strong>IC_ref</strong>: P_promedio_op_ref / P_instalada_efectiva, P_eff = 0,85 × ${P_INST_TOTAL} kW = ${Math.round(P_INST_EFECTIVA)} kW.<br>
  <strong>IDOM_D</strong>: 0,4·IE + 0,3·ID + 0,3·(IC_día/IC_ref). &nbsp;
  <strong>Semáforo</strong>: VERDE ≥ 0,85 · AMARILLO 0,75–0,85 · NARANJA 0,65–0,75 · ROJO &lt; 0,65.
</div>`;
}

// ========= SÍNTESIS OPERATIVA (reemplaza IDOM en el informe) =========

function buildSintesisOperativa(params: {
  u1Horas: number;
  u2Horas: number;
  shareG1: number;
  shareG2: number;
  fp_eff: number;
  idomScore: number | null;
  aut_hfo: number;
  aut_do: number;
  u1_rest: number;
  u2_rest: number;
  esInformeMensual?: boolean;
}): string {
  const { u1Horas, u2Horas, shareG1, shareG2, fp_eff, idomScore, aut_hfo, aut_do, u1_rest, u2_rest, esInformeMensual } = params;
  const fmtP = (x: number) => Number.isFinite(x) && x >= 0 ? x.toFixed(1) + " %" : "—";

  // Disponibilidad
  const unidadesActivas = (u1Horas > 0 ? 1 : 0) + (u2Horas > 0 ? 1 : 0);
  const dispPct = unidadesActivas * 50;
  const dispLectura = dispPct === 100
    ? "Parque de generación disponible durante el período"
    : dispPct === 50
    ? "Una unidad disponible durante el período"
    : "Sin unidades en operación durante el período";

  // Participación por unidad
  const part1Lectura = u1Horas > 0
    ? (shareG1 >= 55 ? "Mayor aporte relativo en la operación del período"
      : shareG1 >= 45 ? "Aporte equilibrado con respecto a la otra unidad"
      : "Menor aporte relativo en la operación del período")
    : "Unidad sin operación en el período";
  const part2Lectura = u2Horas > 0
    ? (shareG2 >= 55 ? "Mayor aporte relativo en la operación del período"
      : shareG2 >= 45 ? "Aporte equilibrado con respecto a la otra unidad"
      : "Menor aporte relativo en la operación del período")
    : "Unidad sin operación en el período";

  // Utilización capacidad efectiva
  const fpPct = fp_eff * 100;
  const fpLectura = fpPct >= 65
    ? "Operación a carga alta, acorde a la demanda atendida"
    : fpPct >= 35
    ? "Nivel de carga acorde a la demanda atendida"
    : fpPct > 0
    ? "Carga reducida en el período"
    : "Sin operación en el período";

  // Condición operativa
  let condicion: string, condicionLectura: string;
  const scoreEfectivo = idomScore !== null ? idomScore : (fp_eff >= 0.60 ? 0.87 : fp_eff >= 0.30 ? 0.76 : 0.60);
  if (scoreEfectivo >= 0.85) {
    condicion = "Estable";
    condicionLectura = "Operación regular durante el período, dentro de los parámetros de referencia";
  } else if (scoreEfectivo >= 0.75) {
    condicion = "Estable con variaciones operativas";
    condicionLectura = "Comportamiento acorde al período, con variaciones dentro del rango operativo";
  } else {
    condicion = "Con seguimiento operativo";
    condicionLectura = "Período con aspectos de seguimiento para el equipo técnico";
  }

  // Aspecto principal
  let aspecto: string, aspectoLectura: string;
  if ((u1Horas > 0 && u1_rest < 500) || (u2Horas > 0 && u2_rest < 500)) {
    aspecto = "Planificación de mantenimiento preventivo";
    aspectoLectura = "Se recomienda programar revisión de horas acumuladas de la unidad próxima al intervalo";
  } else if (aut_hfo > 0 && aut_hfo < 5 && u1Horas + u2Horas > 0) {
    aspecto = "Seguimiento de reservas HFO";
    aspectoLectura = "Se recomienda monitorear niveles de stock de Fuel Oil Pesado";
  } else if (aut_do > 0 && aut_do < 5 && u1Horas + u2Horas > 0) {
    aspecto = "Seguimiento de reservas de Diésel";
    aspectoLectura = "Se recomienda monitorear niveles de stock de Diésel (DO)";
  } else {
    aspecto = "Operación regular";
    aspectoLectura = "Sin aspectos operativos de atención inmediata identificados en el período";
  }

  const periodoLabel = esInformeMensual ? "del mes" : "del día";

  return `<table class="data-table"><thead><tr>
<th style="width:34%">Indicador</th>
<th style="width:16%;text-align:center">Valor</th>
<th>Lectura ejecutiva</th>
</tr></thead><tbody>
<tr><td class="label">Disponibilidad de central</td><td class="num" style="text-align:center;font-weight:700">${fmtP(dispPct)}</td><td>${dispLectura}</td></tr>
<tr><td class="label">Participación operativa U1</td><td class="num" style="text-align:center">${u1Horas > 0 ? fmtP(shareG1) : "—"}</td><td>${part1Lectura}</td></tr>
<tr><td class="label">Participación operativa U2</td><td class="num" style="text-align:center">${u2Horas > 0 ? fmtP(shareG2) : "—"}</td><td>${part2Lectura}</td></tr>
<tr><td class="label">Utilización de capacidad efectiva</td><td class="num" style="text-align:center">${fmtP(fpPct)}</td><td>${fpLectura}</td></tr>
<tr class="rpt-row-total"><td class="label"><strong>Condición operativa ${periodoLabel}</strong></td><td style="text-align:center;font-weight:700;color:#1e3a6e;font-size:14px">${condicion}</td><td>${condicionLectura}</td></tr>
<tr><td class="label">Aspecto principal de seguimiento</td><td style="text-align:center;color:#374151;font-weight:600;font-size:13px">${aspecto}</td><td>${aspectoLectura}</td></tr>
</tbody></table>`;
}

// ========= IDOM GERENCIAL =========

/**
 * Genera la tabla gerencial del IDOM con 5 filas y conclusión.
 * Reemplaza buildSintesisOperativa en la sección 6 de cada informe.
 */
function buildIDOMGerencial(params: {
  idomScore: number | null;
  ie: number;
  driver: string;
  lossKwh: number;
  dispPct: number;
}): string {
  const { idomScore, ie, driver, lossKwh, dispPct } = params;

  const idomVal = idomScore !== null ? idomScore.toFixed(4) : "—";
  const idomLectura = idomScore === null
    ? "Sin datos de referencia disponibles para calcular IDOM"
    : idomScore >= 0.90 ? "Desempeño operativo favorable"
    : idomScore >= 0.80 ? "Desempeño operativo aceptable"
    : "Desempeño con seguimiento requerido";
  const idomClass = idomScore === null ? "" : idomScore >= 0.90 ? "hi" : idomScore >= 0.80 ? "" : "warn";

  const dispStr = dispPct.toFixed(0) + " %";
  const dispLectura = dispPct >= 100
    ? "La central se mantuvo disponible durante el período"
    : dispPct >= 50
    ? "La disponibilidad se vio parcialmente reducida durante el período"
    : "La central no operó durante el período";

  const iePct = Number.isFinite(ie) ? ie * 100 : NaN;
  const ieStr = Number.isFinite(iePct) ? iePct.toFixed(1) + " %" : "—";
  const ieLectura = !Number.isFinite(iePct)
    ? "Sin referencia disponible para comparación"
    : iePct >= 98 ? "Rendimiento en línea con la referencia"
    : iePct >= 92 ? "Rendimiento ligeramente por debajo de la referencia"
    : "Rendimiento por debajo de la referencia";

  const lossStr = Number.isFinite(lossKwh) && lossKwh >= 0
    ? fmt(lossKwh, 0) + " kWh" : "—";
  const lossLectura = !Number.isFinite(lossKwh) || lossKwh < 0
    ? "Sin datos suficientes"
    : lossKwh < 500 ? "Diferencia menor respecto a la referencia"
    : lossKwh < 2000 ? "Diferencia apreciable respecto a la referencia"
    : "Diferencia relevante respecto a la referencia";

  const driverLabels: Record<string, string> = {
    DISPONIBILIDAD: "El factor limitante fue la disponibilidad de unidades en el período",
    CARGA: "El factor dominante fue el nivel de carga atendida en el período",
    RENDIMIENTO: "El factor principal fue la eficiencia en el consumo de combustible",
  };
  const driverStr     = driver || "—";
  const driverLectura = driver
    ? (driverLabels[driver] ?? "Factor no determinado con precisión")
    : "Sin datos suficientes para determinar el factor";

  const conclusionTexto = idomScore === null
    ? "Sin datos de referencia histórica disponibles. Se requieren al menos 20 días previos con datos válidos para calcular el IDOM."
    : `La central operó con ${idomLectura.toLowerCase()}. ${driverLectura}.`;

  return `<table class="data-table"><thead><tr>
<th style="width:32%">Indicador</th>
<th style="width:20%;text-align:center">Valor</th>
<th>Lectura gerencial</th>
</tr></thead><tbody>
<tr><td class="label">IDOM del período</td><td class="num ${idomClass}" style="text-align:center;font-weight:700;font-size:16px">${idomVal}</td><td>${idomLectura}</td></tr>
<tr><td class="label">Disponibilidad de central</td><td class="num" style="text-align:center">${dispStr}</td><td>${dispLectura}</td></tr>
<tr><td class="label">Eficiencia frente a referencia</td><td class="num" style="text-align:center">${ieStr}</td><td>${ieLectura}</td></tr>
<tr><td class="label">Pérdida energética estimada</td><td class="num" style="text-align:center">${lossStr}</td><td>${lossLectura}</td></tr>
<tr class="rpt-row-total"><td class="label"><strong>Factor principal observado</strong></td><td style="text-align:center;font-weight:700;color:#1e3a6e;font-size:13px">${driverStr}</td><td>${driverLectura}</td></tr>
</tbody></table>
<p class="rpt-fuel-causa" style="margin-top:8px;font-size:13px"><strong>Conclusión IDOM:</strong> ${conclusionTexto}</p>`;
}

// ========= INFORME DIARIO =========

export function generarInformeDiario(
  prodBuffer: ArrayBuffer,
  fechaStr: string,
  obs: string
): string {
  const fechaJS = new Date(fechaStr + "T00:00:00");
  const fechaLarga = formatFechaLarga(fechaStr);
  const wbProd = XLSX.read(prodBuffer, { type: "array" });
  const { rows } = getProdSheetAndRows(wbProd, fechaJS);
  const row = findRowByDate(rows, fechaJS);
  if (!row) throw new Error("No se encontró la fecha en el archivo de producción.");

  const aux_kwh = posNum(row[CONFIG.COL_AUX_KWH]);
  const lanec_kwh = posNum(row[CONFIG.COL_LANEC_PARCIAL_KWH]);
  const graca_kwh = posNum(row[CONFIG.COL_GRACA_PARCIAL_KWH]);
  const gen1_kwh = posNum(row[CONFIG.COL_GEN1_KWH]);
  const gen2_kwh = posNum(row[CONFIG.COL_GEN2_KWH]);
  const total_kwh_clientes = lanec_kwh + graca_kwh;
  const total_gen_kwh = total_kwh_clientes + aux_kwh;
  const share_lan = total_kwh_clientes > 0 ? (lanec_kwh / total_kwh_clientes) * 100 : 0;
  const share_gra = total_kwh_clientes > 0 ? (graca_kwh / total_kwh_clientes) * 100 : 0;
  const share_aux = total_gen_kwh > 0 ? (aux_kwh / total_gen_kwh) * 100 : 0;
  const sumGenKwh = gen1_kwh + gen2_kwh;

  const h1i = posNum(row[CONFIG.COL_HG1_LANEC_INI]);
  const h1f = posNum(row[CONFIG.COL_HG1_LANEC_FIN]);
  const h2i = posNum(row[CONFIG.COL_HG2_LANEC_INI]);
  const h2f = posNum(row[CONFIG.COL_HG2_LANEC_FIN]);
  const u1_dia = Math.max(0, h1f - h1i);
  const u2_dia = Math.max(0, h2f - h2i);
  const u1_ac = Math.max(0, h1f - HORO_BASE_U1);
  const u2_ac = Math.max(0, h2f - HORO_BASE_U2);
  const u1_rest = OBJ_MTO_HORAS_U1 - u1_ac;
  const u2_rest = OBJ_MTO_HORAS_U2 - u2_ac;
  const horasOperDia = Math.max(u1_dia, u2_dia);
  const pmed_total = horasOperDia > 0 ? total_gen_kwh / horasOperDia : 0;
  const pmed_cli = horasOperDia > 0 ? total_kwh_clientes / horasOperDia : 0;
  const pmed_aux = horasOperDia > 0 ? aux_kwh / horasOperDia : 0;
  const pmed_lan = horasOperDia > 0 ? lanec_kwh / horasOperDia : 0;
  const pmed_gra = horasOperDia > 0 ? graca_kwh / horasOperDia : 0;
  const pmed_g1 = u1_dia > 0 ? gen1_kwh / u1_dia : 0;
  const pmed_g2 = u2_dia > 0 ? gen2_kwh / u2_dia : 0;
  const shareG1 = sumGenKwh > 0 ? (gen1_kwh / sumGenKwh) * 100 : 0;
  const shareG2 = sumGenKwh > 0 ? (gen2_kwh / sumGenKwh) * 100 : 0;

  const hfo = posNum(row[CONFIG.COL_HFO_GAL]);
  const dsl = posNum(row[CONFIG.COL_DO_GAL]);
  const fuelTot = hfo + dsl;
  const rendimiento = fuelTot > 0 ? total_gen_kwh / fuelTot : 0;
  const stock_hfo = posNum(row[CONFIG.COL_STOCK_HFO_TOTAL]);
  const stock_do = posNum(row[CONFIG.COL_STOCK_DO_TOTAL]);
  const aut_hfo = hfo > 0 ? stock_hfo / hfo : 0;
  const aut_do = dsl > 0 ? stock_do / dsl : 0;

  let html = rptHeader("Reporte Post Operativo Diario", fechaLarga);

  html += seccion(1, "Producción de Energía");
  html += `<table class="data-table"><thead><tr>
<th>Concepto</th><th>Energía [kWh]</th><th>Potencia media [kW]</th></tr></thead>
<tbody>
<tr><td class="label">Energía generada total</td><td class="num hi">${fmt(total_gen_kwh)}</td><td>${fmt(pmed_total, 1)}</td></tr>
<tr><td class="label">Energía a clientes</td><td class="num">${fmt(total_kwh_clientes)}</td><td>${fmt(pmed_cli, 1)}</td></tr>
<tr><td class="label">Auxiliares</td><td class="num">${fmt(aux_kwh)}</td><td>${fmt(pmed_aux, 1)}</td></tr>
</tbody></table>`;

  html += `<table class="data-table"><thead><tr>
<th>Unidad generadora</th><th>Energía [kWh]</th><th>Potencia media [kW]</th><th>Participación [%]</th>
</tr></thead><tbody>
<tr><td class="label">Generador 1</td><td class="num">${fmt(gen1_kwh)}</td><td>${u1_dia > 0 ? fmt(pmed_g1, 1) : "—"}</td><td>${fmt(shareG1, 1)}</td></tr>
<tr><td class="label">Generador 2</td><td class="num">${fmt(gen2_kwh)}</td><td>${u2_dia > 0 ? fmt(pmed_g2, 1) : "—"}</td><td>${fmt(shareG2, 1)}</td></tr>
</tbody></table>`;

  html += seccion(2, "Distribución por Alimentador");
  html += `<table class="data-table"><thead><tr>
<th>Destino</th><th>Energía [kWh]</th><th>Potencia media [kW]</th><th>Participación [%]</th>
</tr></thead><tbody>
<tr><td class="label">LANEC</td><td class="num">${fmt(lanec_kwh)}</td><td>${horasOperDia > 0 ? fmt(pmed_lan, 1) : "—"}</td><td>${fmt(share_lan, 1)}</td></tr>
<tr><td class="label">GRACA</td><td class="num">${fmt(graca_kwh)}</td><td>${horasOperDia > 0 ? fmt(pmed_gra, 1) : "—"}</td><td>${fmt(share_gra, 1)}</td></tr>
<tr><td class="label">Auxiliares</td><td class="num">${fmt(aux_kwh)}</td><td>${horasOperDia > 0 ? fmt(pmed_aux, 1) : "—"}</td><td>${fmt(share_aux, 1)}</td></tr>
</tbody></table>`;

  // Consumo por unidad (prorrateo por energía generada)
  const genTot = gen1_kwh + gen2_kwh;
  const w1 = genTot > 0 ? gen1_kwh / genTot : (u1_dia > 0 ? u1_dia / (u1_dia + u2_dia || 1) : 0);
  const w2 = genTot > 0 ? gen2_kwh / genTot : (u2_dia > 0 ? u2_dia / (u1_dia + u2_dia || 1) : 0);
  const hfo_u1 = hfo * w1;  const dsl_u1 = dsl * w1;  const fuel_u1 = fuelTot * w1;
  const hfo_u2 = hfo * w2;  const dsl_u2 = dsl * w2;  const fuel_u2 = fuelTot * w2;
  const galh_u1     = u1_dia > 0 ? fuel_u1 / u1_dia : NaN;
  const galh_hfo_u1 = u1_dia > 0 ? hfo_u1  / u1_dia : NaN;
  const galh_do_u1  = u1_dia > 0 ? dsl_u1  / u1_dia : NaN;
  const galh_u2     = u2_dia > 0 ? fuel_u2 / u2_dia : NaN;
  const galh_hfo_u2 = u2_dia > 0 ? hfo_u2  / u2_dia : NaN;
  const galh_do_u2  = u2_dia > 0 ? dsl_u2  / u2_dia : NaN;

  html += seccion(3, "Combustible y Eficiencia");
  html += `<table class="data-table"><thead><tr>
<th>Combustible / Unidad</th><th>Total planta [gal]</th><th>U1 est. [gal]</th><th>gal/h U1</th><th>U2 est. [gal]</th><th>gal/h U2</th></tr></thead><tbody>
<tr>
  <td class="label">HFO (Fuel Oil Pesado)</td>
  <td class="num">${fmt(hfo)}</td>
  <td class="num">${u1_dia > 0 ? fmt(hfo_u1) : "—"}</td>
  <td class="num">${u1_dia > 0 ? fmt(galh_hfo_u1, 1) : "—"}</td>
  <td class="num">${u2_dia > 0 ? fmt(hfo_u2) : "—"}</td>
  <td class="num">${u2_dia > 0 ? fmt(galh_hfo_u2, 1) : "—"}</td>
</tr>
<tr>
  <td class="label">Diésel (DO)</td>
  <td class="num">${fmt(dsl)}</td>
  <td class="num">${u1_dia > 0 ? fmt(dsl_u1) : "—"}</td>
  <td class="num">${u1_dia > 0 ? fmt(galh_do_u1, 1) : "—"}</td>
  <td class="num">${u2_dia > 0 ? fmt(dsl_u2) : "—"}</td>
  <td class="num">${u2_dia > 0 ? fmt(galh_do_u2, 1) : "—"}</td>
</tr>
<tr class="rpt-row-total">
  <td class="label"><strong>Total</strong></td>
  <td class="num hi"><strong>${fmt(fuelTot)}</strong></td>
  <td class="num">${u1_dia > 0 ? fmt(fuel_u1) : "—"}</td>
  <td class="num hi">${u1_dia > 0 ? fmt(galh_u1, 1) : "—"}</td>
  <td class="num">${u2_dia > 0 ? fmt(fuel_u2) : "—"}</td>
  <td class="num hi">${u2_dia > 0 ? fmt(galh_u2, 1) : "—"}</td>
</tr>
</tbody></table>
<p class="rpt-muted" style="font-size:10.5px;margin-bottom:6px">* Consumo por unidad estimado por prorrateo proporcional a energía generada (kWh). U1 = Generador 1, U2 = Generador 2.</p>
<div class="rpt-kpi-inline">Rendimiento global: <span class="rpt-kpi-val">${fmt(rendimiento, 2)} kWh/gal</span></div>`;

  html += buildFuelExecutiveHTML(wbProd, fechaJS, "daily");

  html += seccion(4, "Horas de Operación");
  html += `<table class="data-table"><thead><tr>
<th>Unidad</th><th>Horas del día [h]</th><th>Horas acumuladas [h]</th><th>Restantes para mantenimiento [h]</th>
</tr></thead><tbody>
<tr><td class="label">Unidad 1</td><td class="num">${fmt(u1_dia, 1)}</td><td class="num">${fmt(u1_ac, 1)}</td><td class="num ${u1_rest < 500 ? "warn" : ""}">${fmt(u1_rest, 1)}</td></tr>
<tr><td class="label">Unidad 2</td><td class="num">${fmt(u2_dia, 1)}</td><td class="num">${fmt(u2_ac, 1)}</td><td class="num ${u2_rest < 500 ? "warn" : ""}">${fmt(u2_rest, 1)}</td></tr>
</tbody></table>`;

  html += seccion(5, "Stocks y Autonomías");
  html += `<table class="data-table"><thead><tr>
<th>Producto</th><th>Stock [gal]</th><th>Autonomía estimada [días]</th></tr></thead><tbody>
<tr><td class="label">HFO (Fuel Oil Pesado)</td><td class="num">${fmt(stock_hfo)}</td><td class="num ${aut_hfo > 0 && aut_hfo < 3 ? "warn" : ""}">${aut_hfo > 0 ? fmt(aut_hfo, 1) : "—"}</td></tr>
<tr><td class="label">Diésel (DO)</td><td class="num">${fmt(stock_do)}</td><td class="num ${aut_do > 0 && aut_do < 3 ? "warn" : ""}">${aut_do > 0 ? fmt(aut_do, 1) : "—"}</td></tr>
</tbody></table>`;

  const refPrev = calcularReferenciaPromedio3Meses(wbProd, fechaJS);
  const idomDia = calcularIDOMDia(total_gen_kwh, fuelTot, horasOperDia, u1_dia, u2_dia, refPrev);
  const fp_eff_dia = horasOperDia > 0 ? pmed_total / P_INST_EFECTIVA : 0;

  html += seccion(6, "Síntesis Operativa");
  html += buildSintesisOperativa({
    u1Horas: u1_dia,
    u2Horas: u2_dia,
    shareG1,
    shareG2,
    fp_eff: fp_eff_dia,
    idomScore: idomDia ? idomDia.IDOM : null,
    aut_hfo,
    aut_do,
    u1_rest,
    u2_rest,
    esInformeMensual: false,
  });

  html += seccion(7, "Observaciones");
  html += obs
    ? `<div class="rpt-obs">${obs.replace(/\n/g, "<br>")}</div>`
    : `<div class="rpt-obs rpt-obs-empty">Sin novedades operativas relevantes.</div>`;

  return html;
}

// ========= INFORME MENSUAL =========

export function generarInformeMensual(prodBuffer: ArrayBuffer, mesStr: string): string {
  const partes = mesStr.split("-");
  if (partes.length !== 2) throw new Error("Formato de mes inválido.");
  const year = parseInt(partes[0], 10);
  const monthIndex = parseInt(partes[1], 10) - 1;
  if (isNaN(year) || isNaN(monthIndex) || monthIndex < 0 || monthIndex > 11) throw new Error("Mes inválido.");

  const fechaCorte = new Date(year, monthIndex + 1, 0);
  fechaCorte.setHours(0, 0, 0, 0);

  const wbProd = XLSX.read(prodBuffer, { type: "array" });
  const { rows } = getProdSheetAndRows(wbProd, fechaCorte);

  let lan = 0, gra = 0, aux = 0, g1 = 0, g2 = 0, hfo = 0, dsl = 0;
  let first_h1: number | null = null, last_h1: number | null = null;
  let first_h2: number | null = null, last_h2: number | null = null;
  let ultimoDia = 0;

  for (const r of rows) {
    const key = excelDateKey(r[CONFIG.COL_FECHA]);
    if (!key) continue;
    const d = new Date(key + "T00:00:00");
    if (d.getMonth() !== monthIndex || d.getFullYear() !== year) continue;
    if (d.getTime() > fechaCorte.getTime()) continue;
    if (d.getDate() > ultimoDia) ultimoDia = d.getDate();

    lan += posNum(r[CONFIG.COL_LANEC_PARCIAL_KWH]);
    gra += posNum(r[CONFIG.COL_GRACA_PARCIAL_KWH]);
    aux += posNum(r[CONFIG.COL_AUX_KWH]);
    g1 += posNum(r[CONFIG.COL_GEN1_KWH]);
    g2 += posNum(r[CONFIG.COL_GEN2_KWH]);
    hfo += posNum(r[CONFIG.COL_HFO_GAL]);
    dsl += posNum(r[CONFIG.COL_DO_GAL]);

    const h1i = posNum(r[CONFIG.COL_HG1_LANEC_INI]);
    const h1f = posNum(r[CONFIG.COL_HG1_LANEC_FIN]);
    const h2i = posNum(r[CONFIG.COL_HG2_LANEC_INI]);
    const h2f = posNum(r[CONFIG.COL_HG2_LANEC_FIN]);

    if (h1i > 0 && first_h1 === null) first_h1 = h1i;
    if (h1f > 0) last_h1 = h1f;
    if (h2i > 0 && first_h2 === null) first_h2 = h2i;
    if (h2f > 0) last_h2 = h2f;
  }

  const tot_cli = lan + gra;
  const tot_gen = tot_cli + aux;
  const fuelTot = hfo + dsl;
  const rendimiento = fuelTot > 0 ? tot_gen / fuelTot : 0;
  const shareL = tot_cli > 0 ? (lan / tot_cli) * 100 : 0;
  const shareG = tot_cli > 0 ? (gra / tot_cli) * 100 : 0;
  const u1_mes = (first_h1 !== null && last_h1 !== null) ? Math.max(0, last_h1 - first_h1) : 0;
  const u2_mes = (first_h2 !== null && last_h2 !== null) ? Math.max(0, last_h2 - first_h2) : 0;
  const horasOperMes = Math.max(u1_mes, u2_mes);
  const pmed_total = horasOperMes > 0 ? tot_gen / horasOperMes : 0;
  const pmed_cli = horasOperMes > 0 ? tot_cli / horasOperMes : 0;
  const pmed_aux = horasOperMes > 0 ? aux / horasOperMes : 0;
  const pmed_lan = horasOperMes > 0 ? lan / horasOperMes : 0;
  const pmed_gra = horasOperMes > 0 ? gra / horasOperMes : 0;
  const pmed_g1 = u1_mes > 0 ? g1 / u1_mes : 0;
  const pmed_g2 = u2_mes > 0 ? g2 / u2_mes : 0;
  const sumGenKwh = g1 + g2;
  const shareG1 = sumGenKwh > 0 ? (g1 / sumGenKwh) * 100 : 0;
  const shareG2 = sumGenKwh > 0 ? (g2 / sumGenKwh) * 100 : 0;
  // ─── Prorrateo por unidad: suma de prorrateos diarios ────────────────────────
  // Método idéntico al Análisis Ejecutivo de Combustible para garantizar
  // consistencia entre la tabla "Combustible / Unidad" y el análisis.
  // Un único peso mensual (g1/g2 total) da resultados distintos cuando la
  // proporción generador-1 / generador-2 varía día a día.
  const monthlyFuelMetrics = getAllFuelMetrics(wbProd).filter(
    d => d.date.getFullYear() === year && d.date.getMonth() === monthIndex,
  );
  let hfo_u1m = 0, dsl_u1m = 0, fuel_u1m = 0;
  let hfo_u2m = 0, dsl_u2m = 0, fuel_u2m = 0;
  // h1_fuel / h2_fuel: horas de los días con registro completo (kWh+comb.+horas)
  // Son las mismas horas que usa el Análisis Ejecutivo como denominador en gal/h.
  let h1_fuel = 0, h2_fuel = 0;
  for (const d of monthlyFuelMetrics) {
    hfo_u1m += d.hfo_u1;  dsl_u1m += d.dsl_u1;  fuel_u1m += d.fuel_u1;
    hfo_u2m += d.hfo_u2;  dsl_u2m += d.dsl_u2;  fuel_u2m += d.fuel_u2;
    h1_fuel += d.h1;       h2_fuel += d.h2;
  }
  const galh_u1m     = h1_fuel > 0 ? fuel_u1m / h1_fuel : NaN;
  const galh_hfo_u1m = h1_fuel > 0 ? hfo_u1m  / h1_fuel : NaN;
  const galh_do_u1m  = h1_fuel > 0 ? dsl_u1m  / h1_fuel : NaN;
  const galh_u2m     = h2_fuel > 0 ? fuel_u2m / h2_fuel : NaN;
  const galh_hfo_u2m = h2_fuel > 0 ? hfo_u2m  / h2_fuel : NaN;
  const galh_do_u2m  = h2_fuel > 0 ? dsl_u2m  / h2_fuel : NaN;
  const mesTexto = getSheetNameFromDate(fechaCorte);
  const textoPeriodo = ultimoDia > 0 ? `${mesTexto} (hasta el día ${ultimoDia})` : mesTexto;
  const diasPeriodo = ultimoDia > 0 ? ultimoDia : getDaysInMonth(year, monthIndex);
  const horasCalendario = diasPeriodo * 24;
  const fp_inst = horasCalendario > 0 ? tot_gen / (P_INST_TOTAL * horasCalendario) : 0;
  const fp_eff = horasCalendario > 0 ? tot_gen / (P_INST_EFECTIVA * horasCalendario) : 0;
  const aux_lan = tot_cli > 0 ? aux * (lan / tot_cli) : 0;
  const aux_gra = tot_cli > 0 ? aux * (gra / tot_cli) : 0;
  const lan_fact = lan + aux_lan;
  const gra_fact = gra + aux_gra;

  let html = rptHeader("Reporte Post Operativo Mensual", textoPeriodo);

  html += seccion(1, "Producción de Energía");
  html += `<table class="data-table"><thead><tr>
<th>Concepto</th><th>Energía [kWh]</th><th>Potencia media [kW]</th>
</tr></thead><tbody>
<tr><td class="label">Energía generada total</td><td class="num hi">${fmt(tot_gen)}</td><td>${horasOperMes > 0 ? fmt(pmed_total, 1) : "—"}</td></tr>
<tr><td class="label">Energía a clientes</td><td class="num">${fmt(tot_cli)}</td><td>${horasOperMes > 0 ? fmt(pmed_cli, 1) : "—"}</td></tr>
<tr><td class="label">Auxiliares</td><td class="num">${fmt(aux)}</td><td>${horasOperMes > 0 ? fmt(pmed_aux, 1) : "—"}</td></tr>
</tbody></table>`;

  html += `<table class="rpt-kpi-row" style="display:table;width:100%;border-collapse:separate;border-spacing:8px"><tr>
  <td class="rpt-kpi-card"><div class="rpt-kpi-label">Factor de planta (vs instalada)</div><div class="rpt-kpi-big">${(fp_inst * 100).toFixed(1)}<span class="rpt-kpi-unit">%</span></div><div class="rpt-kpi-sub">${P_INST_TOTAL} kW instalados</div></td>
  <td class="rpt-kpi-card"><div class="rpt-kpi-label">Factor de planta (vs efectiva)</div><div class="rpt-kpi-big">${(fp_eff * 100).toFixed(1)}<span class="rpt-kpi-unit">%</span></div><div class="rpt-kpi-sub">${Math.round(P_INST_EFECTIVA)} kW efectivos</div></td>
  <td class="rpt-kpi-card"><div class="rpt-kpi-label">Rendimiento promedio</div><div class="rpt-kpi-big">${fmt(rendimiento, 2)}<span class="rpt-kpi-unit">kWh/gal</span></div><div class="rpt-kpi-sub">Energía por galón consumido</div></td>
</tr></table>`;

  html += `<table class="data-table"><thead><tr>
<th>Unidad generadora</th><th>Energía [kWh]</th><th>Potencia media [kW]</th><th>Participación [%]</th>
</tr></thead><tbody>
<tr><td class="label">Generador 1</td><td class="num">${fmt(g1)}</td><td>${u1_mes > 0 ? fmt(pmed_g1, 1) : "—"}</td><td>${fmt(shareG1, 1)}</td></tr>
<tr><td class="label">Generador 2</td><td class="num">${fmt(g2)}</td><td>${u2_mes > 0 ? fmt(pmed_g2, 1) : "—"}</td><td>${fmt(shareG2, 1)}</td></tr>
</tbody></table>`;

  html += seccion(2, "Distribución Energética");
  html += `<table class="data-table"><thead><tr>
<th>Destino</th><th>Energía [kWh]</th><th>Potencia media [kW]</th><th>Participación [%]</th>
</tr></thead><tbody>
<tr><td class="label">LANEC</td><td class="num">${fmt(lan)}</td><td>${horasOperMes > 0 ? fmt(pmed_lan, 1) : "—"}</td><td>${fmt(shareL, 1)}</td></tr>
<tr><td class="label">GRACA</td><td class="num">${fmt(gra)}</td><td>${horasOperMes > 0 ? fmt(pmed_gra, 1) : "—"}</td><td>${fmt(shareG, 1)}</td></tr>
</tbody></table>`;

  html += seccion(3, "Combustible y Eficiencia");
  html += `<table class="data-table"><thead><tr>
<th>Combustible / Unidad</th><th>Total planta [gal]</th><th>U1 est. [gal]</th><th>gal/h U1</th><th>U2 est. [gal]</th><th>gal/h U2</th></tr></thead><tbody>
<tr>
  <td class="label">HFO (Fuel Oil Pesado)</td>
  <td class="num">${fmt(hfo)}</td>
  <td class="num">${h1_fuel > 0 ? fmt(hfo_u1m) : "—"}</td>
  <td class="num">${h1_fuel > 0 ? fmt(galh_hfo_u1m, 1) : "—"}</td>
  <td class="num">${h2_fuel > 0 ? fmt(hfo_u2m) : "—"}</td>
  <td class="num">${h2_fuel > 0 ? fmt(galh_hfo_u2m, 1) : "—"}</td>
</tr>
<tr>
  <td class="label">Diésel (DO)</td>
  <td class="num">${fmt(dsl)}</td>
  <td class="num">${h1_fuel > 0 ? fmt(dsl_u1m) : "—"}</td>
  <td class="num">${h1_fuel > 0 ? fmt(galh_do_u1m, 1) : "—"}</td>
  <td class="num">${h2_fuel > 0 ? fmt(dsl_u2m) : "—"}</td>
  <td class="num">${h2_fuel > 0 ? fmt(galh_do_u2m, 1) : "—"}</td>
</tr>
<tr class="rpt-row-total">
  <td class="label"><strong>Total</strong></td>
  <td class="num hi"><strong>${fmt(fuelTot)}</strong></td>
  <td class="num">${h1_fuel > 0 ? fmt(fuel_u1m) : "—"}</td>
  <td class="num hi">${h1_fuel > 0 ? fmt(galh_u1m, 1) : "—"}</td>
  <td class="num">${h2_fuel > 0 ? fmt(fuel_u2m) : "—"}</td>
  <td class="num hi">${h2_fuel > 0 ? fmt(galh_u2m, 1) : "—"}</td>
</tr>
</tbody></table>
<p class="rpt-muted" style="font-size:10.5px;margin-bottom:6px">* Consumo por unidad: suma de prorrateos diarios proporcional a energía generada (kWh). gal/h: Σgal_unidad / Σhoras_con_dato. Método idéntico al Análisis Ejecutivo de Combustible.</p>`;
  html += buildFuelExecutiveHTML(wbProd, fechaCorte, "monthly");

  html += seccion(4, "Horas de Operación");
  html += `<table class="data-table"><thead><tr>
<th>Unidad</th><th>Horas del mes [h]</th></tr></thead><tbody>
<tr><td class="label">Unidad 1</td><td class="num">${fmt(u1_mes, 1)}</td></tr>
<tr><td class="label">Unidad 2</td><td class="num">${fmt(u2_mes, 1)}</td></tr>
<tr class="rpt-row-total"><td class="label">Horas sistema (máximo)</td><td class="num hi">${fmt(horasOperMes, 1)}</td></tr>
</tbody></table>`;

  html += seccion(5, "Distribución de Energía Facturable");
  html += `<table class="data-table"><thead><tr>
<th>Cliente</th><th>Energía directa [kWh]</th><th>Part. [%]</th><th>Aux. asignados [kWh]</th><th>Total facturable [kWh]</th>
</tr></thead><tbody>
<tr><td class="label">LANEC</td><td class="num">${fmt(lan)}</td><td>${fmt(shareL, 2)}</td><td class="num">${fmt(aux_lan)}</td><td class="num hi">${fmt(lan_fact)}</td></tr>
<tr><td class="label">GRACA</td><td class="num">${fmt(gra)}</td><td>${fmt(shareG, 2)}</td><td class="num">${fmt(aux_gra)}</td><td class="num hi">${fmt(gra_fact)}</td></tr>
</tbody></table>`;

  const refPrevM = calcularReferenciaPromedio3Meses(wbProd, fechaCorte);
  const idomScoreM: number | null = (() => {
    if (!refPrevM || !refPrevM.R_ref || !refPrevM.IC_ref || !fuelTot || !horasOperMes) return null;
    const IE_mes = rendimiento / refPrevM.R_ref;
    const IC_mes = (tot_gen / horasOperMes) / P_INST_EFECTIVA;
    return 0.4 * IE_mes + 0.3 * 1 + 0.3 * (IC_mes / refPrevM.IC_ref);
  })();

  html += seccion(6, "Síntesis Operativa");
  html += buildSintesisOperativa({
    u1Horas: u1_mes,
    u2Horas: u2_mes,
    shareG1,
    shareG2,
    fp_eff,
    idomScore: idomScoreM,
    aut_hfo: 0,
    aut_do: 0,
    u1_rest: 9999,
    u2_rest: 9999,
    esInformeMensual: true,
  });

  return html;
}

// ========= INFORME DE FACTURACIÓN =========

export function generarInformeFacturacion(
  prodBuffer: ArrayBuffer,
  mesStr: string,
  diasFallaU1: number,
  diasFallaU2: number,
  costoCombTransporte?: number
): string {
  const partes = mesStr.split("-");
  if (partes.length !== 2) throw new Error("Formato de mes inválido.");
  const year = parseInt(partes[0], 10);
  const monthIndex = parseInt(partes[1], 10) - 1;
  if (isNaN(year) || isNaN(monthIndex)) throw new Error("Mes inválido.");

  const diasMes = getDaysInMonth(year, monthIndex);
  const fechaCorte = new Date(year, monthIndex + 1, 0);
  fechaCorte.setHours(0, 0, 0, 0);

  const wbProd = XLSX.read(prodBuffer, { type: "array" });
  const { rows } = getProdSheetAndRows(wbProd, fechaCorte);

  let lan = 0, gra = 0, aux = 0;
  let ultimoDia = 0;

  for (const r of rows) {
    const key = excelDateKey(r[CONFIG.COL_FECHA]);
    if (!key) continue;
    const d = new Date(key + "T00:00:00");
    if (d.getMonth() !== monthIndex || d.getFullYear() !== year) continue;
    if (d.getTime() > fechaCorte.getTime()) continue;
    if (d.getDate() > ultimoDia) ultimoDia = d.getDate();

    lan += posNum(r[CONFIG.COL_LANEC_PARCIAL_KWH]);
    gra += posNum(r[CONFIG.COL_GRACA_PARCIAL_KWH]);
    aux += posNum(r[CONFIG.COL_AUX_KWH]);
  }

  const tot_cli = lan + gra;
  const tot_gen = tot_cli + aux;
  const aux_lan = tot_cli > 0 ? aux * (lan / tot_cli) : 0;
  const aux_gra = tot_cli > 0 ? aux * (gra / tot_cli) : 0;
  const lan_fact = lan + aux_lan;
  const gra_fact = gra + aux_gra;

  const mesNombre = getMesNombreES(monthIndex);
  const textoPeriodo = ultimoDia > 0 ? `${mesNombre} ${year} (hasta el día ${ultimoDia})` : `${mesNombre} ${year}`;

  const dispU1 = Math.max(0, (diasMes - diasFallaU1) / diasMes);
  const dispU2 = Math.max(0, (diasMes - diasFallaU2) / diasMes);
  const fijoU1 = COSTO_FIJO_MENSUAL_POR_UNIDAD * dispU1;
  const fijoU2 = COSTO_FIJO_MENSUAL_POR_UNIDAD * dispU2;
  const fijoTotal = fijoU1 + fijoU2;

  const factorContratoLan = P_CONTR_TOT > 0 ? P_CONTR_LANEC / P_CONTR_TOT : 0;
  const factorContratoGra = P_CONTR_TOT > 0 ? P_CONTR_GRACA / P_CONTR_TOT : 0;

  const fijoLanU1 = fijoU1 * factorContratoLan;
  const fijoLanU2 = fijoU2 * factorContratoLan;
  const fijoGraU1 = fijoU1 * factorContratoGra;
  const fijoGraU2 = fijoU2 * factorContratoGra;
  const fijoLan = fijoLanU1 + fijoLanU2;
  const fijoGra = fijoGraU1 + fijoGraU2;

  const costosEfectivos = { ...COSTOS_VARIABLES };
  if (costoCombTransporte !== undefined) {
    costosEfectivos.combustible_transporte = costoCombTransporte;
  }
  const costoVarTotalEfectivo = Object.values(costosEfectivos).reduce((a, b) => a + b, 0);

  function subtotalVariable(kwh: number): Record<string, number> {
    const subt: Record<string, number> = {};
    for (const [k, v] of Object.entries(costosEfectivos)) { subt[k] = kwh * v; }
    return subt;
  }

  const varLanBy = subtotalVariable(lan_fact);
  const varGraBy = subtotalVariable(gra_fact);
  const varTotBy = subtotalVariable(tot_gen);

  const varLanTotal = lan_fact * costoVarTotalEfectivo;
  const varGraTotal = gra_fact * costoVarTotalEfectivo;
  const varTotTotal = tot_gen * costoVarTotalEfectivo;

  const totalLan = varLanTotal + fijoLan;
  const totalGra = varGraTotal + fijoGra;
  const totalTot = totalLan + totalGra;

  const precioLan = lan_fact > 0 ? totalLan / lan_fact : 0;
  const precioGra = gra_fact > 0 ? totalGra / gra_fact : 0;
  const precioTot = tot_gen > 0 ? totalTot / tot_gen : 0;

  const fijoTotU1 = fijoLanU1 + fijoGraU1;
  const fijoTotU2 = fijoLanU2 + fijoGraU2;
  const fijoTot = fijoLan + fijoGra;

  function tablaCliente(
    secLabel: string, titulo: string, nombre: string,
    energiaConsumida: number, auxAsig: number, totalFact: number,
    varBy: Record<string, number>, varTotal: number,
    fijoAsigU1: number, fijoAsigU2: number, fijoAsig: number,
    totalUSD: number, precioFinal: number
  ): string {
    const energiaLabel = nombre === "TOTAL" ? "Energía consumida total (LANEC + GRACA)" : `Energía consumida – ${nombre}`;
    const auxLabel = "Auxiliares asignados (proporcional)";
    const totalLabel = nombre === "TOTAL" ? "Energía total a facturar (+auxiliares)" : `Total facturable ${nombre} (+aux.)`;
    return `
<div class="rpt-section-title"><span class="rpt-section-num">${secLabel}</span>${titulo}</div>
<table class="data-table">
<thead><tr><th>Rubro</th><th>P. Unit [USD/kWh]</th><th>Subtotal [USD]</th></tr></thead>
<tbody>
<tr class="rpt-row-grupo"><td class="label" colspan="3">Energía facturable</td></tr>
<tr><td class="label">${energiaLabel}</td><td>—</td><td class="num">${fmt(energiaConsumida)} kWh</td></tr>
<tr><td class="label">${auxLabel}</td><td>—</td><td class="num">${fmt(auxAsig)} kWh</td></tr>
<tr class="rpt-row-total"><td class="label">${totalLabel}</td><td>—</td><td class="num hi">${fmt(totalFact)} kWh</td></tr>
<tr class="rpt-row-grupo"><td class="label" colspan="3">Costos variables de producción</td></tr>
<tr><td class="label">Combustible + Transporte</td><td class="num">${fmt(costosEfectivos.combustible_transporte, 4)}</td><td class="num">$ ${fmt(varBy.combustible_transporte)}</td></tr>
<tr><td class="label">Lubricantes + Químicos</td><td class="num">${fmt(costosEfectivos.lubricantes_quimicos, 4)}</td><td class="num">$ ${fmt(varBy.lubricantes_quimicos)}</td></tr>
<tr><td class="label">Agua + Insumos</td><td class="num">${fmt(costosEfectivos.agua_insumos, 4)}</td><td class="num">$ ${fmt(varBy.agua_insumos)}</td></tr>
<tr><td class="label">Repuestos Mantenimiento Predictivo</td><td class="num">${fmt(costosEfectivos.repuestos_predictivo, 4)}</td><td class="num">$ ${fmt(varBy.repuestos_predictivo)}</td></tr>
<tr><td class="label">Impacto Ambiental</td><td class="num">${fmt(costosEfectivos.impacto_ambiental, 4)}</td><td class="num">$ ${fmt(varBy.impacto_ambiental)}</td></tr>
<tr><td class="label">Servicios Auxiliares</td><td class="num">${fmt(costosEfectivos.servicios_auxiliares, 4)}</td><td class="num">$ ${fmt(varBy.servicios_auxiliares)}</td></tr>
<tr><td class="label">Margen Variable</td><td class="num">${fmt(costosEfectivos.margen_variable, 4)}</td><td class="num">$ ${fmt(varBy.margen_variable)}</td></tr>
<tr class="rpt-row-total"><td class="label"><strong>Subtotal costo variable</strong></td><td class="num"><strong>${fmt(costoVarTotalEfectivo, 4)}</strong></td><td class="num"><strong>$ ${fmt(varTotal)}</strong></td></tr>
<tr class="rpt-row-grupo"><td class="label" colspan="3">Costos fijos (por disponibilidad)</td></tr>
<tr><td class="label">Costo fijo asignado U1</td><td>—</td><td class="num">$ ${fmt(fijoAsigU1)}</td></tr>
<tr><td class="label">Costo fijo asignado U2</td><td>—</td><td class="num">$ ${fmt(fijoAsigU2)}</td></tr>
<tr class="rpt-row-total"><td class="label"><strong>Subtotal costo fijo asignado</strong></td><td>—</td><td class="num"><strong>$ ${fmt(fijoAsig)}</strong></td></tr>
<tr class="rpt-row-grand"><td class="label"><strong>TOTAL A FACTURAR</strong></td><td class="num"><strong>USD/kWh: ${fmt(precioFinal, 4)}</strong></td><td class="num"><strong>$ ${fmt(totalUSD)} + IVA</strong></td></tr>
</tbody></table>`;
  }

  let html = rptHeader("Informe de Facturación de Energía", textoPeriodo);

  html += seccion(1, "Resumen de Energía Facturable");
  html += `<table class="data-table">
<thead><tr>
<th>Cliente</th><th>Energía consumida [kWh]</th><th>Auxiliares asignados [kWh]</th><th>Total facturable [kWh]</th>
</tr></thead>
<tbody>
<tr><td class="label">LANEC</td><td class="num">${fmt(lan)}</td><td class="num">${fmt(aux_lan)}</td><td class="num hi">${fmt(lan_fact)}</td></tr>
<tr><td class="label">GRACA</td><td class="num">${fmt(gra)}</td><td class="num">${fmt(aux_gra)}</td><td class="num hi">${fmt(gra_fact)}</td></tr>
<tr class="rpt-row-total"><td class="label"><strong>TOTAL</strong></td><td class="num"><strong>${fmt(tot_cli)}</strong></td><td class="num"><strong>${fmt(aux)}</strong></td><td class="num hi"><strong>${fmt(tot_gen)}</strong></td></tr>
</tbody></table>`;

  html += tablaCliente("2.0", "Costos del Mes — Totales", "TOTAL",
    tot_cli, aux, tot_gen, varTotBy, varTotTotal, fijoTotU1, fijoTotU2, fijoTot, totalTot, precioTot);
  html += tablaCliente("2.1", "Costos del Mes — LANEC", "LANEC",
    lan, aux_lan, lan_fact, varLanBy, varLanTotal, fijoLanU1, fijoLanU2, fijoLan, totalLan, precioLan);
  html += tablaCliente("2.2", "Costos del Mes — GRACA", "GRACA",
    gra, aux_gra, gra_fact, varGraBy, varGraTotal, fijoGraU1, fijoGraU2, fijoGra, totalGra, precioGra);

  html += seccion(3, "Costo Fijo por Disponibilidad (Auditable)");
  html += `<table class="data-table">
<thead><tr>
<th>Unidad</th><th>Días mes</th><th>Días indisp.</th><th>Factor disp.</th><th>CF base [USD]</th><th>CF ajustado [USD]</th>
</tr></thead>
<tbody>
<tr><td class="label">Unidad 1</td><td class="num">${diasMes}</td><td class="num">${diasFallaU1}</td><td class="num">${fmt(dispU1, 4)}</td><td class="num">${fmt(COSTO_FIJO_MENSUAL_POR_UNIDAD)}</td><td class="num hi">${fmt(fijoU1)}</td></tr>
<tr><td class="label">Unidad 2</td><td class="num">${diasMes}</td><td class="num">${diasFallaU2}</td><td class="num">${fmt(dispU2, 4)}</td><td class="num">${fmt(COSTO_FIJO_MENSUAL_POR_UNIDAD)}</td><td class="num hi">${fmt(fijoU2)}</td></tr>
<tr class="rpt-row-total"><td class="label"><strong>TOTAL</strong></td><td colspan="4"></td><td class="num hi"><strong>${fmt(fijoTotal)}</strong></td></tr>
</tbody></table>`;

  html += seccion(4, "Asignación del Costo Fijo a Clientes (Factor Contrato)");
  html += `<table class="data-table">
<thead><tr>
<th>Cliente</th><th>kW contratados</th><th>Factor contrato</th>
<th>CF U1 [USD]</th><th>CF U2 [USD]</th><th>CF total asignado [USD]</th>
</tr></thead>
<tbody>
<tr><td class="label">LANEC</td><td class="num">${fmt(P_CONTR_LANEC, 0)}</td><td class="num">${fmt(factorContratoLan * 100, 2)} %</td><td class="num">${fmt(fijoLanU1)}</td><td class="num">${fmt(fijoLanU2)}</td><td class="num hi">${fmt(fijoLan)}</td></tr>
<tr><td class="label">GRACA</td><td class="num">${fmt(P_CONTR_GRACA, 0)}</td><td class="num">${fmt(factorContratoGra * 100, 2)} %</td><td class="num">${fmt(fijoGraU1)}</td><td class="num">${fmt(fijoGraU2)}</td><td class="num hi">${fmt(fijoGra)}</td></tr>
<tr class="rpt-row-total"><td class="label"><strong>TOTAL</strong></td><td></td><td></td><td class="num">${fmt(fijoU1)}</td><td class="num">${fmt(fijoU2)}</td><td class="num hi"><strong>${fmt(fijoTotal)}</strong></td></tr>
</tbody></table>`;

  return html;
}
