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
    const dt = new Date(Math.round((v - 25569) * 86400 * 1000));
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
  return `<div class="rpt-header">
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
  return `<div class="rpt-section-title"><span class="rpt-section-num">${n}</span>${titulo}</div>`;
}

// ========= ANÁLISIS EJECUTIVO DE COMBUSTIBLE =========

interface FuelMetric {
  date: Date;
  kWh: number;
  hfo: number;
  dsl: number;
  fuel: number;
  pctDO: number;
  gal_h: number;
  horasOp: number;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function buildFuelMetricFromRow(row: any[]): FuelMetric | null {
  const d = parseFechaRobusta(row[CONFIG.COL_FECHA]);
  if (!d) return null;
  const aux = posNum(row[CONFIG.COL_AUX_KWH]);
  const lan = posNum(row[CONFIG.COL_LANEC_PARCIAL_KWH]);
  const gra = posNum(row[CONFIG.COL_GRACA_PARCIAL_KWH]);
  const kWh = aux + lan + gra;
  const hfo = posNum(row[CONFIG.COL_HFO_GAL]);
  const dsl = posNum(row[CONFIG.COL_DO_GAL]);
  const fuel = hfo + dsl;
  const h1 = Math.max(0, posNum(row[CONFIG.COL_HG1_LANEC_FIN]) - posNum(row[CONFIG.COL_HG1_LANEC_INI]));
  const h2 = Math.max(0, posNum(row[CONFIG.COL_HG2_LANEC_FIN]) - posNum(row[CONFIG.COL_HG2_LANEC_INI]));
  const horasOp = Math.max(h1, h2);
  if (!(kWh > 0 && fuel > 0 && horasOp > 0)) return null;
  return {
    date: new Date(d.getFullYear(), d.getMonth(), d.getDate()),
    kWh, hfo, dsl, fuel,
    pctDO: dsl / fuel,
    gal_h: fuel / horasOp,
    horasOp,
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
    const galh_ref = meanSafe(win90.map(d => d.gal_h));
    const pctDO_ref = meanSafe(win90.map(d => d.pctDO));
    const fmtPct = (x: number) => Number.isFinite(x) ? (x * 100).toFixed(1) + "%" : "—";
    const fmt1 = (x: number) => Number.isFinite(x) ? x.toFixed(1) : "—";
    const fmt0 = (x: number) => Number.isFinite(x) ? x.toFixed(0) : "—";

    if (String(mode).toLowerCase() !== "monthly") {
      const key = jsDateKey(fechaJS);
      let today: FuelMetric | null = null;
      for (let i = metrics.length - 1; i >= 0; i--) {
        if (jsDateKey(metrics[i].date) === key) { today = metrics[i]; break; }
        if (metrics[i].date <= fechaJS && !today) today = metrics[i];
      }
      if (!today) return `<div class="rpt-notice">Análisis Ejecutivo de Combustible: No se encontró un día válido para análisis.</div>`;

      const delta_galh = today.gal_h - galh_ref;
      const delta_gal_dia = Number.isFinite(delta_galh) ? Math.max(0, delta_galh) * today.horasOp : NaN;
      const m = today.date.getMonth(), y = today.date.getFullYear();
      const end = new Date(fechaJS.getFullYear(), fechaJS.getMonth(), fechaJS.getDate());
      let delta_mes = 0;
      for (const d of metrics) {
        if (d.date.getFullYear() === y && d.date.getMonth() === m && d.date <= end) {
          const dif = d.gal_h - galh_ref;
          if (dif > 0) delta_mes += dif * d.horasOp;
        }
      }

      let estado = "NORMAL";
      if (today.pctDO > pctDO_ref + 0.20 && today.gal_h > galh_ref * 1.20) { estado = "CRÍTICO"; }
      else if (today.pctDO > pctDO_ref + 0.10 || today.gal_h > galh_ref * 1.10) { estado = "ALERTA OPERATIVA"; }

      let causa = "Operación dentro de los parámetros de referencia.";
      if (estado !== "NORMAL") {
        if (today.pctDO > pctDO_ref + 0.10) causa = "Mayor uso de Diésel respecto a la referencia 90D (evento operacional).";
        else causa = "Consumo por hora elevado respecto a la referencia 90D (revisar operación).";
      }

      return `<div class="rpt-fuel-box">
  <div class="rpt-fuel-header">
    <span class="rpt-fuel-title">Análisis Ejecutivo de Combustible</span>
  </div>
  <p class="rpt-fuel-causa">${causa}</p>
  <table class="data-table">
    <thead><tr>
      <th>Indicador</th>
      <th>Día del informe</th>
      <th>Referencia 90D</th>
    </tr></thead>
    <tbody>
      <tr><td class="label">% Diésel</td><td>${fmtPct(today.pctDO)}</td><td>${fmtPct(pctDO_ref)}</td></tr>
      <tr><td class="label">Consumo (gal/h)</td><td>${fmt1(today.gal_h)}</td><td>${fmt1(galh_ref)}</td></tr>
      <tr><td class="label">Sobreconsumo estimado (gal/h)</td><td>${fmt1(Math.max(0, delta_galh))}</td><td>—</td></tr>
      <tr><td class="label">Sobreconsumo día (gal)</td><td>${fmt0(delta_gal_dia)}</td><td>—</td></tr>
      <tr><td class="label">Sobreconsumo acumulado mes (gal)</td><td>${fmt0(delta_mes)}</td><td>—</td></tr>
    </tbody>
  </table>
</div>`;
    }

    // MENSUAL
    const end = new Date(fechaJS.getFullYear(), fechaJS.getMonth(), fechaJS.getDate());
    const y = end.getFullYear(), m = end.getMonth();
    let sumFuel = 0, sumDsl = 0, sumHoras = 0;
    for (const d of metrics) {
      if (d.date.getFullYear() === y && d.date.getMonth() === m && d.date <= end) {
        sumFuel += d.fuel; sumDsl += d.dsl; sumHoras += d.horasOp;
      }
    }
    if (!(sumFuel > 0 && sumHoras > 0)) {
      return `<div class="rpt-notice">Análisis Ejecutivo de Combustible: Sin datos suficientes del mes.</div>`;
    }
    const pctDO_mes = sumDsl / sumFuel;
    const galh_mes = sumFuel / sumHoras;
    const delta_galh = galh_mes - galh_ref;
    const delta_periodo = Math.max(0, delta_galh) * sumHoras;

    let estado = "NORMAL";
    if (pctDO_mes > pctDO_ref + 0.20 && galh_mes > galh_ref * 1.20) { estado = "CRÍTICO"; }
    else if (pctDO_mes > pctDO_ref + 0.10 || galh_mes > galh_ref * 1.10) { estado = "ALERTA OPERATIVA"; }

    let causa = "Operación dentro de los parámetros de referencia.";
    if (estado !== "NORMAL") {
      if (pctDO_mes > pctDO_ref + 0.10) causa = "Mayor uso de Diésel en el mes respecto a la referencia 90D.";
      else causa = "Consumo por hora del mes elevado respecto a la referencia 90D.";
    }

    return `<div class="rpt-fuel-box">
  <div class="rpt-fuel-header">
    <span class="rpt-fuel-title">Análisis Ejecutivo de Combustible (Mensual)</span>
  </div>
  <p class="rpt-fuel-causa">${causa}</p>
  <table class="data-table">
    <thead><tr>
      <th>Indicador</th>
      <th>Mes (acumulado)</th>
      <th>Referencia 90D</th>
    </tr></thead>
    <tbody>
      <tr><td class="label">% Diésel</td><td>${fmtPct(pctDO_mes)}</td><td>${fmtPct(pctDO_ref)}</td></tr>
      <tr><td class="label">Consumo (gal/h)</td><td>${fmt1(galh_mes)}</td><td>${fmt1(galh_ref)}</td></tr>
      <tr><td class="label">Sobreconsumo periodo (gal)</td><td>${fmt0(delta_periodo)}</td><td>—</td></tr>
    </tbody>
  </table>
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

  html += seccion(3, "Combustible y Eficiencia");
  html += `<table class="data-table"><thead><tr>
<th>Combustible</th><th>Consumo [gal]</th></tr></thead><tbody>
<tr><td class="label">HFO (Fuel Oil Pesado)</td><td class="num">${fmt(hfo)}</td></tr>
<tr><td class="label">Diésel (DO)</td><td class="num">${fmt(dsl)}</td></tr>
<tr><td class="label">Total equivalente</td><td class="num hi">${fmt(fuelTot)}</td></tr>
</tbody></table>
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

  html += seccion(6, "Indicador de Desempeño Operacional (IDOM)");
  if (idomDia) {
    html += `<table class="data-table"><thead><tr><th>Parámetro</th><th>Valor</th></tr></thead><tbody>
<tr><td class="label">Referencia (prom. móvil 3 meses)</td><td>${refPrev.mesRef} <span class="rpt-muted">(${refPrev.diasConDato} días con dato)</span></td></tr>
<tr><td class="label">Rendimiento de referencia (R_ref)</td><td class="num">${fmt(refPrev.R_ref, 2)} kWh/gal</td></tr>
<tr><td class="label">Índice de carga de referencia (IC_ref)</td><td class="num">${fmt(refPrev.IC_ref, 3)}</td></tr>
<tr><td class="label">Rendimiento del día (R_día)</td><td class="num">${fmt(idomDia.R_dia, 2)} kWh/gal</td></tr>
<tr><td class="label">Índice de eficiencia (IE)</td><td class="num">${fmt(idomDia.IE, 3)}</td></tr>
<tr><td class="label">Disponibilidad diaria (ID)</td><td class="num">${fmt(idomDia.ID, 2)}</td></tr>
<tr><td class="label">Índice de carga del día (IC_día)</td><td class="num">${fmt(idomDia.IC_dia, 3)}</td></tr>
<tr class="rpt-row-total"><td class="label"><strong>IDOM_D</strong></td><td class="num"><strong>${fmt(idomDia.IDOM, 4)}</strong></td></tr>
<tr><td class="label">Estado operacional</td><td>${badgeEstado(idomDia.estado)}</td></tr>
<tr><td class="label">Factor limitante principal</td><td><strong>${idomDia.driver}</strong></td></tr>
<tr><td class="label">Pérdida energética vs referencia</td><td class="num">${fmt(idomDia.Loss_kWh, 0)} kWh</td></tr>
</tbody></table>`;
    html += notaPieIDOM(refPrev);
  } else {
    html += `<div class="rpt-notice">No se pudo calcular IDOM: verifique datos completos del día y de la referencia (promedio móvil 3 meses).</div>`;
  }

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

  html += `<div class="rpt-kpi-row">
  <div class="rpt-kpi-card"><div class="rpt-kpi-label">Factor de planta (vs instalada)</div><div class="rpt-kpi-big">${(fp_inst * 100).toFixed(1)}<span class="rpt-kpi-unit">%</span></div><div class="rpt-kpi-sub">${P_INST_TOTAL} kW instalados</div></div>
  <div class="rpt-kpi-card"><div class="rpt-kpi-label">Factor de planta (vs efectiva)</div><div class="rpt-kpi-big">${(fp_eff * 100).toFixed(1)}<span class="rpt-kpi-unit">%</span></div><div class="rpt-kpi-sub">${Math.round(P_INST_EFECTIVA)} kW efectivos</div></div>
  <div class="rpt-kpi-card"><div class="rpt-kpi-label">Rendimiento promedio</div><div class="rpt-kpi-big">${fmt(rendimiento, 2)}<span class="rpt-kpi-unit">kWh/gal</span></div><div class="rpt-kpi-sub">Energía por galón consumido</div></div>
</div>`;

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
<th>Combustible</th><th>Consumo mensual [gal]</th></tr></thead><tbody>
<tr><td class="label">HFO (Fuel Oil Pesado)</td><td class="num">${fmt(hfo)}</td></tr>
<tr><td class="label">Diésel (DO)</td><td class="num">${fmt(dsl)}</td></tr>
<tr><td class="label">Total equivalente</td><td class="num hi">${fmt(fuelTot)}</td></tr>
</tbody></table>`;
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
  html += seccion(6, "Indicador de Desempeño Operacional (IDOM)");
  if (refPrevM && refPrevM.R_ref > 0 && refPrevM.IC_ref > 0 && fuelTot > 0 && horasOperMes > 0) {
    const R_mes = rendimiento;
    const IE_mes = R_mes / refPrevM.R_ref;
    const IC_mes = (tot_gen / horasOperMes) / P_INST_EFECTIVA;
    const ID_mes = 1;
    const IDOM_M = 0.4 * IE_mes + 0.3 * ID_mes + 0.3 * (IC_mes / refPrevM.IC_ref);
    const estadoM = semaforo(IDOM_M);
    const lossM = Math.max(0, fuelTot * refPrevM.R_ref - tot_gen);

    html += `<table class="data-table"><thead><tr><th>Parámetro</th><th>Valor</th></tr></thead><tbody>
<tr><td class="label">Referencia (prom. móvil 3 meses)</td><td>${refPrevM.mesRef} <span class="rpt-muted">(${refPrevM.diasConDato} días)</span></td></tr>
<tr><td class="label">Rendimiento de referencia (R_ref)</td><td class="num">${fmt(refPrevM.R_ref, 2)} kWh/gal</td></tr>
<tr><td class="label">Índice de carga de referencia (IC_ref)</td><td class="num">${fmt(refPrevM.IC_ref, 3)}</td></tr>
<tr><td class="label">Rendimiento del mes (R_mes)</td><td class="num">${fmt(R_mes, 2)} kWh/gal</td></tr>
<tr><td class="label">Índice de eficiencia mensual (IE)</td><td class="num">${fmt(IE_mes, 3)}</td></tr>
<tr><td class="label">Índice de carga mensual (IC_mes)</td><td class="num">${fmt(IC_mes, 3)}</td></tr>
<tr class="rpt-row-total"><td class="label"><strong>IDOM_M</strong></td><td class="num"><strong>${fmt(IDOM_M, 4)}</strong></td></tr>
<tr><td class="label">Estado operacional</td><td>${badgeEstado(estadoM)}</td></tr>
<tr><td class="label">Pérdida energética vs referencia</td><td class="num">${fmt(lossM, 0)} kWh</td></tr>
</tbody></table>`;
    html += notaPieIDOM(refPrevM);
  } else {
    html += `<div class="rpt-notice">No se pudo calcular IDOM_M: verifique datos del mes y de la referencia (promedio móvil 3 meses).</div>`;
  }

  return html;
}

// ========= INFORME DE FACTURACIÓN =========

export function generarInformeFacturacion(
  prodBuffer: ArrayBuffer,
  mesStr: string,
  diasFallaU1: number,
  diasFallaU2: number
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

  function subtotalVariable(kwh: number): Record<string, number> {
    const subt: Record<string, number> = {};
    for (const [k, v] of Object.entries(COSTOS_VARIABLES)) { subt[k] = kwh * v; }
    return subt;
  }

  const varLanBy = subtotalVariable(lan_fact);
  const varGraBy = subtotalVariable(gra_fact);
  const varTotBy = subtotalVariable(tot_gen);

  const varLanTotal = lan_fact * COSTO_VARIABLE_TOTAL;
  const varGraTotal = gra_fact * COSTO_VARIABLE_TOTAL;
  const varTotTotal = tot_gen * COSTO_VARIABLE_TOTAL;

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
<tr><td class="label">Combustible + Transporte</td><td class="num">${fmt(COSTOS_VARIABLES.combustible_transporte, 4)}</td><td class="num">$ ${fmt(varBy.combustible_transporte)}</td></tr>
<tr><td class="label">Lubricantes + Químicos</td><td class="num">${fmt(COSTOS_VARIABLES.lubricantes_quimicos, 4)}</td><td class="num">$ ${fmt(varBy.lubricantes_quimicos)}</td></tr>
<tr><td class="label">Agua + Insumos</td><td class="num">${fmt(COSTOS_VARIABLES.agua_insumos, 4)}</td><td class="num">$ ${fmt(varBy.agua_insumos)}</td></tr>
<tr><td class="label">Repuestos Mantenimiento Predictivo</td><td class="num">${fmt(COSTOS_VARIABLES.repuestos_predictivo, 4)}</td><td class="num">$ ${fmt(varBy.repuestos_predictivo)}</td></tr>
<tr><td class="label">Impacto Ambiental</td><td class="num">${fmt(COSTOS_VARIABLES.impacto_ambiental, 4)}</td><td class="num">$ ${fmt(varBy.impacto_ambiental)}</td></tr>
<tr><td class="label">Servicios Auxiliares</td><td class="num">${fmt(COSTOS_VARIABLES.servicios_auxiliares, 4)}</td><td class="num">$ ${fmt(varBy.servicios_auxiliares)}</td></tr>
<tr><td class="label">Margen Variable</td><td class="num">${fmt(COSTOS_VARIABLES.margen_variable, 4)}</td><td class="num">$ ${fmt(varBy.margen_variable)}</td></tr>
<tr class="rpt-row-total"><td class="label"><strong>Subtotal costo variable</strong></td><td class="num"><strong>${fmt(COSTO_VARIABLE_TOTAL, 4)}</strong></td><td class="num"><strong>$ ${fmt(varTotal)}</strong></td></tr>
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
