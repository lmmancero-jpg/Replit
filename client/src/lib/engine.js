console.log("App Central El Morro cargada.");

// === CONFIGURACIÓN DE COLUMNAS (0 = A, 1 = B, ...) ===
const CONFIG = {
  COL_FECHA: 1,            // B (Fecha)

  // Horómetros solo LANEC
  COL_HG1_LANEC_INI: 2,    // C
  COL_HG1_LANEC_FIN: 3,    // D
  COL_HG2_LANEC_INI: 8,    // I
  COL_HG2_LANEC_FIN: 9,    // J

  // Energía (kWh)
  COL_AUX_KWH:            24, // Y
  COL_LANEC_PARCIAL_KWH:  25, // Z
  COL_GRACA_PARCIAL_KWH:  30, // AE
  COL_TOTAL_LG_KWH:       35, // AJ
  COL_GEN1_KWH:           36, // AK
  COL_GEN2_KWH:           37, // AL

  // Combustibles (gal)
  COL_HFO_GAL:            43,
  COL_DO_GAL:             46,

  // Stocks
  COL_STOCK_HFO_TOTAL:    49,
  COL_STOCK_DO_TOTAL:     51
};

// ========================= FECHA ROBUSTA (Excel serial / dd/mm/yy / ISO) =========================
function parseFechaRobusta(v){
  if (v === null || v === undefined || v === "") return null;

  if (v instanceof Date){
    return isNaN(v) ? null : v;
  }

  // Excel serial (SheetJS suele entregar números)
  if (typeof v === "number" && isFinite(v)){
    const dt = new Date(Math.round((v - 25569) * 86400 * 1000));
    return isNaN(dt) ? null : dt;
  }

  if (typeof v === "string"){
    const s = v.trim();

    // ISO yyyy-mm-dd
    const iso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (iso){
      const y = +iso[1], m = +iso[2], d = +iso[3];
      const dt = new Date(y, m-1, d);
      return isNaN(dt) ? null : dt;
    }

    // dd/mm/yy o mm/dd/yy (heurística: por defecto dd/mm para Ecuador)
    const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2}|\d{4})$/);
    if (m){
      let a = +m[1], b = +m[2], y = +m[3];
      if (y < 100) y += 2000;

      let day, mon;
      if (a > 12 && b <= 12){
        // 13/03/26 -> dd/mm
        day = a; mon = b;
      } else if (b > 12 && a <= 12){
        // 03/13/26 -> mm/dd
        day = b; mon = a;
      } else {
        // Ambiguo (1/3/26): asumir dd/mm
        day = a; mon = b;
      }
      const dt = new Date(y, mon-1, day);
      return isNaN(dt) ? null : dt;
    }

    // Último intento: Date.parse
    const dt2 = new Date(s);
    return isNaN(dt2) ? null : dt2;
  }

  return null;
}



// Horómetros base
const HORO_BASE_U1 = 0;
const HORO_BASE_U2 = 21041;
const OBJ_MTO_HORAS_U1 = 8000;
const OBJ_MTO_HORAS_U2 = 6000;

// ====== COSTOS PARA FACTURACIÓN (USD/kWh y USD/mes) ======
const COSTOS_VARIABLES = {
  combustible_transporte: 0.1153,
  lubricantes_quimicos:  0.0182,
  agua_insumos:          0.0070,
  repuestos_predictivo:  0.0299,
  impacto_ambiental:     0.0029,
  servicios_auxiliares:  0.0034,
  margen_variable:       0.0138
};

const COSTO_VARIABLE_TOTAL = Object.values(COSTOS_VARIABLES).reduce((a,b)=>a+b,0); // 0.1956
const COSTO_FIJO_MENSUAL_POR_UNIDAD = 30720; // USD/mes/unidad


let prodArrayBuffer = null;
let aforoArrayBuffer = null;

// ===== Manejo de archivos =====
document.getElementById("fileProd").addEventListener("change", e => {
  const f = e.target.files[0];
  if (!f) return;
  const r = new FileReader();
  r.onload = ev => prodArrayBuffer = ev.target.result;
  r.readAsArrayBuffer(f);
});

document.getElementById("fileAforo").addEventListener("change", e => {
  const f = e.target.files[0];
  if (!f) return;
  const r = new FileReader();
  r.onload = ev => aforoArrayBuffer = ev.target.result;
  r.readAsArrayBuffer(f);
});

function showError(msg) {
  document.getElementById("msg").textContent = msg || "";
}

// === Utilidades ===
function num(v) {
  if (v === null || v === undefined) return 0;
  if (typeof v === "number") return v;
  if (typeof v === "string") {
    const s = v.replace(/\./g, "").replace(",", ".");
    const n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  }
  return 0;
}

// Reemplaza valores negativos por 0
function posNum(v) {
  const n = num(v);
  return n < 0 ? 0 : n;
}

// =================== ANÁLISIS EJECUTIVO DE COMBUSTIBLE (GERENCIA) ===================
// Siempre usa los últimos 30 días con data completa, cruzando todas las hojas del archivo de producción.
// Salida: estado (NORMAL/ALERTA/CRÍTICO), causa simple y sobreconsumos en gal/h, gal/día y gal mes en curso (sin precios).

function buildFuelMetricFromRow(row){
  const d = parseFechaRobusta(row[CONFIG.COL_FECHA]);
  if(!d) return null;

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

  if(!(kWh>0 && fuel>0 && horasOp>0)) return null;

  return {
    date: new Date(d.getFullYear(), d.getMonth(), d.getDate()),
    kWh,
    hfo, dsl, fuel,
    pctDO: dsl / fuel,
    gal_h: fuel / horasOp,
    horasOp
  };
}

function meanSafe(arr){
  const a = arr.filter(v=>Number.isFinite(v));
  if(!a.length) return NaN;
  return a.reduce((s,x)=>s+x,0)/a.length;
}

function getAllFuelMetricsFromWorkbook(wbProd){
  const out = [];
  for (let si=0; si<wbProd.SheetNames.length; si++){
    const name = wbProd.SheetNames[si];
    const ws = wbProd.Sheets[name];
    if(!ws) continue;
    const rows = XLSX.utils.sheet_to_json(ws,{header:1, raw:true});
    for(let i=0;i<rows.length;i++){
      const m = buildFuelMetricFromRow(rows[i]);
      if(m) out.push(m);
    }
  }
  out.sort((a,b)=>a.date-b.date);
  return out;
}

function lastNDaysWithData(metrics, endDate, n){
  const end = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
  const out = [];
  for (let i=metrics.length-1; i>=0 && out.length<n; i--){
    const d = metrics[i].date;
    if(d<=end) out.push(metrics[i]);
  }
  return out.reverse();
}

function last3MonthsWithData(metrics, fechaJS){

  const end = new Date(fechaJS);
  const start = new Date(fechaJS);
  start.setMonth(start.getMonth()-3);

  const arr = [];

  for(let i=0;i<metrics.length;i++){
    const d = metrics[i].date;
    if(d >= start && d <= end){
      arr.push(metrics[i]);
    }
  }

  return arr;
}

function buildFuelExecutiveHTML(wbProd, fechaJS, mode="daily"){
  try{
    const metrics = getAllFuelMetricsFromWorkbook(wbProd);

    // Ventana de referencia: últimos 90 días con dato, hasta la fecha de corte
    const win90 = lastNDaysWithData(metrics, fechaJS, 90);

    // Validación mínima para referencias estables
    if(!win90 || win90.length < 20){
      return `
<div style="border:1px solid #999;padding:10px;margin-top:8px;">
<strong>Análisis Ejecutivo de Combustible:</strong> Información insuficiente para referencia 90D (días válidos: ${(win90||[]).length}). 
</div>`;
    }

    const galh_ref = meanSafe(win90.map(d=>d.gal_h));
    const pctDO_ref = meanSafe(win90.map(d=>d.pctDO));

    // -------------------- MODO DIARIO --------------------
    if(String(mode).toLowerCase() !== "monthly"){
      // Buscar el día del informe (si no existe, usar el último día <= fecha)
      let today = null;
      const key = jsDateKey(fechaJS);
      for(let i=metrics.length-1;i>=0;i--){
        if(jsDateKey(metrics[i].date)===key){ today = metrics[i]; break; }
        if(metrics[i].date <= fechaJS && !today) today = metrics[i];
      }

      if(!today){
        return `
<div style="border:1px solid #999;padding:10px;margin-top:8px;">
<strong>Análisis Ejecutivo de Combustible:</strong> No se encontró un día válido para análisis. 
</div>`;
      }

      const delta_galh = today.gal_h - galh_ref;
      const delta_gal_dia = Number.isFinite(delta_galh) ? Math.max(0, delta_galh) * today.horasOp : NaN;

      // Sobrec. acumulado mes en curso (solo días hasta fecha del informe)
      const m = today.date.getMonth(), y = today.date.getFullYear();
      const end = new Date(fechaJS.getFullYear(), fechaJS.getMonth(), fechaJS.getDate());
      let delta_mes = 0;
      for(let i=0;i<metrics.length;i++){
        const d=metrics[i];
        if(d.date.getFullYear()===y && d.date.getMonth()===m && d.date<=end){
          const dif = d.gal_h - galh_ref;
          if(dif>0) delta_mes += dif * d.horasOp;
        }
      }

      // Estado (simple, gerencial)
      let estado = "NORMAL", color = "#1f7a1f";
      if(today.pctDO > pctDO_ref + 0.20 && today.gal_h > galh_ref*1.20){
        estado = "CRÍTICO"; color = "#b00020";
      } else if(today.pctDO > pctDO_ref + 0.10 || today.gal_h > galh_ref*1.10){
        estado = "ALERTA OPERATIVA"; color = "#b26a00";
      }

      // Causa simplificada
      let causa = "Operación normal";
      if(estado!=="NORMAL"){
        if(today.pctDO > pctDO_ref + 0.10) causa = "Mayor uso de Diésel respecto a la referencia 90D (evento operacional).";
        else causa = "Consumo por hora elevado respecto a la referencia 90D (revisar operación).";
      }

      const fmtPct = x => Number.isFinite(x)? (x*100).toFixed(1)+"%":"—";
      const fmt1 = x => Number.isFinite(x)? x.toFixed(1):"—";
      const fmt0 = x => Number.isFinite(x)? x.toFixed(0):"—";

      return `
<div style="margin-top:8px;">
  <div style="font-weight:700;margin-bottom:6px;">Análisis Ejecutivo de Combustible</div>
  <div style="font-size:12px;margin-bottom:6px;color:#333;">
    <strong>Estado:</strong> <span style="color:${color};font-weight:700">${estado}</span> — ${causa}
  </div>
  <table style="border-collapse:collapse;width:100%;font-size:13px;">
    <thead>
      <tr>
        <th style="text-align:left;border:1px solid #cfcfcf;padding:6px;background:#f3f4f6;">Indicador</th>
        <th style="text-align:right;border:1px solid #cfcfcf;padding:6px;background:#f3f4f6;">Día del informe</th>
        <th style="text-align:right;border:1px solid #cfcfcf;padding:6px;background:#f3f4f6;">Referencia 90D</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td style="border:1px solid #cfcfcf;padding:6px;">% Diésel</td>
        <td style="border:1px solid #cfcfcf;padding:6px;text-align:right;">${fmtPct(today.pctDO)}</td>
        <td style="border:1px solid #cfcfcf;padding:6px;text-align:right;">${fmtPct(pctDO_ref)}</td>
      </tr>
      <tr>
        <td style="border:1px solid #cfcfcf;padding:6px;">Consumo (gal/h)</td>
        <td style="border:1px solid #cfcfcf;padding:6px;text-align:right;">${fmt1(today.gal_h)}</td>
        <td style="border:1px solid #cfcfcf;padding:6px;text-align:right;">${fmt1(galh_ref)}</td>
      </tr>
      <tr>
        <td style="border:1px solid #cfcfcf;padding:6px;">Sobrec. estimado (gal/h)</td>
        <td style="border:1px solid #cfcfcf;padding:6px;text-align:right;">${fmt1(Math.max(0,delta_galh))}</td>
        <td style="border:1px solid #cfcfcf;padding:6px;text-align:right;">—</td>
      </tr>
      <tr>
        <td style="border:1px solid #cfcfcf;padding:6px;">Sobrec. día (gal)</td>
        <td style="border:1px solid #cfcfcf;padding:6px;text-align:right;">${fmt0(delta_gal_dia)}</td>
        <td style="border:1px solid #cfcfcf;padding:6px;text-align:right;">—</td>
      </tr>
      <tr>
        <td style="border:1px solid #cfcfcf;padding:6px;">Sobrec. acumulado mes (gal)</td>
        <td style="border:1px solid #cfcfcf;padding:6px;text-align:right;">${fmt0(delta_mes)}</td>
        <td style="border:1px solid #cfcfcf;padding:6px;text-align:right;">—</td>
      </tr>
    </tbody>
  </table>
</div>`;
    }

    // -------------------- MODO MENSUAL --------------------
    const end = new Date(fechaJS.getFullYear(), fechaJS.getMonth(), fechaJS.getDate());
    const y = end.getFullYear(), m = end.getMonth();

    let sumFuel=0, sumDsl=0, sumHoras=0, sumKWh=0;
    for(const d of metrics){
      if(d.date.getFullYear()===y && d.date.getMonth()===m && d.date<=end){
        sumFuel += d.fuel;
        sumDsl  += d.dsl;
        sumHoras += d.horasOp;
        sumKWh  += d.kWh;
      }
    }

    if(!(sumFuel>0 && sumHoras>0)){
      return `
<div style="border:1px solid #999;padding:10px;margin-top:8px;">
<strong>Análisis Ejecutivo de Combustible:</strong> Sin datos suficientes del mes para calcular % Diésel mensual. 
</div>`;
    }

    const pctDO_mes = sumDsl / sumFuel;
    const galh_mes = sumFuel / sumHoras;

    const delta_galh = galh_mes - galh_ref;
    const delta_periodo = Math.max(0, delta_galh) * sumHoras;

    // Estado mensual (simple)
    let estado = "NORMAL", color = "#1f7a1f";
    if(pctDO_mes > pctDO_ref + 0.20 && galh_mes > galh_ref*1.20){
      estado = "CRÍTICO"; color = "#b00020";
    } else if(pctDO_mes > pctDO_ref + 0.10 || galh_mes > galh_ref*1.10){
      estado = "ALERTA OPERATIVA"; color = "#b26a00";
    }

    let causa = "Operación normal";
    if(estado!=="NORMAL"){
      if(pctDO_mes > pctDO_ref + 0.10) causa = "Mayor uso de Diésel en el mes respecto a la referencia 90D.";
      else causa = "Consumo por hora del mes elevado respecto a la referencia 90D.";
    }

    const fmtPct = x => Number.isFinite(x)? (x*100).toFixed(1)+"%":"—";
    const fmt1 = x => Number.isFinite(x)? x.toFixed(1):"—";
    const fmt0 = x => Number.isFinite(x)? x.toFixed(0):"—";

    return `
<div style="margin-top:8px;">
  <div style="font-weight:700;margin-bottom:6px;">Análisis Ejecutivo de Combustible (Mensual)</div>
  <div style="font-size:12px;margin-bottom:6px;color:#333;">
    <strong>Estado:</strong> <span style="color:${color};font-weight:700">${estado}</span> — ${causa}
  </div>
  <table style="border-collapse:collapse;width:100%;font-size:13px;">
    <thead>
      <tr>
        <th style="text-align:left;border:1px solid #cfcfcf;padding:6px;background:#f3f4f6;">Indicador</th>
        <th style="text-align:right;border:1px solid #cfcfcf;padding:6px;background:#f3f4f6;">Mes (acumulado)</th>
        <th style="text-align:right;border:1px solid #cfcfcf;padding:6px;background:#f3f4f6;">Referencia 90D</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td style="border:1px solid #cfcfcf;padding:6px;">% Diésel</td>
        <td style="border:1px solid #cfcfcf;padding:6px;text-align:right;">${fmtPct(pctDO_mes)}</td>
        <td style="border:1px solid #cfcfcf;padding:6px;text-align:right;">${fmtPct(pctDO_ref)}</td>
      </tr>
      <tr>
        <td style="border:1px solid #cfcfcf;padding:6px;">Consumo (gal/h)</td>
        <td style="border:1px solid #cfcfcf;padding:6px;text-align:right;">${fmt1(galh_mes)}</td>
        <td style="border:1px solid #cfcfcf;padding:6px;text-align:right;">${fmt1(galh_ref)}</td>
      </tr>
      <tr>
        <td style="border:1px solid #cfcfcf;padding:6px;">Sobrec. periodo (gal)</td>
        <td style="border:1px solid #cfcfcf;padding:6px;text-align:right;">${fmt0(delta_periodo)}</td>
        <td style="border:1px solid #cfcfcf;padding:6px;text-align:right;">—</td>
      </tr>
    </tbody>
  </table>
</div>`;
  }catch(e){
    console.error("FuelExecutive error:", e);
    return `<div style="border:1px solid #b00020;padding:10px;margin-top:8px;"><strong>Análisis Ejecutivo de Combustible:</strong> No disponible por error de cálculo.</div>`;
  }
}

function fmt(v, dec = 2) {
  return Number(v).toLocaleString("es-EC", {
    minimumFractionDigits: dec,
    maximumFractionDigits: dec
  });
}

function pad2(n) { return n < 10 ? "0" + n : "" + n; }

function jsDateKey(d) {
  return d.getFullYear() + "-" + pad2(d.getMonth()+1) + "-" + pad2(d.getDate());
}

function excelDateKey(v) {
  if (v == null) return null;

  if (typeof v === "number") {
    const dc = XLSX.SSF.parse_date_code(v);
    if (!dc) return null;
    return `${dc.y}-${pad2(dc.m)}-${pad2(dc.d)}`;
  }

  if (v instanceof Date) return jsDateKey(v);

  if (typeof v === "string") {
    const s = v.trim();
    const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (m) {
      let d = parseInt(m[1],10), mo = parseInt(m[2],10),
          y = parseInt(m[3],10);
      if (y < 100) y += 2000;
      return `${y}-${pad2(mo)}-${pad2(d)}`;
    }
    const d2 = new Date(s);
    if (!isNaN(d2.getTime())) return jsDateKey(d2);
  }
  return null;
}

// ======================= IDOM (Indicador de Desempeño Operacional) ==========================
const P_INST_TOTAL = 5100;                  // kW instalada
const P_INST_EFECTIVA = 0.85 * P_INST_TOTAL; // kW efectiva (85%)

// Devuelve "YYYY-MM" desde Date
function monthStrFromDate(d){
  const y = d.getFullYear();
  const m = pad2(d.getMonth()+1);
  return `${y}-${m}`;
}

// Devuelve el promedio móvil de 3 meses (meses anteriores) en formato "YYYY-MM"
function prevMonthStr(yyyyMM){
  const [yS,mS] = (yyyyMM||"").split("-");
  let y = parseInt(yS,10), m = parseInt(mS,10);
  if (!y || !m) return null;
  m -= 1;
  if (m === 0){ m = 12; y -= 1; }
  return `${y}-${pad2(m)}`;
}

// Colores para semáforo (inline para que funcione aunque no exista CSS)
function semaforo(v){
  if (v === null || v === undefined || isNaN(v)) return "SIN DATO";
  if (v >= 0.85) return "VERDE";
  if (v >= 0.75) return "AMARILLO";
  if (v >= 0.65) return "NARANJA";
  return "ROJO";
}

function semColor(estado){
  switch(String(estado||"").toUpperCase()){
    case "VERDE": return "#0a7d32";
    case "AMARILLO": return "#b8860b";
    case "NARANJA": return "#e67e00";
    case "ROJO": return "#c0392b";
    default: return "#111";
  }
}

// Calcula referencia (promedio móvil de 3 meses (meses anteriores)) desde el workbook de producción
// Retorna: {mesRef, diasConDato, R_ref, IC_ref}
function calcularReferenciaPromedio3Meses(wbProd, fechaJS){
  const mesActual = monthStrFromDate(fechaJS);
  const m1 = prevMonthStr(mesActual);
  const m2 = m1 ? prevMonthStr(m1) : null;
  const m3 = m2 ? prevMonthStr(m2) : null;

  const meses = [m1,m2,m3].filter(Boolean);
  if (!meses.length) return {mesRef:null,mesesUsados:[],diasConDato:0,R_ref:0,IC_ref:0};

  let sumKWh = 0;
  let sumGal = 0;
  let sumHoras = 0;
  const diasSet = new Set();

  for (const mesStr of meses){
    const [yS,mS] = mesStr.split("-");
    const year = parseInt(yS,10);
    const monthIndex = parseInt(mS,10)-1;

    // Fecha de corte = último día del mes en cuestión (solo para ubicar la hoja)
    const fechaCorte = new Date(year, monthIndex+1, 0);
    fechaCorte.setHours(0,0,0,0);

    const {rows} = getProdSheetAndRows(wbProd, fechaCorte);

    for (let i=0;i<rows.length;i++){
      const r = rows[i];
      const key = excelDateKey(r[CONFIG.COL_FECHA]);
      if (!key) continue;

      const d = new Date(key+"T00:00:00");
      if (d.getFullYear() !== year || d.getMonth() !== monthIndex) continue;

      const aux_kwh   = posNum(r[CONFIG.COL_AUX_KWH]);
      const lanec_kwh = posNum(r[CONFIG.COL_LANEC_PARCIAL_KWH]);
      const graca_kwh = posNum(r[CONFIG.COL_GRACA_PARCIAL_KWH]);
      const total_gen_kwh = lanec_kwh + graca_kwh + aux_kwh;

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

      if (total_gen_kwh > 0 && fuelTot > 0 && horasOperDia > 0){
        sumKWh += total_gen_kwh;
        sumGal += fuelTot;
        sumHoras += horasOperDia;
        diasSet.add(jsDateKey(d));
      }
    }
  }

  const diasConDato = diasSet.size;
  const R_ref = sumGal>0 ? (sumKWh/sumGal) : 0;
  const IC_ref = sumHoras>0 ? ((sumKWh/sumHoras)/P_INST_EFECTIVA) : 0;

  return {mesRef: meses.join(", "), mesesUsados: meses, diasConDato, R_ref, IC_ref};
}

// Calcula IDOM del día con referencia (promedio móvil 3 meses)
function calcularIDOMDia(total_gen_kwh, fuelTot, horasOperDia, u1_dia, u2_dia, ref){
  if (!(total_gen_kwh>0 && fuelTot>0 && horasOperDia>0 && ref && ref.R_ref>0 && ref.IC_ref>0)){
    return null;
  }

  const R_dia = total_gen_kwh / fuelTot; // kWh/gal
  const IE = R_dia / ref.R_ref;

  const Pavg = total_gen_kwh / horasOperDia; // kW durante operación
  const IC_dia = (Pavg / P_INST_EFECTIVA);

  const ID = ((u1_dia>0)?1:0 + (u2_dia>0)?1:0); // cuidado precedencia
  const ID_ok = (((u1_dia>0)?1:0) + ((u2_dia>0)?1:0)) / 2;

  const IDOM = 0.4*IE + 0.3*ID_ok + 0.3*(IC_dia/ref.IC_ref);

  const estado = semaforo(IDOM);

  const Loss_kWh = Math.max(0, (fuelTot*ref.R_ref) - total_gen_kwh);

  const penR = Math.max(0, ref.R_ref - R_dia);
  const penC = Math.max(0, ref.IC_ref - IC_dia);

  let driver = "RENDIMIENTO";
  if (ID_ok === 0.5) driver = "DISPONIBILIDAD";
  else driver = (penC > penR) ? "CARGA" : "RENDIMIENTO";

  return {R_dia, IE, ID: ID_ok, IC_dia, IDOM, estado, driver, Loss_kWh, ref};
}

// Pie de página explicativo (se usa en diario y mensual)
function notaPieIDOM(ref){
  if (!ref || !ref.R_ref || !ref.IC_ref) return "";
  return `
  <div style="font-size:11px; margin-top:12px; color:#333; line-height:1.35;">
    <strong>Nota técnica (KPIs IDOM):</strong><br>
    <strong>Rendimiento de referencia (kWh/gal)</strong>: R_ref = Σ(kWh generados totales)/Σ(gal HFO+Diésel) del promedio móvil de 3 meses (meses anteriores) (${ref.mesRef}, días con dato: ${ref.diasConDato}).<br>
    <strong>Índice de carga de referencia</strong>: IC_ref = (P_promedio_operación_ref / P_instalada_efectiva), con P_promedio_operación_ref = Σ(kWh)/Σ(horas de operación) y P_instalada_efectiva = 0,85×${P_INST_TOTAL} = ${Math.round(P_INST_EFECTIVA)} kW.<br>
    <strong>Rendimiento del día</strong>: R_día = kWh_generados_totales / gal_totales (HFO+Diésel).<br>
    <strong>Índice de eficiencia</strong>: IE = R_día / R_ref.<br>
    <strong>Disponibilidad diaria</strong>: ID = (U1_operó + U2_operó)/2, donde U_operó = 1 si horas>0, caso contrario 0.<br>
    <strong>Índice de carga del día</strong>: IC_día = (kWh_día / horas_operación_día) / P_instalada_efectiva; horas_operación_día = max(horas U1, horas U2).<br>
    <strong>IDOM_D</strong>: 0,4×IE + 0,3×ID + 0,3×(IC_día/IC_ref).<br>
    <strong>Semáforo</strong>: VERDE ≥ 0,85; AMARILLO 0,75–0,85; NARANJA 0,65–0,75; ROJO < 0,65.<br>
    <strong>Loss_kWh</strong>: pérdida energética vs promedio móvil de 3 meses (meses anteriores) = max(0, gal_totales×R_ref − kWh_día).
  </div>`;
}

function formatFechaLarga(str) {
  const d = new Date(str + "T00:00:00");
  if (isNaN(d.getTime())) return "";
  const meses = [
    "enero","febrero","marzo","abril","mayo","junio",
    "julio","agosto","septiembre","octubre","noviembre","diciembre"
  ];
  return `${d.getDate()} de ${meses[d.getMonth()]} de ${d.getFullYear()}`;
}

function getSheetNameFromDate(fechaJS) {
  const meses = [
    "Enero","Febrero","Marzo","Abril","Mayo","Junio",
    "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"
  ];
  return `${meses[fechaJS.getMonth()]} ${fechaJS.getFullYear()}`;
}

function getProdSheetAndRows(wbProd, fechaJS) {
  const name = getSheetNameFromDate(fechaJS);
  let ws = wbProd.Sheets[name];
  if (ws) {
    return { ws, rows: XLSX.utils.sheet_to_json(ws,{header:1,raw:true}) };
  }

  // Buscar hoja por fecha si el nombre no coincide
  const keyTarget = jsDateKey(fechaJS);
  for (const n of wbProd.SheetNames) {
    const wsTry = wbProd.Sheets[n];
    const rows = XLSX.utils.sheet_to_json(wsTry, {header:1,raw:true});
    for (let i=0;i<rows.length;i++){
      if (excelDateKey(rows[i][CONFIG.COL_FECHA]) === keyTarget) {
        return {ws:wsTry, rows};
      }
    }
  }

  const ws0 = wbProd.Sheets[wbProd.SheetNames[0]];
  return {ws:ws0, rows:XLSX.utils.sheet_to_json(ws0,{header:1,raw:true})};
}

function findRowByDate(rows, fechaJS) {
  const k = jsDateKey(fechaJS);
  for (let i=0;i<rows.length;i++){
    if (excelDateKey(rows[i][CONFIG.COL_FECHA])===k) return rows[i];
  }
  return null;
}

// ======================= INFORME DIARIO ==========================
function generarInformeDiario() {
  showError("");

  const fechaStr = document.getElementById("fecha").value;
  if (!fechaStr){ showError("Selecciona la fecha del informe diario."); return; }
  if (!prodArrayBuffer){ showError("Carga el archivo de producción."); return; }
  if (!aforoArrayBuffer){ showError("Carga el archivo de aforo."); return; }

  const fechaJS = new Date(fechaStr+"T00:00:00");
  const fechaLarga = formatFechaLarga(fechaStr);
  const obs = document.getElementById("obsInput").value.trim();

  try {
    const wbProd = XLSX.read(prodArrayBuffer,{type:"array"});
    const {rows} = getProdSheetAndRows(wbProd, fechaJS);
    const row = findRowByDate(rows, fechaJS);
    if (!row){ showError("No se encontró la fecha en el archivo de producción."); return; }

    // Energía (kWh)
    const aux_kwh   = posNum(row[CONFIG.COL_AUX_KWH]);
    const lanec_kwh = posNum(row[CONFIG.COL_LANEC_PARCIAL_KWH]);
    const graca_kwh = posNum(row[CONFIG.COL_GRACA_PARCIAL_KWH]);
    const gen1_kwh  = posNum(row[CONFIG.COL_GEN1_KWH]);
    const gen2_kwh  = posNum(row[CONFIG.COL_GEN2_KWH]);

    const total_kwh_clientes = lanec_kwh + graca_kwh;
    const total_gen_kwh      = total_kwh_clientes + aux_kwh;

    const share_lan = total_kwh_clientes>0? (lanec_kwh/total_kwh_clientes)*100:0;
    const share_gra = total_kwh_clientes>0? (graca_kwh/total_kwh_clientes)*100:0;
    const share_aux = total_gen_kwh>0? (aux_kwh/total_gen_kwh)*100:0;

    const sumGenKwh = gen1_kwh + gen2_kwh;

    // Horas
    const h1i = posNum(row[CONFIG.COL_HG1_LANEC_INI]);
    const h1f = posNum(row[CONFIG.COL_HG1_LANEC_FIN]);
    const h2i = posNum(row[CONFIG.COL_HG2_LANEC_INI]);
    const h2f = posNum(row[CONFIG.COL_HG2_LANEC_FIN]);

    let u1_dia = Math.max(0, h1f - h1i);
    let u2_dia = Math.max(0, h2f - h2i);

    let u1_ac = Math.max(0, h1f - HORO_BASE_U1);
    let u2_ac = Math.max(0, h2f - HORO_BASE_U2);

    const u1_rest = OBJ_MTO_HORAS_U1 - u1_ac;
    const u2_rest = OBJ_MTO_HORAS_U2 - u2_ac;

    // Horas de operación del sistema (al menos una unidad en servicio)
    const horasOperDia = Math.max(u1_dia, u2_dia);

    // Potencias medias (kW) usando horas de operación
    const pmed_total = horasOperDia > 0 ? total_gen_kwh      / horasOperDia : 0;
    const pmed_cli   = horasOperDia > 0 ? total_kwh_clientes / horasOperDia : 0;
    const pmed_aux   = horasOperDia > 0 ? aux_kwh            / horasOperDia : 0;
    const pmed_lan   = horasOperDia > 0 ? lanec_kwh          / horasOperDia : 0;
    const pmed_gra   = horasOperDia > 0 ? graca_kwh          / horasOperDia : 0;

    // Potencias medias por generador (kWh / horas de operación de cada unidad)
    const pmed_g1 = u1_dia > 0 ? gen1_kwh / u1_dia : 0;
    const pmed_g2 = u2_dia > 0 ? gen2_kwh / u2_dia : 0;

    const shareG1 = sumGenKwh>0? (gen1_kwh/sumGenKwh)*100:0;
    const shareG2 = sumGenKwh>0? (gen2_kwh/sumGenKwh)*100:0;

    // Combustible
    const hfo = posNum(row[CONFIG.COL_HFO_GAL]);
    const dsl = posNum(row[CONFIG.COL_DO_GAL]);
    const fuelTot = hfo + dsl;

    const rendimiento = fuelTot>0? total_gen_kwh/fuelTot : 0;

    // Stocks
    const stock_hfo = posNum(row[CONFIG.COL_STOCK_HFO_TOTAL]);
    const stock_do  = posNum(row[CONFIG.COL_STOCK_DO_TOTAL]);

    const aut_hfo = (hfo>0? stock_hfo/hfo : 0);
    const aut_do  = (dsl>0? stock_do/dsl : 0);

    // ================= SALIDA HTML =====================
    let html = "";
    html += `<pre style="margin:0 0 10px 0; white-space:pre-wrap;">
REPORTE POST OPERATIVO DIARIO
CENTRAL EL MORRO – MORRO ENERGY S.A.
Fecha de operación: ${fechaLarga}
</pre>`;

    // 1. Producción de energía
    html += `<div class="section-title">1. PRODUCCIÓN DE ENERGÍA</div>
<table class="data-table"><thead><tr>
<th>Concepto</th><th>Energía [kWh]</th><th>Potencia media [kW]</th></tr></thead><tbody>
<tr><td class="label">Energía generada total</td><td>${fmt(total_gen_kwh)}</td><td>${fmt(pmed_total,1)}</td></tr>
<tr><td class="label">Energía a clientes</td><td>${fmt(total_kwh_clientes)}</td><td>${fmt(pmed_cli,1)}</td></tr>
<tr><td class="label">Auxiliares</td><td>${fmt(aux_kwh)}</td><td>${fmt(pmed_aux,1)}</td></tr>
</tbody></table>`;

    html += `<table class="data-table"><thead><tr>
<th>Unidad</th><th>Energía [kWh]</th><th>Potencia media [kW]</th><th>Participación [%]</th>
</tr></thead><tbody>
<tr>
  <td class="label">Generador 1</td>
  <td>${fmt(gen1_kwh)}</td>
  <td>${u1_dia>0 ? fmt(pmed_g1,1) : "N/A"}</td>
  <td>${fmt(shareG1,1)}</td>
</tr>
<tr>
  <td class="label">Generador 2</td>
  <td>${fmt(gen2_kwh)}</td>
  <td>${u2_dia>0 ? fmt(pmed_g2,1) : "N/A"}</td>
  <td>${fmt(shareG2,1)}</td>
</tr>
</tbody></table>`;

    // 2. Distribución por alimentador
    html += `<div class="section-title">2. DISTRIBUCIÓN POR ALIMENTADOR</div>
<table class="data-table"><thead><tr>
<th>Destino</th><th>Energía [kWh]</th><th>Potencia media [kW]</th><th>Participación [%]</th>
</tr></thead><tbody>
<tr><td class="label">LANEC</td><td>${fmt(lanec_kwh)}</td><td>${horasOperDia>0?fmt(pmed_lan,1):"N/A"}</td><td>${fmt(share_lan,1)}</td></tr>
<tr><td class="label">GRACA</td><td>${fmt(graca_kwh)}</td><td>${horasOperDia>0?fmt(pmed_gra,1):"N/A"}</td><td>${fmt(share_gra,1)}</td></tr>
<tr><td class="label">Auxiliares</td><td>${fmt(aux_kwh)}</td><td>${horasOperDia>0?fmt(pmed_aux,1):"N/A"}</td><td>${fmt(share_aux,1)}</td></tr>
</tbody></table>`;

    // 3. Combustible
    html += `<div class="section-title">3. COMBUSTIBLE Y EFICIENCIA</div>
<table class="data-table"><thead><tr>
<th>Combustible</th><th>Consumo [gal]</th></tr></thead><tbody>
<tr><td class="label">HFO</td><td>${fmt(hfo)}</td></tr>
<tr><td class="label">Diésel</td><td>${fmt(dsl)}</td></tr>
<tr><td class="label">Total equivalente</td><td>${fmt(fuelTot)}</td></tr>
</tbody></table>


<p>Rendimiento global de la central (energía total generada / combustible total):  <strong>${fmt(rendimiento,2)} kWh/gal</strong></p>`;

// 3.X Análisis Ejecutivo de Combustible (Gerencia) – últimos 30 días con data (cross-hojas)
    html += buildFuelExecutiveHTML(wbProd, fechaJS, "daily");

    // 4. Horas
    html += `<div class="section-title">4. HORAS DE OPERACIÓN</div>
<table class="data-table"><thead><tr>
<th>Unidad</th><th>Día [h]</th><th>Acumuladas [h]</th><th>Restantes para próximo mantenimiento [h]</th>
</tr></thead><tbody>
<tr><td class="label">Unidad 1</td><td>${fmt(u1_dia,1)}</td><td>${fmt(u1_ac,1)}</td><td>${fmt(u1_rest,1)}</td></tr>
<tr><td class="label">Unidad 2</td><td>${fmt(u2_dia,1)}</td><td>${fmt(u2_ac,1)}</td><td>${fmt(u2_rest,1)}</td></tr>
</tbody></table>`;

 
    // 5. Stocks
    html += `<div class="section-title">5. STOCKS Y AUTONOMÍAS</div>
<table class="data-table"><thead><tr>
<th>Producto</th><th>Stock [gal]</th><th>Autonomía [días]</th></tr></thead><tbody>
<tr><td class="label">HFO</td><td>${fmt(stock_hfo)}</td><td>${aut_hfo>0?fmt(aut_hfo,2):"N/A"}</td></tr>
<tr><td class="label">Diésel</td><td>${fmt(stock_do)}</td><td>${aut_do>0?fmt(aut_do,2):"N/A"}</td></tr>
</tbody></table>`;

    
    // 6. INDICADOR DE DESEMPEÑO OPERACIONAL (IDOM)
    const refPrev = calcularReferenciaPromedio3Meses(wbProd, fechaJS);
    const idomDia = calcularIDOMDia(total_gen_kwh, fuelTot, horasOperDia, u1_dia, u2_dia, refPrev);

    html += `<div class="section-title">6. INDICADOR DE DESEMPEÑO OPERACIONAL (IDOM)</div>`;

    if (idomDia) {
      const c = semColor(idomDia.estado);
      html += `
<table class="data-table"><thead><tr><th>Parámetro</th><th>Valor</th></tr></thead><tbody>
<tr><td class="label">Referencia (promedio móvil de 3 meses (meses anteriores))</td><td>${refPrev.mesRef} (días con dato: ${refPrev.diasConDato})</td></tr>
<tr><td class="label">Rendimiento de referencia</td><td>${fmt(refPrev.R_ref,2)} kWh/gal</td></tr>
<tr><td class="label">Índice de carga de referencia</td><td>${fmt(refPrev.IC_ref,3)}</td></tr>
<tr><td class="label">Rendimiento del día</td><td>${fmt(idomDia.R_dia,2)} kWh/gal</td></tr>
<tr><td class="label">Índice de eficiencia</td><td>${fmt(idomDia.IE,3)}</td></tr>
<tr><td class="label">Disponibilidad diaria</td><td>${fmt(idomDia.ID,2)}</td></tr>
<tr><td class="label">Índice de carga del día</td><td>${fmt(idomDia.IC_dia,3)}</td></tr>
<tr><td class="label"><strong>IDOM_D</strong></td><td><strong>${fmt(idomDia.IDOM,4)}</strong></td></tr>
<tr><td class="label">Estado operacional</td><td><strong style="color:${c}">${idomDia.estado}</strong></td></tr>
<tr><td class="label">Causa principal</td><td><strong>${idomDia.driver}</strong></td></tr>
<tr><td class="label">Pérdida energética vs promedio móvil de 3 meses (meses anteriores)</td><td>${fmt(idomDia.Loss_kWh,0)} kWh</td></tr>
</tbody></table>
`;
      html += notaPieIDOM(refPrev);
    } else {
      html += `<p>No se pudo calcular IDOM: verifique que existan datos completos del día (energía, combustible y horas) y del promedio móvil de 3 meses (meses anteriores) (referencia).</p>`;
    }

// 7. OBSERVACIONES
    html += `<div class="section-title">7. OBSERVACIONES</div>`;
    html += obs ? `<p>${obs.replace(/\n/g,"<br>")}</p>` :
                  `<p>Sin novedades operativas relevantes.</p>`;

    document.getElementById("output").innerHTML = html;

  } catch (err) {
    console.error(err);
    showError("Error en informe diario.");
  }
}


// ======================= HELPERS PARA INFORME MENSUAL / FACTURACIÓN ==========================
function getDaysInMonth(year, monthIndex){
  // monthIndex: 0-11
  return new Date(year, monthIndex + 1, 0).getDate();
}

function getMesNombreES(monthIndex){
  const meses = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
  return meses[monthIndex] || "";
}

function calcularAcumuladosMensuales(mesStr){
  // mesStr: "YYYY-MM"
  if (!prodArrayBuffer) throw new Error("No hay archivo de producción cargado.");

  const partes = (mesStr || "").split("-");
  if (partes.length !== 2) throw new Error("Formato de mes inválido.");
  const year = parseInt(partes[0], 10);
  const monthIndex = parseInt(partes[1], 10) - 1;
  if (isNaN(year) || isNaN(monthIndex) || monthIndex < 0 || monthIndex > 11) throw new Error("Mes inválido.");

  const diasMes = getDaysInMonth(year, monthIndex);

  // Fecha de corte = último día del mes seleccionado
  const fechaCorte = new Date(year, monthIndex + 1, 0);
  fechaCorte.setHours(0,0,0,0);

  const wbProd = XLSX.read(prodArrayBuffer,{type:"array"});
  const {rows} = getProdSheetAndRows(wbProd, fechaCorte);

  let lan=0, gra=0, aux=0, g1=0, g2=0, hfo=0, dsl=0;
  let first_h1=null, last_h1=null, first_h2=null, last_h2=null;
  let ultimoDia = 0;

  for (let i=0;i<rows.length;i++) {
    const r = rows[i];
    const key = excelDateKey(r[CONFIG.COL_FECHA]);
    if (!key) continue;

    const d = new Date(key+"T00:00:00");
    if (d.getMonth() !== monthIndex || d.getFullYear() !== year) continue;
    if (d.getTime() > fechaCorte.getTime()) continue;

    if (d.getDate() > ultimoDia) ultimoDia = d.getDate();

    const LAN = posNum(r[CONFIG.COL_LANEC_PARCIAL_KWH]);
    const GRA = posNum(r[CONFIG.COL_GRACA_PARCIAL_KWH]);
    const AUX = posNum(r[CONFIG.COL_AUX_KWH]);

    const G1  = posNum(r[CONFIG.COL_GEN1_KWH]);
    const G2  = posNum(r[CONFIG.COL_GEN2_KWH]);

    const HFO = posNum(r[CONFIG.COL_HFO_GAL]);
    const DSL = posNum(r[CONFIG.COL_DO_GAL]);

    const h1i = posNum(r[CONFIG.COL_HG1_LANEC_INI]);
    const h1f = posNum(r[CONFIG.COL_HG1_LANEC_FIN]);
    const h2i = posNum(r[CONFIG.COL_HG2_LANEC_INI]);
    const h2f = posNum(r[CONFIG.COL_HG2_LANEC_FIN]);

    lan+=LAN; gra+=GRA; aux+=AUX; g1+=G1; g2+=G2; hfo+=HFO; dsl+=DSL;

    if (h1i>0 && first_h1===null) first_h1=h1i;
    if (h1f>0) last_h1=h1f;
    if (h2i>0 && first_h2===null) first_h2=h2i;
    if (h2f>0) last_h2=h2f;
  }

  const tot_cli = lan + gra;
  const tot_gen = tot_cli + aux;

  const fuelTot = hfo + dsl;
  const rendimiento = fuelTot>0? tot_gen/fuelTot : 0;

  const shareL = tot_cli>0? (lan/tot_cli)*100 : 0;
  const shareG = tot_cli>0? (gra/tot_cli)*100 : 0;

  const aux_lan = tot_cli>0 ? aux * (lan/tot_cli) : 0;
  const aux_gra = tot_cli>0 ? aux * (gra/tot_cli) : 0;

  const lan_fact = lan + aux_lan;
  const gra_fact = gra + aux_gra;

  let u1_mes = (first_h1!==null && last_h1!==null)? Math.max(0,last_h1-first_h1):0;
  let u2_mes = (first_h2!==null && last_h2!==null)? Math.max(0,last_h2-first_h2):0;

  const horasOperMes = Math.max(u1_mes, u2_mes);

  const mesNombre = getMesNombreES(monthIndex);
  const textoPeriodo = ultimoDia > 0 ? `${mesNombre} ${year} (hasta el día ${ultimoDia})` : `${mesNombre} ${year}`;

  return {
    year, monthIndex, diasMes,
    textoPeriodo,
    lan, gra, aux, tot_cli, tot_gen,
    g1, g2, hfo, dsl, fuelTot, rendimiento,
    shareL, shareG,
    aux_lan, aux_gra, lan_fact, gra_fact,
    u1_mes, u2_mes, horasOperMes
  };
}

// ======================= INFORME MENSUAL ==========================
// Ahora usa el mes seleccionado en el campo <input type="month" id="mesMensual">
function generarInformeMensual() {
  showError("");
  try{

  if (!prodArrayBuffer){ 
    showError("Carga el archivo de producción."); 
    return; 
  }

  const mesStr = document.getElementById("mesMensual").value;
  if (!mesStr) {
    showError("Selecciona el mes del informe mensual.");
    return;
  }

  // mesStr viene como "YYYY-MM"
  const partes = mesStr.split("-");
  if (partes.length !== 2) {
    showError("Formato de mes inválido.");
    return;
  }

  const year = parseInt(partes[0], 10);
  const monthIndex = parseInt(partes[1], 10) - 1; // 0–11
  if (isNaN(year) || isNaN(monthIndex) || monthIndex < 0 || monthIndex > 11) {
    showError("Mes del informe mensual inválido.");
    return;
  }

  // Fecha de corte = último día del mes seleccionado
  const fechaCorte = new Date(year, monthIndex + 1, 0);
  fechaCorte.setHours(0,0,0,0);

  const wbProd = XLSX.read(prodArrayBuffer,{type:"array"});
  const {rows} = getProdSheetAndRows(wbProd, fechaCorte);

  const mesTexto = getSheetNameFromDate(fechaCorte); // "Diciembre 2025", etc.

  let lan=0, gra=0, aux=0, g1=0, g2=0, hfo=0, dsl=0;
  let first_h1=null, last_h1=null, first_h2=null, last_h2=null;
  let ultimoDia = 0; // último día del mes con datos

  for (let i=0;i<rows.length;i++) {
    const r = rows[i];
    const key = excelDateKey(r[CONFIG.COL_FECHA]);
    if (!key) continue;

    const d = new Date(key+"T00:00:00");
    if (d.getMonth() !== fechaCorte.getMonth() || d.getFullYear() !== fechaCorte.getFullYear()) continue;
    if (d.getTime() > fechaCorte.getTime()) continue;

    // Actualizar último día con datos
    if (d.getDate() > ultimoDia) ultimoDia = d.getDate();

    const LAN = posNum(r[CONFIG.COL_LANEC_PARCIAL_KWH]);
    const GRA = posNum(r[CONFIG.COL_GRACA_PARCIAL_KWH]);
    const AUX = posNum(r[CONFIG.COL_AUX_KWH]);

    const G1  = posNum(r[CONFIG.COL_GEN1_KWH]);
    const G2  = posNum(r[CONFIG.COL_GEN2_KWH]);
    const HFO = posNum(r[CONFIG.COL_HFO_GAL]);
    const DSL = posNum(r[CONFIG.COL_DO_GAL]);

    const h1i = posNum(r[CONFIG.COL_HG1_LANEC_INI]);
    const h1f = posNum(r[CONFIG.COL_HG1_LANEC_FIN]);
    const h2i = posNum(r[CONFIG.COL_HG2_LANEC_INI]);
    const h2f = posNum(r[CONFIG.COL_HG2_LANEC_FIN]);

    lan+=LAN; gra+=GRA; aux+=AUX; g1+=G1; g2+=G2; hfo+=HFO; dsl+=DSL;

    if (h1i>0 && first_h1===null) first_h1=h1i;
    if (h1f>0) last_h1=h1f;
    if (h2i>0 && first_h2===null) first_h2=h2i;
    if (h2f>0) last_h2=h2f;
  }

  const tot_cli = lan + gra;
  const tot_gen = tot_cli + aux;

  const fuelTot = hfo + dsl;
  const rendimiento = fuelTot>0? tot_gen/fuelTot : 0;

  const shareL = tot_cli>0? (lan/tot_cli)*100 : 0;
  const shareG = tot_cli>0? (gra/tot_cli)*100 : 0;

  let u1_mes = (first_h1!==null && last_h1!==null)? Math.max(0,last_h1-first_h1):0;
  let u2_mes = (first_h2!==null && last_h2!==null)? Math.max(0,last_h2-first_h2):0;

  // Horas de operación del sistema en el mes
  const horasOperMes = Math.max(u1_mes, u2_mes);

  // Potencias medias (kW) usando horas de operación del sistema
  const pmed_total = horasOperMes>0 ? tot_gen/horasOperMes : 0;
  const pmed_cli   = horasOperMes>0 ? tot_cli/horasOperMes : 0;
  const pmed_aux   = horasOperMes>0 ? aux/horasOperMes : 0;
  const pmed_lan   = horasOperMes>0 ? lan/horasOperMes : 0;
  const pmed_gra   = horasOperMes>0 ? gra/horasOperMes : 0;

  // Potencias medias por unidad
  const pmed_g1 = u1_mes>0 ? g1/u1_mes : 0;
  const pmed_g2 = u2_mes>0 ? g2/u2_mes : 0;

  const sumGenKwh = g1 + g2;
  const shareG1 = sumGenKwh>0 ? (g1/sumGenKwh)*100 : 0;
  const shareG2 = sumGenKwh>0 ? (g2/sumGenKwh)*100 : 0;

  // Encabezado: si tenemos último día, lo mostramos
  const textoPeriodo = ultimoDia > 0
    ? `${mesTexto} (hasta el día ${ultimoDia})`
    : `${mesTexto}`;

  // === HTML ===
  let html = "";
  html += `<pre style="margin:0 0 10px 0; white-space:pre-wrap;">
REPORTE POST OPERATIVO MENSUAL
CENTRAL EL MORRO – MORRO ENERGY S.A.
Período de operación: ${textoPeriodo}
</pre>`;

  html += `<div class="section-title">1. PRODUCCIÓN DE ENERGÍA</div>
<table class="data-table"><thead><tr>
<th>Concepto</th><th>Energía [kWh]</th><th>Potencia media [kW]</th>
</tr></thead><tbody>
<tr><td class="label">Energía generada total</td><td>${fmt(tot_gen)}</td><td>${horasOperMes>0?fmt(pmed_total,1):"N/A"}</td></tr>
<tr><td class="label">Energía a clientes</td><td>${fmt(tot_cli)}</td><td>${horasOperMes>0?fmt(pmed_cli,1):"N/A"}</td></tr>
<tr><td class="label">Auxiliares</td><td>${fmt(aux)}</td><td>${horasOperMes>0?fmt(pmed_aux,1):"N/A"}</td></tr>
</tbody></table>`;

  // Factor de planta (respecto a potencia instalada y potencia efectiva)
  const diasPeriodo = (ultimoDia > 0) ? ultimoDia : getDaysInMonth(year, monthIndex);
  const horasCalendario = diasPeriodo * 24;
  const fp_inst = (horasCalendario>0) ? (tot_gen / (P_INST_TOTAL * horasCalendario)) : 0;
  const fp_eff  = (horasCalendario>0) ? (tot_gen / (P_INST_EFECTIVA * horasCalendario)) : 0;

  html += `<p>
    <strong>Factor de planta (período reportado)</strong>: 
    ${(fp_inst*100).toFixed(1)}% (vs instalada ${P_INST_TOTAL} kW) — 
    ${(fp_eff*100).toFixed(1)}% (vs efectiva ${Math.round(P_INST_EFECTIVA)} kW)
  </p>`;


  html += `<table class="data-table"><thead><tr>
<th>Unidad</th><th>Energía [kWh]</th><th>Potencia media [kW]</th><th>Participación [%]</th>
</tr></thead><tbody>
<tr>
  <td class="label">Generador 1</td>
  <td>${fmt(g1)}</td>
  <td>${u1_mes>0 ? fmt(pmed_g1,1) : "N/A"}</td>
  <td>${fmt(shareG1,1)}</td>
</tr>
<tr>
  <td class="label">Generador 2</td>
  <td>${fmt(g2)}</td>
  <td>${u2_mes>0 ? fmt(pmed_g2,1) : "N/A"}</td>
  <td>${fmt(shareG2,1)}</td>
</tr>
</tbody></table>`;

  html += `<div class="section-title">2. DISTRIBUCIÓN ENERGÉTICA</div>
<table class="data-table"><thead><tr>
<th>Destino</th><th>Energía [kWh]</th><th>Potencia media [kW]</th><th>Participación [%]</th>
</tr></thead><tbody>
<tr><td class="label">LANEC</td><td>${fmt(lan)}</td><td>${horasOperMes>0?fmt(pmed_lan,1):"N/A"}</td><td>${fmt(shareL,1)}</td></tr>
<tr><td class="label">GRACA</td><td>${fmt(gra)}</td><td>${horasOperMes>0?fmt(pmed_gra,1):"N/A"}</td><td>${fmt(shareG,1)}</td></tr>
</tbody></table>`;

  html += `<div class="section-title">3. COMBUSTIBLE Y EFICIENCIA</div>
<table class="data-table"><thead><tr>
<th>Combustible</th><th>Consumo [gal]</th>
</tr></thead><tbody>
<tr><td class="label">HFO</td><td>${fmt(hfo)}</td></tr>
<tr><td class="label">Diésel</td><td>${fmt(dsl)}</td></tr>
<tr><td class="label">Total equivalente</td><td>${fmt(fuelTot)}</td></tr>
</tbody></table>
<p>Rendimiento promedio: <strong>${fmt(rendimiento,2)} kWh/gal</strong></p>`;

// 3.X Análisis Ejecutivo de Combustible (Gerencia) – últimos 30 días con data (cross-hojas)
    html += buildFuelExecutiveHTML(wbProd, fechaCorte, "monthly");  

  html += `<div class="section-title">4. HORAS DE OPERACIÓN (MENSUAL)</div>
<table class="data-table"><thead><tr>
<th>Unidad</th><th>Horas [h]</th>
</tr></thead><tbody>
<tr><td class="label">Unidad 1</td><td>${fmt(u1_mes,1)}</td></tr>
<tr><td class="label">Unidad 2</td><td>${fmt(u2_mes,1)}</td></tr>
</tbody></table>`;

    


// ======================== 5. ENERGÍA A FACTURAR POR CLIENTE ==========================
  const aux_lan  = aux * (shareL / 100);
  const aux_gra  = aux * (shareG / 100);

  const lan_fact = lan + aux_lan;
  const gra_fact = gra + aux_gra;

  html += `<div class="section-title">5. ENERGÍA A FACTURAR POR CLIENTE</div>
<table class="data-table">
<thead>
<tr>
  <th>Cliente</th>
  <th>Energía directa [kWh]</th>
  <th>Participación aux [%]</th>
  <th>Auxiliares asignados [kWh]</th>
  <th>Total facturable [kWh]</th>
</tr>
</thead>
<tbody>
<tr>
  <td class="label">LANEC</td>
  <td>${fmt(lan)}</td>
  <td>${fmt(shareL,2)}</td>
  <td>${fmt(aux_lan)}</td>
  <td><strong>${fmt(lan_fact)}</strong></td>
</tr>
<tr>
  <td class="label">GRACA</td>
  <td>${fmt(gra)}</td>
  <td>${fmt(shareG,2)}</td>
  <td>${fmt(aux_gra)}</td>
  <td><strong>${fmt(gra_fact)}</strong></td>
</tr>
</tbody>
</table>`;

  
  // 6. INDICADOR DE DESEMPEÑO OPERACIONAL (IDOM) - MENSUAL (referencia = promedio móvil de 3 meses (meses anteriores))
  const refPrevM = calcularReferenciaPromedio3Meses(wbProd, fechaCorte);
  html += `<div class="section-title">6. INDICADOR DE DESEMPEÑO OPERACIONAL (IDOM)</div>`;

  if (refPrevM && refPrevM.R_ref>0 && refPrevM.IC_ref>0 && fuelTot>0 && horasOperMes>0) {
    const R_mes = rendimiento; // kWh/gal
    const IE_mes = R_mes / refPrevM.R_ref;
    const IC_mes = ((tot_gen / horasOperMes) / P_INST_EFECTIVA);
    const ID_mes = 1; // mensual: disponibilidad se refleja principalmente por horas del sistema; mantener ID=1 para comparabilidad
    const IDOM_M = 0.4*IE_mes + 0.3*ID_mes + 0.3*(IC_mes / refPrevM.IC_ref);
    const estadoM = semaforo(IDOM_M);
    const cM = semColor(estadoM);
    const lossM = Math.max(0, (fuelTot*refPrevM.R_ref) - tot_gen);

    html += `
<table class="data-table"><thead><tr><th>Parámetro</th><th>Valor</th></tr></thead><tbody>
<tr><td class="label">Referencia (promedio móvil de 3 meses (meses anteriores))</td><td>${refPrevM.mesRef} (días con dato: ${refPrevM.diasConDato})</td></tr>
<tr><td class="label">Rendimiento de referencia</td><td>${fmt(refPrevM.R_ref,2)} kWh/gal</td></tr>
<tr><td class="label">Índice de carga de referencia</td><td>${fmt(refPrevM.IC_ref,3)}</td></tr>
<tr><td class="label">Rendimiento del mes</td><td>${fmt(R_mes,2)} kWh/gal</td></tr>
<tr><td class="label">Índice de eficiencia mensual</td><td>${fmt(IE_mes,3)}</td></tr>
<tr><td class="label">Índice de carga mensual</td><td>${fmt(IC_mes,3)}</td></tr>
<tr><td class="label"><strong>IDOM_M</strong></td><td><strong>${fmt(IDOM_M,4)}</strong></td></tr>
<tr><td class="label">Estado operacional</td><td><strong style="color:${cM}">${estadoM}</strong></td></tr>
<tr><td class="label">Pérdida energética vs promedio móvil de 3 meses (meses anteriores)</td><td>${fmt(lossM,0)} kWh</td></tr>
</tbody></table>
`;
    html += notaPieIDOM(refPrevM);
  } else {
    html += `<p>No se pudo calcular IDOM_M: verifique datos completos del mes seleccionado y del promedio móvil de 3 meses (meses anteriores) (referencia).</p>`;
  }

document.getElementById("output").innerHTML = html;
    document.getElementById("output").innerHTML = html;
  } catch (err) {
    console.error(err);
    showError("Error en informe mensual: " + (err && err.message ? err.message : err));
  }
}



// ======================= INFORME DE FACTURACIÓN (MENSUAL) ==========================
function generarInformeFacturacion(){
  showError("");

  if (!prodArrayBuffer){
    showError("Carga el archivo de producción.");
    return;
  }

  const mesStr = document.getElementById("mesMensual").value;
  if (!mesStr){
    showError("Selecciona el mes para la facturación (campo Mes del informe mensual).");
    return;
  }

  let data;
  try {
    data = calcularAcumuladosMensuales(mesStr);
  } catch (e){
    console.error(e);
    showError(e.message || "Error al calcular acumulados del mes.");
    return;
  }

  // Días indisponibles (manual)
  const diasFallaU1 = Math.max(0, Math.floor(posNum(document.getElementById("diasFallaU1").value)));
  const diasFallaU2 = Math.max(0, Math.floor(posNum(document.getElementById("diasFallaU2").value)));

  const diasMes = data.diasMes;

  const dispU1 = Math.max(0, (diasMes - diasFallaU1) / diasMes);
  const dispU2 = Math.max(0, (diasMes - diasFallaU2) / diasMes);

  const fijoU1 = COSTO_FIJO_MENSUAL_POR_UNIDAD * dispU1;
  const fijoU2 = COSTO_FIJO_MENSUAL_POR_UNIDAD * dispU2;
  const fijoTotal = fijoU1 + fijoU2;

  // === COSTO FIJO: reparto por potencia contratada + factor de utilización (OPCIÓN A) ===
  // Contratos (kW)
  const P_CONTR_LANEC = 3800;
  const P_CONTR_GRACA = 1000;
  const P_CONTR_TOT   = P_CONTR_LANEC + P_CONTR_GRACA;

  const factorContratoLan = P_CONTR_TOT > 0 ? (P_CONTR_LANEC / P_CONTR_TOT) : 0;
  const factorContratoGra = P_CONTR_TOT > 0 ? (P_CONTR_GRACA / P_CONTR_TOT) : 0;

  // === COSTO FIJO: reparto SOLO por % de Factor contrato (sin factor de utilización) ===
  // Costo fijo por unidad ya considera disponibilidad (dispU1/dispU2)
  const fijoLanU1 = fijoU1 * factorContratoLan;
  const fijoLanU2 = fijoU2 * factorContratoLan;
  const fijoGraU1 = fijoU1 * factorContratoGra;
  const fijoGraU2 = fijoU2 * factorContratoGra;

  const fijoLan = fijoLanU1 + fijoLanU2;
  const fijoGra = fijoGraU1 + fijoGraU2;

  // Auditoría (debe ser igual a fijoTotal)
  const fijoSumClientes = fijoLan + fijoGra;
// Costos variables (desglosados)
  function subtotalVariable(kwh){
    const subt = {};
    for (const [k,v] of Object.entries(COSTOS_VARIABLES)){
      subt[k] = kwh * v;
    }
    return subt;
  }

  const varLanBy = subtotalVariable(data.lan_fact);
  const varGraBy = subtotalVariable(data.gra_fact);

  const varLanTotal = data.lan_fact * COSTO_VARIABLE_TOTAL;
  const varGraTotal = data.gra_fact * COSTO_VARIABLE_TOTAL;

  const totalLan = varLanTotal + fijoLan;
  const totalGra = varGraTotal + fijoGra;

  const precioLan = data.lan_fact > 0 ? totalLan / data.lan_fact : 0;
  const precioGra = data.gra_fact > 0 ? totalGra / data.gra_fact : 0;
  const varTotBy = subtotalVariable(data.tot_gen);
  const varTotTotal = data.tot_gen * COSTO_VARIABLE_TOTAL;

  // Totales auditables (deben ser iguales a LANEC + GRACA)
  const fijoTotU1 = fijoLanU1 + fijoGraU1;
  const fijoTotU2 = fijoLanU2 + fijoGraU2;
  const fijoTot   = fijoLan + fijoGra;

  const totalTot  = totalLan + totalGra;
  const precioTot = data.tot_gen > 0 ? totalTot / data.tot_gen : 0;


  // ===== HTML CELEC-style =====
  let html = "";
  html += `<pre style="margin:0 0 10px 0; white-space:pre-wrap;">
INFORME DE FACTURACIÓN DE ENERGÍA (MENSUAL)
CENTRAL EL MORRO – MORRO ENERGY S.A.
Período de facturación: ${data.textoPeriodo}
</pre>`;

  // 1. Resumen de energía facturable
  html += `<div class="section-title">1. RESUMEN DE ENERGÍA FACTURABLE</div>
<table class="data-table">
<thead><tr>
<th>Cliente</th>
<th>Energía consumida [kWh]</th>
<th>Auxiliares asignados [kWh]</th>
<th>Total facturable [kWh]</th>
</tr></thead>
<tbody>
<tr>
<td class="label">LANEC</td>
<td>${fmt(data.lan)}</td>
<td>${fmt(data.aux_lan)}</td>
<td><strong>${fmt(data.lan_fact)}</strong></td>
</tr>
<tr>
<td class="label">GRACA</td>
<td>${fmt(data.gra)}</td>
<td>${fmt(data.aux_gra)}</td>
<td><strong>${fmt(data.gra_fact)}</strong></td>
</tr>
<tr>
<td class="label">TOTAL</td>
<td>${fmt(data.tot_cli)}</td>
<td>${fmt(data.aux)}</td>
<td><strong>${fmt(data.tot_gen)}</strong></td>
</tr>
</tbody></table>`;

  // 2. Costos variables (tabla de P. Unit)
  //html += `<div class="section-title">2. COSTO VARIABLE (USD/kWh)</div>
//<table class="data-table">
//<thead><tr><th>Rubro</th><th>P. Unit [USD/kWh]</th></tr></thead>
//<tbody>
//<tr><td class="label">Combustible + Transporte</td><td>${fmt(COSTOS_VARIABLES.combustible_transporte,4)}</td></tr>
//<tr><td class="label">Lubricantes + Químicos</td><td>${fmt(COSTOS_VARIABLES.lubricantes_quimicos,4)}</td></tr>
//<tr><td class="label">Agua + Insumos</td><td>${fmt(COSTOS_VARIABLES.agua_insumos,4)}</td></tr>
//<tr><td class="label">Repuestos Mantenimiento Predictivo</td><td>${fmt(COSTOS_VARIABLES.repuestos_predictivo,4)}</td></tr>
//<tr><td class="label">Impacto Ambiental</td><td>${fmt(COSTOS_VARIABLES.impacto_ambiental,4)}</td></tr>
//<tr><td class="label">Servicios Auxiliares</td><td>${fmt(COSTOS_VARIABLES.servicios_auxiliares,4)}</td></tr>
//<tr><td class="label">Margen Variable</td><td>${fmt(COSTOS_VARIABLES.margen_variable,4)}</td></tr>
//<tr><td class="label"><strong>Costo variable de producción</strong></td><td><strong>${fmt(COSTO_VARIABLE_TOTAL,4)}</strong></td></tr>
//</tbody></table>`;

  // 2. Detalle por cliente (similar a tus ejemplos)
    function tablaCliente(secLabel, titulo, nombre, energiaConsumida, auxAsig, totalFact, varBy, varTotal, fijoAsigU1, fijoAsigU2, fijoAsig, totalUSD, precioFinal){
    const energiaLabel = (nombre === "TOTAL")
      ? "Energía consumida total (LANEC + GRACA)"
      : `Energía consumida por ${nombre}`;
    const auxLabel = (nombre === "TOTAL")
      ? "Energía consumida por auxiliares (total)"
      : "Energía consumida por auxiliares";
    const totalLabel = (nombre === "TOTAL")
      ? "Energía total a facturar (+auxiliares)"
      : `Energía total ${nombre} (+auxiliares)`;

    return `
<div class="section-title">${secLabel} ${titulo}</div>
<table class="data-table">
<thead><tr><th>Rubro</th><th>P. Unit</th><th>Subtotal</th></tr></thead>
<tbody>
<tr><td class="label">${energiaLabel}</td><td></td><td>${fmt(energiaConsumida)} kWh</td></tr>
<tr><td class="label">${auxLabel}</td><td></td><td>${fmt(auxAsig)} kWh</td></tr>
<tr><td class="label">${totalLabel}</td><td></td><td><strong>${fmt(totalFact)} kWh</strong></td></tr>

<tr><td class="label">Costo Combustible + Transporte</td><td>$ ${fmt(COSTOS_VARIABLES.combustible_transporte,4)}</td><td>$ ${fmt(varBy.combustible_transporte,2)}</td></tr>
<tr><td class="label">Costo Lubricantes + Químicos</td><td>$ ${fmt(COSTOS_VARIABLES.lubricantes_quimicos,4)}</td><td>$ ${fmt(varBy.lubricantes_quimicos,2)}</td></tr>
<tr><td class="label">Costo Agua + Insumos</td><td>$ ${fmt(COSTOS_VARIABLES.agua_insumos,4)}</td><td>$ ${fmt(varBy.agua_insumos,2)}</td></tr>
<tr><td class="label">Costo Repuestos Mantenimiento Predictivo</td><td>$ ${fmt(COSTOS_VARIABLES.repuestos_predictivo,4)}</td><td>$ ${fmt(varBy.repuestos_predictivo,2)}</td></tr>
<tr><td class="label">Costo Impacto Ambiental</td><td>$ ${fmt(COSTOS_VARIABLES.impacto_ambiental,4)}</td><td>$ ${fmt(varBy.impacto_ambiental,2)}</td></tr>
<tr><td class="label">Costo Servicios Auxiliares</td><td>$ ${fmt(COSTOS_VARIABLES.servicios_auxiliares,4)}</td><td>$ ${fmt(varBy.servicios_auxiliares,2)}</td></tr>
<tr><td class="label">Margen Variable</td><td>$ ${fmt(COSTOS_VARIABLES.margen_variable,4)}</td><td>$ ${fmt(varBy.margen_variable,2)}</td></tr>

<tr><td class="label"><strong>Costo Variable de Producción</strong></td><td><strong>$ ${fmt(COSTO_VARIABLE_TOTAL,4)}</strong></td><td><strong>$ ${fmt(varTotal,2)}</strong></td></tr>
<tr><td class="label">Costo fijo asignado U1 (disponibilidad)</td><td></td><td>$ ${fmt(fijoAsigU1,2)}</td></tr>
<tr><td class="label">Costo fijo asignado U2 (disponibilidad)</td><td></td><td>$ ${fmt(fijoAsigU2,2)}</td></tr>
<tr><td class="label"><strong>Costo fijo (disponibilidad) asignado</strong></td><td></td><td><strong>$ ${fmt(fijoAsig,2)}</strong></td></tr>
<tr><td class="label"><strong>Subtotal</strong></td><td></td><td><strong>$ ${fmt(totalUSD,2)}</strong> &nbsp; MAS IVA</td></tr>
<tr><td class="label"><strong>Precio final USD/kWh</strong></td><td></td><td><strong>$ ${fmt(precioFinal,4)}</strong> &nbsp; MAS IVA</td></tr>
</tbody></table>`;
  }

    html += tablaCliente("2.0", "COSTOS DEL MES TOTALES", "TOTAL",
    data.tot_cli, data.aux, data.tot_gen,
    varTotBy, varTotTotal,
    fijoTotU1, fijoTotU2, fijoTot,
    totalTot, precioTot
  );

  html += tablaCliente("2.1", "COSTOS DEL MES LANEC", "LANEC",
    data.lan, data.aux_lan, data.lan_fact,
    varLanBy, varLanTotal,
    fijoLanU1, fijoLanU2, fijoLan,
    totalLan, precioLan
  );

  html += tablaCliente("2.2", "COSTOS DEL MES GRACA", "GRACA",
    data.gra, data.aux_gra, data.gra_fact,
    varGraBy, varGraTotal,
    fijoGraU1, fijoGraU2, fijoGra,
    totalGra, precioGra
  );

  // 4. Costo fijo y disponibilidad (auditable)
  html += `<div class="section-title">3. COSTO FIJO POR DISPONIBILIDAD (AUDITABLE)</div>
<table class="data-table">
<thead><tr>
<th>Unidad</th><th>Días mes</th><th>Días indisponibles</th><th>Factor disponibilidad</th><th>Costo fijo base [USD]</th><th>Costo fijo ajustado [USD]</th>
</tr></thead>
<tbody>
<tr>
<td class="label">Unidad 1</td>
<td>${diasMes}</td>
<td>${diasFallaU1}</td>
<td>${fmt(dispU1,4)}</td>
<td>${fmt(COSTO_FIJO_MENSUAL_POR_UNIDAD,2)}</td>
<td><strong>${fmt(fijoU1,2)}</strong></td>
</tr>
<tr>
<td class="label">Unidad 2</td>
<td>${diasMes}</td>
<td>${diasFallaU2}</td>
<td>${fmt(dispU2,4)}</td>
<td>${fmt(COSTO_FIJO_MENSUAL_POR_UNIDAD,2)}</td>
<td><strong>${fmt(fijoU2,2)}</strong></td>
</tr>
<tr>
<td class="label"><strong>TOTAL</strong></td>
<td></td><td></td><td></td><td></td>
<td><strong>${fmt(fijoTotal,2)}</strong></td>
</tr>
</tbody></table>`;

  // 5. Reparto fijo
  html += `<div class="section-title">4. ASIGNACIÓN DEL COSTO FIJO A CLIENTES (POR FACTOR CONTRATO)</div>
<table class="data-table">
<thead>
<tr>
  <th>Cliente</th>
  <th>kW contratados</th>
  <th>Factor contrato</th>
  <th>CF U1 (disp) [USD]</th>
  <th>CF U2 (disp) [USD]</th>
  <th><strong>CF total asignado [USD]</strong></th>
</tr>
</thead>
<tbody>
<tr>
  <td class="label">LANEC</td>
  <td>${fmt(P_CONTR_LANEC,0)}</td>
  <td>${fmt(factorContratoLan*100,2)} %</td>
  <td>${fmt(fijoLanU1,2)}</td>
  <td>${fmt(fijoLanU2,2)}</td>
  <td><strong>${fmt(fijoLan,2)}</strong></td>
</tr>
<tr>
  <td class="label">GRACA</td>
  <td>${fmt(P_CONTR_GRACA,0)}</td>
  <td>${fmt(factorContratoGra*100,2)} %</td>
  <td>${fmt(fijoGraU1,2)}</td>
  <td>${fmt(fijoGraU2,2)}</td>
  <td><strong>${fmt(fijoGra,2)}</strong></td>
</tr>
<tr>
  <td class="label"><strong>TOTAL</strong></td>
  <td></td>
  <td></td>
  <td>${fmt(fijoU1,2)}</td>
  <td>${fmt(fijoU2,2)}</td>
  <td><strong>${fmt(fijoTotal,2)}</strong></td>
</tr>
</tbody>
</table>

`;


  document.getElementById("output").innerHTML = html;
}


// ======================= GENERAR PDF ===============================
async function generarPDF() {
  const contenido = document.getElementById("output");
  const fechaStr = document.getElementById("fecha").value;

  if (!contenido || !contenido.innerHTML.trim()) {
    showError("No hay contenido para exportar a PDF.");
    return;
  }

  if (!fechaStr &&
      !contenido.innerHTML.includes("REPORTE POST OPERATIVO MENSUAL") &&
      !contenido.innerHTML.includes("INFORME DE FACTURACIÓN")) {
    showError("Selecciona una fecha antes de generar el PDF diario.");
    return;
  }

  if (typeof html2pdf === "undefined") {
    showError("Falta la librería html2pdf.js en el HTML.");
    return;
  }

  showError("");

  const esMensual = contenido.innerHTML.includes("REPORTE POST OPERATIVO MENSUAL");
  const esFacturacion = contenido.innerHTML.includes("INFORME DE FACTURACIÓN");

  let nombreArchivo = "";

  const meses = [
    "Enero","Febrero","Marzo","Abril","Mayo","Junio",
    "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"
  ];

  if (esFacturacion) {
    const mesStr = document.getElementById("mesMensual").value;
    if (mesStr) {
      const partes = mesStr.split("-");
      const year = parseInt(partes[0], 10);
      const monthIndex = parseInt(partes[1], 10) - 1;
      if (!isNaN(year) && !isNaN(monthIndex) && monthIndex >= 0 && monthIndex < 12) {
        nombreArchivo = `Facturacion_ElMorro_${meses[monthIndex]}_${year}.pdf`;
      }
    }
    if (!nombreArchivo) nombreArchivo = "Facturacion_ElMorro.pdf";
  } else if (esMensual) {
    const mesStr = document.getElementById("mesMensual").value;
    if (mesStr) {
      const partes = mesStr.split("-");
      const year = parseInt(partes[0], 10);
      const monthIndex = parseInt(partes[1], 10) - 1;
      if (!isNaN(year) && !isNaN(monthIndex) && monthIndex >= 0 && monthIndex < 12) {
        nombreArchivo = `Reporte_Mensual_ElMorro_${meses[monthIndex]}_${year}.pdf`;
      }
    }
    if (!nombreArchivo) {
      const hoy = new Date();
      hoy.setHours(0,0,0,0);
      const fechaCorte = new Date(hoy);
      fechaCorte.setDate(hoy.getDate() - 1);
      nombreArchivo = `Reporte_Mensual_ElMorro_${meses[fechaCorte.getMonth()]}_${fechaCorte.getFullYear()}.pdf`;
    }
  } else {
    nombreArchivo = `Reporte_Diario_ElMorro_${fechaStr}.pdf`;
  }

  // Activar estilo especial de PDF
  contenido.classList.add("pdf-export");

  const opt = {
    margin:       5,
    filename:     nombreArchivo,
    image:        { type: "jpeg", quality: 0.98 },
    html2canvas:  { scale: 2, scrollY: 0 },
    jsPDF:        { unit: "mm", format: "a4", orientation: "portrait" }
  };

  try {
    await html2pdf().set(opt).from(contenido).save();
  } finally {
    contenido.classList.remove("pdf-export");
  }
}