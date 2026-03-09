/* =========================================================
   Central El Morro – Sistema Integrado de Gráficas
   Integra módulos:
   1) Producción • Clientes • Eficiencia (gerencial)
   2) Combustible (HFO/DO) • Tanques • Cisterna 2
   ========================================================= */

let wbProd = null;
let wbAforo = null;
let prodLoaded = false;
let aforoLoaded = false;

const charts = {};
let lastData = null; // {sheetName, month, year, prod, aforo, resumen}

const elStatus = document.getElementById("statusText");
const elMes = document.getElementById("mesSeleccionado");
const elBtnProcesar = document.getElementById("btnProcesar");
const elBtnPdf = document.getElementById("btnPdf");
const elPdfTipo = document.getElementById("pdfTipo");

const elKpiProd = document.getElementById("kpiProduccion");
const elKpiComb = document.getElementById("kpiCombustible") || document.getElementById("fuelKpis");

/* =========================
   Tabs
   ========================= */
document.querySelectorAll(".tab").forEach((btn) => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".tab").forEach((b) => {
      b.classList.remove("active");
      b.setAttribute("aria-selected", "false");
    });
    btn.classList.add("active");
    btn.setAttribute("aria-selected", "true");

    const id = btn.dataset.tab;
    document.querySelectorAll(".tab-panel").forEach((p) => p.classList.remove("active"));
    document.getElementById(id).classList.add("active");

    // PDF por defecto según tab activa
    elPdfTipo.value = id === "tab-combustible" ? "combustible" : "produccion";

    // Re-render charts when a hidden tab becomes visible (Chart.js needs visible canvas size)
    if (lastData) {
      if (id === "tab-combustible") {
        setTimeout(() => {
          try { renderModuleCombustible(lastData.prod, lastData.aforo, lastData.resumen); } catch(e) { console.error(e); }
        }, 60);
      }
      if (id === "tab-produccion") {
        setTimeout(() => {
          try { renderModuleProduccion(lastData.prod, lastData.aforo, lastData.resumen); } catch(e) { console.error(e); }
        }, 60);
      }
    }
  });
});

function setStatus(msg, isError = false) {
  elStatus.textContent = msg || "";
  elStatus.style.color = isError ? "#b91c1c" : "";
}

/* =========================
   Utilidades numéricas / fecha
   ========================= */
function toNumber(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return isFinite(value) ? value : 0;

  let str = String(value).trim();
  if (!str) return 0;
  str = str.replace(/\s+/g, "");

  // 1.234.567,89
  if (/^\d{1,3}(\.\d{3})*(,\d+)?$/.test(str)) {
    str = str.replace(/\./g, "").replace(",", ".");
  }
  // 1,234,567.89
  else if (/^\d{1,3}(,\d{3})*(\.\d+)?$/.test(str)) {
    str = str.replace(/,/g, "");
  }
  // 12345,67
  else if (/^\d+,\d+$/.test(str)) {
    str = str.replace(",", ".");
  }

  const n = Number(str);
  return isFinite(n) ? n : 0;
}

// Excel (Date, serial, string) -> Date
function parseExcelDate(value) {
  if (value === null || value === undefined || value === "") return null;

  if (Object.prototype.toString.call(value) === "[object Date]") {
    return isNaN(value.getTime()) ? null : value;
  }

  if (typeof value === "number") {
    const o = XLSX.SSF.parse_date_code(value);
    if (!o) return null;
    return new Date(o.y, o.m - 1, o.d);
  }

  const s = String(value).trim();
  if (!s) return null;

  const d1 = new Date(s);
  if (!isNaN(d1.getTime())) return d1;

  // dd/mm/yyyy
  const m = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (m) {
    let dd = parseInt(m[1], 10);
    let mm = parseInt(m[2], 10);
    let yy = parseInt(m[3], 10);
    if (yy < 100) yy = 2000 + yy;
    const d2 = new Date(yy, mm - 1, dd);
    return isNaN(d2.getTime()) ? null : d2;
  }

  return null;
}

function dayLabel(date) {
  const d = String(date.getDate()).padStart(2, "0");
  const m = String(date.getMonth() + 1).padStart(2, "0");
  return `${d}-${m}`;
}

function formatNumber(value, decimals = 0) {
  if (value === null || value === undefined || !isFinite(value)) return "-";
  try {
    return value.toLocaleString("es-EC", {
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals
    });
  } catch {
    return Number(value).toFixed(decimals);
  }
}

/* =========================
   Chart helper
   ========================= */
function createOrUpdateChart(canvasId, config) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;

  // Aplicar estilo consistente (como el informe de combustible)
  config = applyChartDefaults(config);

  if (charts[canvasId]) charts[canvasId].destroy();
  charts[canvasId] = new Chart(canvas.getContext("2d"), config);
}

function applyChartDefaults(config){
  if (!config || !config.type || !config.data) return config;

  // Defaults para líneas (puntos visibles, líneas suaves)
  if (config.type === "line") {
    config.options = config.options || {};
    config.options.responsive = true;
    config.options.maintainAspectRatio = false;
    config.options.interaction = config.options.interaction || { mode: "index", intersect: false };

    config.options.elements = config.options.elements || {};
    config.options.elements.line = Object.assign({ tension: 0.25, borderWidth: 2 }, config.options.elements.line || {});
    config.options.elements.point = Object.assign({ radius: 2, hoverRadius: 4 }, config.options.elements.point || {});

    config.options.plugins = config.options.plugins || {};
    config.options.plugins.legend = config.options.plugins.legend || {};
    config.options.plugins.legend.position = config.options.plugins.legend.position || "top";
    config.options.plugins.legend.labels = Object.assign({ boxWidth: 14, boxHeight: 10 }, config.options.plugins.legend.labels || {});

    (config.data.datasets || []).forEach((ds) => {
      if (ds.pointRadius == null) ds.pointRadius = 2;
      if (ds.pointHoverRadius == null) ds.pointHoverRadius = 4;
      if (ds.borderWidth == null) ds.borderWidth = 2;
      // NO forzamos fill (para respetar colores default sin manchar); si quieres fill tipo área, lo activamos luego.
    });
  }

  return config;
}

/* =========================
   Carga de archivos
   ========================= */
document.getElementById("fileProduccion").addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      wbProd = XLSX.read(evt.target.result, { type: "array", cellDates: true });
      prodLoaded = true;
      setStatus("Archivo GEN cargado. Selecciona la hoja/mes.");
      populateMonths();
      updateControls();
    } catch (err) {
      console.error(err);
      wbProd = null;
      prodLoaded = false;
      setStatus("Error al leer el archivo de producción (GEN).", true);
      updateControls();
    }
  };
  reader.readAsArrayBuffer(file);
});

document.getElementById("fileAforo").addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      wbAforo = XLSX.read(evt.target.result, { type: "array", cellDates: true });
      aforoLoaded = true;
      setStatus('Archivo de aforo cargado. (Se usa la hoja "Sondas 00 00 hrs")');
      updateControls();
    } catch (err) {
      console.error(err);
      wbAforo = null;
      aforoLoaded = false;
      setStatus("Error al leer el archivo de aforo.", true);
      updateControls();
    }
  };
  reader.readAsArrayBuffer(file);
});

function populateMonths() {
  elMes.innerHTML = "";
  const opt0 = document.createElement("option");
  opt0.value = "";
  opt0.textContent = "-- Seleccionar --";
  elMes.appendChild(opt0);

  (wbProd?.SheetNames || []).forEach((name) => {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name.trim();
    elMes.appendChild(opt);
  });

  elMes.disabled = !prodLoaded;
}

elMes.addEventListener("change", updateControls);

function updateControls() {
  elMes.disabled = !prodLoaded;

  const ok = prodLoaded && aforoLoaded && !!elMes.value;
  elBtnProcesar.disabled = !ok;

  const hasLast = !!lastData;
  elBtnPdf.disabled = !hasLast;
  elPdfTipo.disabled = !hasLast;
}

/* =========================
   Extracción de datos
   ========================= */
const COL = {
  FECHA: 1, // B
  H_U1: 4, // E
  H_U2: 10, // K
  AUX: 24, // Y
  LANEC: 25, // Z
  GRACA: 30, // AE
  ETOTAL: 35, // AJ
  E_U1: 36, // AK
  E_U2: 37, // AL
  REND: 38, // AM (kWh/gal)
  HFO_TOT: 43, // AR (gal)
  DO_TOT: 46 // AU (gal)
};

// Mes/año dominante para ignorar filas del mes anterior dentro de la hoja
function dominantMonthYear(rows) {
  const counts = {};
  for (let i = 0; i < rows.length; i++) {
    const d = parseExcelDate(rows[i]?.[COL.FECHA]);
    if (!d) continue;
    const key = `${d.getFullYear()}-${d.getMonth() + 1}`;
    counts[key] = (counts[key] || 0) + 1;
  }
  let bestKey = null,
    best = 0;
  Object.entries(counts).forEach(([k, c]) => {
    if (c > best) {
      best = c;
      bestKey = k;
    }
  });
  if (!bestKey) return null;
  const [y, m] = bestKey.split("-").map(Number);
  return { year: y, month: m };
}

function extractProduction(sheetName) {
  const ws = wbProd?.Sheets?.[sheetName];
  if (!ws) return null;

  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  if (!rows?.length) return null;

  const dom = dominantMonthYear(rows);
  if (!dom) return null;

  const labels = [];
  const dates = [];

  const etotal = [];
  const aux = [];
  const lanec = [];
  const graca = [];
  const e_u1 = [];
  const e_u2 = [];
  const h_u1 = [];
  const h_u2 = [];
  const pot_u1 = [];
  const pot_u2 = [];
  const rend = [];

  const hfoTot = [];
  const doTot = [];
  const hfoG1 = [];
  const hfoG2 = [];
  const doG1 = [];
  const doG2 = [];

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const d = parseExcelDate(r?.[COL.FECHA]);
    if (!d) continue;

    const m = d.getMonth() + 1;
    const y = d.getFullYear();
    if (m !== dom.month || y !== dom.year) continue;

    const eTot = toNumber(r[COL.ETOTAL]);
    const e1 = toNumber(r[COL.E_U1]);
    const e2 = toNumber(r[COL.E_U2]);
    const a = toNumber(r[COL.AUX]);
    const l = toNumber(r[COL.LANEC]);
    const g = toNumber(r[COL.GRACA]);
    const hu1 = toNumber(r[COL.H_U1]);
    const hu2 = toNumber(r[COL.H_U2]);
    const re = toNumber(r[COL.REND]);

    const hf = toNumber(r[COL.HFO_TOT]);
    const dox = toNumber(r[COL.DO_TOT]);

    // Fila vacía -> saltar
    if (!eTot && !e1 && !e2 && !a && !l && !g && !hu1 && !hu2 && !re && !hf && !dox) continue;

    labels.push(dayLabel(d));
    dates.push(d);

    etotal.push(eTot || null);
    aux.push(a || null);
    lanec.push(l || null);
    graca.push(g || null);
    e_u1.push(e1 || null);
    e_u2.push(e2 || null);
    h_u1.push(hu1 || null);
    h_u2.push(hu2 || null);
    rend.push(re || null);

    const pu1 = e1 > 0 && hu1 > 0 ? e1 / hu1 : null;
    const pu2 = e2 > 0 && hu2 > 0 ? e2 / hu2 : null;
    pot_u1.push(pu1);
    pot_u2.push(pu2);

    hfoTot.push(hf || 0);
    doTot.push(dox || 0);

    const eSum = (e1 || 0) + (e2 || 0);
    if (eSum > 0) {
      hfoG1.push((hf || 0) * ((e1 || 0) / eSum));
      hfoG2.push((hf || 0) * ((e2 || 0) / eSum));
      doG1.push((dox || 0) * ((e1 || 0) / eSum));
      doG2.push((dox || 0) * ((e2 || 0) / eSum));
    } else {
      hfoG1.push(0);
      hfoG2.push(0);
      doG1.push(0);
      doG2.push(0);
    }
  }

  return {
    sheetName,
    targetMonth: dom.month,
    targetYear: dom.year,
    labels,
    dates,
    // producción
    etotal,
    aux,
    lanec,
    graca,
    e_u1,
    e_u2,
    h_u1,
    h_u2,
    pot_u1,
    pot_u2,
    rend,
    // combustible
    hfoTot,
    doTot,
    hfoG1,
    hfoG2,
    doG1,
    doG2
  };
}

// AFORO: hoja "Sondas 00 00 hrs" con tags y Cisterna 2 en Tipo
function extractAforo(targetMonth, targetYear) {
  const ws = wbAforo?.Sheets?.["Sondas 00 00 hrs"] || null;
  if (!ws) return null;

  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  if (!rows?.length) return null;

  const labels = [];
  const t601 = [];
  const t602 = [];
  const t610 = [];
  const t611 = [];
  const cisterna2 = [];

  let currentDate = null;

  const COL_FECHA = 0;
  const COL_TAG = 1;
  const COL_TIPO = 2;
  const COL_VOL = 5;

  for (let i = 2; i < rows.length; i++) {
    const r = rows[i];

    const maybeDate = parseExcelDate(r?.[COL_FECHA]);
    if (maybeDate) currentDate = maybeDate;
    if (!currentDate) continue;

    const m = currentDate.getMonth() + 1;
    const y = currentDate.getFullYear();
    if (m !== targetMonth || y !== targetYear) continue;

    const label = dayLabel(currentDate);

    const rawTag = String(r?.[COL_TAG] ?? "").trim().toUpperCase();
    const tagNorm = rawTag.replace(/\s+/g, ""); // "T610 V" -> "T610V"
    const tipoNorm = String(r?.[COL_TIPO] ?? "")
      .trim()
      .toUpperCase()
      .replace(/\s+/g, "");

    const vol = toNumber(r?.[COL_VOL]);
    if (!vol) continue;

    function idxForLabel() {
      let idx = labels.indexOf(label);
      if (idx === -1) {
        labels.push(label);
        t601.push(null);
        t602.push(null);
        t610.push(null);
        t611.push(null);
        cisterna2.push(null);
        idx = labels.length - 1;
      }
      return idx;
    }

    if (tagNorm.startsWith("T601")) {
      const idx = idxForLabel();
      t601[idx] = vol;
    } else if (tagNorm.startsWith("T602")) {
      const idx = idxForLabel();
      t602[idx] = vol;
    } else if (tagNorm.startsWith("T610")) {
      const idx = idxForLabel();
      t610[idx] = vol;
    } else if (tagNorm.startsWith("T611")) {
      const idx = idxForLabel();
      t611[idx] = vol;
    } else if (tipoNorm === "CISTERNA2") {
      const idx = idxForLabel();
      cisterna2[idx] = vol;
    }
  }

  return { labels, t601, t602, t610, t611, cisterna2 };
}

/* =========================
   Render KPI cards
   ========================= */
function renderKpis(container, items) {
  if (!container) return;
  container.innerHTML = "";
  items.forEach((it) => {
    const div = document.createElement("div");
    div.className = "kpi";
    div.innerHTML = `
      <div class="kpi-title">${it.title}</div>
      <div class="kpi-value">${it.value}</div>
      ${it.sub ? `<div class="kpi-sub">${it.sub}</div>` : ""}
    `;
    container.appendChild(div);
  });
}

/* =========================
   Botón PROCESAR
   ========================= */
elBtnProcesar.addEventListener("click", () => {
  const sheet = elMes.value;
  if (!sheet) return;

  try {
    const prod = extractProduction(sheet);
    if (!prod || !prod.labels.length) {
      setStatus("No se encontraron datos válidos en esa hoja/mes.", true);
      return;
    }

    const aforo = extractAforo(prod.targetMonth, prod.targetYear);
    if (!aforo || !aforo.labels.length) {
      setStatus('Producción OK. Nota: no se hallaron datos de aforo en "Sondas 00 00 hrs" para ese mes.', false);
    } else {
      setStatus(`Procesado: "${sheet.trim()}" (mes ${String(prod.targetMonth).padStart(2, "0")}/${prod.targetYear}).`);
    }

    // Resúmenes
    const sum = (arr) => (arr || []).reduce((a, b) => a + (b || 0), 0);
    const avgNoZero = (arr) => {
      const f = (arr || []).filter((v) => v != null && v !== 0);
      if (!f.length) return null;
      return f.reduce((a, b) => a + b, 0) / f.length;
    };

    const sumEt = sum(prod.etotal);
    const sumU1 = sum(prod.e_u1);
    const sumU2 = sum(prod.e_u2);
    const sumLan = sum(prod.lanec);
    const sumGra = sum(prod.graca);
    const sumAux = sum(prod.aux);
    const sumHU1 = sum(prod.h_u1);
    const sumHU2 = sum(prod.h_u2);
    const potPromU1 = sumU1 > 0 && sumHU1 > 0 ? sumU1 / sumHU1 : null;
    const potPromU2 = sumU2 > 0 && sumHU2 > 0 ? sumU2 / sumHU2 : null;
    const eficProm = avgNoZero(prod.rend);

    const totalHfo = sum(prod.hfoTot);
    const totalDo = sum(prod.doTot);
    const dias = prod.labels.length;

    const resumen = {
      energiaTotalMWh: sumEt / 1000,
      energiaU1MWh: sumU1 / 1000,
      energiaU2MWh: sumU2 / 1000,
      energiaLanecMWh: sumLan / 1000,
      energiaGracaMWh: sumGra / 1000,
      energiaAuxMWh: sumAux / 1000,
      horasU1: sumHU1,
      horasU2: sumHU2,
      potPromU1,
      potPromU2,
      eficProm,
      hfoGal: totalHfo,
      doGal: totalDo,
      dias
    };

    lastData = { sheetName: sheet, month: prod.targetMonth, year: prod.targetYear, prod, aforo, resumen };
    updateControls();

    // Render módulos (aislar errores para no bloquear todo el procesamiento)
    try { renderModuleProduccion(prod, aforo, resumen); }
    catch (e) { console.error("Render Producción:", e); }

    try { renderModuleCombustible(prod, aforo, resumen); }
    catch (e) {
      console.error("Render Combustible:", e);
      setStatus(`Procesado "${sheet.trim()}", pero hubo un error al graficar Combustible: ${e?.message || e}`, true);
      return;
    }
} catch (err) {
    console.error(err);
    setStatus(`Error al procesar: ${err?.message || err}. Revise estructura de archivos/hojas.`, true);
  }
});

/* =========================
   Render: PRODUCCIÓN
   ========================= */
function renderModuleProduccion(prod, aforo, resumen) {
  // KPIs
  renderKpis(elKpiProd, [
    { title: "Energía total", value: `${formatNumber(resumen.energiaTotalMWh, 1)} MWh`, sub: `U1: ${formatNumber(resumen.energiaU1MWh, 1)} • U2: ${formatNumber(resumen.energiaU2MWh, 1)}` },
    { title: "Energía por cliente", value: `LANEC ${formatNumber(resumen.energiaLanecMWh, 1)} MWh`, sub: `GRACA ${formatNumber(resumen.energiaGracaMWh, 1)} • AUX ${formatNumber(resumen.energiaAuxMWh, 1)}` },
    { title: "Horas", value: `U1 ${formatNumber(resumen.horasU1, 1)} h`, sub: `U2 ${formatNumber(resumen.horasU2, 1)} h` },
    { title: "Potencia promedio", value: `U1 ${formatNumber(resumen.potPromU1, 0)} kW`, sub: `U2 ${formatNumber(resumen.potPromU2, 0)} kW` },
    { title: "Eficiencia", value: `${resumen.eficProm != null ? formatNumber(resumen.eficProm, 2) : "-"} kWh/gal`, sub: `Mes ${String(prod.targetMonth).padStart(2, "0")}/${prod.targetYear}` }
  ]);

  // Charts
  createOrUpdateChart("chartEnergiaTotal", {
    type: "line",
    data: {
      labels: prod.labels,
      datasets: [{ label: "Energía total (kWh)", data: prod.etotal }]
    },
    options: commonLineOptions("kWh")
  });

  createOrUpdateChart("chartEnergiaAlimentadores", {
    type: "line",
    data: {
      labels: prod.labels,
      datasets: [
        { label: "LANEC (kWh)", data: prod.lanec },
        { label: "GRACA (kWh)", data: prod.graca },
        { label: "Auxiliares (kWh)", data: prod.aux }
      ]
    },
    options: commonLineOptions("kWh")
  });

  createOrUpdateChart("chartEnergiaUnidades", {
    type: "line",
    data: {
      labels: prod.labels,
      datasets: [
        { label: "Unidad 1 (kWh)", data: prod.e_u1 },
        { label: "Unidad 2 (kWh)", data: prod.e_u2 }
      ]
    },
    options: commonLineOptions("kWh")
  });

  createOrUpdateChart("chartPotenciaProm", {
    type: "line",
    data: {
      labels: prod.labels,
      datasets: [
        { label: "Potencia promedio U1 (kW)", data: prod.pot_u1 },
        { label: "Potencia promedio U2 (kW)", data: prod.pot_u2 }
      ]
    },
    options: commonLineOptions("kW")
  });

  createOrUpdateChart("chartHorasOp", {
    type: "line",
    data: {
      labels: prod.labels,
      datasets: [
        { label: "Horas U1 (h)", data: prod.h_u1 },
        { label: "Horas U2 (h)", data: prod.h_u2 }
      ]
    },
    options: commonLineOptions("h")
  });

  createOrUpdateChart("chartEficiencia", {
    type: "line",
    data: {
      labels: prod.labels,
      datasets: [{ label: "Eficiencia (kWh/gal)", data: prod.rend }]
    },
    options: commonLineOptions("kWh/gal")
  });

  // Tanques (si hay datos)
  if (aforo && aforo.labels?.length) {
    createOrUpdateChart("chartTanquesAll", {
      type: "line",
      data: {
        labels: aforo.labels,
        datasets: [
          { label: "T601 (HFO, gal)", data: aforo.t601 },
          { label: "T602 (HFO, gal)", data: aforo.t602 },
          { label: "T610 (Diesel, gal)", data: aforo.t610 },
          { label: "T611 (Diesel, gal)", data: aforo.t611 },
          { label: "Cisterna 2 (gal)", data: aforo.cisterna2 }
        ]
      },
      options: commonLineOptions("gal")
    });
  } else {
    // limpiar gráfico si no hay aforo
    if (charts["chartTanquesAll"]) charts["chartTanquesAll"].destroy();
  }
}

/* =========================
   Render: COMBUSTIBLE
   ========================= */
function renderModuleCombustible(prod, aforo, resumen) {
  // Encabezado (como el PDF)
  const elMes = document.getElementById("fuelMes");
  const elFecha = document.getElementById("fuelFecha");
  if (elMes) elMes.textContent = `${String(prod.targetMonth).padStart(2, "0")}/${prod.targetYear}`;
  if (elFecha) elFecha.textContent = new Date().toLocaleDateString("es-EC");

  renderKpis(elKpiComb, [
    { title: "HFO total Consumido del mes", value: `${formatNumber(resumen.hfoGal, 0)} gal` },
    { title: "Diesel total Consumido del mes", value: `${formatNumber(resumen.doGal, 0)} gal` },
    { title: "Días con registro", value: `${formatNumber(resumen.dias, 0)}` }
  ]);

  createOrUpdateChart("chartTotal", {
    type: "line",
    data: {
      labels: prod.labels,
      datasets: [
        { label: "HFO total (gal)", data: prod.hfoTot },
        { label: "Diesel total (gal)", data: prod.doTot }
      ]
    },
    options: commonLineOptions("gal")
  });

  createOrUpdateChart("chartHFOUnidad", {
    type: "line",
    data: {
      labels: prod.labels,
      datasets: [
        { label: "HFO G1 (gal)", data: prod.hfoG1 },
        { label: "HFO G2 (gal)", data: prod.hfoG2 }
      ]
    },
    options: commonLineOptions("gal")
  });

  createOrUpdateChart("chartDOUnidad", {
    type: "line",
    data: {
      labels: prod.labels,
      datasets: [
        { label: "Diesel G1 (gal)", data: prod.doG1 },
        { label: "Diesel G2 (gal)", data: prod.doG2 }
      ]
    },
    options: commonLineOptions("gal")
  });

  if (aforo && aforo.labels?.length) {
    createOrUpdateChart("chartHFOtanques", {
      type: "line",
      data: {
        labels: aforo.labels,
        datasets: [
          { label: "T601 (HFO, gal)", data: aforo.t601 },
          { label: "T602 (HFO, gal)", data: aforo.t602 }
        ]
      },
      options: commonLineOptions("gal")
    });

    createOrUpdateChart("chartDOtanques", {
      type: "line",
      data: {
        labels: aforo.labels,
        datasets: [
          { label: "T610 (Diesel, gal)", data: aforo.t610 },
          { label: "T611 (Diesel, gal)", data: aforo.t611 }
        ]
      },
      options: commonLineOptions("gal")
    });

    createOrUpdateChart("chartCisterna", {
      type: "line",
      data: {
        labels: aforo.labels,
        datasets: [{ label: "Cisterna 2 (gal)", data: aforo.cisterna2 }]
      },
      options: commonLineOptions("gal")
    });
  } else {
    ["chartHFOtanques", "chartDOtanques", "chartCisterna"].forEach((id) => {
      if (charts[id]) charts[id].destroy();
    });
  }
}

function commonLineOptions(yTitle) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    interaction: { mode: "index", intersect: false },
    plugins: { legend: { position: "top" } },
    scales: {
      x: { title: { display: true, text: "Día" } },
      y: { title: { display: true, text: yTitle } }
    }
  };
}

/* =========================
   Exportar PDF
   ========================= */
elBtnPdf.addEventListener("click", async () => {
  if (!lastData) return;

  try {
    if (elPdfTipo.value === "combustible") {
      exportPdfCombustible(lastData);
    } else {
      exportPdfProduccion(lastData);
    }
  } catch (err) {
    console.error(err);
    setStatus("Error al generar el PDF.", true);
  }
});

function exportPdfCombustible(ctx) {
  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });

  const mes = String(ctx.month).padStart(2, "0");
  const anio = ctx.year;
  const hoy = new Date();
  const fechaStr = `${String(hoy.getDate()).padStart(2, "0")}/${String(hoy.getMonth() + 1).padStart(2, "0")}/${hoy.getFullYear()}`;

  // Header
  pdf.setFillColor(15, 23, 42);
  pdf.rect(0, 0, 210, 18, "F");
  pdf.setTextColor(255, 255, 255);
  pdf.setFontSize(12);
  pdf.text("Central El Morro - Informe Gerencial de Combustible", 6, 11);

  pdf.setFontSize(9);
  pdf.setTextColor(230, 230, 230);
  pdf.text(`Mes analizado: ${mes}/${anio}      Fecha emisión: ${fechaStr}`, 6, 16);

  // KPI box
  pdf.setTextColor(0, 0, 0);
  pdf.setFontSize(9);
  pdf.setDrawColor(60, 60, 60);
  pdf.rect(6, 22, 198, 14);

  pdf.text(`HFO total mes: ${formatNumber(ctx.resumen.hfoGal, 0)} gal`, 10, 28);
  pdf.text(`Diesel total mes: ${formatNumber(ctx.resumen.doGal, 0)} gal`, 75, 28);
  pdf.text(`Días con registro: ${ctx.resumen.dias}`, 150, 28);

  // Layout 2x3
  const positions = [
    { id: "chartTotal", x: 6, y: 40, w: 96, h: 50, title: "Consumo total HFO vs Diesel" },
    { id: "chartHFOUnidad", x: 108, y: 40, w: 96, h: 50, title: "Consumo HFO por unidad" },
    { id: "chartDOUnidad", x: 6, y: 95, w: 96, h: 50, title: "Consumo Diesel por unidad" },
    { id: "chartHFOtanques", x: 108, y: 95, w: 96, h: 50, title: "Tanques HFO (T601/T602)" },
    { id: "chartDOtanques", x: 6, y: 150, w: 96, h: 50, title: "Tanques Diesel (T610/T611)" },
    { id: "chartCisterna", x: 108, y: 150, w: 96, h: 50, title: "Cisterna 2 (sludge)" }
  ];

  pdf.setFontSize(8);
  positions.forEach((p) => {
    const canvas = document.getElementById(p.id);
    if (!canvas) return;
    const img = canvas.toDataURL("image/png", 1.0);
    pdf.text(p.title, p.x, p.y - 2);
    pdf.addImage(img, "PNG", p.x, p.y, p.w, p.h);
  });

  pdf.save(`Informe_Combustible_Central_El_Morro_${mes}-${anio}.pdf`);
}

async function exportPdfProduccion(ctx) {
  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF("portrait", "pt", "a4");

  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();
  const margin = 40;
  let cursorY = 40;

  pdf.setFontSize(16);
  pdf.setFont("helvetica", "bold");
  pdf.text("CENTRAL EL MORRO – RESUMEN GERENCIAL", margin, cursorY);
  cursorY += 18;

  pdf.setFontSize(11);
  pdf.setFont("helvetica", "normal");
  pdf.text(`Reporte mensual – ${ctx.sheetName}`, margin, cursorY);
  cursorY += 14;

  const r = ctx.resumen;
  pdf.setFontSize(10);
  const lines = [
    `Energía total: ${formatNumber(r.energiaTotalMWh, 1)} MWh (U1: ${formatNumber(r.energiaU1MWh, 1)} • U2: ${formatNumber(r.energiaU2MWh, 1)})`,
    `Clientes: LANEC ${formatNumber(r.energiaLanecMWh, 1)} MWh • GRACA ${formatNumber(r.energiaGracaMWh, 1)} MWh • AUX ${formatNumber(r.energiaAuxMWh, 1)} MWh`,
    `Horas: U1 ${formatNumber(r.horasU1, 1)} h • U2 ${formatNumber(r.horasU2, 1)} h`,
    `Potencia promedio: U1 ${formatNumber(r.potPromU1, 0)} kW • U2 ${formatNumber(r.potPromU2, 0)} kW`,
    `Eficiencia: ${r.eficProm != null ? formatNumber(r.eficProm, 2) + " kWh/gal" : "-"}`
  ];
  lines.forEach((t) => {
    pdf.text(t, margin, cursorY);
    cursorY += 14;
  });
  cursorY += 10;

  const chartIds = [
    "chartEnergiaTotal",
    "chartEnergiaAlimentadores",
    "chartEnergiaUnidades",
    "chartPotenciaProm",
    "chartHorasOp",
    "chartEficiencia",
    "chartTanquesAll"
  ];

  for (const id of chartIds) {
    const canvas = document.getElementById(id);
    if (!canvas) continue;

    // Aumentar resolución temporalmente (si existe chart)
    const chartObj = charts[id];
    let originalW = null, originalH = null;
    if (chartObj) {
      originalW = chartObj.width;
      originalH = chartObj.height;
      chartObj.resize(1000, 360);
    }

    const img = canvas.toDataURL("image/png", 1.0);
    const canvasW = canvas.width || 1000;
    const canvasH = canvas.height || 360;
    const aspect = canvasH / canvasW;

    const imgW = pageWidth - margin * 2;
    const imgH = imgW * aspect;

    if (cursorY + imgH > pageHeight - margin) {
      pdf.addPage();
      cursorY = margin;
    }
    pdf.addImage(img, "PNG", margin, cursorY, imgW, imgH);
    cursorY += imgH + 18;

    if (chartObj && originalW && originalH) chartObj.resize(originalW, originalH);
  }

  pdf.save("Resumen_Gerencial_Central_El_Morro.pdf");
}

/* =========================
   Init status
   ========================= */
setStatus("Cargue el archivo GEN y el archivo de aforo. Luego seleccione el mes y procese.");
updateControls();
