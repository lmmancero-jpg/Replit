import * as XLSX from "xlsx";
import {
  CONFIG, COSTOS_VARIABLES, COSTO_FIJO_MENSUAL_POR_UNIDAD, CBMT_U1_MENSUAL,
  P_CONTR_LANEC, P_CONTR_GRACA, P_CONTR_TOT,
  posNum, fmt, excelDateKey, getDaysInMonth, getMesNombreES,
  getProdSheetAndRows, rptHeader, seccion,
} from "./reportEngine";
import { INVOICE_CATEGORIES, INVOICE_CATEGORY_LABELS, FUEL_TYPE_LABELS, type Invoice, type InvoiceCategory } from "@shared/schema";

// ── Tipos ─────────────────────────────────────────────────────────────────────

export interface MonthlyProductionSummary {
  totalKwh: number;
  lanecKwh: number;
  gracaKwh: number;
  auxKwh: number;
  lan_fact: number;
  gra_fact: number;
  tot_gen: number;
  ultimoDia: number;
  textoPeriodo: string;
  diasMes: number;
  hfoGal: number;
  doGal: number;
  kWhHFO: number;
  kWhDO: number;
}

export interface InvoiceCategorySummary {
  totalByCategory: Record<string, number>;
  cvRealByCategory: Record<string, number>;
  grandTotal: number;
  cvTotalReal: number;
}

// ── Helpers de producción ─────────────────────────────────────────────────────

export function getMonthlyProductionSummary(
  wbProd: XLSX.WorkBook,
  mesStr: string
): MonthlyProductionSummary {
  const [yS, mS] = (mesStr || "").split("-");
  const year = parseInt(yS, 10);
  const monthIndex = parseInt(mS, 10) - 1;
  if (isNaN(year) || isNaN(monthIndex)) throw new Error("Mes inválido: " + mesStr);

  const diasMes = getDaysInMonth(year, monthIndex);
  const fechaCorte = new Date(year, monthIndex + 1, 0);
  const { rows } = getProdSheetAndRows(wbProd, fechaCorte);

  let lan = 0, gra = 0, aux = 0, ultimoDia = 0;
  let hfoGal = 0, doGal = 0;

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
    hfoGal += posNum(r[CONFIG.COL_HFO_GAL]);
    doGal  += posNum(r[CONFIG.COL_DO_GAL]);
  }

  const tot_cli = lan + gra;
  const aux_lan = tot_cli > 0 ? aux * (lan / tot_cli) : 0;
  const aux_gra = tot_cli > 0 ? aux * (gra / tot_cli) : 0;
  const lan_fact = lan + aux_lan;
  const gra_fact = gra + aux_gra;
  const tot_gen = lan_fact + gra_fact;

  const fuelTotal = hfoGal + doGal;
  const kWhHFO = fuelTotal > 0 ? tot_gen * (hfoGal / fuelTotal) : 0;
  const kWhDO  = fuelTotal > 0 ? tot_gen * (doGal  / fuelTotal) : 0;

  const mesNombre = getMesNombreES(monthIndex);
  const textoPeriodo = ultimoDia > 0
    ? `${mesNombre} ${year} (hasta el día ${ultimoDia})`
    : `${mesNombre} ${year}`;

  return {
    totalKwh: tot_cli + aux,
    lanecKwh: lan,
    gracaKwh: gra,
    auxKwh: aux,
    lan_fact,
    gra_fact,
    tot_gen,
    ultimoDia,
    textoPeriodo,
    diasMes,
    hfoGal,
    doGal,
    kWhHFO,
    kWhDO,
  };
}

export function getInvoiceCategorySummary(
  invoiceList: Invoice[],
  productionKwh: number
): InvoiceCategorySummary {
  const totalByCategory: Record<string, number> = {};
  const cvRealByCategory: Record<string, number> = {};

  for (const cat of INVOICE_CATEGORIES) {
    totalByCategory[cat] = 0;
  }

  for (const inv of invoiceList) {
    const cat = inv.category;
    if (cat in totalByCategory) {
      totalByCategory[cat] += parseFloat(inv.total ?? "0");
    }
  }

  let grandTotal = 0;
  for (const cat of INVOICE_CATEGORIES) {
    grandTotal += totalByCategory[cat];
    cvRealByCategory[cat] = productionKwh > 0
      ? totalByCategory[cat] / productionKwh
      : 0;
  }

  const cvTotalReal = productionKwh > 0 ? grandTotal / productionKwh : 0;

  return { totalByCategory, cvRealByCategory, grandTotal, cvTotalReal };
}

export function getContractualVariableCosts(): Record<string, number> {
  return { ...COSTOS_VARIABLES };
}

// ── Fijo helpers ──────────────────────────────────────────────────────────────

function calcFixed(diasMes: number, diasFallaU1: number, diasFallaU2: number) {
  const dispU1 = Math.max(0, (diasMes - diasFallaU1) / diasMes);
  const dispU2 = Math.max(0, (diasMes - diasFallaU2) / diasMes);
  const fijoU1 = COSTO_FIJO_MENSUAL_POR_UNIDAD * dispU1;
  const fijoU2 = COSTO_FIJO_MENSUAL_POR_UNIDAD * dispU2;
  const fijoTotal = fijoU1 + fijoU2;
  const cbmtU1 = CBMT_U1_MENSUAL * dispU1;

  const factorLan = P_CONTR_TOT > 0 ? P_CONTR_LANEC / P_CONTR_TOT : 0;
  const factorGra = P_CONTR_TOT > 0 ? P_CONTR_GRACA / P_CONTR_TOT : 0;

  return { dispU1, dispU2, fijoU1, fijoU2, fijoTotal, cbmtU1, factorLan, factorGra };
}

// ── Informe clientes ──────────────────────────────────────────────────────────

export function buildClientBillingReport(params: {
  wbProd: XLSX.WorkBook;
  mesStr: string;
  diasFallaU1: number;
  diasFallaU2: number;
  invoiceCombTotal: number;
  hasProduction: boolean;
}): string {
  const { wbProd, mesStr, diasFallaU1, diasFallaU2, invoiceCombTotal, hasProduction } = params;

  const prod = getMonthlyProductionSummary(wbProd, mesStr);
  const { lan_fact, gra_fact, tot_gen, lanecKwh: lan, gracaKwh: gra, auxKwh: aux,
    textoPeriodo, diasMes } = prod;

  const aux_lan = tot_gen > 0 ? aux * (lan / (lan + gra || 1)) : 0;
  const aux_gra = tot_gen > 0 ? aux * (gra / (lan + gra || 1)) : 0;

  const fixed = calcFixed(diasMes, diasFallaU1, diasFallaU2);
  const { dispU1, dispU2, fijoU1, fijoU2, fijoTotal, cbmtU1, factorLan, factorGra } = fixed;

  const cvCombAjustado = tot_gen > 0 && hasProduction
    ? invoiceCombTotal / tot_gen
    : COSTOS_VARIABLES.combustible_transporte;

  const costosEfectivos: Record<string, number> = {
    ...COSTOS_VARIABLES,
    combustible_transporte: cvCombAjustado,
  };
  const costoVarTotal = Object.values(costosEfectivos).reduce((a, b) => a + b, 0);

  function subtotalVar(kwh: number): Record<string, number> {
    const r: Record<string, number> = {};
    for (const [k, v] of Object.entries(costosEfectivos)) r[k] = kwh * v;
    return r;
  }

  const varLanBy = subtotalVar(lan_fact);
  const varGraBy = subtotalVar(gra_fact);
  const varTotBy = subtotalVar(tot_gen);
  const varLanTotal = lan_fact * costoVarTotal;
  const varGraTotal = gra_fact * costoVarTotal;
  const varTotTotal = tot_gen * costoVarTotal;

  const fijoLanU1 = fijoU1 * factorLan, fijoLanU2 = fijoU2 * factorLan;
  const fijoGraU1 = fijoU1 * factorGra, fijoGraU2 = fijoU2 * factorGra;
  const fijoLan = fijoLanU1 + fijoLanU2, fijoGra = fijoGraU1 + fijoGraU2;
  const fijoTotU1 = fijoLanU1 + fijoGraU1, fijoTotU2 = fijoLanU2 + fijoGraU2;

  const cbmtLanU1 = cbmtU1 * factorLan, cbmtGraU1 = cbmtU1 * factorGra;

  const totalLan = varLanTotal + fijoLan + cbmtLanU1;
  const totalGra = varGraTotal + fijoGra + cbmtGraU1;
  const totalTot = totalLan + totalGra;

  const precioLan = lan_fact > 0 ? totalLan / lan_fact : 0;
  const precioGra = gra_fact > 0 ? totalGra / gra_fact : 0;
  const precioTot = tot_gen > 0 ? totalTot / tot_gen : 0;

  const tot_cli = lan + gra;

  function tablaCliente(
    secLabel: string, titulo: string, nombre: string,
    energiaConsumida: number, auxAsig: number, totalFact: number,
    varBy: Record<string, number>, varTotal: number,
    fijoAsigU1: number, fijoAsigU2: number, fijoAsig: number,
    cbmtAsigU1: number,
    totalUSD: number, precioFinal: number
  ): string {
    const energiaLabel = nombre === "TOTAL"
      ? "Energía consumida total (LANEC + GRACA)"
      : `Energía consumida – ${nombre}`;
    const totalLabel = nombre === "TOTAL"
      ? "Energía total a facturar (+auxiliares)"
      : `Total facturable ${nombre} (+aux.)`;
    const totalFijo = fijoAsig + cbmtAsigU1;
    return `
<div class="rpt-section-title"><span class="rpt-section-num">${secLabel}</span>${titulo}</div>
<table class="data-table">
<thead><tr><th>Rubro</th><th>P. Unit [USD/kWh]</th><th>Subtotal [USD]</th></tr></thead>
<tbody>
<tr class="rpt-row-grupo"><td class="label" colspan="3">Energía facturable</td></tr>
<tr><td class="label">${energiaLabel}</td><td>—</td><td class="num">${fmt(energiaConsumida)} kWh</td></tr>
<tr><td class="label">Auxiliares asignados (proporcional)</td><td>—</td><td class="num">${fmt(auxAsig)} kWh</td></tr>
<tr class="rpt-row-total"><td class="label">${totalLabel}</td><td>—</td><td class="num hi">${fmt(totalFact)} kWh</td></tr>
<tr class="rpt-row-grupo"><td class="label" colspan="3">Costos variables de producción</td></tr>
<tr><td class="label">Combustible + Transporte <small style="color:#6b7280">(ajustado por facturas reales)</small></td><td class="num">${fmt(costosEfectivos.combustible_transporte, 4)}</td><td class="num">$ ${fmt(varBy.combustible_transporte)}</td></tr>
<tr><td class="label">Lubricantes + Químicos</td><td class="num">${fmt(costosEfectivos.lubricantes_quimicos, 4)}</td><td class="num">$ ${fmt(varBy.lubricantes_quimicos)}</td></tr>
<tr><td class="label">Agua + Insumos</td><td class="num">${fmt(costosEfectivos.agua_insumos, 4)}</td><td class="num">$ ${fmt(varBy.agua_insumos)}</td></tr>
<tr><td class="label">Repuestos Mantenimiento Predictivo</td><td class="num">${fmt(costosEfectivos.repuestos_predictivo, 4)}</td><td class="num">$ ${fmt(varBy.repuestos_predictivo)}</td></tr>
<tr><td class="label">Impacto Ambiental</td><td class="num">${fmt(costosEfectivos.impacto_ambiental, 4)}</td><td class="num">$ ${fmt(varBy.impacto_ambiental)}</td></tr>
<tr><td class="label">Servicios Auxiliares</td><td class="num">${fmt(costosEfectivos.servicios_auxiliares, 4)}</td><td class="num">$ ${fmt(varBy.servicios_auxiliares)}</td></tr>
<tr><td class="label">Margen Variable</td><td class="num">${fmt(costosEfectivos.margen_variable, 4)}</td><td class="num">$ ${fmt(varBy.margen_variable)}</td></tr>
<tr class="rpt-row-total"><td class="label"><strong>Subtotal costo variable</strong></td><td class="num"><strong>${fmt(costoVarTotal, 4)}</strong></td><td class="num"><strong>$ ${fmt(varTotal)}</strong></td></tr>
<tr class="rpt-row-grupo"><td class="label" colspan="3">Costos fijos</td></tr>
<tr><td class="label">Costo fijo por disponibilidad U1</td><td>—</td><td class="num">$ ${fmt(fijoAsigU1)}</td></tr>
<tr><td class="label">Costo fijo por disponibilidad U2</td><td>—</td><td class="num">$ ${fmt(fijoAsigU2)}</td></tr>
<tr><td class="label">CBMT Unidad 1</td><td>—</td><td class="num">$ ${fmt(cbmtAsigU1)}</td></tr>
<tr class="rpt-row-total"><td class="label"><strong>Subtotal costo fijo asignado</strong></td><td>—</td><td class="num"><strong>$ ${fmt(totalFijo)}</strong></td></tr>
<tr class="rpt-row-grand"><td class="label"><strong>TOTAL A FACTURAR</strong></td><td class="num"><strong>USD/kWh: ${fmt(precioFinal, 4)}</strong></td><td class="num"><strong>$ ${fmt(totalUSD)} + IVA</strong></td></tr>
</tbody></table>`;
  }

  let html = rptHeader("Informe de Facturación – Clientes", textoPeriodo);

  if (!hasProduction) {
    html += `<div class="rpt-notice rpt-notice-warn">⚠ Sin archivo de producción cargado — los costos de combustible muestran la tarifa contractual de referencia.</div>`;
  }
  if (invoiceCombTotal === 0) {
    html += `<div class="rpt-notice rpt-notice-warn">⚠ No hay facturas de Combustible + Transporte para este período — se usa tarifa contractual como referencia.</div>`;
  }

  html += seccion(1, "Resumen de Energía Facturable");
  html += `<table class="data-table">
<thead><tr><th>Cliente</th><th>Energía consumida [kWh]</th><th>Auxiliares asignados [kWh]</th><th>Total facturable [kWh]</th></tr></thead>
<tbody>
<tr><td class="label">LANEC</td><td class="num">${fmt(lan)}</td><td class="num">${fmt(aux_lan)}</td><td class="num hi">${fmt(lan_fact)}</td></tr>
<tr><td class="label">GRACA</td><td class="num">${fmt(gra)}</td><td class="num">${fmt(aux_gra)}</td><td class="num hi">${fmt(gra_fact)}</td></tr>
<tr class="rpt-row-total"><td class="label"><strong>TOTAL</strong></td><td class="num"><strong>${fmt(tot_cli)}</strong></td><td class="num"><strong>${fmt(aux)}</strong></td><td class="num hi"><strong>${fmt(tot_gen)}</strong></td></tr>
</tbody></table>`;

  html += tablaCliente("2.0", "Costos del Mes — Totales", "TOTAL",
    tot_cli, aux, tot_gen, varTotBy, varTotTotal, fijoTotU1, fijoTotU2, fijoTotal, cbmtU1, totalTot, precioTot);
  html += tablaCliente("2.1", "Costos del Mes — LANEC", "LANEC",
    lan, aux_lan, lan_fact, varLanBy, varLanTotal, fijoLanU1, fijoLanU2, fijoLan, cbmtLanU1, totalLan, precioLan);
  html += tablaCliente("2.2", "Costos del Mes — GRACA", "GRACA",
    gra, aux_gra, gra_fact, varGraBy, varGraTotal, fijoGraU1, fijoGraU2, fijoGra, cbmtGraU1, totalGra, precioGra);

  html += seccion(3, "Costos Fijos (Auditable)");
  html += `<table class="data-table">
<thead><tr><th>Rubro / Unidad</th><th>Días mes</th><th>Días indisp.</th><th>Factor disp.</th><th>Base mensual [USD]</th><th>Valor ajustado [USD]</th></tr></thead>
<tbody>
<tr class="rpt-row-grupo"><td colspan="6" class="label">Costo Fijo por Disponibilidad (Cargo Base por Reserva)</td></tr>
<tr><td class="label">Unidad 1</td><td class="num">${diasMes}</td><td class="num">${diasFallaU1}</td><td class="num">${fmt(dispU1, 4)}</td><td class="num">${fmt(COSTO_FIJO_MENSUAL_POR_UNIDAD)}</td><td class="num hi">${fmt(fijoU1)}</td></tr>
<tr><td class="label">Unidad 2</td><td class="num">${diasMes}</td><td class="num">${diasFallaU2}</td><td class="num">${fmt(dispU2, 4)}</td><td class="num">${fmt(COSTO_FIJO_MENSUAL_POR_UNIDAD)}</td><td class="num hi">${fmt(fijoU2)}</td></tr>
<tr class="rpt-row-total"><td class="label"><strong>Subtotal CF disponibilidad</strong></td><td colspan="4"></td><td class="num hi"><strong>${fmt(fijoTotal)}</strong></td></tr>
<tr class="rpt-row-grupo"><td colspan="6" class="label">CBMT Unidad 1 (Cargo Base por Mantenimiento y Transmisión — 2 400 kW × 7,36 USD/kW-mes)</td></tr>
<tr><td class="label">Unidad 1</td><td class="num">${diasMes}</td><td class="num">${diasFallaU1}</td><td class="num">${fmt(dispU1, 4)}</td><td class="num">${fmt(CBMT_U1_MENSUAL)}</td><td class="num hi">${fmt(cbmtU1)}</td></tr>
<tr class="rpt-row-total"><td class="label"><strong>Subtotal CBMT U1</strong></td><td colspan="4"></td><td class="num hi"><strong>${fmt(cbmtU1)}</strong></td></tr>
<tr class="rpt-row-grand"><td class="label"><strong>TOTAL COSTOS FIJOS</strong></td><td colspan="4"></td><td class="num hi"><strong>${fmt(fijoTotal + cbmtU1)}</strong></td></tr>
</tbody></table>`;

  html += seccion(4, "Asignación de Costos Fijos a Clientes (Factor Contrato)");
  html += `<table class="data-table">
<thead><tr><th>Cliente</th><th>kW contratados</th><th>Factor contrato</th>
<th>CF disp. U1 [USD]</th><th>CF disp. U2 [USD]</th><th>CF disponib. total [USD]</th><th>CBMT U1 [USD]</th><th>Total fijo asignado [USD]</th></tr></thead>
<tbody>
<tr><td class="label">LANEC</td><td class="num">${fmt(P_CONTR_LANEC, 0)}</td><td class="num">${fmt(factorLan * 100, 2)} %</td><td class="num">${fmt(fijoLanU1)}</td><td class="num">${fmt(fijoLanU2)}</td><td class="num">${fmt(fijoLan)}</td><td class="num">${fmt(cbmtLanU1)}</td><td class="num hi">${fmt(fijoLan + cbmtLanU1)}</td></tr>
<tr><td class="label">GRACA</td><td class="num">${fmt(P_CONTR_GRACA, 0)}</td><td class="num">${fmt(factorGra * 100, 2)} %</td><td class="num">${fmt(fijoGraU1)}</td><td class="num">${fmt(fijoGraU2)}</td><td class="num">${fmt(fijoGra)}</td><td class="num">${fmt(cbmtGraU1)}</td><td class="num hi">${fmt(fijoGra + cbmtGraU1)}</td></tr>
<tr class="rpt-row-total"><td class="label"><strong>TOTAL</strong></td><td></td><td></td><td class="num">${fmt(fijoU1)}</td><td class="num">${fmt(fijoU2)}</td><td class="num">${fmt(fijoTotal)}</td><td class="num">${fmt(cbmtU1)}</td><td class="num hi"><strong>${fmt(fijoTotal + cbmtU1)}</strong></td></tr>
</tbody></table>`;

  return html;
}

// ── Desglose de combustible HFO / Diesel ──────────────────────────────────────

export interface FuelBreakdown {
  hfoTotal: number;
  dieselTotal: number;
  transporteTotal: number;
}

export function parseFuelBreakdown(invoiceList: Invoice[]): FuelBreakdown {
  let hfoTotal = 0, dieselTotal = 0, transporteTotal = 0;
  for (const inv of invoiceList) {
    if (inv.category !== "combustible_transporte") continue;
    try {
      if (inv.lineItems && inv.lineItems !== "[]") {
        const items = JSON.parse(inv.lineItems) as Array<Record<string, unknown>>;
        for (const item of items) {
          const t = parseFloat(String(item.total ?? "0"));
          if (item.isTransporte) {
            transporteTotal += t;
          } else if (item.fuelType === "diesel_2") {
            dieselTotal += t;
          } else {
            hfoTotal += t;
          }
        }
      } else {
        hfoTotal += parseFloat(inv.total ?? "0");
      }
    } catch {
      hfoTotal += parseFloat(inv.total ?? "0");
    }
  }
  return { hfoTotal, dieselTotal, transporteTotal };
}

// ── Informe real ──────────────────────────────────────────────────────────────

export function buildRealBillingReport(params: {
  wbProd: XLSX.WorkBook;
  mesStr: string;
  diasFallaU1: number;
  diasFallaU2: number;
  invoiceList: Invoice[];
  hasProduction: boolean;
}): string {
  const { wbProd, mesStr, diasFallaU1, diasFallaU2, invoiceList, hasProduction } = params;

  const prod = getMonthlyProductionSummary(wbProd, mesStr);
  const { lan_fact, gra_fact, tot_gen, lanecKwh: lan, gracaKwh: gra, auxKwh: aux,
    textoPeriodo, diasMes, hfoGal, doGal, kWhHFO, kWhDO } = prod;
  const tot_cli = lan + gra;

  const fixed = calcFixed(diasMes, diasFallaU1, diasFallaU2);
  const { dispU1, dispU2, fijoU1, fijoU2, fijoTotal, cbmtU1, factorLan, factorGra } = fixed;
  const fijoLan = (fijoU1 + fijoU2) * factorLan;
  const fijoGra = (fijoU1 + fijoU2) * factorGra;
  const cbmtLanU1 = cbmtU1 * factorLan;
  const cbmtGraU1 = cbmtU1 * factorGra;

  const invSummary = getInvoiceCategorySummary(invoiceList, tot_gen);
  const { totalByCategory, cvRealByCategory } = invSummary;

  const margenVariable = COSTOS_VARIABLES.margen_variable;
  const margenTotal = margenVariable * tot_gen;

  const costoVarRealTotal = Object.values(totalByCategory).reduce((a, b) => a + b, 0);
  const costoVarRealTotalConMargen = costoVarRealTotal + margenTotal;
  const costoVarRealCvTotal = tot_gen > 0
    ? costoVarRealTotalConMargen / tot_gen
    : 0;

  const costoFijoTotal = fijoTotal + cbmtU1;
  const costoRealTotal = costoVarRealTotalConMargen + costoFijoTotal;
  const costoRealTotalKwh = tot_gen > 0 ? costoRealTotal / tot_gen : 0;

  let html = rptHeader("Informe de Facturación Real", textoPeriodo);

  if (!hasProduction) {
    html += `<div class="rpt-notice rpt-notice-warn">⚠ Sin archivo de producción cargado — los cálculos de USD/kWh no están disponibles.</div>`;
  }
  if (invoiceList.length === 0) {
    html += `<div class="rpt-notice rpt-notice-warn">⚠ No hay facturas registradas para este período.</div>`;
  }

  html += seccion(1, "Resumen de Producción del Período");
  html += `<table class="data-table">
<thead><tr><th>Cliente</th><th>Energía consumida [kWh]</th><th>Auxiliares [kWh]</th><th>Total facturable [kWh]</th></tr></thead>
<tbody>
<tr><td class="label">LANEC</td><td class="num">${fmt(lan)}</td><td class="num">${fmt(lan_fact - lan)}</td><td class="num hi">${fmt(lan_fact)}</td></tr>
<tr><td class="label">GRACA</td><td class="num">${fmt(gra)}</td><td class="num">${fmt(gra_fact - gra)}</td><td class="num hi">${fmt(gra_fact)}</td></tr>
<tr class="rpt-row-total"><td class="label"><strong>TOTAL</strong></td><td class="num"><strong>${fmt(tot_cli)}</strong></td><td class="num"><strong>${fmt(aux)}</strong></td><td class="num hi"><strong>${fmt(tot_gen)}</strong></td></tr>
</tbody></table>`;

  const fuelBreakdown = parseFuelBreakdown(invoiceList);
  const { hfoTotal, dieselTotal, transporteTotal } = fuelBreakdown;
  const combTotal = totalByCategory["combustible_transporte"] ?? 0;

  const cvHFO    = kWhHFO   > 0 ? hfoTotal    / kWhHFO   : 0;
  const cvDiesel = kWhDO    > 0 ? dieselTotal  / kWhDO    : 0;
  const cvComb   = tot_gen  > 0 ? combTotal    / tot_gen  : 0;
  const hasFuelAttrib = (hfoGal + doGal) > 0;

  html += seccion(2, "Facturas del Período por Rubro");
  html += `<table class="data-table">
<thead><tr><th>Rubro</th><th class="num">kWh atribuidos</th><th class="num">Total facturado [USD]</th><th class="num">CV real [USD/kWh]</th></tr></thead>
<tbody>`;

  for (const cat of INVOICE_CATEGORIES) {
    const label = INVOICE_CATEGORY_LABELS[cat as InvoiceCategory] ?? cat;
    if (cat === "combustible_transporte") {
      html += `<tr class="rpt-row-grupo"><td class="label" colspan="4">${label}</td></tr>`;
      if (hfoTotal > 0 || dieselTotal > 0 || transporteTotal > 0) {
        if (hfoTotal > 0)
          html += `<tr><td class="label" style="padding-left:1.5em">↳ ${FUEL_TYPE_LABELS.hfo}</td>
            <td class="num">${hasFuelAttrib && kWhHFO > 0 ? fmt(kWhHFO, 0) : "—"}</td>
            <td class="num">$ ${fmt(hfoTotal)}</td>
            <td class="num ${cvHFO > 0 ? "hi" : ""}">${cvHFO > 0 ? fmt(cvHFO, 4) : "—"}</td></tr>`;
        if (dieselTotal > 0)
          html += `<tr><td class="label" style="padding-left:1.5em">↳ ${FUEL_TYPE_LABELS.diesel_2}</td>
            <td class="num">${hasFuelAttrib && kWhDO > 0 ? fmt(kWhDO, 0) : "—"}</td>
            <td class="num">$ ${fmt(dieselTotal)}</td>
            <td class="num ${cvDiesel > 0 ? "hi" : ""}">${cvDiesel > 0 ? fmt(cvDiesel, 4) : "—"}</td></tr>`;
        if (transporteTotal > 0)
          html += `<tr><td class="label" style="padding-left:1.5em">↳ Transporte (sin IVA)</td>
            <td class="num">—</td>
            <td class="num">$ ${fmt(transporteTotal)}</td>
            <td class="num">—</td></tr>`;
      }
      html += `<tr class="rpt-row-total"><td class="label"><strong>Subtotal Combustible + Transporte</strong></td>
        <td class="num"><strong>${fmt(tot_gen, 0)}</strong></td>
        <td class="num"><strong>$ ${fmt(combTotal)}</strong></td>
        <td class="num"><strong>${fmt(cvComb, 4)}</strong></td></tr>`;
    } else {
      html += `<tr><td class="label">${label}</td>
        <td class="num">—</td>
        <td class="num">$ ${fmt(totalByCategory[cat])}</td>
        <td class="num">${fmt(cvRealByCategory[cat], 4)}</td></tr>`;
    }
  }

  html += `<tr class="rpt-row-total"><td class="label"><strong>Subtotal rubros facturados</strong></td>
    <td></td>
    <td class="num"><strong>$ ${fmt(costoVarRealTotal)}</strong></td>
    <td class="num"><strong>${fmt(tot_gen > 0 ? costoVarRealTotal / tot_gen : 0, 4)}</strong></td></tr>
<tr class="rpt-row-grupo"><td class="label" colspan="4">Margen variable (contractual fijo)</td></tr>
<tr><td class="label">Margen Variable</td><td></td><td class="num">$ ${fmt(margenTotal)}</td><td class="num">${fmt(margenVariable, 4)}</td></tr>
<tr class="rpt-row-grand"><td class="label"><strong>CV Total Real (incluye margen)</strong></td><td></td>
  <td class="num"><strong>$ ${fmt(costoVarRealTotalConMargen)}</strong></td>
  <td class="num"><strong>${fmt(costoVarRealCvTotal, 4)} USD/kWh</strong></td></tr>
</tbody></table>
${hasFuelAttrib
  ? `<p class="rpt-muted">* kWh atribuidos por combustible calculados proporcionalmente al consumo del período: HFO ${fmt(hfoGal, 0)} gal (${fmt(hfoGal / (hfoGal + doGal) * 100, 1)}%), Diesel 2 ${fmt(doGal, 0)} gal (${fmt(doGal / (hfoGal + doGal) * 100, 1)}%).</p>`
  : `<div class="rpt-notice rpt-notice-warn">⚠ Sin datos de consumo de galones — los kWh atribuidos por combustible no están disponibles.</div>`
}`;

  html += seccion(3, "Costos Fijos");
  html += `<table class="data-table">
<thead><tr><th>Rubro / Unidad</th><th>Días mes</th><th>Días indisp.</th><th>Factor disp.</th><th>Base mensual [USD]</th><th>Valor ajustado [USD]</th></tr></thead>
<tbody>
<tr class="rpt-row-grupo"><td colspan="6" class="label">Costo Fijo por Disponibilidad</td></tr>
<tr><td class="label">Unidad 1</td><td class="num">${diasMes}</td><td class="num">${diasFallaU1}</td><td class="num">${fmt(dispU1, 4)}</td><td class="num">${fmt(COSTO_FIJO_MENSUAL_POR_UNIDAD)}</td><td class="num hi">${fmt(fijoU1)}</td></tr>
<tr><td class="label">Unidad 2</td><td class="num">${diasMes}</td><td class="num">${diasFallaU2}</td><td class="num">${fmt(dispU2, 4)}</td><td class="num">${fmt(COSTO_FIJO_MENSUAL_POR_UNIDAD)}</td><td class="num hi">${fmt(fijoU2)}</td></tr>
<tr class="rpt-row-total"><td class="label"><strong>Subtotal CF disponibilidad</strong></td><td colspan="4"></td><td class="num hi"><strong>${fmt(fijoTotal)}</strong></td></tr>
<tr class="rpt-row-grupo"><td colspan="6" class="label">CBMT Unidad 1</td></tr>
<tr><td class="label">Unidad 1</td><td class="num">${diasMes}</td><td class="num">${diasFallaU1}</td><td class="num">${fmt(dispU1, 4)}</td><td class="num">${fmt(CBMT_U1_MENSUAL)}</td><td class="num hi">${fmt(cbmtU1)}</td></tr>
<tr class="rpt-row-grand"><td class="label"><strong>TOTAL COSTOS FIJOS</strong></td><td colspan="4"></td><td class="num hi"><strong>${fmt(costoFijoTotal)}</strong></td></tr>
</tbody></table>`;

  html += seccion(4, "Asignación de Costos Fijos a Clientes");
  html += `<table class="data-table">
<thead><tr><th>Cliente</th><th>Factor contrato</th><th>CF disponib. [USD]</th><th>CBMT U1 [USD]</th><th>Total fijo [USD]</th></tr></thead>
<tbody>
<tr><td class="label">LANEC</td><td class="num">${fmt(factorLan * 100, 2)} %</td><td class="num">${fmt(fijoLan)}</td><td class="num">${fmt(cbmtLanU1)}</td><td class="num hi">${fmt(fijoLan + cbmtLanU1)}</td></tr>
<tr><td class="label">GRACA</td><td class="num">${fmt(factorGra * 100, 2)} %</td><td class="num">${fmt(fijoGra)}</td><td class="num">${fmt(cbmtGraU1)}</td><td class="num hi">${fmt(fijoGra + cbmtGraU1)}</td></tr>
<tr class="rpt-row-total"><td class="label"><strong>TOTAL</strong></td><td></td><td class="num">${fmt(fijoTotal)}</td><td class="num">${fmt(cbmtU1)}</td><td class="num hi"><strong>${fmt(costoFijoTotal)}</strong></td></tr>
</tbody></table>`;

  html += seccion(5, "Resumen Costo Real Total del Período");
  html += `<table class="data-table">
<thead><tr><th>Componente</th><th>USD/kWh</th><th>USD</th></tr></thead>
<tbody>
<tr><td class="label">Costos variables reales (rubros facturados)</td><td class="num">${fmt(tot_gen > 0 ? costoVarRealTotal / tot_gen : 0, 4)}</td><td class="num">$ ${fmt(costoVarRealTotal)}</td></tr>
<tr><td class="label">Margen variable (contractual)</td><td class="num">${fmt(margenVariable, 4)}</td><td class="num">$ ${fmt(margenTotal)}</td></tr>
<tr class="rpt-row-total"><td class="label"><strong>Subtotal variable</strong></td><td class="num"><strong>${fmt(costoVarRealCvTotal, 4)}</strong></td><td class="num"><strong>$ ${fmt(costoVarRealTotalConMargen)}</strong></td></tr>
<tr><td class="label">Costos fijos</td><td class="num">—</td><td class="num">$ ${fmt(costoFijoTotal)}</td></tr>
<tr class="rpt-row-grand"><td class="label"><strong>COSTO REAL TOTAL</strong></td><td class="num"><strong>${fmt(costoRealTotalKwh, 4)} USD/kWh</strong></td><td class="num"><strong>$ ${fmt(costoRealTotal)} + IVA</strong></td></tr>
</tbody></table>
<p class="rpt-muted">* Energía base del período: ${fmt(tot_gen, 0)} kWh (LANEC + GRACA + Auxiliares asignados).</p>`;

  return html;
}
