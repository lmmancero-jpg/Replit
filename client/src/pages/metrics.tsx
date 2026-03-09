import { useState, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  Chart as ChartJS,
  CategoryScale, LinearScale, PointElement, LineElement,
  Title, Tooltip, Legend, Filler,
} from "chart.js";
import { Line } from "react-chartjs-2";
import { Layout } from "@/components/layout";
import { extractProduction, extractAforo, buildResumen, fmt } from "@/lib/metricsEngine";
import type { ProdData, AforoData, Resumen } from "@/lib/metricsEngine";

ChartJS.register(CategoryScale, LinearScale, PointElement, LineElement, Title, Tooltip, Legend, Filler);

// ─── Paleta de colores ───────────────────────────────────────────────────────
const COLORS = ["#0063a6", "#00a65a", "#f59e0b", "#ef4444", "#8b5cf6", "#06b6d4", "#ec4899"];

function lineDataset(label: string, data: (number | null)[], idx = 0) {
  return {
    label,
    data,
    borderColor: COLORS[idx % COLORS.length],
    backgroundColor: COLORS[idx % COLORS.length] + "22",
    borderWidth: 2,
    pointRadius: 2,
    pointHoverRadius: 4,
    tension: 0.25,
    spanGaps: true,
  };
}

function lineOpts(yLabel: string) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    interaction: { mode: "index" as const, intersect: false },
    plugins: {
      legend: { position: "top" as const, labels: { boxWidth: 14, boxHeight: 10, font: { size: 11 } } },
    },
    scales: {
      x: { ticks: { font: { size: 10 } }, title: { display: true, text: "Día", font: { size: 10 } } },
      y: { ticks: { font: { size: 10 } }, title: { display: true, text: yLabel, font: { size: 10 } } },
    },
  };
}

// ─── KPI Card ────────────────────────────────────────────────────────────────
function KpiCard({ title, value, sub }: { title: string; value: string; sub?: string }) {
  return (
    <div className="bg-white border border-gray-200 rounded-xl shadow-sm p-4 border-l-4 border-l-blue-600 flex flex-col gap-1 min-h-[100px]">
      <div className="text-xs font-bold uppercase tracking-wider text-gray-500">{title}</div>
      <div className="text-2xl font-black text-gray-800 mt-1 flex-1 flex items-end">{value}</div>
      {sub && <div className="text-sm font-semibold text-gray-500">{sub}</div>}
    </div>
  );
}

// ─── Chart Card ──────────────────────────────────────────────────────────────
function ChartCard({ title, subtitle, children }: { title: string; subtitle?: string; children: React.ReactNode }) {
  return (
    <div className="border border-gray-200 rounded-xl bg-white shadow-sm overflow-hidden">
      <div className="px-4 pt-3 pb-1">
        <div className="font-black text-sm text-gray-800">{title}</div>
        {subtitle && <div className="text-xs text-gray-400 mt-0.5">{subtitle}</div>}
      </div>
      <div className="relative w-full h-[300px] px-3 pb-3">
        {children}
      </div>
    </div>
  );
}

// ─── Placeholder cuando no hay datos de aforo ─────────────────────────────────
function NoAforo() {
  return (
    <div className="flex items-center justify-center h-full text-sm text-gray-400 italic">
      Sin datos de aforo cargados
    </div>
  );
}

// ─── Tabs ────────────────────────────────────────────────────────────────────
type Tab = "produccion" | "combustible";

// ─── Estado de procesamiento ──────────────────────────────────────────────────
interface ProcessedData {
  prod: ProdData;
  aforo: AforoData | null;
  resumen: Resumen;
}

export default function Metrics() {
  const [wbProd, setWbProd] = useState<XLSX.WorkBook | null>(null);
  const [wbAforo, setWbAforo] = useState<XLSX.WorkBook | null>(null);
  const [sheets, setSheets] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [status, setStatus] = useState("Cargue el archivo GEN y el archivo de aforo. Luego seleccione el mes y procese.");
  const [statusError, setStatusError] = useState(false);
  const [data, setData] = useState<ProcessedData | null>(null);
  const [activeTab, setActiveTab] = useState<Tab>("produccion");

  const fileProdRef = useRef<HTMLInputElement>(null);
  const fileAforoRef = useRef<HTMLInputElement>(null);

  const setMsg = (msg: string, err = false) => { setStatus(msg); setStatusError(err); };

  const onProdFile = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target!.result as ArrayBuffer, { type: "array", cellDates: true });
        setWbProd(wb);
        setSheets(wb.SheetNames);
        setSelectedSheet(wb.SheetNames[0] || "");
        setMsg("Archivo GEN cargado. Selecciona la hoja/mes.");
      } catch {
        setWbProd(null); setSheets([]);
        setMsg("Error al leer el archivo de producción (GEN).", true);
      }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const onAforoFile = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target!.result as ArrayBuffer, { type: "array", cellDates: true });
        setWbAforo(wb);
        setMsg('Archivo de aforo cargado. (Se usa la hoja "Sondas 00 00 hrs")');
      } catch {
        setWbAforo(null);
        setMsg("Error al leer el archivo de aforo.", true);
      }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const onProcess = useCallback(() => {
    if (!wbProd || !selectedSheet) return;
    try {
      const prod = extractProduction(wbProd, selectedSheet);
      if (!prod || !prod.labels.length) {
        setMsg("No se encontraron datos válidos en esa hoja/mes.", true); return;
      }
      const aforo = wbAforo ? extractAforo(wbAforo, prod.targetMonth, prod.targetYear) : null;
      const resumen = buildResumen(prod);
      setData({ prod, aforo, resumen });
      setMsg(`Procesado: "${selectedSheet.trim()}" (mes ${String(prod.targetMonth).padStart(2, "0")}/${prod.targetYear}).${!aforo ? " Nota: sin datos de aforo." : ""}`);
    } catch (e: unknown) {
      setMsg(`Error al procesar: ${e instanceof Error ? e.message : String(e)}`, true);
    }
  }, [wbProd, wbAforo, selectedSheet]);

  const canProcess = !!wbProd && !!wbAforo && !!selectedSheet;

  return (
    <Layout>
      <div className="p-6 space-y-5 max-w-[1280px] mx-auto">

        {/* ── Encabezado ─────────────────────────────────────────────── */}
        <div>
          <h1 className="text-xl font-black text-gray-800 tracking-wide uppercase">
            Métricas – Central El Morro
          </h1>
          <p className="text-sm text-gray-500 mt-0.5">
            Producción · Clientes · Eficiencia · Combustible · Tanques
          </p>
        </div>

        {/* ── Panel de controles ────────────────────────────────────────── */}
        <div className="bg-white border border-gray-200 rounded-xl shadow-sm p-4">
          <div className="text-xs font-extrabold uppercase tracking-widest text-gray-400 mb-3">
            Parámetros de consulta
          </div>
          <div className="flex flex-wrap gap-4 items-end">

            {/* GEN */}
            <div className="flex flex-col gap-1 min-w-[220px]">
              <label className="text-xs font-semibold text-gray-500" htmlFor="metricsProdFile">
                Archivo de Producción (GEN):
              </label>
              <input
                id="metricsProdFile"
                data-testid="input-metrics-gen"
                ref={fileProdRef}
                type="file"
                accept=".xlsx,.xls"
                onChange={onProdFile}
                className="text-sm border border-gray-300 rounded-md px-2 py-1.5 bg-white min-h-[36px]"
              />
            </div>

            {/* Aforo */}
            <div className="flex flex-col gap-1 min-w-[220px]">
              <label className="text-xs font-semibold text-gray-500" htmlFor="metricsAforoFile">
                Archivo de Aforo (Tanques):
              </label>
              <input
                id="metricsAforoFile"
                data-testid="input-metrics-aforo"
                ref={fileAforoRef}
                type="file"
                accept=".xlsx,.xls"
                onChange={onAforoFile}
                className="text-sm border border-gray-300 rounded-md px-2 py-1.5 bg-white min-h-[36px]"
              />
            </div>

            {/* Mes */}
            <div className="flex flex-col gap-1 min-w-[200px]">
              <label className="text-xs font-semibold text-gray-500" htmlFor="metricsSheet">
                Mes (Hoja del GEN):
              </label>
              <select
                id="metricsSheet"
                data-testid="select-metrics-sheet"
                value={selectedSheet}
                onChange={e => setSelectedSheet(e.target.value)}
                disabled={!sheets.length}
                className="text-sm border border-gray-300 rounded-md px-2 py-1.5 bg-white min-h-[36px] disabled:opacity-50"
              >
                {!sheets.length && <option value="">-- Cargue el GEN --</option>}
                {sheets.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>

            {/* Botón */}
            <button
              data-testid="button-metrics-process"
              onClick={onProcess}
              disabled={!canProcess}
              className="px-6 py-2 rounded-full bg-blue-600 text-white font-bold text-sm shadow hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed transition-all"
            >
              Procesar
            </button>
          </div>

          <div className={`text-xs mt-3 pt-3 border-t border-dashed border-gray-200 ${statusError ? "text-red-600" : "text-gray-500"}`}>
            {status}
          </div>
        </div>

        {data && (
          <>
            {/* ── Tabs ──────────────────────────────────────────────────── */}
            <div className="flex gap-2">
              {(["produccion", "combustible"] as Tab[]).map(tab => (
                <button
                  key={tab}
                  data-testid={`tab-metrics-${tab}`}
                  onClick={() => setActiveTab(tab)}
                  className={`px-5 py-2 rounded-full text-sm font-bold transition-all ${activeTab === tab
                    ? "bg-blue-600 text-white shadow"
                    : "bg-white border border-gray-200 text-gray-600 hover:bg-gray-50"}`}
                >
                  {tab === "produccion" ? "Producción" : "Combustible & Tanques"}
                </button>
              ))}
            </div>

            {/* ══════════════════════════════════════════════════════
                TAB: PRODUCCIÓN
               ══════════════════════════════════════════════════════ */}
            {activeTab === "produccion" && (
              <div className="space-y-5">
                {/* KPIs */}
                <div className="grid grid-cols-2 md:grid-cols-3 xl:grid-cols-5 gap-3">
                  <KpiCard
                    title="Energía total"
                    value={`${fmt(data.resumen.energiaTotalMWh, 1)} MWh`}
                    sub={`U1: ${fmt(data.resumen.energiaU1MWh, 1)} · U2: ${fmt(data.resumen.energiaU2MWh, 1)}`}
                  />
                  <KpiCard
                    title="LANEC"
                    value={`${fmt(data.resumen.energiaLanecMWh, 1)} MWh`}
                    sub={`GRACA: ${fmt(data.resumen.energiaGracaMWh, 1)} MWh`}
                  />
                  <KpiCard
                    title="Auxiliares"
                    value={`${fmt(data.resumen.energiaAuxMWh, 1)} MWh`}
                  />
                  <KpiCard
                    title="Horas"
                    value={`U1 ${fmt(data.resumen.horasU1, 1)} h`}
                    sub={`U2 ${fmt(data.resumen.horasU2, 1)} h`}
                  />
                  <KpiCard
                    title="Eficiencia"
                    value={data.resumen.eficProm != null ? `${fmt(data.resumen.eficProm, 2)} kWh/gal` : "—"}
                    sub={`Mes ${String(data.prod.targetMonth).padStart(2, "0")}/${data.prod.targetYear}`}
                  />
                </div>

                {/* Gráficos 2×3 + tanques */}
                <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
                  <ChartCard title="Energía Total" subtitle="kWh generados por día">
                    <Line
                      data={{ labels: data.prod.labels, datasets: [lineDataset("Energía total (kWh)", data.prod.etotal, 0)] }}
                      options={lineOpts("kWh")}
                    />
                  </ChartCard>

                  <ChartCard title="Energía por Cliente" subtitle="LANEC · GRACA · Auxiliares">
                    <Line
                      data={{ labels: data.prod.labels, datasets: [
                        lineDataset("LANEC (kWh)", data.prod.lanec, 0),
                        lineDataset("GRACA (kWh)", data.prod.graca, 1),
                        lineDataset("Auxiliares (kWh)", data.prod.aux, 2),
                      ]}}
                      options={lineOpts("kWh")}
                    />
                  </ChartCard>

                  <ChartCard title="Energía por Unidad" subtitle="U1 (9L26) · U2 (9SW280)">
                    <Line
                      data={{ labels: data.prod.labels, datasets: [
                        lineDataset("Unidad 1 (kWh)", data.prod.e_u1, 0),
                        lineDataset("Unidad 2 (kWh)", data.prod.e_u2, 1),
                      ]}}
                      options={lineOpts("kWh")}
                    />
                  </ChartCard>

                  <ChartCard title="Potencia Promedio" subtitle="kW promedio por unidad (energía / horas)">
                    <Line
                      data={{ labels: data.prod.labels, datasets: [
                        lineDataset("Potencia U1 (kW)", data.prod.pot_u1, 0),
                        lineDataset("Potencia U2 (kW)", data.prod.pot_u2, 1),
                      ]}}
                      options={lineOpts("kW")}
                    />
                  </ChartCard>

                  <ChartCard title="Horas de Operación" subtitle="Horas diarias por unidad">
                    <Line
                      data={{ labels: data.prod.labels, datasets: [
                        lineDataset("Horas U1 (h)", data.prod.h_u1, 0),
                        lineDataset("Horas U2 (h)", data.prod.h_u2, 1),
                      ]}}
                      options={lineOpts("h")}
                    />
                  </ChartCard>

                  <ChartCard title="Eficiencia" subtitle="kWh por galón (columna AM)">
                    <Line
                      data={{ labels: data.prod.labels, datasets: [lineDataset("Eficiencia (kWh/gal)", data.prod.rend, 0)] }}
                      options={lineOpts("kWh/gal")}
                    />
                  </ChartCard>

                  <ChartCard title="Tanques & Sludge" subtitle="T601/T602/T610/T611/Cisterna 2 (00H00)">
                    {data.aforo ? (
                      <Line
                        data={{ labels: data.aforo.labels, datasets: [
                          lineDataset("T601 (HFO, gal)", data.aforo.t601, 0),
                          lineDataset("T602 (HFO, gal)", data.aforo.t602, 1),
                          lineDataset("T610 (Diesel, gal)", data.aforo.t610, 2),
                          lineDataset("T611 (Diesel, gal)", data.aforo.t611, 3),
                          lineDataset("Cisterna 2 (gal)", data.aforo.cisterna2, 4),
                        ]}}
                        options={lineOpts("gal")}
                      />
                    ) : <NoAforo />}
                  </ChartCard>
                </div>
              </div>
            )}

            {/* ══════════════════════════════════════════════════════
                TAB: COMBUSTIBLE & TANQUES
               ══════════════════════════════════════════════════════ */}
            {activeTab === "combustible" && (
              <div className="space-y-5">
                {/* Encabezado tipo informe */}
                <div className="bg-gray-900 text-white rounded-xl px-5 py-4">
                  <div className="font-black text-base tracking-wide">
                    Central El Morro — Informe Gerencial de Combustible
                  </div>
                  <div className="flex flex-wrap gap-5 mt-2 text-xs opacity-90">
                    <span>Mes analizado: <strong>{String(data.prod.targetMonth).padStart(2, "0")}/{data.prod.targetYear}</strong></span>
                    <span>Fecha emisión: <strong>{new Date().toLocaleDateString("es-EC")}</strong></span>
                  </div>
                </div>

                {/* KPIs combustible */}
                <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                  <KpiCard title="HFO total consumido del mes" value={`${fmt(data.resumen.hfoGal, 0)} gal`} />
                  <KpiCard title="Diésel total consumido del mes" value={`${fmt(data.resumen.doGal, 0)} gal`} />
                  <KpiCard title="Días con registro" value={`${data.resumen.dias}`} />
                </div>

                {/* Gráficos 2×3 */}
                <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
                  <ChartCard title="Consumo total HFO vs Diésel" subtitle="Galones diarios totales de combustible">
                    <Line
                      data={{ labels: data.prod.labels, datasets: [
                        lineDataset("HFO total (gal)", data.prod.hfoTot.map(v => v || null), 0),
                        lineDataset("Diésel total (gal)", data.prod.doTot.map(v => v || null), 1),
                      ]}}
                      options={lineOpts("gal")}
                    />
                  </ChartCard>

                  <ChartCard title="Consumo HFO por Unidad" subtitle="Reparto HFO entre G1 y G2 según energía diaria">
                    <Line
                      data={{ labels: data.prod.labels, datasets: [
                        lineDataset("HFO G1 (gal)", data.prod.hfoG1.map(v => v || null), 0),
                        lineDataset("HFO G2 (gal)", data.prod.hfoG2.map(v => v || null), 1),
                      ]}}
                      options={lineOpts("gal")}
                    />
                  </ChartCard>

                  <ChartCard title="Consumo Diésel por Unidad" subtitle="Reparto Diésel entre G1 y G2 según energía diaria">
                    <Line
                      data={{ labels: data.prod.labels, datasets: [
                        lineDataset("Diésel G1 (gal)", data.prod.doG1.map(v => v || null), 0),
                        lineDataset("Diésel G2 (gal)", data.prod.doG2.map(v => v || null), 1),
                      ]}}
                      options={lineOpts("gal")}
                    />
                  </ChartCard>

                  <ChartCard title="Tanques HFO T601 y T602" subtitle="Volumen (gal) a las 00H00">
                    {data.aforo ? (
                      <Line
                        data={{ labels: data.aforo.labels, datasets: [
                          lineDataset("T601 (HFO, gal)", data.aforo.t601, 0),
                          lineDataset("T602 (HFO, gal)", data.aforo.t602, 1),
                        ]}}
                        options={lineOpts("gal")}
                      />
                    ) : <NoAforo />}
                  </ChartCard>

                  <ChartCard title="Tanques Diésel T610 y T611" subtitle="Volumen (gal) a las 00H00">
                    {data.aforo ? (
                      <Line
                        data={{ labels: data.aforo.labels, datasets: [
                          lineDataset("T610 (Diesel, gal)", data.aforo.t610, 2),
                          lineDataset("T611 (Diesel, gal)", data.aforo.t611, 3),
                        ]}}
                        options={lineOpts("gal")}
                      />
                    ) : <NoAforo />}
                  </ChartCard>

                  <ChartCard title="Cisterna 2 (Sludge)" subtitle='Volumen (gal) registrado como "CISTERNA 2"'>
                    {data.aforo ? (
                      <Line
                        data={{ labels: data.aforo.labels, datasets: [
                          lineDataset("Cisterna 2 (gal)", data.aforo.cisterna2, 4),
                        ]}}
                        options={lineOpts("gal")}
                      />
                    ) : <NoAforo />}
                  </ChartCard>
                </div>
              </div>
            )}
          </>
        )}

        {/* Estado vacío inicial */}
        {!data && (
          <div className="flex flex-col items-center justify-center py-20 text-center text-gray-400">
            <svg className="w-16 h-16 mb-4 opacity-30" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z" />
            </svg>
            <p className="text-base font-semibold">Cargue los archivos y procese para ver las gráficas</p>
            <p className="text-sm mt-1">GEN (producción) + Aforo (tanques) → seleccione el mes → Procesar</p>
          </div>
        )}
      </div>
    </Layout>
  );
}
