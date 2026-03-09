import { useState, useRef, useCallback, useEffect } from "react";
import {
  Chart as ChartJS,
  CategoryScale, LinearScale, PointElement, LineElement,
  Title, Tooltip, Legend, Filler,
} from "chart.js";
import { Line } from "react-chartjs-2";
import { Layout } from "@/components/layout";
import { useFileStore } from "@/lib/fileStore";
import {
  UploadCloud, CheckCircle2, AlertCircle, Zap, Droplets,
  Clock, TrendingUp, Fuel, Database, BarChart2, FlameKindling,
  ChevronRight, FileDown,
} from "lucide-react";
import { extractProduction, extractAforo, buildResumen, fmt } from "@/lib/metricsEngine";
import type { ProdData, AforoData, Resumen } from "@/lib/metricsEngine";
import { useToast } from "@/hooks/use-toast";

ChartJS.register(CategoryScale, LinearScale, PointElement, LineElement, Title, Tooltip, Legend, Filler);

// ─── Paleta ──────────────────────────────────────────────────────────────────
const P = {
  blue:   "#2563eb",
  green:  "#16a34a",
  amber:  "#d97706",
  red:    "#dc2626",
  purple: "#7c3aed",
  cyan:   "#0891b2",
  rose:   "#e11d48",
  teal:   "#0d9488",
};
const PALETTE = [P.blue, P.green, P.amber, P.red, P.purple, P.cyan, P.rose, P.teal];

function hex(color: string, alpha = 1): string {
  const r = parseInt(color.slice(1, 3), 16);
  const g = parseInt(color.slice(3, 5), 16);
  const b = parseInt(color.slice(5, 7), 16);
  return `rgba(${r},${g},${b},${alpha})`;
}

function makeGradient(color: string) {
  return (ctx: CanvasRenderingContext2D, chartArea: { top: number; bottom: number }) => {
    if (!chartArea) return color;
    const gradient = ctx.createLinearGradient(0, chartArea.top, 0, chartArea.bottom);
    gradient.addColorStop(0, hex(color, 0.32));
    gradient.addColorStop(1, hex(color, 0.02));
    return gradient;
  };
}

function lineDataset(label: string, data: (number | null)[], idx = 0) {
  const color = PALETTE[idx % PALETTE.length];
  return {
    label,
    data,
    borderColor: color,
    backgroundColor: (ctx: { chart: { ctx: CanvasRenderingContext2D; chartArea: { top: number; bottom: number } } }) =>
      makeGradient(color)(ctx.chart.ctx, ctx.chart.chartArea),
    fill: true,
    borderWidth: 2,
    pointRadius: 2.5,
    pointHoverRadius: 5,
    pointBackgroundColor: color,
    tension: 0.3,
    spanGaps: true,
  };
}

function lineOpts(yLabel: string) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    animation: { duration: 0 } as { duration: number },
    interaction: { mode: "index" as const, intersect: false },
    plugins: {
      legend: {
        position: "top" as const,
        labels: { boxWidth: 10, boxHeight: 10, padding: 10, font: { size: 10, family: "system-ui" } },
      },
      tooltip: {
        backgroundColor: "rgba(15,23,42,0.9)",
        titleFont: { size: 11, weight: "bold" as const },
        bodyFont: { size: 10 },
        padding: 8,
        cornerRadius: 6,
      },
    },
    scales: {
      x: {
        grid: { color: "rgba(0,0,0,0.04)" },
        ticks: { font: { size: 9 }, color: "#6b7280" },
      },
      y: {
        grid: { color: "rgba(0,0,0,0.04)" },
        ticks: { font: { size: 9 }, color: "#6b7280" },
        title: { display: true, text: yLabel, font: { size: 9 }, color: "#9ca3af" },
      },
    },
  };
}

// ─── File indicator ───────────────────────────────────────────────────────────
function FileIndicator({
  testId, label, loaded, fileName, onChange,
}: {
  testId: string; label: string; loaded: boolean; fileName: string;
  onChange: (e: React.ChangeEvent<HTMLInputElement>) => void;
}) {
  const ref = useRef<HTMLInputElement>(null);
  return (
    <div
      onClick={() => ref.current?.click()}
      className={`relative cursor-pointer rounded-xl border-2 border-dashed p-4 flex items-center gap-3 transition-all select-none
        ${loaded
          ? "border-green-400 bg-green-50 hover:bg-green-100"
          : "border-gray-300 bg-gray-50 hover:border-blue-400 hover:bg-blue-50"}`}
    >
      <input ref={ref} data-testid={testId} type="file" accept=".xlsx,.xls" onChange={onChange} className="hidden" />
      <div className={`w-10 h-10 rounded-lg flex items-center justify-center shrink-0 ${loaded ? "bg-green-500" : "bg-gray-200"}`}>
        {loaded ? <CheckCircle2 className="w-5 h-5 text-white" /> : <UploadCloud className="w-5 h-5 text-gray-500" />}
      </div>
      <div className="min-w-0">
        <div className={`text-xs font-bold uppercase tracking-wider mb-0.5 ${loaded ? "text-green-700" : "text-gray-500"}`}>{label}</div>
        <div className={`text-sm font-medium truncate max-w-[200px] ${loaded ? "text-green-800" : "text-gray-400"}`}>
          {loaded ? fileName : "Seleccionar archivo…"}
        </div>
        {loaded && <div className="text-xs text-green-600 mt-0.5">Clic para cambiar</div>}
      </div>
    </div>
  );
}

// ─── KPI Card (pantalla) ──────────────────────────────────────────────────────
function KpiCard({ title, value, unit, sub, icon: Icon, accent = P.blue }: {
  title: string; value: string; unit?: string; sub?: string;
  icon?: React.ElementType; accent?: string;
}) {
  return (
    <div className="bg-white rounded-xl border border-gray-100 shadow-sm overflow-hidden">
      <div className="h-1 w-full" style={{ background: accent }} />
      <div className="p-4 flex flex-col gap-2">
        <div className="flex items-center justify-between">
          <span className="text-xs font-bold uppercase tracking-wider text-gray-400">{title}</span>
          {Icon && (
            <div className="w-7 h-7 rounded-lg flex items-center justify-center" style={{ background: hex(accent, 0.1) }}>
              <Icon className="w-4 h-4" style={{ color: accent }} />
            </div>
          )}
        </div>
        <div className="flex items-baseline gap-1.5">
          <span className="text-2xl font-black text-gray-900">{value}</span>
          {unit && <span className="text-sm font-semibold text-gray-400">{unit}</span>}
        </div>
        {sub && <div className="text-xs text-gray-500 font-medium">{sub}</div>}
      </div>
    </div>
  );
}

// ─── Chart Card (pantalla) ────────────────────────────────────────────────────
function ChartCard({ title, subtitle, accent = P.blue, height = 280, children }: {
  title: string; subtitle?: string; accent?: string; height?: number; children: React.ReactNode;
}) {
  return (
    <div className="bg-white rounded-xl border border-gray-100 shadow-sm overflow-hidden flex flex-col">
      <div className="flex items-start gap-3 px-4 pt-4 pb-3">
        <div className="w-1 self-stretch rounded-full shrink-0" style={{ background: accent }} />
        <div>
          <div className="font-bold text-sm text-gray-800">{title}</div>
          {subtitle && <div className="text-xs text-gray-400 mt-0.5">{subtitle}</div>}
        </div>
      </div>
      <div className="relative w-full px-3 pb-4" style={{ height }}>
        {children}
      </div>
    </div>
  );
}

function NoAforo() {
  return (
    <div className="flex flex-col items-center justify-center h-full gap-2 text-gray-300">
      <Database className="w-8 h-8" />
      <span className="text-xs font-medium">Sin datos de aforo</span>
    </div>
  );
}

function SectionHeader({ icon: Icon, title, sub, accent = P.blue }: {
  icon: React.ElementType; title: string; sub?: string; accent?: string;
}) {
  return (
    <div className="flex items-center gap-3 mb-4">
      <div className="w-8 h-8 rounded-lg flex items-center justify-center shrink-0"
        style={{ background: hex(accent, 0.12) }}>
        <Icon className="w-4 h-4" style={{ color: accent }} />
      </div>
      <div>
        <div className="font-black text-sm text-gray-800">{title}</div>
        {sub && <div className="text-xs text-gray-400">{sub}</div>}
      </div>
    </div>
  );
}

// ─── Componente de impresión PDF ──────────────────────────────────────────────
// Renderiza AMBAS secciones con layout fijo para captura con html2canvas.
// Se monta fuera de pantalla, espera a que Chart.js pinte, luego se captura.
function PdfPrintContent({ data }: { data: ProcessedData }) {
  const { prod, aforo, resumen } = data;
  const mm = String(prod.targetMonth).padStart(2, "0");
  const period = `${mm}/${prod.targetYear}`;
  const today = new Date().toLocaleDateString("es-EC");

  // Estilo de KPI para el PDF (inline, sin Tailwind)
  const kpi = (label: string, val: string, unit: string, sub?: string, color = P.blue) => (
    <div style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, overflow: "hidden", minWidth: 0 }}>
      <div style={{ height: 4, background: color }} />
      <div style={{ padding: "10px 14px" }}>
        <div style={{ fontSize: 9, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.06em", color: "#9ca3af", marginBottom: 4 }}>{label}</div>
        <div style={{ fontSize: 20, fontWeight: 900, color: "#111827" }}>{val} <span style={{ fontSize: 11, fontWeight: 600, color: "#9ca3af" }}>{unit}</span></div>
        {sub && <div style={{ fontSize: 10, color: "#6b7280", marginTop: 2 }}>{sub}</div>}
      </div>
    </div>
  );

  const chartStyle = { position: "relative" as const, height: 155, marginTop: 4 };

  return (
    <div style={{ fontFamily: "system-ui, sans-serif", background: "#fff", padding: "0 24px 24px", width: "100%" }}>

      {/* ── SECCIÓN PRODUCCIÓN ────────────────────────────────────────── */}
      <div data-pdf-section="produccion" style={{ paddingTop: 24 }}>
        {/* Cabecera */}
        <div style={{ background: "linear-gradient(90deg,#0f172a,#1e3a5f)", borderRadius: 12, padding: "14px 20px", marginBottom: 18, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div>
            <div style={{ color: "#93c5fd", fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: 2 }}>Métricas de Producción</div>
            <div style={{ color: "#fff", fontSize: 18, fontWeight: 900 }}>Central El Morro — Morro Energy S.A.</div>
            <div style={{ color: "#94a3b8", fontSize: 11, marginTop: 2 }}>Período: <b style={{ color: "#e2e8f0" }}>{period}</b> · Hoja: {prod.sheetName} · Emisión: {today}</div>
          </div>
        </div>

        {/* KPIs */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(5,1fr)", gap: 10, marginBottom: 18 }}>
          {kpi("Energía total", fmt(resumen.energiaTotalMWh, 1), "MWh", `U1 ${fmt(resumen.energiaU1MWh, 1)} · U2 ${fmt(resumen.energiaU2MWh, 1)} MWh`, P.blue)}
          {kpi("LANEC", fmt(resumen.energiaLanecMWh, 1), "MWh", `${fmt(resumen.energiaTotalMWh > 0 ? resumen.energiaLanecMWh / resumen.energiaTotalMWh * 100 : 0, 1)}% del total`, P.green)}
          {kpi("GRACA", fmt(resumen.energiaGracaMWh, 1), "MWh", `${fmt(resumen.energiaTotalMWh > 0 ? resumen.energiaGracaMWh / resumen.energiaTotalMWh * 100 : 0, 1)}% del total`, P.amber)}
          {kpi("Horas U1", fmt(resumen.horasU1, 0), "h", `U2: ${fmt(resumen.horasU2, 0)} h`, P.purple)}
          {kpi("Eficiencia", resumen.eficProm != null ? fmt(resumen.eficProm, 2) : "—", "kWh/gal", period, P.cyan)}
        </div>

        {/* Gráficos fila 1 */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 12 }}>
          {[
            { title: "Energía Total (kWh/día)", ds: [lineDataset("Energía total (kWh)", prod.etotal, 0)], y: "kWh" },
            { title: "Energía por Cliente (kWh/día)", ds: [lineDataset("LANEC", prod.lanec, 0), lineDataset("GRACA", prod.graca, 1), lineDataset("Auxiliares", prod.aux, 2)], y: "kWh" },
          ].map(({ title, ds, y }) => (
            <div key={title} style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, padding: "10px 12px 12px" }}>
              <div style={{ fontSize: 11, fontWeight: 700, color: "#1f2937", marginBottom: 2 }}>{title}</div>
              <div style={chartStyle}><Line data={{ labels: prod.labels, datasets: ds }} options={lineOpts(y)} /></div>
            </div>
          ))}
        </div>

        {/* Gráficos fila 2 */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 12 }}>
          {[
            { title: "Energía por Unidad (kWh/día)", ds: [lineDataset("U1 (kWh)", prod.e_u1, 0), lineDataset("U2 (kWh)", prod.e_u2, 4)], y: "kWh" },
            { title: "Potencia Promedio (kW/día)", ds: [lineDataset("Potencia U1", prod.pot_u1, 2), lineDataset("Potencia U2", prod.pot_u2, 3)], y: "kW" },
          ].map(({ title, ds, y }) => (
            <div key={title} style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, padding: "10px 12px 12px" }}>
              <div style={{ fontSize: 11, fontWeight: 700, color: "#1f2937", marginBottom: 2 }}>{title}</div>
              <div style={chartStyle}><Line data={{ labels: prod.labels, datasets: ds }} options={lineOpts(y)} /></div>
            </div>
          ))}
        </div>

        {/* Gráficos fila 3 */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: aforo ? 12 : 0 }}>
          {[
            { title: "Horas de Operación (h/día)", ds: [lineDataset("Horas U1", prod.h_u1, 5), lineDataset("Horas U2", prod.h_u2, 1)], y: "h" },
            { title: "Eficiencia (kWh/gal·día)", ds: [lineDataset("Eficiencia", prod.rend, 7)], y: "kWh/gal" },
          ].map(({ title, ds, y }) => (
            <div key={title} style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, padding: "10px 12px 12px" }}>
              <div style={{ fontSize: 11, fontWeight: 700, color: "#1f2937", marginBottom: 2 }}>{title}</div>
              <div style={chartStyle}><Line data={{ labels: prod.labels, datasets: ds }} options={lineOpts(y)} /></div>
            </div>
          ))}
        </div>

      </div>

      {/* ── SECCIÓN COMBUSTIBLE ───────────────────────────────────────── */}
      <div data-pdf-section="combustible" style={{ paddingTop: 24 }}>
        {/* Cabecera */}
        <div style={{ background: "linear-gradient(90deg,#0f172a,#292524)", borderRadius: 12, padding: "14px 20px", marginBottom: 18, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div>
            <div style={{ color: "#fcd34d", fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: 2 }}>Informe de Combustible & Tanques</div>
            <div style={{ color: "#fff", fontSize: 18, fontWeight: 900 }}>Central El Morro — Morro Energy S.A.</div>
            <div style={{ color: "#94a3b8", fontSize: 11, marginTop: 2 }}>Período: <b style={{ color: "#e2e8f0" }}>{period}</b> · Emisión: {today}</div>
          </div>
        </div>

        {/* KPIs combustible */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 10, marginBottom: 18 }}>
          {kpi("HFO consumido", fmt(resumen.hfoGal, 0), "gal", `${fmt(resumen.hfoGal + resumen.doGal > 0 ? resumen.hfoGal / (resumen.hfoGal + resumen.doGal) * 100 : 0, 1)}% del total`, P.blue)}
          {kpi("Diésel consumido", fmt(resumen.doGal, 0), "gal", `${fmt(resumen.hfoGal + resumen.doGal > 0 ? resumen.doGal / (resumen.hfoGal + resumen.doGal) * 100 : 0, 1)}% del total`, P.amber)}
          {kpi("Días con registro", `${resumen.dias}`, "días", `Total: ${fmt(resumen.hfoGal + resumen.doGal, 0)} gal`, P.green)}
        </div>

        {/* Gráficos combustible fila 1 */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 12 }}>
          {[
            { title: "Consumo total HFO vs Diésel (gal/día)", ds: [lineDataset("HFO total", prod.hfoTot.map(v => v || null), 0), lineDataset("Diésel total", prod.doTot.map(v => v || null), 2)], y: "gal" },
            { title: "Consumo HFO por Unidad (gal/día)", ds: [lineDataset("HFO G1", prod.hfoG1.map(v => v || null), 0), lineDataset("HFO G2", prod.hfoG2.map(v => v || null), 5)], y: "gal" },
          ].map(({ title, ds, y }) => (
            <div key={title} style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, padding: "10px 12px 12px" }}>
              <div style={{ fontSize: 11, fontWeight: 700, color: "#1f2937", marginBottom: 2 }}>{title}</div>
              <div style={chartStyle}><Line data={{ labels: prod.labels, datasets: ds }} options={lineOpts(y)} /></div>
            </div>
          ))}
        </div>

        {/* Gráficos combustible fila 2 */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 12 }}>
          {[
            { title: "Consumo Diésel por Unidad (gal/día)", ds: [lineDataset("Diésel G1", prod.doG1.map(v => v || null), 2), lineDataset("Diésel G2", prod.doG2.map(v => v || null), 4)], y: "gal" },
            ...(aforo
              ? [{ title: "Tanques HFO — T601 y T602 (gal · 00H00)", ds: [lineDataset("T601", aforo.t601, 0), lineDataset("T602", aforo.t602, 5)], y: "gal", labels: aforo.labels }]
              : []),
          ].map(({ title, ds, y, labels: lbl }) => (
            <div key={title} style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, padding: "10px 12px 12px" }}>
              <div style={{ fontSize: 11, fontWeight: 700, color: "#1f2937", marginBottom: 2 }}>{title}</div>
              <div style={chartStyle}><Line data={{ labels: lbl ?? prod.labels, datasets: ds }} options={lineOpts(y)} /></div>
            </div>
          ))}
        </div>

        {/* Gráficos combustible fila 3 */}
        {aforo && (
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            {[
              { title: "Tanques Diésel — T610 y T611 (gal · 00H00)", ds: [lineDataset("T610", aforo.t610, 2), lineDataset("T611", aforo.t611, 3)] },
              { title: "Cisterna 2 — Sludge (gal · 00H00)", ds: [lineDataset("Cisterna 2", aforo.cisterna2, 4)] },
            ].map(({ title, ds }) => (
              <div key={title} style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, padding: "10px 12px 12px" }}>
                <div style={{ fontSize: 11, fontWeight: 700, color: "#1f2937", marginBottom: 2 }}>{title}</div>
                <div style={chartStyle}><Line data={{ labels: aforo.labels, datasets: ds }} options={lineOpts("gal")} /></div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

// ─── Tipos ───────────────────────────────────────────────────────────────────
type Tab = "produccion" | "combustible";
interface ProcessedData { prod: ProdData; aforo: AforoData | null; resumen: Resumen; }

// ─── PÁGINA PRINCIPAL ─────────────────────────────────────────────────────────
export default function Metrics() {
  const { toast } = useToast();
  const { wbProd, wbAforo, fileNameProd, fileNameAforo, setProdEntry, setAforoEntry, prodLoading, aforoLoading } = useFileStore();

  const [sheets, setSheets] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [status, setStatus] = useState("");
  const [statusError, setStatusError] = useState(false);
  const [data, setData] = useState<ProcessedData | null>(null);
  const [activeTab, setActiveTab] = useState<Tab>("produccion");
  const [processing, setProcessing] = useState(false);
  const [exportingPdf, setExportingPdf] = useState(false);
  const [showPrintContainer, setShowPrintContainer] = useState(false);

  const printRef = useRef<HTMLDivElement>(null);

  const setMsg = (msg: string, err = false) => { setStatus(msg); setStatusError(err); };

  useEffect(() => {
    if (wbProd) {
      setSheets(wbProd.SheetNames);
      setSelectedSheet(prev => wbProd.SheetNames.includes(prev) ? prev : wbProd.SheetNames[0] || "");
    } else {
      setSheets([]);
      setSelectedSheet("");
    }
  }, [wbProd]);

  const onProdFile = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]; if (file) setProdEntry(file);
  }, [setProdEntry]);

  const onAforoFile = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]; if (file) setAforoEntry(file);
  }, [setAforoEntry]);

  const onProcess = useCallback(() => {
    if (!wbProd || !selectedSheet) return;
    setProcessing(true);
    setTimeout(() => {
      try {
        const prod = extractProduction(wbProd, selectedSheet);
        if (!prod || !prod.labels.length) {
          setMsg("No se encontraron datos válidos en esa hoja/mes.", true);
          setProcessing(false); return;
        }
        const aforo = wbAforo ? extractAforo(wbAforo, prod.targetMonth, prod.targetYear) : null;
        const resumen = buildResumen(prod);
        setData({ prod, aforo, resumen });
        setMsg(`${String(prod.targetMonth).padStart(2, "0")}/${prod.targetYear} — ${prod.labels.length} días procesados${!aforo ? " (sin aforo)" : ""}`);
      } catch (e: unknown) {
        setMsg(`Error: ${e instanceof Error ? e.message : String(e)}`, true);
      }
      setProcessing(false);
    }, 50);
  }, [wbProd, wbAforo, selectedSheet]);

  const canProcess = !!wbProd && !!wbAforo && !!selectedSheet && !prodLoading && !aforoLoading;

  const handleExportPDF = useCallback(async () => {
    if (!data) return;
    setExportingPdf(true);

    // 1. Mostrar el contenedor de impresión (fuera de pantalla pero en el DOM)
    setShowPrintContainer(true);

    // 2. Esperar a que React renderice + Chart.js pinte todos los canvas
    await new Promise(r => setTimeout(r, 3000));

    try {
      // html2canvas + jsPDF directamente: sin clonar DOM → canvas de Chart.js se preserva.
      // Cada sección [data-pdf-section] se captura individualmente → sin cortes en gráficos.
      const html2canvas = (await import("html2canvas")).default;
      const { jsPDF }   = await import("jspdf");

      if (!printRef.current) return;
      const period  = `${String(data.prod.targetMonth).padStart(2, "0")}_${data.prod.targetYear}`;
      const sections = Array.from(
        printRef.current.querySelectorAll<HTMLElement>("[data-pdf-section]")
      );

      // A4 landscape: 297 × 210 mm
      const pdfW     = 297;
      const pdfH     = 210;
      const margin   = 6;
      const contentW = pdfW - margin * 2;
      const contentH = pdfH - margin * 2;

      const pdf = new jsPDF({ unit: "mm", format: "a4", orientation: "landscape" });
      let firstPage = true;

      for (const section of sections) {
        // Capturar la sección como canvas a alta resolución
        const canvas = await html2canvas(section, {
          scale: 2,
          useCORS: true,
          allowTaint: true,
          logging: false,
          width: 1122,
          windowWidth: 1122,
          backgroundColor: "#ffffff",
        });

        const imgW        = canvas.width;
        const imgH        = canvas.height;
        const pxPerMmW    = imgW / contentW;
        const pageHeightPx = contentH * pxPerMmW;

        let yPx = 0;
        while (yPx < imgH) {
          if (!firstPage) pdf.addPage();
          firstPage = false;

          const sliceH   = Math.min(pageHeightPx, imgH - yPx);
          const sliceMmH = sliceH / pxPerMmW;

          const slice = document.createElement("canvas");
          slice.width  = imgW;
          slice.height = Math.ceil(sliceH);
          const ctx = slice.getContext("2d")!;
          ctx.drawImage(canvas, 0, yPx, imgW, Math.ceil(sliceH), 0, 0, imgW, Math.ceil(sliceH));

          pdf.addImage(slice.toDataURL("image/png"), "PNG", margin, margin, contentW, sliceMmH);
          yPx += pageHeightPx;
        }
      }

      pdf.save(`Metricas_ElMorro_${period}.pdf`);
      toast({ title: "PDF generado", description: "Producción + Combustible exportados correctamente." });
    } catch (err) {
      console.error("PDF metrics error:", err);
      toast({ title: "Error al generar PDF", description: "Intenta de nuevo.", variant: "destructive" });
    } finally {
      setShowPrintContainer(false);
      setExportingPdf(false);
    }
  }, [data, toast]);

  const TABS: { id: Tab; label: string; icon: React.ElementType }[] = [
    { id: "produccion",  label: "Producción",           icon: Zap },
    { id: "combustible", label: "Combustible & Tanques", icon: FlameKindling },
  ];

  return (
    <Layout>
      <div className="min-h-full bg-gray-50/60">

        {/* ── Contenedor de impresión PDF ──────────────────────────────── */}
        {showPrintContainer && data && (
          <>
            {/* Overlay oscuro sobre toda la pantalla */}
            <div style={{
              position: "fixed", inset: 0, zIndex: 9999,
              background: "rgba(10,20,40,0.93)",
              display: "flex", flexDirection: "column",
              alignItems: "center", justifyContent: "center", gap: 14,
            }}>
              <div style={{ color: "#fff", fontSize: 17, fontWeight: 700, letterSpacing: "0.02em" }}>
                Generando PDF…
              </div>
              <div style={{ color: "#94a3b8", fontSize: 13 }}>
                Renderizando gráficos · Por favor espere
              </div>
            </div>
            {/* Contenedor de captura: visible en pantalla, cubierto por el overlay */}
            <div
              ref={printRef}
              style={{
                position: "fixed",
                top: 0,
                left: 0,
                width: "1122px",
                background: "#fff",
                zIndex: 9998,
                pointerEvents: "none",
                overflow: "hidden",
              }}
            >
              <PdfPrintContent data={data} />
            </div>
          </>
        )}

        {/* ── Hero ─────────────────────────────────────────────────────── */}
        <div className="bg-gradient-to-r from-slate-900 via-blue-950 to-slate-900 px-6 py-6">
          <div className="max-w-[1280px] mx-auto flex items-center justify-between gap-4 flex-wrap">
            <div>
              <div className="flex items-center gap-2 mb-1">
                <BarChart2 className="w-5 h-5 text-blue-400" />
                <span className="text-blue-400 text-xs font-bold uppercase tracking-widest">Panel de Métricas</span>
              </div>
              <h1 className="text-white text-2xl font-black tracking-tight">Central El Morro</h1>
              <p className="text-slate-400 text-sm mt-0.5">Producción · Clientes · Eficiencia · Combustible · Tanques</p>
            </div>
            {data && (
              <div className="bg-white/10 backdrop-blur border border-white/10 rounded-xl px-5 py-3 text-white">
                <div className="text-xs text-slate-300 uppercase tracking-widest font-bold mb-1">Período activo</div>
                <div className="text-2xl font-black">
                  {String(data.prod.targetMonth).padStart(2, "0")} / {data.prod.targetYear}
                </div>
                <div className="text-xs text-slate-400 mt-0.5">{data.prod.labels.length} días · {data.prod.sheetName}</div>
              </div>
            )}
          </div>
        </div>

        <div className="max-w-[1280px] mx-auto px-6 py-6 space-y-5">

          {/* ── Panel de carga ─────────────────────────────────────────── */}
          <div className="bg-white rounded-2xl border border-gray-100 shadow-sm p-5">
            <div className="text-xs font-extrabold uppercase tracking-widest text-gray-400 mb-1">
              Archivos de entrada
            </div>
            <p className="text-xs text-gray-400 mb-4">
              Si ya cargaste los archivos en el Generador, aparecen listos aquí automáticamente.
            </p>
            <div className="flex flex-wrap gap-4 items-end">
              <FileIndicator testId="input-metrics-gen" label="Producción (GEN)"
                loaded={!!wbProd} fileName={fileNameProd} onChange={onProdFile} />
              <FileIndicator testId="input-metrics-aforo" label="Aforo (Tanques)"
                loaded={!!wbAforo} fileName={fileNameAforo} onChange={onAforoFile} />

              <div className="flex flex-col gap-1.5 min-w-[180px]">
                <label className="text-xs font-bold uppercase tracking-wider text-gray-400">Mes / Hoja</label>
                <select
                  id="metricsSheet" data-testid="select-metrics-sheet"
                  value={selectedSheet} onChange={e => setSelectedSheet(e.target.value)}
                  disabled={!sheets.length}
                  className="text-sm border border-gray-200 rounded-lg px-3 py-2 bg-white min-h-[42px]
                    disabled:opacity-40 focus:outline-none focus:ring-2 focus:ring-blue-500/30 focus:border-blue-400"
                >
                  {!sheets.length && <option value="">— cargue el GEN —</option>}
                  {sheets.map(s => <option key={s} value={s}>{s}</option>)}
                </select>
              </div>

              <button
                data-testid="button-metrics-process"
                onClick={onProcess} disabled={!canProcess || processing}
                className="flex items-center gap-2 px-7 py-2.5 rounded-xl bg-blue-600 text-white font-bold text-sm
                  shadow-lg shadow-blue-600/25 hover:bg-blue-700 disabled:opacity-40 disabled:cursor-not-allowed
                  disabled:shadow-none transition-all duration-150 min-h-[42px]"
              >
                {processing
                  ? <svg className="w-4 h-4 animate-spin" fill="none" viewBox="0 0 24 24">
                      <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"/>
                      <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"/>
                    </svg>
                  : <ChevronRight className="w-4 h-4" />}
                {processing ? "Procesando…" : "Procesar"}
              </button>
            </div>

            {status && (
              <div className={`flex items-center gap-2 mt-4 pt-4 border-t border-gray-100 text-xs font-medium
                ${statusError ? "text-red-600" : "text-gray-500"}`}>
                {statusError
                  ? <AlertCircle className="w-3.5 h-3.5 shrink-0" />
                  : <CheckCircle2 className="w-3.5 h-3.5 shrink-0 text-green-500" />}
                {status}
              </div>
            )}
          </div>

          {/* ── Resultados ─────────────────────────────────────────────── */}
          {data ? (
            <>
              {/* Tabs + botón PDF */}
              <div className="flex items-center justify-between gap-3 flex-wrap">
                <div className="flex gap-2 p-1 bg-gray-100 rounded-xl w-fit">
                  {TABS.map(({ id, label, icon: Icon }) => (
                    <button key={id} data-testid={`tab-metrics-${id}`}
                      onClick={() => setActiveTab(id)}
                      className={`flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-semibold transition-all
                        ${activeTab === id ? "bg-white shadow text-gray-900" : "text-gray-500 hover:text-gray-700"}`}
                    >
                      <Icon className="w-4 h-4" />{label}
                    </button>
                  ))}
                </div>

                <button
                  data-testid="button-metrics-pdf"
                  onClick={handleExportPDF} disabled={exportingPdf}
                  className="flex items-center gap-2 px-5 py-2 rounded-xl border border-gray-200 bg-white
                    text-sm font-semibold text-gray-700 hover:bg-gray-50 shadow-sm
                    disabled:opacity-60 disabled:cursor-not-allowed transition-all"
                >
                  {exportingPdf
                    ? <>
                        <svg className="w-4 h-4 animate-spin text-blue-600" fill="none" viewBox="0 0 24 24">
                          <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"/>
                          <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"/>
                        </svg>
                        Generando PDF…
                      </>
                    : <><FileDown className="w-4 h-4 text-blue-600" /> Exportar PDF completo</>}
                </button>
              </div>

              {/* ══ TAB PRODUCCIÓN ══════════════════════════════════════ */}
              {activeTab === "produccion" && (
                <div className="space-y-6">
                  <section>
                    <SectionHeader icon={Zap} title="Energía del período" sub="Totales y distribución mensual" accent={P.blue} />
                    <div className="grid grid-cols-2 md:grid-cols-3 xl:grid-cols-5 gap-3">
                      <KpiCard icon={Zap} accent={P.blue} title="Energía total"
                        value={fmt(data.resumen.energiaTotalMWh, 1)} unit="MWh"
                        sub={`U1 ${fmt(data.resumen.energiaU1MWh, 1)} · U2 ${fmt(data.resumen.energiaU2MWh, 1)} MWh`} />
                      <KpiCard icon={TrendingUp} accent={P.green} title="LANEC"
                        value={fmt(data.resumen.energiaLanecMWh, 1)} unit="MWh"
                        sub={`${fmt(data.resumen.energiaTotalMWh > 0 ? data.resumen.energiaLanecMWh / data.resumen.energiaTotalMWh * 100 : 0, 1)}% del total`} />
                      <KpiCard icon={TrendingUp} accent={P.amber} title="GRACA"
                        value={fmt(data.resumen.energiaGracaMWh, 1)} unit="MWh"
                        sub={`${fmt(data.resumen.energiaTotalMWh > 0 ? data.resumen.energiaGracaMWh / data.resumen.energiaTotalMWh * 100 : 0, 1)}% del total`} />
                      <KpiCard icon={Clock} accent={P.purple} title="Horas operación"
                        value={fmt(data.resumen.horasU1, 0)} unit="h U1"
                        sub={`U2: ${fmt(data.resumen.horasU2, 0)} h`} />
                      <KpiCard icon={TrendingUp} accent={P.cyan} title="Eficiencia"
                        value={data.resumen.eficProm != null ? fmt(data.resumen.eficProm, 2) : "—"} unit="kWh/gal"
                        sub={`Mes ${String(data.prod.targetMonth).padStart(2, "0")}/${data.prod.targetYear}`} />
                    </div>
                  </section>

                  <section>
                    <SectionHeader icon={BarChart2} title="Gráficas diarias" sub="Tendencias del período" accent={P.blue} />
                    <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
                      <ChartCard title="Energía Total" subtitle="kWh generados por día" accent={P.blue}>
                        <Line data={{ labels: data.prod.labels, datasets: [lineDataset("Energía total (kWh)", data.prod.etotal, 0)] }} options={lineOpts("kWh")} />
                      </ChartCard>
                      <ChartCard title="Energía por Cliente" subtitle="LANEC · GRACA · Auxiliares" accent={P.green}>
                        <Line data={{ labels: data.prod.labels, datasets: [lineDataset("LANEC", data.prod.lanec, 0), lineDataset("GRACA", data.prod.graca, 1), lineDataset("Auxiliares", data.prod.aux, 2)] }} options={lineOpts("kWh")} />
                      </ChartCard>
                      <ChartCard title="Energía por Unidad" subtitle="U1 (9L26) · U2 (9SW280)" accent={P.purple}>
                        <Line data={{ labels: data.prod.labels, datasets: [lineDataset("Unidad 1 (kWh)", data.prod.e_u1, 0), lineDataset("Unidad 2 (kWh)", data.prod.e_u2, 4)] }} options={lineOpts("kWh")} />
                      </ChartCard>
                      <ChartCard title="Potencia Promedio" subtitle="kW por unidad (energía / horas)" accent={P.amber}>
                        <Line data={{ labels: data.prod.labels, datasets: [lineDataset("Potencia U1 (kW)", data.prod.pot_u1, 2), lineDataset("Potencia U2 (kW)", data.prod.pot_u2, 3)] }} options={lineOpts("kW")} />
                      </ChartCard>
                      <ChartCard title="Horas de Operación" subtitle="Horas diarias por unidad" accent={P.cyan}>
                        <Line data={{ labels: data.prod.labels, datasets: [lineDataset("Horas U1 (h)", data.prod.h_u1, 5), lineDataset("Horas U2 (h)", data.prod.h_u2, 1)] }} options={lineOpts("h")} />
                      </ChartCard>
                      <ChartCard title="Eficiencia" subtitle="kWh por galón consumido" accent={P.teal}>
                        <Line data={{ labels: data.prod.labels, datasets: [lineDataset("Eficiencia (kWh/gal)", data.prod.rend, 7)] }} options={lineOpts("kWh/gal")} />
                      </ChartCard>
                    </div>
                  </section>

                  {data.aforo && (
                    <section>
                      <SectionHeader icon={Database} title="Niveles de tanques" sub="Volúmenes a las 00H00 diarios" accent={P.cyan} />
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <ChartCard title="Todos los Tanques & Sludge" subtitle="T601 · T602 · T610 · T611 · Cisterna 2" accent={P.cyan} height={300}>
                          <Line data={{ labels: data.aforo.labels, datasets: [
                            lineDataset("T601 HFO (gal)", data.aforo.t601, 0),
                            lineDataset("T602 HFO (gal)", data.aforo.t602, 1),
                            lineDataset("T610 Diesel (gal)", data.aforo.t610, 2),
                            lineDataset("T611 Diesel (gal)", data.aforo.t611, 3),
                            lineDataset("Cisterna 2 (gal)", data.aforo.cisterna2, 4),
                          ]}} options={lineOpts("gal")} />
                        </ChartCard>
                      </div>
                    </section>
                  )}
                </div>
              )}

              {/* ══ TAB COMBUSTIBLE ══════════════════════════════════════ */}
              {activeTab === "combustible" && (
                <div className="space-y-6">
                  <div className="rounded-xl overflow-hidden border border-slate-200">
                    <div className="bg-gradient-to-r from-slate-900 to-slate-800 px-5 py-3 flex items-center justify-between gap-3 flex-wrap">
                      <div>
                        <div className="text-white font-black text-sm tracking-wide">Informe Gerencial de Combustible — Central El Morro</div>
                        <div className="text-slate-400 text-xs mt-0.5">
                          Mes: <strong className="text-slate-200">{String(data.prod.targetMonth).padStart(2, "0")}/{data.prod.targetYear}</strong>
                          &nbsp;·&nbsp; Emisión: <strong className="text-slate-200">{new Date().toLocaleDateString("es-EC")}</strong>
                        </div>
                      </div>
                      <Fuel className="w-8 h-8 text-amber-400 opacity-70" />
                    </div>
                  </div>

                  <section>
                    <SectionHeader icon={Droplets} title="Consumo del período" sub="Totales mensuales de combustible" accent={P.amber} />
                    <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                      <KpiCard icon={Droplets} accent={P.blue} title="HFO consumido"
                        value={fmt(data.resumen.hfoGal, 0)} unit="gal"
                        sub={`${fmt(data.resumen.hfoGal + data.resumen.doGal > 0 ? data.resumen.hfoGal / (data.resumen.hfoGal + data.resumen.doGal) * 100 : 0, 1)}% del total`} />
                      <KpiCard icon={Fuel} accent={P.amber} title="Diésel consumido"
                        value={fmt(data.resumen.doGal, 0)} unit="gal"
                        sub={`${fmt(data.resumen.hfoGal + data.resumen.doGal > 0 ? data.resumen.doGal / (data.resumen.hfoGal + data.resumen.doGal) * 100 : 0, 1)}% del total`} />
                      <KpiCard icon={BarChart2} accent={P.green} title="Días con registro"
                        value={`${data.resumen.dias}`} unit="días"
                        sub={`Total: ${fmt(data.resumen.hfoGal + data.resumen.doGal, 0)} gal`} />
                    </div>
                  </section>

                  <section>
                    <SectionHeader icon={FlameKindling} title="Consumo por tipo y unidad" sub="Prorrateo por energía generada" accent={P.amber} />
                    <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
                      <ChartCard title="Consumo total HFO vs Diésel" subtitle="Galones diarios totales" accent={P.blue}>
                        <Line data={{ labels: data.prod.labels, datasets: [lineDataset("HFO total (gal)", data.prod.hfoTot.map(v => v || null), 0), lineDataset("Diésel total (gal)", data.prod.doTot.map(v => v || null), 2)] }} options={lineOpts("gal")} />
                      </ChartCard>
                      <ChartCard title="HFO por Unidad" subtitle="Reparto entre G1 y G2 por energía" accent={P.blue}>
                        <Line data={{ labels: data.prod.labels, datasets: [lineDataset("HFO G1 (gal)", data.prod.hfoG1.map(v => v || null), 0), lineDataset("HFO G2 (gal)", data.prod.hfoG2.map(v => v || null), 5)] }} options={lineOpts("gal")} />
                      </ChartCard>
                      <ChartCard title="Diésel por Unidad" subtitle="Reparto entre G1 y G2 por energía" accent={P.amber}>
                        <Line data={{ labels: data.prod.labels, datasets: [lineDataset("Diésel G1 (gal)", data.prod.doG1.map(v => v || null), 2), lineDataset("Diésel G2 (gal)", data.prod.doG2.map(v => v || null), 4)] }} options={lineOpts("gal")} />
                      </ChartCard>
                      <ChartCard title="Tanques HFO — T601 y T602" subtitle="Volumen (gal) a las 00H00" accent={P.blue}>
                        {data.aforo ? <Line data={{ labels: data.aforo.labels, datasets: [lineDataset("T601 (HFO, gal)", data.aforo.t601, 0), lineDataset("T602 (HFO, gal)", data.aforo.t602, 5)] }} options={lineOpts("gal")} /> : <NoAforo />}
                      </ChartCard>
                      <ChartCard title="Tanques Diésel — T610 y T611" subtitle="Volumen (gal) a las 00H00" accent={P.amber}>
                        {data.aforo ? <Line data={{ labels: data.aforo.labels, datasets: [lineDataset("T610 (Diesel, gal)", data.aforo.t610, 2), lineDataset("T611 (Diesel, gal)", data.aforo.t611, 3)] }} options={lineOpts("gal")} /> : <NoAforo />}
                      </ChartCard>
                      <ChartCard title="Cisterna 2 — Sludge" subtitle='Volumen registrado como "CISTERNA 2"' accent={P.purple}>
                        {data.aforo ? <Line data={{ labels: data.aforo.labels, datasets: [lineDataset("Cisterna 2 (gal)", data.aforo.cisterna2, 4)] }} options={lineOpts("gal")} /> : <NoAforo />}
                      </ChartCard>
                    </div>
                  </section>
                </div>
              )}
            </>
          ) : (
            /* ── Estado vacío ────────────────────────────────────────── */
            <div className="bg-white rounded-2xl border border-gray-100 shadow-sm">
              <div className="flex flex-col items-center justify-center py-20 px-6 text-center">
                <div className="w-16 h-16 rounded-2xl bg-blue-50 flex items-center justify-center mb-5">
                  <BarChart2 className="w-8 h-8 text-blue-400" />
                </div>
                <h3 className="text-base font-bold text-gray-700 mb-2">Sin datos procesados</h3>
                <p className="text-sm text-gray-400 max-w-sm mb-6">
                  {wbProd && wbAforo
                    ? "Los archivos ya están cargados. Selecciona el mes y presiona Procesar."
                    : "Cargue el archivo GEN y el de Aforo, seleccione el mes y presione Procesar."}
                </p>
                <div className="flex items-center gap-3 text-xs text-gray-400 flex-wrap justify-center">
                  <div className={`flex items-center gap-1.5 rounded-lg px-3 py-2 border
                    ${wbProd ? "bg-green-50 border-green-200 text-green-700" : "bg-gray-50 border-gray-200"}`}>
                    {wbProd ? <CheckCircle2 className="w-3.5 h-3.5" /> : <UploadCloud className="w-3.5 h-3.5" />}
                    GEN {wbProd ? "listo" : "pendiente"}
                  </div>
                  <ChevronRight className="w-3.5 h-3.5" />
                  <div className={`flex items-center gap-1.5 rounded-lg px-3 py-2 border
                    ${wbAforo ? "bg-green-50 border-green-200 text-green-700" : "bg-gray-50 border-gray-200"}`}>
                    {wbAforo ? <CheckCircle2 className="w-3.5 h-3.5" /> : <UploadCloud className="w-3.5 h-3.5" />}
                    Aforo {wbAforo ? "listo" : "pendiente"}
                  </div>
                  <ChevronRight className="w-3.5 h-3.5" />
                  <div className="flex items-center gap-1.5 bg-blue-50 border border-blue-200 rounded-lg px-3 py-2 text-blue-600">
                    <ChevronRight className="w-3.5 h-3.5" /> Procesar
                  </div>
                </div>
              </div>
            </div>
          )}
        </div>
      </div>
    </Layout>
  );
}
