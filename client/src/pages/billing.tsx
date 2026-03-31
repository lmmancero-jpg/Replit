import { useState, useCallback } from "react";
import { format } from "date-fns";
import {
  FileDown, Calendar, AlertCircle, CheckCircle2,
  FileSpreadsheet, DollarSign, Save, Activity,
} from "lucide-react";
import { Layout } from "@/components/layout";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Tabs, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { useToast } from "@/hooks/use-toast";
import { useFileStore } from "@/lib/fileStore";
import { useInvoices } from "@/hooks/use-invoices";
import { buildClientBillingReport, buildRealBillingReport, getMonthlyProductionSummary } from "@/lib/billingEngine";
import { apiRequest } from "@/lib/queryClient";
import { queryClient } from "@/lib/queryClient";
import { openPrintWindow } from "@/lib/printPDF";
import { fmt } from "@/lib/reportEngine";

type ReportMode = "clientes" | "real";

function readLS(key: string, fallback: number): number {
  try {
    const v = parseFloat(localStorage.getItem(key) ?? "");
    return isNaN(v) ? fallback : v;
  } catch { return fallback; }
}

export default function BillingPage() {
  const { toast } = useToast();
  const { wbProd, prodFile, fileNameProd, setProdEntry } = useFileStore();

  const [period, setPeriod] = useState<string>(format(new Date(), "yyyy-MM"));
  const [mode, setMode] = useState<ReportMode>("clientes");
  const [u1Downtime, setU1Downtime] = useState<number>(() => readLS("nexus_u1dt", 0));
  const [u2Downtime, setU2Downtime] = useState<number>(() => readLS("nexus_u2dt", 0));
  const [generatedHtml, setGeneratedHtml] = useState<string | null>(null);
  const [currentMode, setCurrentMode] = useState<ReportMode>("clientes");
  const [isGenerating, setIsGenerating] = useState(false);
  const [isSaving, setIsSaving] = useState(false);

  const { data: invoiceList = [] } = useInvoices(period);

  // Producción del período
  let prodSummary = null as ReturnType<typeof getMonthlyProductionSummary> | null;
  if (wbProd) {
    try { prodSummary = getMonthlyProductionSummary(wbProd, period); } catch { prodSummary = null; }
  }

  const hasProd = !!wbProd && prodSummary !== null;
  const hasInvoices = invoiceList.length > 0;

  // Suma de combustible para informe clientes
  const invoiceCombTotal = invoiceList
    .filter((i) => i.category === "combustible_transporte")
    .reduce((acc, i) => acc + parseFloat(i.total ?? "0"), 0);

  const handleGenerate = useCallback(async () => {
    if (!wbProd) {
      toast({ title: "Archivo de producción requerido", description: "Carga el archivo de producción en el Generador.", variant: "destructive" });
      return;
    }
    setIsGenerating(true);
    setCurrentMode(mode);
    try {
      let html = "";
      if (mode === "clientes") {
        html = buildClientBillingReport({
          wbProd,
          mesStr: period,
          diasFallaU1: u1Downtime,
          diasFallaU2: u2Downtime,
          invoiceCombTotal,
          hasProduction: hasProd,
        });
      } else {
        html = buildRealBillingReport({
          wbProd,
          mesStr: period,
          diasFallaU1: u1Downtime,
          diasFallaU2: u2Downtime,
          invoiceList,
          hasProduction: hasProd,
        });
      }
      setGeneratedHtml(html);
      toast({ title: "Informe generado", description: "Revisa la previsualización antes de exportar." });
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : "Error desconocido";
      toast({ title: "Error al generar", description: msg, variant: "destructive" });
    } finally {
      setIsGenerating(false);
    }
  }, [wbProd, mode, period, u1Downtime, u2Downtime, invoiceCombTotal, invoiceList, hasProd, toast]);

  const handleExportPDF = async () => {
    if (!generatedHtml) return;
    const modeLabel = currentMode === "clientes" ? "Facturacion_Clientes" : "Facturacion_Real";
    const filename = `${modeLabel}_ElMorro_${period}.pdf`;
    try {
      const res = await fetch("/api/export/pdf", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ html: generatedHtml, title: filename }),
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = filename; a.click();
      URL.revokeObjectURL(url);
      toast({ title: "PDF generado" });
    } catch {
      openPrintWindow(generatedHtml, filename.replace(".pdf", ""));
      toast({ title: "Ventana de impresión abierta", description: 'Selecciona "Guardar como PDF".' });
    }
  };

  const handleSave = async () => {
    if (!generatedHtml) return;
    setIsSaving(true);
    const reportType = currentMode === "clientes" ? "facturacion_clientes" : "facturacion_real";
    const modeLabel = currentMode === "clientes" ? "Facturación Clientes" : "Facturación Real";
    try {
      await apiRequest("POST", "/api/reports", {
        title: `${modeLabel} – ${period}`,
        reportType,
        date: period,
        content: generatedHtml,
      });
      queryClient.invalidateQueries({ queryKey: ["/api/reports"] });
      toast({ title: "Reporte guardado correctamente" });
    } catch (err) {
      toast({ title: "Error al guardar", description: String(err), variant: "destructive" });
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <Layout>
      <div className="flex flex-col xl:flex-row gap-6 h-full">
        {/* Columna izquierda */}
        <div className="w-full xl:w-[380px] shrink-0 flex flex-col gap-4">
          {/* Parámetros */}
          <Card className="border-border/60 shadow-md">
            <div className="h-1 w-full rounded-t-lg bg-gradient-to-r from-primary to-primary/50" />
            <CardHeader className="pb-3">
              <CardTitle className="text-base flex items-center gap-2">
                <DollarSign className="w-4 h-4 text-primary" />
                Parámetros de facturación
              </CardTitle>
              <CardDescription className="text-xs">Central El Morro – Morro Energy S.A.</CardDescription>
            </CardHeader>
            <CardContent className="space-y-4">
              {/* Período */}
              <div>
                <label className="text-xs font-medium flex items-center gap-1 mb-1">
                  <Calendar className="w-3 h-3" /> Período (AAAA-MM)
                </label>
                <Input
                  data-testid="input-billing-period"
                  type="month"
                  value={period}
                  onChange={(e) => { setPeriod(e.target.value); setGeneratedHtml(null); }}
                  className="text-sm h-8"
                />
              </div>

              {/* Indisponibilidad */}
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-xs font-medium mb-1 block">Días indisp. U1</label>
                  <Input
                    data-testid="input-billing-u1"
                    type="number" min="0" max="31"
                    value={u1Downtime}
                    onChange={(e) => {
                      const v = parseInt(e.target.value, 10) || 0;
                      setU1Downtime(v);
                      localStorage.setItem("nexus_u1dt", String(v));
                    }}
                    className="text-sm h-8"
                  />
                </div>
                <div>
                  <label className="text-xs font-medium mb-1 block">Días indisp. U2</label>
                  <Input
                    data-testid="input-billing-u2"
                    type="number" min="0" max="31"
                    value={u2Downtime}
                    onChange={(e) => {
                      const v = parseInt(e.target.value, 10) || 0;
                      setU2Downtime(v);
                      localStorage.setItem("nexus_u2dt", String(v));
                    }}
                    className="text-sm h-8"
                  />
                </div>
              </div>

              {/* Archivo de producción */}
              <div className="rounded-md border border-border/50 bg-muted/30 p-3 space-y-2">
                <label className="text-xs font-medium flex items-center gap-1">
                  <FileSpreadsheet className="w-3 h-3 text-green-600" />
                  Archivo de producción (.xlsx)
                </label>
                {fileNameProd ? (
                  <p className="text-xs text-green-600 flex items-center gap-1">
                    <CheckCircle2 className="w-3 h-3 shrink-0" /> {fileNameProd}
                    {prodSummary && (
                      <span className="text-muted-foreground ml-1">({fmt(prodSummary.tot_gen, 0)} kWh)</span>
                    )}
                  </p>
                ) : (
                  <>
                    <p className="text-xs text-amber-600 flex items-center gap-1">
                      <AlertCircle className="w-3 h-3" /> Sin archivo cargado
                    </p>
                    <Input
                      data-testid="input-billing-prod-file"
                      type="file" accept=".xlsx,.xls"
                      className="text-xs h-8 cursor-pointer"
                      onChange={(e) => { const f = e.target.files?.[0]; if (f) setProdEntry(f); }}
                    />
                  </>
                )}
              </div>

              {/* Estado de datos */}
              <div className="rounded-md border border-border/40 bg-muted/20 p-3 space-y-1.5">
                <p className="text-xs font-semibold text-muted-foreground uppercase tracking-wider mb-2">Estado del período</p>
                <DataStatus ok={hasProd} label="Producción cargada" />
                <DataStatus ok={hasInvoices} label={`Facturas del período (${invoiceList.length})`} />
                {!hasProd && (
                  <p className="text-[10px] text-muted-foreground mt-1">Carga el archivo de producción en Generador o aquí arriba.</p>
                )}
                {!hasInvoices && (
                  <p className="text-[10px] text-muted-foreground mt-1">Registra las facturas del período en el módulo Facturas.</p>
                )}
              </div>

              {/* Tipo de informe */}
              <div>
                <label className="text-xs font-medium mb-2 block">Tipo de informe</label>
                <Tabs value={mode} onValueChange={(v) => { setMode(v as ReportMode); setGeneratedHtml(null); }}>
                  <TabsList className="w-full">
                    <TabsTrigger data-testid="tab-clientes" value="clientes" className="flex-1 text-xs">Informe Clientes</TabsTrigger>
                    <TabsTrigger data-testid="tab-real" value="real" className="flex-1 text-xs">Informe Real</TabsTrigger>
                  </TabsList>
                </Tabs>
                <p className="text-[10px] text-muted-foreground mt-1.5">
                  {mode === "clientes"
                    ? "Combustible ajustado por facturas reales. Demás rubros: tarifa contractual."
                    : "Todos los rubros calculados desde facturas reales + margen variable contractual."}
                </p>
              </div>

              <Button
                data-testid="button-generate-billing"
                className="w-full"
                onClick={handleGenerate}
                disabled={isGenerating || !wbProd}
              >
                {isGenerating ? "Procesando..." : `Generar informe ${mode === "clientes" ? "clientes" : "real"}`}
              </Button>
            </CardContent>
          </Card>

          {/* Combustible del período (solo modo clientes) */}
          {mode === "clientes" && (
            <Card className="border-border/60 shadow-sm">
              <CardHeader className="pb-2">
                <CardTitle className="text-sm flex items-center gap-2">
                  <Activity className="w-4 h-4 text-muted-foreground" />
                  Combustible + Transporte
                </CardTitle>
              </CardHeader>
              <CardContent>
                <p className="text-xs text-muted-foreground mb-1">Total facturado período:</p>
                <p className="text-xl font-bold text-primary">$ {invoiceCombTotal.toLocaleString("es-EC", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</p>
                {hasProd && prodSummary && prodSummary.tot_gen > 0 && (
                  <p className="text-xs text-muted-foreground mt-1">
                    → CV ajustado: {fmt(invoiceCombTotal / prodSummary.tot_gen, 4)} USD/kWh
                  </p>
                )}
                {!hasInvoices && (
                  <p className="text-xs text-amber-600 mt-1 flex items-center gap-1">
                    <AlertCircle className="w-3 h-3" /> Sin facturas — se usará tarifa contractual (0,1153 USD/kWh)
                  </p>
                )}
              </CardContent>
            </Card>
          )}
        </div>

        {/* Columna derecha – previsualización */}
        <div className="flex-1 flex flex-col min-h-[600px] xl:min-h-0 rounded-lg border border-border/60 overflow-hidden shadow-inner">
          <div className="h-12 bg-card border-b border-border/60 flex items-center justify-between px-4 shrink-0">
            <h3 className="text-sm font-semibold flex items-center gap-2">
              <span className={`inline-flex h-2 w-2 rounded-full ${generatedHtml ? "bg-green-500" : "bg-muted-foreground/40"}`} />
              {generatedHtml
                ? `Informe ${currentMode === "clientes" ? "Clientes" : "Real"} – ${period}`
                : "Previsualización"}
            </h3>
            <div className="flex items-center gap-2">
              <Button
                data-testid="button-save-billing"
                variant="outline" size="sm"
                onClick={handleSave}
                disabled={!generatedHtml || isSaving}
              >
                <Save className="w-3.5 h-3.5 mr-1" />
                {isSaving ? "Guardando..." : "Guardar"}
              </Button>
              <Button
                data-testid="button-export-pdf-billing"
                variant="outline" size="sm"
                onClick={handleExportPDF}
                disabled={!generatedHtml}
              >
                <FileDown className="w-3.5 h-3.5 mr-1" /> PDF
              </Button>
            </div>
          </div>
          <div className="flex-1 overflow-auto bg-slate-50 dark:bg-slate-900/50 p-6">
            {isGenerating ? (
              <div className="h-full flex flex-col items-center justify-center text-muted-foreground">
                <div className="w-12 h-12 border-4 border-primary/20 border-t-primary rounded-full animate-spin mb-3" />
                <p className="text-sm font-medium">Procesando datos del período...</p>
              </div>
            ) : generatedHtml ? (
              <div
                className="report-wrapper bg-white shadow-sm rounded-md p-6 max-w-5xl mx-auto report-content"
                dangerouslySetInnerHTML={{ __html: generatedHtml }}
              />
            ) : (
              <div className="h-full flex flex-col items-center justify-center text-muted-foreground/60 border-2 border-dashed border-border/60 rounded-lg mx-auto max-w-md">
                <AlertCircle className="w-12 h-12 mb-3 opacity-40" />
                <p className="text-sm font-medium text-foreground/50">Selecciona el período y genera el informe</p>
                <p className="text-xs mt-1 text-center max-w-xs">
                  Asegúrate de tener el archivo de producción cargado y las facturas del período registradas.
                </p>
              </div>
            )}
          </div>
        </div>
      </div>
    </Layout>
  );
}

function DataStatus({ ok, label }: { ok: boolean; label: string }) {
  return (
    <div className="flex items-center gap-2 text-xs">
      {ok ? (
        <CheckCircle2 className="w-3.5 h-3.5 text-green-500 shrink-0" />
      ) : (
        <AlertCircle className="w-3.5 h-3.5 text-amber-500 shrink-0" />
      )}
      <span className={ok ? "text-foreground" : "text-muted-foreground"}>{label}</span>
    </div>
  );
}
