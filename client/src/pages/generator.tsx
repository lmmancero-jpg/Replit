import { useState, useRef, useCallback } from "react";
import { z } from "zod";
import { useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { format } from "date-fns";
import { FileDown, Save, Calendar, FileSpreadsheet, Activity, Factory, FileText, AlertCircle, Settings, CheckCircle2 } from "lucide-react";
import { useFileStore } from "@/lib/fileStore";

import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Form, FormControl, FormField, FormItem, FormLabel, FormMessage } from "@/components/ui/form";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { useToast } from "@/hooks/use-toast";
import { Layout } from "@/components/layout";
import { useCreateReport } from "@/hooks/use-reports";
import {
  generarInformeDiario,
  generarInformeMensual,
  generarInformeFacturacion,
} from "@/lib/reportEngine";

const generatorSchema = z.object({
  reportDate: z.string().min(1, "La fecha es requerida"),
  reportMonth: z.string().min(1, "El mes es requerido"),
  u1Downtime: z.coerce.number().min(0).max(31).default(0),
  u2Downtime: z.coerce.number().min(0).max(31).default(0),
  observations: z.string().optional(),
});

type GeneratorValues = z.infer<typeof generatorSchema>;

export default function Generator() {
  const { toast } = useToast();
  const createReport = useCreateReport();
  const { prodFile, aforoFile, fileNameProd, fileNameAforo, setProdEntry, setAforoEntry } = useFileStore();
  const [isGenerating, setIsGenerating] = useState(false);
  const [generatedHtml, setGeneratedHtml] = useState<string | null>(null);
  const [currentReportType, setCurrentReportType] = useState<string>("");
  const previewRef = useRef<HTMLDivElement>(null);

  const form = useForm<GeneratorValues>({
    resolver: zodResolver(generatorSchema),
    defaultValues: {
      reportDate: format(new Date(), "yyyy-MM-dd"),
      reportMonth: format(new Date(), "yyyy-MM"),
      u1Downtime: 0,
      u2Downtime: 0,
      observations: "",
    },
  });

  const readFileAsArrayBuffer = (file: File): Promise<ArrayBuffer> =>
    new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target!.result as ArrayBuffer);
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });

  const handleGenerate = useCallback(async (type: "diario" | "mensual" | "facturacion") => {
    const data = form.getValues();
    setIsGenerating(true);
    setCurrentReportType(type);

    try {
      if (!prodFile) throw new Error("Carga el archivo de producción (.xlsx).");
      const prodBuffer = await readFileAsArrayBuffer(prodFile);

      let html = "";

      if (type === "diario") {
        if (!data.reportDate) throw new Error("Selecciona la fecha del informe diario.");
        html = generarInformeDiario(prodBuffer, data.reportDate, data.observations || "");
      } else if (type === "mensual") {
        if (!data.reportMonth) throw new Error("Selecciona el mes del informe mensual.");
        html = generarInformeMensual(prodBuffer, data.reportMonth);
      } else if (type === "facturacion") {
        if (!data.reportMonth) throw new Error("Selecciona el mes para la facturación.");
        html = generarInformeFacturacion(
          prodBuffer,
          data.reportMonth,
          data.u1Downtime,
          data.u2Downtime
        );
      }

      setGeneratedHtml(html);
      toast({
        title: "Informe generado",
        description: "Revisa la previsualización antes de guardar o exportar.",
      });
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : "Error desconocido.";
      toast({ title: "Error al generar informe", description: msg, variant: "destructive" });
    } finally {
      setIsGenerating(false);
    }
  }, [form, prodFile, toast]);

  const handleSave = () => {
    if (!generatedHtml) return;
    const data = form.getValues();
    const date = currentReportType === "diario" ? data.reportDate : data.reportMonth;

    createReport.mutate({
      title: `${currentReportType === "diario" ? "Informe Diario" : currentReportType === "mensual" ? "Informe Mensual" : "Facturación"} – ${date}`,
      reportType: currentReportType,
      date,
      content: generatedHtml,
    }, {
      onSuccess: () => {
        toast({
          title: "Guardado exitoso",
          description: "El informe ha sido almacenado en el historial.",
        });
      },
      onError: () => {
        toast({ title: "Error al guardar", description: "No se pudo guardar el informe.", variant: "destructive" });
      },
    });
  };

  const handleExportPDF = async () => {
    if (!generatedHtml) return;
    const data = form.getValues();
    const date = currentReportType === "diario" ? data.reportDate : data.reportMonth;
    const typeLabel =
      currentReportType === "facturacion" ? "Facturacion"
      : currentReportType === "mensual"   ? "Reporte_Mensual"
      : "Reporte_Diario";
    const filename = `${typeLabel}_ElMorro_${date}.pdf`;

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
      a.href = url;
      a.download = filename;
      a.click();
      URL.revokeObjectURL(url);
      toast({ title: "PDF generado", description: "El archivo se está descargando." });
    } catch (err) {
      console.error("PDF export error:", err);
      toast({ title: "Error al generar PDF", description: "Intenta de nuevo.", variant: "destructive" });
    }
  };

  return (
    <Layout>
      <div className="flex flex-col xl:flex-row gap-6 h-full">
        {/* Columna izquierda */}
        <div className="w-full xl:w-[380px] shrink-0 flex flex-col gap-4">
          <Card className="border-border/60 shadow-md">
            <div className="h-1 w-full rounded-t-lg bg-gradient-to-r from-primary to-primary/50" />
            <CardHeader className="pb-3">
              <CardTitle className="text-base flex items-center gap-2">
                <Settings className="w-4 h-4 text-primary" />
                Parámetros de entrada
              </CardTitle>
              <CardDescription className="text-xs">
                Central El Morro – Morro Energy S.A.
              </CardDescription>
            </CardHeader>
            <CardContent>
              <Form {...form}>
                <form className="space-y-4">
                  <FormField
                    control={form.control}
                    name="reportDate"
                    render={({ field }) => (
                      <FormItem>
                        <FormLabel className="text-xs flex items-center gap-1">
                          <Calendar className="w-3 h-3" /> Fecha del informe diario
                        </FormLabel>
                        <FormControl>
                          <Input type="date" className="text-sm h-8" {...field} />
                        </FormControl>
                        <FormMessage />
                      </FormItem>
                    )}
                  />

                  <FormField
                    control={form.control}
                    name="reportMonth"
                    render={({ field }) => (
                      <FormItem>
                        <FormLabel className="text-xs flex items-center gap-1">
                          <Calendar className="w-3 h-3" /> Mes del informe mensual (AAAA-MM)
                        </FormLabel>
                        <FormControl>
                          <Input type="month" className="text-sm h-8" {...field} />
                        </FormControl>
                        <FormMessage />
                      </FormItem>
                    )}
                  />

                  <div className="grid grid-cols-2 gap-3">
                    <FormField
                      control={form.control}
                      name="u1Downtime"
                      render={({ field }) => (
                        <FormItem>
                          <FormLabel className="text-xs">Días indisp. U1</FormLabel>
                          <FormControl>
                            <Input type="number" min="0" className="text-sm h-8" {...field} />
                          </FormControl>
                          <FormMessage />
                        </FormItem>
                      )}
                    />
                    <FormField
                      control={form.control}
                      name="u2Downtime"
                      render={({ field }) => (
                        <FormItem>
                          <FormLabel className="text-xs">Días indisp. U2</FormLabel>
                          <FormControl>
                            <Input type="number" min="0" className="text-sm h-8" {...field} />
                          </FormControl>
                          <FormMessage />
                        </FormItem>
                      )}
                    />
                  </div>

                  <div className="rounded-md border border-border/50 bg-muted/30 p-3 space-y-3">
                    <div className="space-y-1">
                      <FormLabel className="text-xs flex items-center gap-1">
                        <FileSpreadsheet className="w-3 h-3 text-green-600" />
                        Archivo de producción (.xlsx) <span className="text-destructive">*</span>
                      </FormLabel>
                      <Input
                        data-testid="input-prod-file"
                        type="file"
                        accept=".xlsx,.xls"
                        className="text-xs h-8 cursor-pointer"
                        onChange={(e) => { const f = e.target.files?.[0]; if (f) setProdEntry(f); }}
                      />
                      {fileNameProd && (
                        <p className="text-xs text-green-600 truncate flex items-center gap-1">
                          <CheckCircle2 className="w-3 h-3 shrink-0" /> {fileNameProd}
                        </p>
                      )}
                    </div>
                    <div className="space-y-1">
                      <FormLabel className="text-xs flex items-center gap-1">
                        <Activity className="w-3 h-3 text-blue-600" />
                        Archivo de aforo de tanques (.xlsx)
                      </FormLabel>
                      <Input
                        data-testid="input-aforo-file"
                        type="file"
                        accept=".xlsx,.xls"
                        className="text-xs h-8 cursor-pointer"
                        onChange={(e) => { const f = e.target.files?.[0]; if (f) setAforoEntry(f); }}
                      />
                      {fileNameAforo && (
                        <p className="text-xs text-blue-600 truncate flex items-center gap-1">
                          <CheckCircle2 className="w-3 h-3 shrink-0" /> {fileNameAforo}
                        </p>
                      )}
                    </div>
                    {(fileNameProd || fileNameAforo) && (
                      <p className="text-xs text-muted-foreground border-t border-border/40 pt-2 mt-1">
                        Los archivos cargados aquí también están disponibles en Métricas.
                      </p>
                    )}
                  </div>

                  <FormField
                    control={form.control}
                    name="observations"
                    render={({ field }) => (
                      <FormItem>
                        <FormLabel className="text-xs">Observaciones operativas (informe diario)</FormLabel>
                        <FormControl>
                          <Textarea
                            placeholder="Ingresa las novedades operativas del día..."
                            className="resize-none text-xs min-h-[70px]"
                            {...field}
                          />
                        </FormControl>
                        <FormMessage />
                      </FormItem>
                    )}
                  />
                  <p className="text-xs text-muted-foreground">
                    * Los valores negativos del archivo de producción se ignoran (se consideran como 0).
                  </p>
                </form>
              </Form>
            </CardContent>
          </Card>

          <Card className="border-border/60 shadow-md">
            <CardHeader className="pb-2">
              <CardTitle className="text-sm flex items-center gap-2">
                <Factory className="w-4 h-4 text-muted-foreground" />
                Generar informe
              </CardTitle>
            </CardHeader>
            <CardContent className="space-y-2">
              {[
                { type: "diario" as const, label: "Informe diario", desc: "Datos del día seleccionado" },
                { type: "mensual" as const, label: "Informe mensual", desc: "Acumulados del mes" },
                { type: "facturacion" as const, label: "Informe facturación", desc: "Costos y energía facturable" },
              ].map(({ type, label, desc }) => (
                <Button
                  key={type}
                  data-testid={`button-generate-${type}`}
                  variant="outline"
                  className="w-full justify-start h-auto py-2 px-3 group"
                  onClick={() => handleGenerate(type)}
                  disabled={isGenerating}
                >
                  <div className="w-7 h-7 rounded bg-primary/10 flex items-center justify-center mr-3 shrink-0 group-hover:bg-primary/20 transition-colors">
                    <FileText className="w-3.5 h-3.5 text-primary" />
                  </div>
                  <div className="text-left">
                    <div className="text-xs font-semibold">
                      {isGenerating && currentReportType === type ? "Procesando..." : label}
                    </div>
                    <div className="text-xs text-muted-foreground">{desc}</div>
                  </div>
                </Button>
              ))}
            </CardContent>
          </Card>
        </div>

        {/* Columna derecha – previsualización */}
        <div className="flex-1 flex flex-col min-h-[600px] xl:min-h-0 rounded-lg border border-border/60 overflow-hidden shadow-inner">
          {/* Barra superior */}
          <div className="h-12 bg-card border-b border-border/60 flex items-center justify-between px-4 shrink-0">
            <h3 className="text-sm font-semibold flex items-center gap-2">
              <span className={`inline-flex h-2 w-2 rounded-full ${generatedHtml ? "bg-green-500" : "bg-muted-foreground/40"}`} />
              Previsualización del informe
            </h3>
            <div className="flex items-center gap-2">
              <Button
                data-testid="button-export-pdf"
                variant="outline"
                size="sm"
                onClick={handleExportPDF}
                disabled={!generatedHtml}
              >
                <FileDown className="w-3.5 h-3.5 mr-1" />
                PDF
              </Button>
              <Button
                data-testid="button-save-report"
                size="sm"
                onClick={handleSave}
                disabled={!generatedHtml || createReport.isPending}
              >
                {createReport.isPending ? "Guardando..." : (
                  <>
                    <Save className="w-3.5 h-3.5 mr-1" />
                    Guardar
                  </>
                )}
              </Button>
            </div>
          </div>

          {/* Área de contenido */}
          <div
            className="flex-1 overflow-auto bg-slate-50 dark:bg-slate-900/50 p-6"
            ref={previewRef}
          >
            {isGenerating ? (
              <div className="h-full flex flex-col items-center justify-center text-muted-foreground">
                <div className="w-12 h-12 border-4 border-primary/20 border-t-primary rounded-full animate-spin mb-3" />
                <p className="text-sm font-medium">Procesando datos del Excel...</p>
              </div>
            ) : generatedHtml ? (
              <div
                className="report-wrapper bg-white shadow-sm rounded-md p-6 max-w-5xl mx-auto report-content"
                dangerouslySetInnerHTML={{ __html: generatedHtml }}
              />
            ) : (
              <div className="h-full flex flex-col items-center justify-center text-muted-foreground/60 border-2 border-dashed border-border/60 rounded-lg mx-auto max-w-md">
                <AlertCircle className="w-12 h-12 mb-3 opacity-40" />
                <p className="text-sm font-medium text-foreground/50">Área de visualización vacía</p>
                <p className="text-xs mt-1 text-center max-w-xs">
                  Carga el archivo de producción, configura los parámetros y selecciona un tipo de informe.
                </p>
              </div>
            )}
          </div>
        </div>
      </div>
    </Layout>
  );
}
