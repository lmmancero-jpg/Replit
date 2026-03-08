import { useState, useRef } from "react";
import { z } from "zod";
import { useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { format } from "date-fns";
import { es } from "date-fns/locale";
import { FileDown, Save, Calendar, FileSpreadsheet, Activity, Factory, FileText, AlertCircle, CheckCircle2 } from "lucide-react";
import html2pdf from "html2pdf.js";
import { Settings } from "lucide-react";

import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle, CardFooter } from "@/components/ui/card";
import { Form, FormControl, FormDescription, FormField, FormItem, FormLabel, FormMessage } from "@/components/ui/form";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { useToast } from "@/hooks/use-toast";
import { Layout } from "@/components/layout";
import { useCreateReport } from "@/hooks/use-reports";

const generatorSchema = z.object({
  reportDate: z.string().min(1, "La fecha es requerida"),
  u1Downtime: z.coerce.number().min(0).max(31).default(0),
  u2Downtime: z.coerce.number().min(0).max(31).default(0),
  observations: z.string().optional(),
});

type GeneratorValues = z.infer<typeof generatorSchema>;

export default function Generator() {
  const { toast } = useToast();
  const createReport = useCreateReport();
  const [isGenerating, setIsGenerating] = useState(false);
  const [generatedHtml, setGeneratedHtml] = useState<string | null>(null);
  const [currentReportType, setCurrentReportType] = useState<string>("");
  const previewRef = useRef<HTMLDivElement>(null);

  const form = useForm<GeneratorValues>({
    resolver: zodResolver(generatorSchema),
    defaultValues: {
      reportDate: format(new Date(), 'yyyy-MM-dd'),
      u1Downtime: 0,
      u2Downtime: 0,
      observations: "",
    },
  });

  const handleGenerate = async (type: string, data: GeneratorValues) => {
    setIsGenerating(true);
    setCurrentReportType(type);
    
    // Simulate generation delay
    await new Promise(r => setTimeout(r, 1500));
    
    const formattedDate = format(new Date(data.reportDate), "dd 'de' MMMM, yyyy", { locale: es });
    
    // Mock robust HTML report generation
    const html = `
      <div class="report-document">
        <div style="display: flex; justify-content: space-between; align-items: flex-end; border-bottom: 2px solid #1e293b; padding-bottom: 1rem; margin-bottom: 2rem;">
          <div>
            <h1 style="margin:0; border:none; padding:0;">INFORME DE ${type.toUpperCase()}</h1>
            <p style="color: #64748b; font-size: 1.125rem; margin: 0.5rem 0 0 0; font-family: var(--font-display);">Central Eléctrica NEXUS</p>
          </div>
          <div style="text-align: right; color: #475569;">
            <strong>Fecha:</strong> ${formattedDate}
          </div>
        </div>
        
        <h2>Resumen Operativo</h2>
        <table>
          <thead>
            <tr>
              <th>Parámetro</th>
              <th>Unidad 1</th>
              <th>Unidad 2</th>
              <th>Total Planta</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>Días Indisponibles</td>
              <td>${data.u1Downtime} días</td>
              <td>${data.u2Downtime} días</td>
              <td>${data.u1Downtime + data.u2Downtime} días</td>
            </tr>
            <tr>
              <td>Producción Bruta (MWh)</td>
              <td>${(Math.random() * 5000 + 1000).toFixed(2)}</td>
              <td>${(Math.random() * 5000 + 1000).toFixed(2)}</td>
              <td><strong>${(Math.random() * 10000 + 2000).toFixed(2)}</strong></td>
            </tr>
            <tr>
              <td>Factor de Planta (%)</td>
              <td>${(Math.random() * 20 + 75).toFixed(1)}%</td>
              <td>${(Math.random() * 20 + 75).toFixed(1)}%</td>
              <td><strong>${(Math.random() * 20 + 75).toFixed(1)}%</strong></td>
            </tr>
          </tbody>
        </table>

        <h2>Datos de Aforo y Consumo</h2>
        <table>
          <thead>
            <tr>
              <th>Recurso</th>
              <th>Volumen Declarado</th>
              <th>Volumen Verificado</th>
              <th>Desviación</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>Combustible Principal (Ton)</td>
              <td>${(Math.random() * 1000 + 500).toFixed(2)}</td>
              <td>${(Math.random() * 1000 + 500).toFixed(2)}</td>
              <td style="color: #0284c7;">-1.2%</td>
            </tr>
            <tr>
              <td>Agua de Refrigeración (m³)</td>
              <td>${(Math.random() * 50000 + 20000).toFixed(2)}</td>
              <td>${(Math.random() * 50000 + 20000).toFixed(2)}</td>
              <td style="color: #16a34a;">+0.5%</td>
            </tr>
          </tbody>
        </table>

        <h2>Observaciones del Turno</h2>
        <div style="background-color: #f8fafc; padding: 1rem; border-left: 4px solid #0284c7; margin-bottom: 2rem;">
          <p style="margin: 0; white-space: pre-wrap;">${data.observations || "Operación normal sin incidencias destacables."}</p>
        </div>

        <div style="margin-top: 4rem; display: flex; justify-content: space-around;">
          <div style="text-align: center; border-top: 1px solid #cbd5e1; padding-top: 1rem; width: 40%;">
            <strong>Jefe de Turno</strong><br/>
            Firma
          </div>
          <div style="text-align: center; border-top: 1px solid #cbd5e1; padding-top: 1rem; width: 40%;">
            <strong>Gerente de Planta</strong><br/>
            Firma
          </div>
        </div>
      </div>
    `;
    
    setGeneratedHtml(html);
    setIsGenerating(false);
    toast({
      title: "Informe generado",
      description: "Revisa la previsualización antes de guardar o exportar.",
    });
  };

  const handleSave = () => {
    if (!generatedHtml) return;
    
    createReport.mutate({
      title: `Informe ${currentReportType} - ${form.getValues().reportDate}`,
      reportType: currentReportType,
      date: form.getValues().reportDate,
      content: generatedHtml,
    }, {
      onSuccess: () => {
        toast({
          title: "Guardado exitoso",
          description: "El informe ha sido almacenado en el historial.",
        });
      }
    });
  };

  const handleExportPDF = () => {
    if (!previewRef.current) return;
    
    const element = previewRef.current.querySelector('.report-document');
    if (!element) return;

    const opt = {
      margin:       10,
      filename:     `Informe_${currentReportType}_${form.getValues().reportDate}.pdf`,
      image:        { type: 'jpeg', quality: 0.98 },
      html2canvas:  { scale: 2, useCORS: true },
      jsPDF:        { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };

    html2pdf().set(opt).from(element).save();
    
    toast({
      title: "PDF en descarga",
      description: "El archivo se está guardando en tu equipo.",
    });
  };

  return (
    <Layout>
      <div className="flex flex-col xl:flex-row gap-6 lg:gap-8 h-full">
        
        {/* Left Column: Form Setup */}
        <div className="w-full xl:w-[400px] 2xl:w-[450px] shrink-0 flex flex-col gap-6">
          <Card className="border-border/60 shadow-lg shadow-black/5 bg-card/80 backdrop-blur-sm overflow-hidden">
            <div className="h-1.5 w-full bg-gradient-to-r from-primary via-primary/80 to-accent"></div>
            <CardHeader className="pb-4">
              <CardTitle className="font-display text-2xl flex items-center gap-2">
                <Settings className="w-6 h-6 text-primary" />
                Configuración
              </CardTitle>
              <CardDescription>
                Ingresa los parámetros y archivos base para calcular la producción.
              </CardDescription>
            </CardHeader>
            <CardContent>
              <Form {...form}>
                <form className="space-y-5">
                  <FormField
                    control={form.control}
                    name="reportDate"
                    render={({ field }) => (
                      <FormItem>
                        <FormLabel className="flex items-center gap-2"><Calendar className="w-4 h-4"/> Fecha de Análisis</FormLabel>
                        <FormControl>
                          <Input type="date" className="bg-background/50 focus:bg-background transition-colors" {...field} />
                        </FormControl>
                        <FormMessage />
                      </FormItem>
                    )}
                  />

                  <div className="grid grid-cols-2 gap-4">
                    <FormField
                      control={form.control}
                      name="u1Downtime"
                      render={({ field }) => (
                        <FormItem>
                          <FormLabel className="text-xs font-semibold">Días Indisp. U1</FormLabel>
                          <FormControl>
                            <Input type="number" min="0" className="bg-background/50" {...field} />
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
                          <FormLabel className="text-xs font-semibold">Días Indisp. U2</FormLabel>
                          <FormControl>
                            <Input type="number" min="0" className="bg-background/50" {...field} />
                          </FormControl>
                          <FormMessage />
                        </FormItem>
                      )}
                    />
                  </div>

                  <div className="p-4 rounded-xl bg-secondary/50 border border-border/50 space-y-4">
                    <div className="space-y-2">
                      <FormLabel className="flex items-center gap-2"><FileSpreadsheet className="w-4 h-4 text-green-600"/> Archivo Producción</FormLabel>
                      <Input type="file" accept=".xlsx,.xls" className="bg-card cursor-pointer file:bg-primary file:text-primary-foreground file:border-0 file:rounded-md file:px-4 file:py-1 file:mr-4 file:font-medium file:cursor-pointer hover:file:bg-primary/90" />
                    </div>
                    <div className="space-y-2">
                      <FormLabel className="flex items-center gap-2"><Activity className="w-4 h-4 text-blue-600"/> Archivo Aforo</FormLabel>
                      <Input type="file" accept=".xlsx,.xls" className="bg-card cursor-pointer file:bg-primary file:text-primary-foreground file:border-0 file:rounded-md file:px-4 file:py-1 file:mr-4 file:font-medium file:cursor-pointer hover:file:bg-primary/90" />
                    </div>
                  </div>

                  <FormField
                    control={form.control}
                    name="observations"
                    render={({ field }) => (
                      <FormItem>
                        <FormLabel>Observaciones Operativas</FormLabel>
                        <FormControl>
                          <Textarea 
                            placeholder="Incidencias, mantenimientos..." 
                            className="resize-none h-24 bg-background/50 focus:bg-background" 
                            {...field} 
                          />
                        </FormControl>
                        <FormMessage />
                      </FormItem>
                    )}
                  />
                </form>
              </Form>
            </CardContent>
          </Card>

          <Card className="border-border/60 shadow-lg shadow-black/5 flex-1">
            <CardHeader className="pb-4">
              <CardTitle className="font-display text-lg flex items-center gap-2">
                <Factory className="w-5 h-5 text-muted-foreground" />
                Acciones de Generación
              </CardTitle>
            </CardHeader>
            <CardContent className="grid grid-cols-1 gap-3">
              <Button 
                variant="outline" 
                className="w-full justify-start h-12 font-medium bg-card hover:bg-secondary/80 border-border hover:border-primary/50 transition-all group"
                onClick={() => handleGenerate('diario', form.getValues())}
                disabled={isGenerating}
              >
                <div className="w-8 h-8 rounded bg-primary/10 flex items-center justify-center mr-3 group-hover:bg-primary/20 transition-colors">
                  <FileText className="w-4 h-4 text-primary" />
                </div>
                {isGenerating && currentReportType === 'diario' ? "Procesando..." : "Generar Informe Diario"}
              </Button>
              <Button 
                variant="outline" 
                className="w-full justify-start h-12 font-medium bg-card hover:bg-secondary/80 border-border hover:border-primary/50 transition-all group"
                onClick={() => handleGenerate('mensual', form.getValues())}
                disabled={isGenerating}
              >
                <div className="w-8 h-8 rounded bg-primary/10 flex items-center justify-center mr-3 group-hover:bg-primary/20 transition-colors">
                  <FileText className="w-4 h-4 text-primary" />
                </div>
                {isGenerating && currentReportType === 'mensual' ? "Procesando..." : "Generar Informe Mensual"}
              </Button>
              <Button 
                variant="outline" 
                className="w-full justify-start h-12 font-medium bg-card hover:bg-secondary/80 border-border hover:border-primary/50 transition-all group"
                onClick={() => handleGenerate('facturacion', form.getValues())}
                disabled={isGenerating}
              >
                <div className="w-8 h-8 rounded bg-accent/10 flex items-center justify-center mr-3 group-hover:bg-accent/20 transition-colors">
                  <FileText className="w-4 h-4 text-accent" />
                </div>
                {isGenerating && currentReportType === 'facturacion' ? "Procesando..." : "Generar Facturación"}
              </Button>
            </CardContent>
          </Card>
        </div>

        {/* Right Column: Preview */}
        <div className="flex-1 flex flex-col min-h-[600px] xl:min-h-0 bg-secondary/30 rounded-2xl border border-border/60 overflow-hidden relative shadow-inner">
          <div className="h-14 bg-card border-b border-border/60 flex items-center justify-between px-6 shrink-0 z-10 shadow-sm">
            <h3 className="font-display font-semibold text-foreground flex items-center gap-2">
              <span className="relative flex h-2 w-2">
                <span className={`absolute inline-flex h-full w-full rounded-full opacity-75 ${generatedHtml ? 'bg-green-400 animate-ping' : 'bg-slate-400'}`}></span>
                <span className={`relative inline-flex rounded-full h-2 w-2 ${generatedHtml ? 'bg-green-500' : 'bg-slate-500'}`}></span>
              </span>
              Previsualización
            </h3>
            
            <div className="flex items-center gap-2">
              <Button 
                variant="outline" 
                size="sm" 
                onClick={handleExportPDF} 
                disabled={!generatedHtml}
                className="h-9 font-medium"
              >
                <FileDown className="w-4 h-4 mr-2" />
                PDF
              </Button>
              <Button 
                size="sm" 
                onClick={handleSave} 
                disabled={!generatedHtml || createReport.isPending}
                className="h-9 font-medium shadow-md shadow-primary/20"
              >
                {createReport.isPending ? (
                  "Guardando..."
                ) : (
                  <>
                    <Save className="w-4 h-4 mr-2" />
                    Guardar Base de Datos
                  </>
                )}
              </Button>
            </div>
          </div>

          <div className="flex-1 overflow-auto p-4 md:p-8 bg-slate-200/50 dark:bg-slate-900/50" ref={previewRef}>
            {isGenerating ? (
              <div className="h-full w-full flex flex-col items-center justify-center text-muted-foreground animate-pulse">
                <div className="w-16 h-16 border-4 border-primary/20 border-t-primary rounded-full animate-spin mb-4"></div>
                <p className="font-display font-medium text-lg">Procesando motores de cálculo...</p>
              </div>
            ) : generatedHtml ? (
              <div 
                dangerouslySetInnerHTML={{ __html: generatedHtml }} 
                className="transition-opacity duration-500 animate-in fade-in"
              />
            ) : (
              <div className="h-full w-full flex flex-col items-center justify-center text-muted-foreground/60 border-2 border-dashed border-border/60 rounded-xl m-auto max-w-lg bg-card/30">
                <AlertCircle className="w-16 h-16 mb-4 opacity-50" />
                <p className="font-display font-medium text-lg text-foreground/50">Área de visualización vacía</p>
                <p className="text-sm mt-2 text-center max-w-sm">
                  Configura los parámetros en el panel izquierdo y selecciona un tipo de informe para generar la vista previa.
                </p>
              </div>
            )}
          </div>
        </div>
      </div>
    </Layout>
  );
}
