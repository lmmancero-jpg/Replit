import { useState } from "react";
import { format } from "date-fns";
import { es } from "date-fns/locale";
import { Trash2, Eye, FileText, Search, Download } from "lucide-react";
import html2pdf from "html2pdf.js";

import { useReports, useDeleteReport } from "@/hooks/use-reports";
import { Layout } from "@/components/layout";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import {
  AlertDialog,
  AlertDialogAction,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
  AlertDialogTrigger,
} from "@/components/ui/alert-dialog";
import { Badge } from "@/components/ui/badge";

export default function History() {
  const { data: reports, isLoading } = useReports();
  const deleteReport = useDeleteReport();
  const [searchTerm, setSearchTerm] = useState("");
  const [selectedReport, setSelectedReport] = useState<any>(null);

  const filteredReports = reports?.filter(report => 
    report.title.toLowerCase().includes(searchTerm.toLowerCase()) ||
    report.reportType.toLowerCase().includes(searchTerm.toLowerCase())
  ) || [];

  const handleExportPDF = async (report: any) => {
    const tempDiv = document.createElement('div');
    tempDiv.className = 'report-content';
    tempDiv.style.cssText = 'position:absolute;left:-9999px;background:#fff;padding:20px;font-family:system-ui,sans-serif;font-size:13px;color:#1b2134;';
    tempDiv.innerHTML = report.content;
    document.body.appendChild(tempDiv);

    const opt = {
      margin: 8,
      filename: `${report.title.replace(/\s+/g, '_')}.pdf`,
      image: { type: 'jpeg', quality: 0.98 },
      html2canvas: { scale: 2, useCORS: true },
      jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' },
    };

    try {
      await html2pdf().set(opt).from(tempDiv).save();
    } finally {
      document.body.removeChild(tempDiv);
    }
  };

  return (
    <Layout>
      <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-500">
        
        {/* Header Section */}
        <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 bg-card p-6 rounded-2xl border border-border/60 shadow-lg shadow-black/5">
          <div>
            <h1 className="text-3xl font-display font-bold text-foreground">Historial de Informes</h1>
            <p className="text-muted-foreground mt-1">
              Registro inmutable de producción y facturación almacenado en el sistema.
            </p>
          </div>
          
          <div className="relative w-full sm:w-72">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground" />
            <Input 
              placeholder="Buscar por título o tipo..." 
              className="pl-9 bg-background/50 border-border/60 focus:bg-background"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
          </div>
        </div>

        {/* Data Table */}
        <div className="bg-card rounded-2xl border border-border/60 shadow-lg shadow-black/5 overflow-hidden">
          {isLoading ? (
            <div className="p-12 flex flex-col items-center justify-center text-muted-foreground">
              <div className="w-12 h-12 border-4 border-primary/20 border-t-primary rounded-full animate-spin mb-4"></div>
              <p>Cargando registros...</p>
            </div>
          ) : filteredReports.length === 0 ? (
            <div className="p-16 flex flex-col items-center justify-center text-muted-foreground bg-secondary/20">
              <FileText className="w-16 h-16 mb-4 opacity-20" />
              <p className="text-lg font-display font-medium text-foreground/70">No se encontraron informes</p>
              <p className="text-sm mt-1">Ajusta tu búsqueda o genera un nuevo informe en la pestaña principal.</p>
            </div>
          ) : (
            <Table>
              <TableHeader className="bg-secondary/50">
                <TableRow className="hover:bg-transparent">
                  <TableHead className="w-[100px] font-semibold text-foreground/80">ID</TableHead>
                  <TableHead className="font-semibold text-foreground/80">Título</TableHead>
                  <TableHead className="font-semibold text-foreground/80">Tipo</TableHead>
                  <TableHead className="font-semibold text-foreground/80">Fecha de Datos</TableHead>
                  <TableHead className="font-semibold text-foreground/80">Generado</TableHead>
                  <TableHead className="text-right font-semibold text-foreground/80">Acciones</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {filteredReports.map((report) => (
                  <TableRow key={report.id} className="group hover:bg-secondary/30 transition-colors">
                    <TableCell className="font-mono text-xs text-muted-foreground">
                      #{String(report.id).padStart(5, '0')}
                    </TableCell>
                    <TableCell className="font-medium">
                      {report.title}
                    </TableCell>
                    <TableCell>
                      <Badge variant="outline" className={`
                        capitalize font-medium
                        ${report.reportType === 'diario' ? 'bg-blue-500/10 text-blue-700 border-blue-200' : ''}
                        ${report.reportType === 'mensual' ? 'bg-purple-500/10 text-purple-700 border-purple-200' : ''}
                        ${report.reportType === 'facturacion' ? 'bg-green-500/10 text-green-700 border-green-200' : ''}
                      `}>
                        {report.reportType}
                      </Badge>
                    </TableCell>
                    <TableCell className="text-muted-foreground">
                      {report.date}
                    </TableCell>
                    <TableCell className="text-muted-foreground">
                      {report.createdAt ? format(new Date(report.createdAt), "dd MMM yyyy, HH:mm", { locale: es }) : 'N/A'}
                    </TableCell>
                    <TableCell className="text-right">
                      <div className="flex justify-end gap-2 opacity-60 group-hover:opacity-100 transition-opacity">
                        <Button 
                          variant="ghost" 
                          size="icon"
                          className="h-8 w-8 hover:bg-primary/10 hover:text-primary"
                          onClick={() => setSelectedReport(report)}
                          title="Ver Informe"
                        >
                          <Eye className="h-4 w-4" />
                        </Button>
                        <Button 
                          variant="ghost" 
                          size="icon"
                          className="h-8 w-8 hover:bg-green-500/10 hover:text-green-600"
                          onClick={() => handleExportPDF(report)}
                          title="Descargar PDF"
                        >
                          <Download className="h-4 w-4" />
                        </Button>
                        <AlertDialog>
                          <AlertDialogTrigger asChild>
                            <Button variant="ghost" size="icon" className="h-8 w-8 hover:bg-destructive/10 hover:text-destructive" title="Eliminar">
                              <Trash2 className="h-4 w-4" />
                            </Button>
                          </AlertDialogTrigger>
                          <AlertDialogContent>
                            <AlertDialogHeader>
                              <AlertDialogTitle>¿Eliminar este informe?</AlertDialogTitle>
                              <AlertDialogDescription>
                                Esta acción no se puede deshacer. Se eliminará permanentemente el registro #{report.id} de la base de datos.
                              </AlertDialogDescription>
                            </AlertDialogHeader>
                            <AlertDialogFooter>
                              <AlertDialogCancel>Cancelar</AlertDialogCancel>
                              <AlertDialogAction 
                                onClick={() => deleteReport.mutate(report.id)}
                                className="bg-destructive text-destructive-foreground hover:bg-destructive/90"
                              >
                                Eliminar Permanentemente
                              </AlertDialogAction>
                            </AlertDialogFooter>
                          </AlertDialogContent>
                        </AlertDialog>
                      </div>
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          )}
        </div>
      </div>

      {/* Viewer Dialog */}
      <Dialog open={!!selectedReport} onOpenChange={(open) => !open && setSelectedReport(null)}>
        <DialogContent className="max-w-4xl max-h-[90vh] flex flex-col overflow-hidden bg-slate-200 dark:bg-slate-900 border-none shadow-2xl p-0">
          <DialogHeader className="p-4 bg-card border-b shrink-0 flex flex-row items-center justify-between">
            <DialogTitle className="font-display text-xl text-foreground">
              {selectedReport?.title}
            </DialogTitle>
            <Button size="sm" onClick={() => handleExportPDF(selectedReport)} className="font-medium mr-6">
              <Download className="w-4 h-4 mr-2"/>
              Exportar a PDF
            </Button>
          </DialogHeader>
          <div className="flex-1 overflow-y-auto p-8">
            {selectedReport && (
              <div className="report-wrapper bg-white shadow-sm rounded-md p-6 max-w-5xl mx-auto report-content">
                <div dangerouslySetInnerHTML={{ __html: selectedReport.content }} />
              </div>
            )}
          </div>
        </DialogContent>
      </Dialog>
    </Layout>
  );
}
