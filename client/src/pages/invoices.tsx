import { useState, useCallback } from "react";
import { useForm, useFieldArray, useWatch } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { z } from "zod";
import { format } from "date-fns";
import {
  PlusCircle, Pencil, Trash2, FileDown, FileSpreadsheet,
  AlertCircle, CheckCircle2, ReceiptText, Plus, X, Truck,
} from "lucide-react";
import { Layout } from "@/components/layout";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { Form, FormControl, FormField, FormItem, FormLabel, FormMessage } from "@/components/ui/form";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Switch } from "@/components/ui/switch";
import { useToast } from "@/hooks/use-toast";
import { useFileStore } from "@/lib/fileStore";
import { INVOICE_CATEGORIES, INVOICE_CATEGORY_LABELS, FUEL_TYPE_LABELS, type InvoiceCategory, type FuelType, type Invoice } from "@shared/schema";
import { useInvoices, useInvoiceSummary, useCreateInvoice, useUpdateInvoice, useDeleteInvoice } from "@/hooks/use-invoices";
import { exportInvoicesExcel } from "@/lib/invoiceExcel";
import { getMonthlyProductionSummary } from "@/lib/billingEngine";
import { openPrintWindow } from "@/lib/printPDF";
import { fmt } from "@/lib/reportEngine";

// ── Constantes ────────────────────────────────────────────────────────────────

const IVA_RATE = 0.15;
const CATEGORIES_WITH_TRANSPORT: InvoiceCategory[] = ["combustible_transporte", "agua_insumos"];

function canBeTransporte(category: string): boolean {
  return CATEGORIES_WITH_TRANSPORT.includes(category as InvoiceCategory);
}

function calcIva(subtotal: number, isTransporte: boolean): number {
  if (isTransporte) return 0;
  return Math.round(subtotal * IVA_RATE * 100) / 100;
}

function calcTotal(subtotal: number, iva: number): number {
  return Math.round((subtotal + iva) * 100) / 100;
}

// ── Schemas ───────────────────────────────────────────────────────────────────

const lineItemSchema = z.object({
  description: z.string().default(""),
  subtotal: z.coerce.number().min(0, "Debe ser ≥ 0"),
  isTransporte: z.boolean().default(false),
  fuelType: z.enum(["hfo", "diesel_2"]).nullable().optional(),
});

const invoiceFormSchema = z.object({
  issueDate: z.string().min(1, "La fecha de emisión es requerida"),
  supplier: z.string().min(1, "El proveedor es requerido"),
  invoiceNumber: z.string().min(1, "El número de factura es requerido"),
  category: z.enum(INVOICE_CATEGORIES, { required_error: "Selecciona un rubro" }),
  observations: z.string().default(""),
  items: z.array(lineItemSchema).min(1, "Al menos un ítem es requerido"),
});

type InvoiceFormValues = z.infer<typeof invoiceFormSchema>;

interface LineItemComputed {
  description: string;
  subtotal: number;
  isTransporte: boolean;
  fuelType?: FuelType | null;
  iva: number;
  total: number;
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function fmtMoney(v: string | number | null | undefined): string {
  return parseFloat(String(v ?? "0")).toLocaleString("es-EC", {
    minimumFractionDigits: 2, maximumFractionDigits: 2,
  });
}

function parseLineItems(inv: Invoice): LineItemComputed[] {
  try {
    if (inv.lineItems && inv.lineItems !== "[]") {
      const parsed = JSON.parse(inv.lineItems);
      if (Array.isArray(parsed) && parsed.length > 0) {
        return parsed.map((i: Record<string, unknown>) => ({
          description: String(i.description ?? ""),
          subtotal: parseFloat(String(i.subtotal ?? "0")),
          isTransporte: !!i.isTransporte,
          fuelType: (i.fuelType as FuelType | null | undefined) ?? null,
          iva: parseFloat(String(i.iva ?? "0")),
          total: parseFloat(String(i.total ?? "0")),
        }));
      }
    }
  } catch {}
  return [{
    description: inv.description ?? "",
    subtotal: parseFloat(inv.subtotal ?? "0"),
    isTransporte: false,
    fuelType: null,
    iva: parseFloat(inv.iva ?? "0"),
    total: parseFloat(inv.total ?? "0"),
  }];
}

function buildInvoiceHtml(invoiceList: Invoice[], period: string, summary: Record<string, number>, prodKwh?: number): string {
  const rows = invoiceList.map((inv) => {
    const items = parseLineItems(inv);
    const hasMulti = items.length > 1;
    if (hasMulti) {
      const itemRows = items.map(item => `
<tr style="background:#fafcff">
  <td class="label" style="padding-left:28px;font-size:14px;color:#6b7280">↳ ${item.isTransporte ? "Transporte" : ""}${item.description ? (item.isTransporte ? " – " : "") + item.description : (!item.isTransporte ? INVOICE_CATEGORY_LABELS[inv.category as InvoiceCategory] ?? inv.category : "")}</td>
  <td></td><td></td><td></td>
  <td class="num">$ ${fmtMoney(item.subtotal)}</td>
  <td class="num">${item.isTransporte ? '<span style="color:#6b7280;font-size:12px">Exento</span>' : "$ " + fmtMoney(item.iva)}</td>
  <td class="num">$ ${fmtMoney(item.total)}</td>
</tr>`).join("");
      return `<tr style="border-bottom:none">
  <td class="label" style="font-weight:600">${inv.issueDate}</td>
  <td class="label">${inv.supplier}</td>
  <td class="label" style="font-family:monospace">${inv.invoiceNumber}</td>
  <td class="label"><span style="padding:2px 8px;background:#e8f0fd;border-radius:4px;font-size:12px">${INVOICE_CATEGORY_LABELS[inv.category as InvoiceCategory] ?? inv.category}</span></td>
  <td class="num">$ ${fmtMoney(inv.subtotal)}</td>
  <td class="num">$ ${fmtMoney(inv.iva)}</td>
  <td class="num" style="font-weight:700">$ ${fmtMoney(inv.total)}</td>
</tr>${itemRows}`;
    }
    return `<tr>
  <td class="label">${inv.issueDate}</td>
  <td class="label">${inv.supplier}</td>
  <td class="label" style="font-family:monospace">${inv.invoiceNumber}</td>
  <td class="label"><span style="padding:2px 8px;background:#e8f0fd;border-radius:4px;font-size:12px">${INVOICE_CATEGORY_LABELS[inv.category as InvoiceCategory] ?? inv.category}</span></td>
  <td class="num">$ ${fmtMoney(inv.subtotal)}</td>
  <td class="num">${parseFloat(inv.iva ?? "0") === 0 ? '<span style="color:#6b7280;font-size:12px">Exento</span>' : "$ " + fmtMoney(inv.iva)}</td>
  <td class="num" style="font-weight:700">$ ${fmtMoney(inv.total)}</td>
</tr>`;
  }).join("");

  const totalPeriod = Object.values(summary).reduce((a, b) => a + b, 0);
  const summaryRows = INVOICE_CATEGORIES.map((cat) => {
    const cv = prodKwh && prodKwh > 0 ? (summary[cat] ?? 0) / prodKwh : null;
    return `<tr>
  <td class="label">${INVOICE_CATEGORY_LABELS[cat]}</td>
  <td class="num">$ ${fmtMoney(summary[cat] ?? 0)}</td>
  <td class="num">${cv !== null ? fmt(cv, 4) : "—"}</td>
</tr>`;
  }).join("");

  return `<div class="rpt-header">
  <div class="rpt-header-body">
    <div class="rpt-header-left">
      <div class="rpt-logo-circle">
        <svg viewBox="0 0 32 32" fill="none" xmlns="http://www.w3.org/2000/svg" width="28" height="28">
          <circle cx="16" cy="16" r="14" fill="rgba(255,255,255,0.15)" stroke="rgba(255,255,255,0.5)" stroke-width="1.5"/>
          <path d="M16 7v4M16 21v4M7 16h4M21 16h4" stroke="rgba(255,255,255,0.9)" stroke-width="1.5" stroke-linecap="round"/>
          <circle cx="16" cy="16" r="3.5" fill="rgba(255,255,255,0.9)"/>
        </svg>
      </div>
      <div>
        <div class="rpt-empresa">Central El Morro &mdash; Morro Energy S.A.</div>
        <div class="rpt-tipo">Registro de Facturas del Período</div>
      </div>
    </div>
    <div class="rpt-header-right">
      <div class="rpt-subtitulo-label">Período</div>
      <div class="rpt-subtitulo">${period}</div>
    </div>
  </div>
  <div class="rpt-header-stripe"></div>
</div>
<div class="rpt-section-title"><span class="rpt-section-num">1</span><span class="rpt-section-label">Detalle de Facturas</span></div>
<table class="data-table">
<thead><tr><th>Fecha</th><th>Proveedor</th><th>N° Factura</th><th>Rubro</th><th>Subtotal</th><th>IVA 15%</th><th>Total</th></tr></thead>
<tbody>${rows}
<tr class="rpt-row-grand"><td colspan="6" class="label"><strong>TOTAL PERÍODO</strong></td><td class="num"><strong>$ ${fmtMoney(totalPeriod)}</strong></td></tr>
</tbody></table>
<div class="rpt-section-title"><span class="rpt-section-num">2</span><span class="rpt-section-label">Resumen por Rubro</span></div>
<table class="data-table">
<thead><tr><th>Rubro</th><th>Total facturado [USD]</th><th>CV real [USD/kWh]</th></tr></thead>
<tbody>${summaryRows}
<tr class="rpt-row-grand"><td class="label"><strong>TOTAL</strong></td><td class="num"><strong>$ ${fmtMoney(totalPeriod)}</strong></td><td class="num">${prodKwh && prodKwh > 0 ? fmt(totalPeriod / prodKwh, 4) : "—"}</td></tr>
</tbody></table>
${prodKwh && prodKwh > 0 ? `<p class="rpt-muted">* Producción del período: ${fmt(prodKwh, 0)} kWh.</p>` : `<div class="rpt-notice rpt-notice-warn">⚠ Sin producción cargada — CV real no disponible.</div>`}`;
}

// ── LineItem Row Component ─────────────────────────────────────────────────────

function LineItemRow({
  index,
  category,
  onRemove,
  canRemove,
  control,
}: {
  index: number;
  category: string;
  onRemove: () => void;
  canRemove: boolean;
  control: ReturnType<typeof useForm<InvoiceFormValues>>["control"];
}) {
  const subtotalVal = useWatch({ control, name: `items.${index}.subtotal` as const });
  const isTransporteVal = useWatch({ control, name: `items.${index}.isTransporte` as const });
  const showTransporte = canBeTransporte(category);

  const sub = Number(subtotalVal) || 0;
  const iva = calcIva(sub, !!isTransporteVal);
  const total = calcTotal(sub, iva);

  return (
    <div className={`rounded-md border p-3 space-y-2 ${isTransporteVal ? "bg-amber-50/50 border-amber-200/70" : "bg-muted/20 border-border/50"}`}>
      <div className="flex items-center justify-between gap-2">
        <span className="text-xs font-semibold text-muted-foreground">Ítem {index + 1}</span>
        <div className="flex items-center gap-3">
          {showTransporte && (
            <FormField
              control={control}
              name={`items.${index}.isTransporte`}
              render={({ field }) => (
                <FormItem className="flex items-center gap-2 m-0 space-y-0">
                  <FormControl>
                    <Switch
                      data-testid={`switch-transport-${index}`}
                      checked={!!field.value}
                      onCheckedChange={field.onChange}
                      className="scale-90"
                    />
                  </FormControl>
                  <FormLabel className="text-xs flex items-center gap-1 cursor-pointer m-0">
                    <Truck className="w-3 h-3 text-amber-600" />
                    <span className={isTransporteVal ? "text-amber-700 font-semibold" : "text-muted-foreground"}>
                      Transporte (sin IVA)
                    </span>
                  </FormLabel>
                </FormItem>
              )}
            />
          )}
          {canRemove && (
            <Button
              data-testid={`button-remove-item-${index}`}
              type="button" variant="ghost" size="icon"
              className="h-6 w-6 text-muted-foreground hover:text-destructive"
              onClick={onRemove}
            >
              <X className="w-3 h-3" />
            </Button>
          )}
        </div>
      </div>

      {/* Tipo de combustible — solo para combustible_transporte y no transporte */}
      {category === "combustible_transporte" && !isTransporteVal && (
        <FormField
          control={control}
          name={`items.${index}.fuelType`}
          render={({ field }) => (
            <FormItem className="space-y-1">
              <FormLabel className="text-xs font-semibold text-primary">Tipo de combustible</FormLabel>
              <Select
                onValueChange={field.onChange}
                value={field.value ?? "hfo"}
              >
                <FormControl>
                  <SelectTrigger data-testid={`select-fuel-type-${index}`} className="text-sm h-8 border-primary/30">
                    <SelectValue placeholder="Selecciona tipo" />
                  </SelectTrigger>
                </FormControl>
                <SelectContent>
                  <SelectItem value="hfo">⛽ HFO (Bunker)</SelectItem>
                  <SelectItem value="diesel_2">🔵 Diesel 2</SelectItem>
                </SelectContent>
              </Select>
              <FormMessage />
            </FormItem>
          )}
        />
      )}

      <FormField
        control={control}
        name={`items.${index}.description`}
        render={({ field }) => (
          <FormItem className="space-y-1">
            <FormLabel className="text-xs">Descripción</FormLabel>
            <FormControl>
              <Input
                data-testid={`input-item-description-${index}`}
                className="text-sm h-8"
                placeholder={isTransporteVal ? "Descripción del servicio de transporte" : "Descripción del bien o servicio"}
                {...field}
              />
            </FormControl>
            <FormMessage />
          </FormItem>
        )}
      />

      <div className="grid grid-cols-3 gap-2">
        <FormField
          control={control}
          name={`items.${index}.subtotal`}
          render={({ field }) => (
            <FormItem className="space-y-1">
              <FormLabel className="text-xs">Subtotal USD</FormLabel>
              <FormControl>
                <Input
                  data-testid={`input-item-subtotal-${index}`}
                  type="number" step="0.01" min="0"
                  className="text-sm h-8 font-mono"
                  {...field}
                />
              </FormControl>
              <FormMessage />
            </FormItem>
          )}
        />
        <div className="space-y-1">
          <label className="text-xs font-medium leading-none">
            IVA 15% {isTransporteVal && <span className="text-amber-600">(exento)</span>}
          </label>
          <div className={`h-8 px-3 rounded-md border text-sm font-mono flex items-center ${isTransporteVal ? "bg-amber-50 border-amber-200 text-amber-700" : "bg-muted/40 border-border/60 text-muted-foreground"}`}>
            {isTransporteVal ? "$ 0.00" : `$ ${fmtMoney(iva)}`}
          </div>
        </div>
        <div className="space-y-1">
          <label className="text-xs font-medium leading-none">Total USD</label>
          <div className="h-8 px-3 rounded-md border bg-primary/5 border-primary/20 text-sm font-mono font-semibold text-primary flex items-center">
            $ {fmtMoney(total)}
          </div>
        </div>
      </div>
    </div>
  );
}

// ── Totals Summary (computed from watched items) ──────────────────────────────

function InvoiceTotals({
  control,
  category,
}: {
  control: ReturnType<typeof useForm<InvoiceFormValues>>["control"];
  category: string;
}) {
  const items = useWatch({ control, name: "items" });
  const computed = (items ?? []).map((item) => {
    const sub = Number(item?.subtotal) || 0;
    const iva = calcIva(sub, !!item?.isTransporte);
    return { sub, iva, total: calcTotal(sub, iva) };
  });
  const grandSub = computed.reduce((s, i) => s + i.sub, 0);
  const grandIva = computed.reduce((s, i) => s + i.iva, 0);
  const grandTotal = computed.reduce((s, i) => s + i.total, 0);

  if (computed.length === 0) return null;

  return (
    <div className="rounded-md border border-primary/20 bg-primary/5 p-3 space-y-1.5">
      <p className="text-xs font-semibold text-muted-foreground uppercase tracking-wider mb-2">Totales de la factura</p>
      <div className="grid grid-cols-3 gap-2 text-sm">
        <div>
          <p className="text-xs text-muted-foreground">Subtotal</p>
          <p className="font-mono font-semibold">$ {fmtMoney(grandSub)}</p>
        </div>
        <div>
          <p className="text-xs text-muted-foreground">IVA 15%</p>
          <p className="font-mono font-semibold text-muted-foreground">$ {fmtMoney(grandIva)}</p>
        </div>
        <div>
          <p className="text-xs text-muted-foreground font-medium">Total</p>
          <p className="font-mono font-bold text-primary text-base">$ {fmtMoney(grandTotal)}</p>
        </div>
      </div>
      {canBeTransporte(category) && computed.some(i => i.iva === 0 && i.sub > 0) && (
        <p className="text-[10px] text-amber-700 flex items-center gap-1 mt-1">
          <Truck className="w-3 h-3" /> Transporte sin IVA incluido en los ítems marcados
        </p>
      )}
    </div>
  );
}

// ── Componente principal ──────────────────────────────────────────────────────

export default function InvoicesPage() {
  const { toast } = useToast();
  const { wbProd, prodFile, fileNameProd, setProdEntry } = useFileStore();

  const [period, setPeriod] = useState<string>(format(new Date(), "yyyy-MM"));
  const [editingId, setEditingId] = useState<number | null>(null);
  const [isSubmitting, setIsSubmitting] = useState(false);

  const { data: invoiceList = [], isLoading } = useInvoices(period);
  const { data: summary = {} } = useInvoiceSummary(period);
  const createMutation = useCreateInvoice(period);
  const updateMutation = useUpdateInvoice(period);
  const deleteMutation = useDeleteInvoice(period);

  const form = useForm<InvoiceFormValues>({
    resolver: zodResolver(invoiceFormSchema),
    defaultValues: {
      issueDate: format(new Date(), "yyyy-MM-dd"),
      supplier: "",
      invoiceNumber: "",
      category: "combustible_transporte",
      observations: "",
      items: [{ description: "", subtotal: 0, isTransporte: false }],
    },
  });

  const { fields, append, remove } = useFieldArray({
    control: form.control,
    name: "items",
  });

  const watchedCategory = useWatch({ control: form.control, name: "category" });

  // Producción del período
  let prodKwh: number | undefined;
  if (wbProd) {
    try {
      const ps = getMonthlyProductionSummary(wbProd, period);
      prodKwh = ps.tot_gen;
    } catch {
      prodKwh = undefined;
    }
  }

  const totalPeriod = Object.values(summary).reduce((a, b) => a + b, 0);

  const resetForm = useCallback(() => {
    setEditingId(null);
    form.reset({
      issueDate: format(new Date(), "yyyy-MM-dd"),
      supplier: "",
      invoiceNumber: "",
      category: "combustible_transporte",
      observations: "",
      items: [{ description: "", subtotal: 0, isTransporte: false }],
    });
  }, [form]);

  function startEdit(inv: Invoice) {
    setEditingId(inv.id);
    const items = parseLineItems(inv);
    form.reset({
      issueDate: inv.issueDate,
      supplier: inv.supplier,
      invoiceNumber: inv.invoiceNumber,
      category: inv.category as InvoiceCategory,
      observations: inv.observations ?? "",
      items: items.map(i => ({
        description: i.description,
        subtotal: i.subtotal,
        isTransporte: i.isTransporte,
        fuelType: i.fuelType ?? null,
      })),
    });
  }

  function buildPayload(values: InvoiceFormValues) {
    const computed = values.items.map(item => {
      const sub = Number(item.subtotal) || 0;
      const iva = calcIva(sub, !!item.isTransporte);
      const total = calcTotal(sub, iva);
      return {
        description: item.description ?? "",
        subtotal: sub,
        isTransporte: !!item.isTransporte,
        fuelType: item.isTransporte ? null : (item.fuelType ?? null),
        iva,
        total,
      };
    });
    const grandSub = computed.reduce((s, i) => s + i.subtotal, 0);
    const grandIva = computed.reduce((s, i) => s + i.iva, 0);
    const grandTotal = computed.reduce((s, i) => s + i.total, 0);
    const desc = computed.map(i => i.description).filter(Boolean).join("; ");
    return {
      period,
      issueDate: values.issueDate,
      supplier: values.supplier,
      invoiceNumber: values.invoiceNumber,
      category: values.category,
      description: desc,
      subtotal: String(Math.round(grandSub * 100) / 100),
      iva: String(Math.round(grandIva * 100) / 100),
      total: String(Math.round(grandTotal * 100) / 100),
      lineItems: JSON.stringify(computed),
      observations: values.observations ?? "",
    };
  }

  async function onSubmit(values: InvoiceFormValues) {
    setIsSubmitting(true);
    try {
      const payload = buildPayload(values);
      if (editingId !== null) {
        await updateMutation.mutateAsync({ id: editingId, data: payload });
        toast({ title: "Factura actualizada correctamente" });
      } else {
        await createMutation.mutateAsync(payload);
        toast({ title: "Factura registrada correctamente" });
      }
      resetForm();
    } catch (err) {
      toast({ title: "Error al guardar", description: String(err), variant: "destructive" });
    } finally {
      setIsSubmitting(false);
    }
  }

  async function handleDelete(id: number) {
    if (!confirm("¿Eliminar esta factura?")) return;
    try {
      await deleteMutation.mutateAsync(id);
      toast({ title: "Factura eliminada" });
      if (editingId === id) resetForm();
    } catch (err) {
      toast({ title: "Error al eliminar", description: String(err), variant: "destructive" });
    }
  }

  async function handleExportPDF() {
    const html = buildInvoiceHtml(invoiceList, period, summary, prodKwh);
    const filename = `Facturas_${period}`;
    try {
      const res = await fetch("/api/export/pdf", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ html, title: filename + ".pdf" }),
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a"); a.href = url; a.download = filename + ".pdf"; a.click();
      URL.revokeObjectURL(url);
      toast({ title: "PDF generado" });
    } catch {
      openPrintWindow(html, filename);
      toast({ title: "Ventana de impresión abierta", description: 'Selecciona "Guardar como PDF".' });
    }
  }

  function handleExportExcel() {
    exportInvoicesExcel(invoiceList, period, prodKwh);
    toast({ title: "Excel descargado", description: `Facturas_${period}.xlsx` });
  }

  return (
    <Layout>
      <div className="space-y-6">
        {/* Encabezado */}
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-lg bg-primary/10 flex items-center justify-center">
              <ReceiptText className="w-5 h-5 text-primary" />
            </div>
            <div>
              <h1 className="text-lg font-bold">Registro de Facturas</h1>
              <p className="text-xs text-muted-foreground">Central El Morro – Morro Energy S.A.</p>
            </div>
          </div>
          <div className="flex gap-2">
            <Button data-testid="button-export-excel" variant="outline" size="sm" onClick={handleExportExcel} disabled={invoiceList.length === 0}>
              <FileSpreadsheet className="w-3.5 h-3.5 mr-1" /> Excel
            </Button>
            <Button data-testid="button-export-pdf-invoices" variant="outline" size="sm" onClick={handleExportPDF} disabled={invoiceList.length === 0}>
              <FileDown className="w-3.5 h-3.5 mr-1" /> PDF
            </Button>
          </div>
        </div>

        {/* Selector de período */}
        <Card className="border-border/60 shadow-sm">
          <CardContent className="pt-4 pb-4">
            <div className="flex flex-wrap gap-4 items-end">
              <div className="flex-1 min-w-[200px] max-w-xs">
                <label className="text-xs font-medium text-muted-foreground mb-1 block">Período (AAAA-MM)</label>
                <Input
                  data-testid="input-period-invoices"
                  type="month" value={period}
                  onChange={(e) => { setPeriod(e.target.value); resetForm(); }}
                  className="text-sm h-8"
                />
              </div>
              <div className="flex items-center gap-3 text-sm">
                <span className="text-muted-foreground">Total período:</span>
                <span className="font-bold text-lg text-primary">$ {fmtMoney(totalPeriod)}</span>
              </div>
              {!prodFile && (
                <div className="flex-1 min-w-[220px]">
                  <label className="text-xs font-medium text-muted-foreground mb-1 block">
                    <FileSpreadsheet className="w-3 h-3 inline mr-1" />Producción (opcional)
                  </label>
                  <Input data-testid="input-prod-file-invoices" type="file" accept=".xlsx,.xls" className="text-xs h-8 cursor-pointer"
                    onChange={(e) => { const f = e.target.files?.[0]; if (f) setProdEntry(f); }} />
                </div>
              )}
              {fileNameProd && (
                <div className="flex items-center gap-1 text-xs text-green-600">
                  <CheckCircle2 className="w-3 h-3" /> {fileNameProd}
                  {prodKwh !== undefined && <span className="text-muted-foreground ml-1">({fmt(prodKwh, 0)} kWh)</span>}
                </div>
              )}
            </div>
          </CardContent>
        </Card>

        <div className="grid grid-cols-1 xl:grid-cols-[440px_1fr] gap-6">
          {/* Formulario */}
          <Card className="border-border/60 shadow-md h-fit">
            <div className="h-1 w-full rounded-t-lg bg-gradient-to-r from-primary to-primary/50" />
            <CardHeader className="pb-3">
              <CardTitle className="text-base flex items-center gap-2">
                {editingId !== null ? <Pencil className="w-4 h-4 text-primary" /> : <PlusCircle className="w-4 h-4 text-primary" />}
                {editingId !== null ? "Editar factura" : "Nueva factura"}
              </CardTitle>
              <CardDescription className="text-xs">Período: {period} · IVA 15% calculado automáticamente</CardDescription>
            </CardHeader>
            <CardContent>
              <Form {...form}>
                <form onSubmit={form.handleSubmit(onSubmit)} className="space-y-3">
                  {/* Encabezado de factura */}
                  <div className="grid grid-cols-2 gap-3">
                    <FormField control={form.control} name="issueDate" render={({ field }) => (
                      <FormItem>
                        <FormLabel className="text-xs">Fecha emisión</FormLabel>
                        <FormControl><Input data-testid="input-issue-date" type="date" className="text-sm h-8" {...field} /></FormControl>
                        <FormMessage />
                      </FormItem>
                    )} />
                    <FormField control={form.control} name="invoiceNumber" render={({ field }) => (
                      <FormItem>
                        <FormLabel className="text-xs">N° Factura</FormLabel>
                        <FormControl><Input data-testid="input-invoice-number" className="text-sm h-8" placeholder="001-001-0001234" {...field} /></FormControl>
                        <FormMessage />
                      </FormItem>
                    )} />
                  </div>

                  <FormField control={form.control} name="supplier" render={({ field }) => (
                    <FormItem>
                      <FormLabel className="text-xs">Proveedor</FormLabel>
                      <FormControl><Input data-testid="input-supplier" className="text-sm h-8" placeholder="Nombre del proveedor" {...field} /></FormControl>
                      <FormMessage />
                    </FormItem>
                  )} />

                  <FormField control={form.control} name="category" render={({ field }) => (
                    <FormItem>
                      <FormLabel className="text-xs">Rubro</FormLabel>
                      <Select onValueChange={field.onChange} value={field.value}>
                        <FormControl>
                          <SelectTrigger data-testid="select-category" className="text-sm h-8">
                            <SelectValue placeholder="Selecciona un rubro" />
                          </SelectTrigger>
                        </FormControl>
                        <SelectContent>
                          {INVOICE_CATEGORIES.map((cat) => (
                            <SelectItem key={cat} value={cat}>{INVOICE_CATEGORY_LABELS[cat]}</SelectItem>
                          ))}
                        </SelectContent>
                      </Select>
                      <FormMessage />
                    </FormItem>
                  )} />

                  {/* Aviso de transporte */}
                  {canBeTransporte(watchedCategory) && (
                    <div className="flex items-start gap-2 rounded-md border border-amber-200/70 bg-amber-50/60 px-3 py-2">
                      <Truck className="w-3.5 h-3.5 text-amber-600 mt-0.5 shrink-0" />
                      <p className="text-xs text-amber-800">
                        Este rubro puede incluir transporte. Activa el switch en cada ítem para marcar los valores que no gravan IVA.
                      </p>
                    </div>
                  )}

                  {/* Ítems dinámicos */}
                  <div className="space-y-2">
                    <div className="flex items-center justify-between">
                      <label className="text-xs font-semibold">Ítems de la factura</label>
                      <Button
                        data-testid="button-add-item"
                        type="button" variant="outline" size="sm"
                        className="h-7 text-xs gap-1"
                        onClick={() => append({ description: "", subtotal: 0, isTransporte: false })}
                      >
                        <Plus className="w-3 h-3" /> Añadir ítem
                      </Button>
                    </div>

                    <div className="space-y-2 max-h-[400px] overflow-y-auto pr-1">
                      {fields.map((field, index) => (
                        <LineItemRow
                          key={field.id}
                          index={index}
                          category={watchedCategory}
                          onRemove={() => remove(index)}
                          canRemove={fields.length > 1}
                          control={form.control}
                        />
                      ))}
                    </div>
                  </div>

                  {/* Totales computados */}
                  <InvoiceTotals control={form.control} category={watchedCategory} />

                  <FormField control={form.control} name="observations" render={({ field }) => (
                    <FormItem>
                      <FormLabel className="text-xs">Observaciones</FormLabel>
                      <FormControl>
                        <Textarea data-testid="input-observations" className="resize-none text-xs min-h-[50px]" placeholder="Notas opcionales..." {...field} />
                      </FormControl>
                      <FormMessage />
                    </FormItem>
                  )} />

                  <div className="flex gap-2 pt-1">
                    <Button data-testid="button-submit-invoice" type="submit" size="sm" className="flex-1" disabled={isSubmitting}>
                      {isSubmitting ? "Guardando..." : editingId !== null ? "Actualizar" : "Registrar factura"}
                    </Button>
                    {editingId !== null && (
                      <Button data-testid="button-cancel-edit" type="button" variant="outline" size="sm" onClick={resetForm}>
                        Cancelar
                      </Button>
                    )}
                  </div>
                </form>
              </Form>
            </CardContent>
          </Card>

          {/* Tabla + resumen */}
          <div className="space-y-4">
            <Card className="border-border/60 shadow-sm">
              <CardHeader className="pb-2">
                <CardTitle className="text-sm">Facturas del período {period}</CardTitle>
                <CardDescription className="text-xs">{invoiceList.length} factura{invoiceList.length !== 1 ? "s" : ""} registradas</CardDescription>
              </CardHeader>
              <CardContent className="p-0">
                {isLoading ? (
                  <div className="flex items-center justify-center h-24 text-muted-foreground text-sm">Cargando...</div>
                ) : invoiceList.length === 0 ? (
                  <div className="flex flex-col items-center justify-center h-24 text-muted-foreground/60">
                    <AlertCircle className="w-8 h-8 mb-2 opacity-40" />
                    <p className="text-xs">No hay facturas para este período</p>
                  </div>
                ) : (
                  <div className="overflow-x-auto">
                    <table className="w-full text-xs">
                      <thead>
                        <tr className="border-b border-border/60 bg-muted/30">
                          <th className="text-left px-3 py-2 font-semibold">Fecha</th>
                          <th className="text-left px-3 py-2 font-semibold">Proveedor</th>
                          <th className="text-left px-3 py-2 font-semibold">N°</th>
                          <th className="text-left px-3 py-2 font-semibold">Rubro</th>
                          <th className="text-right px-3 py-2 font-semibold">Subtotal</th>
                          <th className="text-right px-3 py-2 font-semibold">IVA</th>
                          <th className="text-right px-3 py-2 font-semibold">Total</th>
                          <th className="px-3 py-2 font-semibold text-center">Acc.</th>
                        </tr>
                      </thead>
                      <tbody>
                        {invoiceList.map((inv) => {
                          const items = parseLineItems(inv);
                          const hasMulti = items.length > 1;
                          return (
                            <>
                              <tr key={inv.id} data-testid={`row-invoice-${inv.id}`} className={`border-b border-border/40 hover:bg-muted/20 transition-colors ${editingId === inv.id ? "bg-primary/5" : ""}`}>
                                <td className="px-3 py-2 text-muted-foreground">{inv.issueDate}</td>
                                <td className="px-3 py-2 font-medium max-w-[100px] truncate">{inv.supplier}</td>
                                <td className="px-3 py-2 font-mono text-muted-foreground text-[10px]">{inv.invoiceNumber}</td>
                                <td className="px-3 py-2">
                                  <span className="inline-block px-1.5 py-0.5 rounded text-[10px] bg-primary/10 text-primary font-medium">
                                    {INVOICE_CATEGORY_LABELS[inv.category as InvoiceCategory] ?? inv.category}
                                  </span>
                                </td>
                                <td className="px-3 py-2 text-right font-mono">$ {fmtMoney(inv.subtotal)}</td>
                                <td className="px-3 py-2 text-right font-mono text-muted-foreground">
                                  {parseFloat(inv.iva ?? "0") === 0
                                    ? <span className="text-amber-600 text-[10px]">Exento</span>
                                    : `$ ${fmtMoney(inv.iva)}`}
                                </td>
                                <td className="px-3 py-2 text-right font-mono font-semibold">$ {fmtMoney(inv.total)}</td>
                                <td className="px-3 py-2 text-center">
                                  <div className="flex items-center justify-center gap-1">
                                    {hasMulti && (
                                      <span className="text-[9px] bg-muted px-1 rounded text-muted-foreground">{items.length} íts.</span>
                                    )}
                                    <Button data-testid={`button-edit-invoice-${inv.id}`} variant="ghost" size="icon" className="h-6 w-6" onClick={() => startEdit(inv)}>
                                      <Pencil className="w-3 h-3" />
                                    </Button>
                                    <Button data-testid={`button-delete-invoice-${inv.id}`} variant="ghost" size="icon" className="h-6 w-6 text-destructive hover:text-destructive" onClick={() => handleDelete(inv.id)}>
                                      <Trash2 className="w-3 h-3" />
                                    </Button>
                                  </div>
                                </td>
                              </tr>
                              {hasMulti && items.map((item, iIdx) => (
                                <tr key={`${inv.id}-item-${iIdx}`} className="border-b border-border/20 bg-muted/5">
                                  <td colSpan={4} className="px-3 py-1 pl-8 text-[10px] text-muted-foreground">
                                    ↳ {item.isTransporte
                                        ? <span className="text-amber-600"><Truck className="w-2.5 h-2.5 inline mr-0.5" />Transporte</span>
                                        : item.fuelType
                                          ? <span className="text-blue-700 font-medium">{FUEL_TYPE_LABELS[item.fuelType]}</span>
                                          : null}
                                    {item.description ? " – " + item.description : ""}
                                  </td>
                                  <td className="px-3 py-1 text-right text-[10px] font-mono text-muted-foreground">$ {fmtMoney(item.subtotal)}</td>
                                  <td className="px-3 py-1 text-right text-[10px] font-mono text-muted-foreground">
                                    {item.isTransporte ? <span className="text-amber-600">Exento</span> : `$ ${fmtMoney(item.iva)}`}
                                  </td>
                                  <td className="px-3 py-1 text-right text-[10px] font-mono text-muted-foreground">$ {fmtMoney(item.total)}</td>
                                  <td />
                                </tr>
                              ))}
                            </>
                          );
                        })}
                        <tr className="bg-primary/5 border-t-2 border-primary/20">
                          <td colSpan={6} className="px-3 py-2 font-semibold text-xs">TOTAL PERÍODO</td>
                          <td className="px-3 py-2 text-right font-bold font-mono text-primary">$ {fmtMoney(totalPeriod)}</td>
                          <td />
                        </tr>
                      </tbody>
                    </table>
                  </div>
                )}
              </CardContent>
            </Card>

            {/* Resumen por rubro */}
            <Card className="border-border/60 shadow-sm">
              <CardHeader className="pb-2">
                <CardTitle className="text-sm">Resumen mensual por rubro</CardTitle>
                {!prodFile && (
                  <CardDescription className="text-xs text-amber-600 flex items-center gap-1">
                    <AlertCircle className="w-3 h-3" /> Sin producción cargada — CV real no disponible
                  </CardDescription>
                )}
              </CardHeader>
              <CardContent className="p-0">
                <div className="overflow-x-auto">
                  <table className="w-full text-xs">
                    <thead>
                      <tr className="border-b border-border/60 bg-muted/30">
                        <th className="text-left px-3 py-2 font-semibold">Rubro</th>
                        <th className="text-right px-3 py-2 font-semibold">Total USD</th>
                        <th className="text-right px-3 py-2 font-semibold">CV real USD/kWh</th>
                      </tr>
                    </thead>
                    <tbody>
                      {INVOICE_CATEGORIES.map((cat) => {
                        const total = summary[cat] ?? 0;
                        const cv = prodKwh && prodKwh > 0 ? total / prodKwh : null;
                        return (
                          <tr key={cat} data-testid={`summary-row-${cat}`} className="border-b border-border/40">
                            <td className="px-3 py-2 font-medium">{INVOICE_CATEGORY_LABELS[cat]}</td>
                            <td className="px-3 py-2 text-right font-mono">$ {fmtMoney(total)}</td>
                            <td className="px-3 py-2 text-right font-mono text-muted-foreground">{cv !== null ? fmt(cv, 4) : "—"}</td>
                          </tr>
                        );
                      })}
                      <tr className="bg-primary/5 border-t-2 border-primary/20">
                        <td className="px-3 py-2 font-bold text-xs">TOTAL</td>
                        <td className="px-3 py-2 text-right font-bold font-mono text-primary">$ {fmtMoney(totalPeriod)}</td>
                        <td className="px-3 py-2 text-right font-mono font-semibold">{prodKwh && prodKwh > 0 ? fmt(totalPeriod / prodKwh, 4) : "—"}</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </CardContent>
            </Card>
          </div>
        </div>
      </div>
    </Layout>
  );
}
