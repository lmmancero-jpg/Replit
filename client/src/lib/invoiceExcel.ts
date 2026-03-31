import * as XLSX from "xlsx";
import { INVOICE_CATEGORIES, INVOICE_CATEGORY_LABELS, type Invoice, type InvoiceCategory } from "@shared/schema";

function fmtMoney(v: string | number | null | undefined): number {
  return parseFloat(String(v ?? "0")) || 0;
}

export function exportInvoicesExcel(invoices: Invoice[], period: string, productionKwh?: number): void {
  const wb = XLSX.utils.book_new();

  // ── Hoja 1: Facturas ────────────────────────────────────────────────────────
  const facturasHeader = [
    "ID", "Período", "Fecha emisión", "Proveedor", "N° Factura",
    "Rubro", "Descripción", "Subtotal USD", "IVA USD", "Total USD", "Observaciones",
  ];

  const facturasData = invoices.map((inv) => [
    inv.id,
    inv.period,
    inv.issueDate,
    inv.supplier,
    inv.invoiceNumber,
    INVOICE_CATEGORY_LABELS[inv.category as InvoiceCategory] ?? inv.category,
    inv.description ?? "",
    fmtMoney(inv.subtotal),
    fmtMoney(inv.iva),
    fmtMoney(inv.total),
    inv.observations ?? "",
  ]);

  const wsFact = XLSX.utils.aoa_to_sheet([facturasHeader, ...facturasData]);
  wsFact["!cols"] = [
    { wch: 6 }, { wch: 10 }, { wch: 14 }, { wch: 30 }, { wch: 16 },
    { wch: 32 }, { wch: 40 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 30 },
  ];
  XLSX.utils.book_append_sheet(wb, wsFact, "Facturas");

  // ── Hoja 2: Resumen ──────────────────────────────────────────────────────────
  const totByCategory: Record<string, number> = {};
  for (const cat of INVOICE_CATEGORIES) totByCategory[cat] = 0;
  for (const inv of invoices) {
    const cat = inv.category;
    if (cat in totByCategory) totByCategory[cat] += fmtMoney(inv.total);
  }
  const grandTotal = Object.values(totByCategory).reduce((a, b) => a + b, 0);

  const resumHeader = ["Rubro", "Total USD"];
  if (productionKwh !== undefined && productionKwh > 0) {
    resumHeader.push("CV real USD/kWh");
  }

  const resumData: (string | number)[][] = [];
  for (const cat of INVOICE_CATEGORIES) {
    const row: (string | number)[] = [
      INVOICE_CATEGORY_LABELS[cat as InvoiceCategory] ?? cat,
      totByCategory[cat],
    ];
    if (productionKwh !== undefined && productionKwh > 0) {
      row.push(totByCategory[cat] / productionKwh);
    }
    resumData.push(row);
  }
  resumData.push([]);
  resumData.push(["TOTAL", grandTotal, ...(productionKwh && productionKwh > 0 ? [grandTotal / productionKwh] : [])]);

  if (productionKwh !== undefined) {
    resumData.push([]);
    resumData.push(["Producción del período (kWh)", productionKwh]);
  }

  const wsResum = XLSX.utils.aoa_to_sheet([resumHeader, ...resumData]);
  wsResum["!cols"] = [{ wch: 36 }, { wch: 16 }, { wch: 18 }];
  XLSX.utils.book_append_sheet(wb, wsResum, "Resumen");

  XLSX.writeFile(wb, `Facturas_${period}.xlsx`);
}
