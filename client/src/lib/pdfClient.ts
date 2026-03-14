/**
 * pdfClient.ts — Morro Energy S.A.
 *
 * Cliente para exportación PDF vía servicio HTTP externo configurable
 * con la variable de entorno VITE_PDF_SERVICE_URL.
 *
 * El servicio externo recibe POST { url: string } y devuelve application/pdf
 * (Puppeteer remoto o equivalente: browserless.io, html-pdf-node, etc.).
 *
 * Estrategia de exportación:
 *   1. Si VITE_PDF_SERVICE_URL está configurada: POST URL de vista print → PDF
 *   2. Fallback: exportReportPDF (html2canvas + pdfmake, 100% cliente)
 */

import { exportReportPDF } from "./pdfExporter";

const PDF_SERVICE_URL = (import.meta.env.VITE_PDF_SERVICE_URL ?? "") as string;

export function isPdfServiceConfigured(): boolean {
  return PDF_SERVICE_URL.trim().length > 0;
}

function storePrintHTML(html: string, title: string): string {
  const key = `morro_print_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
  try {
    sessionStorage.setItem(key, html);
  } catch {
    console.warn("sessionStorage lleno; no se puede almacenar vista de impresión.");
  }
  const params = new URLSearchParams({ key, title });
  return `${window.location.origin}/print-view?${params.toString()}`;
}

async function downloadFromService(printUrl: string, filename: string): Promise<void> {
  const response = await fetch(PDF_SERVICE_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ url: printUrl }),
  });
  if (!response.ok) {
    throw new Error(`Servicio PDF respondió ${response.status}: ${response.statusText}`);
  }
  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

export type ExportResult = { usedFallback: boolean };

async function exportarConFallback(
  html: string,
  title: string,
  filename: string,
  fallbackElement?: HTMLElement,
): Promise<ExportResult> {
  if (isPdfServiceConfigured()) {
    try {
      const printUrl = storePrintHTML(html, title);
      await downloadFromService(printUrl, filename);
      return { usedFallback: false };
    } catch (err) {
      console.warn("Servicio PDF externo falló; usando fallback local:", err);
    }
  }
  if (fallbackElement) {
    await exportReportPDF(fallbackElement, filename);
  }
  return { usedFallback: true };
}

export async function exportarReporteDiarioPDF(
  html: string,
  fecha: string,
  fallbackElement?: HTMLElement,
): Promise<ExportResult> {
  return exportarConFallback(
    html,
    `Reporte Diario – ${fecha}`,
    `Reporte_Diario_ElMorro_${fecha}.pdf`,
    fallbackElement,
  );
}

export async function exportarReporteMensualPDF(
  html: string,
  mes: string,
  fallbackElement?: HTMLElement,
): Promise<ExportResult> {
  return exportarConFallback(
    html,
    `Reporte Mensual – ${mes}`,
    `Reporte_Mensual_ElMorro_${mes}.pdf`,
    fallbackElement,
  );
}

export async function exportarFacturacionPDF(
  html: string,
  mes: string,
  fallbackElement?: HTMLElement,
): Promise<ExportResult> {
  return exportarConFallback(
    html,
    `Facturación – ${mes}`,
    `Facturacion_ElMorro_${mes}.pdf`,
    fallbackElement,
  );
}

export function abrirVistaPrint(html: string, title: string): void {
  const url = storePrintHTML(html, title);
  window.open(url, "_blank", "noopener,noreferrer");
}
