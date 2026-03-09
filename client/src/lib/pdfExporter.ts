/**
 * pdfExporter.ts  —  Morro Energy S.A.
 *
 * Módulo de exportación PDF unificado para informes y métricas.
 *
 * Arquitectura:
 *  1. html2canvas captura el DOM REAL (sin clonar) a escala 3× → nitidez de
 *     impresión. No se usa html2pdf para evitar la clonación de DOM que borra
 *     el contenido de los <canvas> de Chart.js.
 *
 *  2. Algoritmo de saltos de página por escaneo de píxeles:
 *     En lugar de depender de CSS break-inside o medición de DOM (que varía
 *     según el viewport), se escanea horizontalmente el canvas ya renderizado.
 *     Para cada límite teórico de página se busca la línea de píxeles con mayor
 *     luminosidad media dentro de ±SEARCH_RANGE px → ese espacio en blanco entre
 *     filas/bloques es el punto de corte real. Nunca corta dentro de una celda.
 *
 *  3. Las métricas capturan cada sección [data-pdf-section] de forma
 *     independiente → ningún gráfico queda partido entre páginas.
 *
 *  4. Imágenes exportadas en JPEG 0.93 → excelente nitidez, archivo compacto.
 */

// ─── Constantes de calidad y layout ───────────────────────────────────────────

/** Factor de resolución: 3× da ~300dpi efectivos → calidad de imprenta. */
const RENDER_SCALE = 3;

/** Calidad JPEG [0-1]. 0.93 preserva bordes y texto sin artefactos visibles. */
const JPEG_QUALITY = 0.93;

/**
 * windowWidth para informes: renderiza el HTML como si el viewport fuese
 * 1250px → el contenido se comprime proporcionalmente al exportar a A4,
 * resultando en texto más pequeño y más filas por página.
 */
const REPORT_WIN_W = 1250;

/**
 * Ancho de captura para informes en px (A4 @96dpi = 794px).
 * html2canvas usa este valor para ajustar el layout antes de capturar.
 */
const REPORT_CAPTURE_W = 794;

/** Ancho de ventana simulado para métricas: A4 landscape @96dpi = 1122px. */
const METRICS_WIN_W = 1122;

/** Rango de búsqueda ±px alrededor del límite teórico de página. */
const SEARCH_RANGE = 90;

/** Muestras horizontales por línea escaneada (velocidad vs. precisión). */
const SCAN_SAMPLES = 150;

// ─── Carga diferida de dependencias ───────────────────────────────────────────

/* eslint-disable @typescript-eslint/no-explicit-any */
let _html2canvas: any = null;
let _jsPDF:       any = null;

async function loadDeps(): Promise<{ html2canvas: any; jsPDF: any }> {
  if (!_html2canvas || !_jsPDF) {
    const [h2c, jspdf] = await Promise.all([
      import("html2canvas"),
      import("jspdf"),
    ]);
    _html2canvas = h2c.default;
    _jsPDF       = jspdf.jsPDF;
  }
  return { html2canvas: _html2canvas, jsPDF: _jsPDF };
}

// ─── Captura de elemento ───────────────────────────────────────────────────────

/**
 * Captura un elemento del DOM real a escala RENDER_SCALE.
 * allowTaint:false → canvas no taintado → podemos leer sus píxeles después.
 * Todos los canvas de Chart.js son mismo origen, no causan taint.
 */
async function captureElement(
  el:        HTMLElement,
  h2c:       any,
  captureW?: number,
  windowW?:  number,
): Promise<HTMLCanvasElement> {
  const opts: Record<string, unknown> = {
    scale:           RENDER_SCALE,
    useCORS:         true,
    allowTaint:      false,
    logging:         false,
    backgroundColor: "#ffffff",
  };
  if (captureW !== undefined) opts.width       = captureW;
  if (windowW  !== undefined) opts.windowWidth = windowW;
  return h2c(el, opts) as Promise<HTMLCanvasElement>;
}

// ─── Algoritmo de corte seguro por escaneo de píxeles ─────────────────────────

/**
 * Dado el canvas renderizado y un punto teórico de corte (targetY en px),
 * busca dentro de ±SEARCH_RANGE px la línea horizontal con mayor luminosidad
 * media (= espacio en blanco entre filas/bloques). Devuelve el Y óptimo.
 *
 * Coeficientes de luminancia perceptual Rec.709 (idénticos al ojo humano).
 * Penaliza distancia al punto teórico → el corte no se desplaza demasiado.
 */
function findSafeCut(canvas: HTMLCanvasElement, targetY: number): number {
  const ctx = canvas.getContext("2d");
  if (!ctx) return targetY;

  const w    = canvas.width;
  const minY = Math.max(1,               targetY - SEARCH_RANGE);
  const maxY = Math.min(canvas.height - 2, targetY + SEARCH_RANGE);
  const step = Math.max(1, Math.floor(w / SCAN_SAMPLES));

  let bestY     = targetY;
  let bestScore = -Infinity;

  for (let y = minY; y <= maxY; y++) {
    const row = ctx.getImageData(0, y, w, 1).data;
    let lum = 0;
    let n   = 0;
    for (let x = 0; x < w; x += step) {
      const i = x * 4;
      lum += row[i] * 0.2126 + row[i + 1] * 0.7152 + row[i + 2] * 0.0722;
      n++;
    }
    lum /= n;
    const score = lum - Math.abs(y - targetY) * 1.8;
    if (score > bestScore) { bestScore = score; bestY = y; }
  }
  return bestY;
}

// ─── Ensamblado de páginas en jsPDF ───────────────────────────────────────────

/**
 * Recorre el canvas completo cortándolo en franjas (= páginas PDF).
 * Cada punto de corte se calcula con findSafeCut para caer en espacios blancos.
 * El objeto isFirstPage se pasa por referencia para manejar múltiples secciones.
 */
function assemblePages(
  pdf:          any,
  canvas:       HTMLCanvasElement,
  contentMmW:   number,
  contentMmH:   number,
  marginMm:     number,
  isFirstPage:  { value: boolean },
): void {
  const pxPerMm  = canvas.width / contentMmW;
  const pageHPx  = Math.round(contentMmH * pxPerMm);
  const totalH   = canvas.height;
  let   yPx      = 0;

  while (yPx < totalH) {
    if (!isFirstPage.value) pdf.addPage();
    isFirstPage.value = false;

    // Punto de corte seguro: busca espacio blanco entre filas
    const rawEnd = yPx + pageHPx;
    const cutY   = rawEnd < totalH ? findSafeCut(canvas, rawEnd) : totalH;

    const sliceH   = Math.max(1, cutY - yPx);
    const sliceMmH = sliceH / pxPerMm;

    // Crear franja del canvas para esta página
    const slice   = document.createElement("canvas");
    slice.width   = canvas.width;
    slice.height  = Math.ceil(sliceH);
    const ctx     = slice.getContext("2d")!;
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, slice.width, slice.height);
    ctx.drawImage(
      canvas,
      0, yPx, canvas.width, Math.ceil(sliceH),
      0, 0,   canvas.width, Math.ceil(sliceH),
    );

    pdf.addImage(
      slice.toDataURL("image/jpeg", JPEG_QUALITY),
      "JPEG",
      marginMm, marginMm,
      contentMmW, sliceMmH,
    );

    yPx = cutY;
  }
}

// ─── API PÚBLICA ───────────────────────────────────────────────────────────────

/**
 * Exporta un informe (diario / mensual / facturación) a PDF A4 portrait.
 *
 * Estrategia:
 * - html2canvas captura el .report-wrapper al ancho A4 (794px) con viewport
 *   simulado de 1250px → fuente más compacta → más contenido por página.
 * - Escala 3× → nitidez de impresión (~300dpi equivalente).
 * - Cortes de página detectados por luminosidad de píxeles → nunca corta
 *   dentro de una celda de tabla.
 *
 * @param element  El .report-wrapper visible en el DOM del generador.
 * @param filename Nombre del archivo resultante (incluye .pdf).
 */
export async function exportReportPDF(
  element:  HTMLElement,
  filename: string,
): Promise<void> {
  const { html2canvas, jsPDF } = await loadDeps();

  // A4 portrait: 210 × 297 mm, margen 8 mm → área útil 194 × 281 mm
  const marginMm   = 8;
  const contentMmW = 210 - marginMm * 2;
  const contentMmH = 297 - marginMm * 2;

  const canvas = await captureElement(element, html2canvas, REPORT_CAPTURE_W, REPORT_WIN_W);

  const pdf         = new jsPDF({ unit: "mm", format: "a4", orientation: "portrait" });
  const isFirstPage = { value: true };

  assemblePages(pdf, canvas, contentMmW, contentMmH, marginMm, isFirstPage);
  pdf.save(filename);
}

/**
 * Exporta las métricas a PDF A4 landscape.
 *
 * Estrategia:
 * - Cada sección [data-pdf-section] se captura con html2canvas de forma
 *   independiente → Chart.js pinta sus canvases antes de la captura (no hay
 *   clonación de DOM) → gráficos siempre completos.
 * - Si una sección supera la altura de una página A4 landscape, el algoritmo
 *   de píxeles busca el corte en el espacio entre filas de gráficos.
 * - Escala 3× → nitidez de impresión.
 *
 * @param container El div ref del print container en metrics.tsx.
 * @param period    Período del informe, e.g. "03_2026".
 */
export async function exportMetricsPDF(
  container: HTMLElement,
  period:    string,
): Promise<void> {
  const { html2canvas, jsPDF } = await loadDeps();

  // A4 landscape: 297 × 210 mm, margen 7 mm → área útil 283 × 196 mm
  const marginMm   = 7;
  const contentMmW = 297 - marginMm * 2;
  const contentMmH = 210 - marginMm * 2;

  const sections = Array.from(
    container.querySelectorAll<HTMLElement>("[data-pdf-section]"),
  );

  const pdf         = new jsPDF({ unit: "mm", format: "a4", orientation: "landscape" });
  const isFirstPage = { value: true };

  for (const section of sections) {
    /*
     * No pasamos captureW: html2canvas usa el ancho real del elemento en el
     * DOM (ya fijado por el contenedor position:fixed de 1122px).
     * windowWidth:1122 asegura correcto cálculo de layout CSS.
     */
    const canvas = await captureElement(section, html2canvas, undefined, METRICS_WIN_W);
    assemblePages(pdf, canvas, contentMmW, contentMmH, marginMm, isFirstPage);
  }

  pdf.save(`Metricas_ElMorro_${period}.pdf`);
}
