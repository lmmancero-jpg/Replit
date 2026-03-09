/**
 * pdfExporter.ts  —  Morro Energy S.A.
 *
 * Módulo de exportación PDF unificado para informes y métricas.
 *
 * Arquitectura:
 *  1. html2canvas captura el DOM REAL (sin clonar) a escala 3× → nitidez de
 *     impresión. NUNCA se pasa `width` a html2canvas para evitar que clips
 *     trunquen contenido; assemblePages escala cualquier ancho al A4.
 *
 *  2. Algoritmo de saltos de página hacia atrás:
 *     findSafeCut busca el espacio en blanco más cercano buscando SOLO hacia
 *     atrás desde el límite teórico (máx. BACK_RANGE px). Esto garantiza:
 *       (a) cutY ≤ rawEnd → la imagen NUNCA desborda la página.
 *       (b) 300px de margen hacia atrás alcanza el hueco entre filas de
 *           gráficos (≈215px de alto), evitando cortes dentro de un gráfico.
 *
 *  3. Las métricas capturan cada sección [data-pdf-section] de forma
 *     independiente → ningún gráfico queda partido entre secciones.
 *
 *  4. Imágenes exportadas en JPEG 0.92 → excelente nitidez, archivo compacto.
 */

// ─── Constantes ───────────────────────────────────────────────────────────────

const RENDER_SCALE  = 3;
const JPEG_QUALITY  = 0.92;
const REPORT_WIN_W  = 1250;
const METRICS_WIN_W = 794;   // A4 portrait @96dpi
/**
 * Distancia máxima hacia atrás para buscar un espacio en blanco antes del
 * límite teórico de página. 300px canvas ≈ 100px DOM @scale-3.
 * Las filas de gráficos de 155px (altura DOM) dan huecos de ≈36-72px canvas;
 * el hueco más alejado del límite teórico es ≈220px canvas → 300px garantiza
 * encontrarlo.
 */
const BACK_RANGE    = 300;
/** Muestras horizontales por línea escaneada. */
const SCAN_SAMPLES  = 150;

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
 * NO se pasa `width` para no truncar el contenido — el escalado al A4
 * lo realiza assemblePages mediante la relación pxPerMm.
 */
async function captureElement(
  el:       HTMLElement,
  h2c:      any,
  windowW?: number,
): Promise<HTMLCanvasElement> {
  const opts: Record<string, unknown> = {
    scale:           RENDER_SCALE,
    useCORS:         true,
    allowTaint:      false,
    logging:         false,
    backgroundColor: "#ffffff",
  };
  if (windowW !== undefined) opts.windowWidth = windowW;
  return h2c(el, opts) as Promise<HTMLCanvasElement>;
}

// ─── Algoritmo de corte seguro (solo hacia atrás) ─────────────────────────────

/**
 * Busca hacia atrás desde `targetY` (límite teórico de página) la línea
 * horizontal de mayor luminosidad media (= espacio en blanco entre filas).
 *
 * Restricción crítica: maxY = targetY  →  cutY ≤ rawEnd  →  la franja de
 * imagen NUNCA excede la altura de la página PDF.
 *
 * Con BACK_RANGE = 300px canvas (@scale-3 = 100px DOM), el algoritmo alcanza
 * el hueco entre filas de gráficos de 155px de altura DOM.
 */
function findSafeCut(canvas: HTMLCanvasElement, targetY: number): number {
  const ctx = canvas.getContext("2d");
  if (!ctx) return targetY;

  const w    = canvas.width;
  const minY = Math.max(1,               targetY - BACK_RANGE);
  const maxY = Math.min(canvas.height - 2, targetY);          // ← solo atrás
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
    // Penalizar distancia: favorecer cortes cerca del límite teórico
    // pero sólo hacia atrás (mayor Y = más cercano al límite = menor penalización)
    const score = lum - (targetY - y) * 0.6;
    if (score > bestScore) { bestScore = score; bestY = y; }
  }
  return bestY;
}

// ─── Ensamblado de páginas en jsPDF ───────────────────────────────────────────

/**
 * Divide el canvas en franjas verticales y las distribuye en páginas PDF.
 * Cada corte se localiza en espacio en blanco mediante findSafeCut.
 * isFirstPage se pasa por referencia para encadenar secciones.
 */
function assemblePages(
  pdf:         any,
  canvas:      HTMLCanvasElement,
  contentMmW:  number,
  contentMmH:  number,
  marginMm:    number,
  isFirstPage: { value: boolean },
): void {
  const pxPerMm = canvas.width / contentMmW;
  const pageHPx = Math.round(contentMmH * pxPerMm);
  const totalH  = canvas.height;
  let   yPx     = 0;

  while (yPx < totalH) {
    if (!isFirstPage.value) pdf.addPage();
    isFirstPage.value = false;

    const rawEnd = yPx + pageHPx;
    // findSafeCut devuelve ≤ rawEnd → sliceH ≤ pageHPx → sliceMmH ≤ contentMmH
    const cutY   = rawEnd < totalH ? findSafeCut(canvas, rawEnd) : totalH;
    const sliceH   = Math.max(1, cutY - yPx);
    const sliceMmH = sliceH / pxPerMm;

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
 * Captura el elemento .report-wrapper a su ancho natural en el DOM (sin
 * clip de width). assemblePages lo escala a los 194mm de ancho útil del A4.
 */
export async function exportReportPDF(
  element:  HTMLElement,
  filename: string,
): Promise<void> {
  const { html2canvas, jsPDF } = await loadDeps();

  const marginMm   = 8;
  const contentMmW = 210 - marginMm * 2;   // 194 mm
  const contentMmH = 297 - marginMm * 2;   // 281 mm

  const canvas = await captureElement(element, html2canvas, REPORT_WIN_W);

  const pdf         = new jsPDF({ unit: "mm", format: "a4", orientation: "portrait" });
  const isFirstPage = { value: true };

  assemblePages(pdf, canvas, contentMmW, contentMmH, marginMm, isFirstPage);
  pdf.save(filename);
}

/**
 * Coloca un canvas en UNA sola página A4, escalando si es necesario.
 *
 * Garantiza que el contenido nunca sea cortado por salto de página:
 *  – Si el canvas cabe con ancho completo → se coloca tal cual.
 *  – Si la altura resultante excede el área útil → se escala uniformemente
 *    para ajustarse a la altura disponible (centered horizontally).
 */
function placeOnePage(
  pdf:         any,
  canvas:      HTMLCanvasElement,
  contentMmW:  number,
  contentMmH:  number,
  marginMm:    number,
  isFirstPage: { value: boolean },
): void {
  if (!isFirstPage.value) pdf.addPage();
  isFirstPage.value = false;

  const aspectRatio = canvas.height / canvas.width;

  let drawW = contentMmW;
  let drawH = drawW * aspectRatio;

  if (drawH > contentMmH) {
    // La sección es más alta que la página → escalar para que quepa en altura
    drawH = contentMmH;
    drawW = drawH / aspectRatio;
  }

  // Centrar horizontalmente si se escaló para ajuste de altura
  const xOffset = marginMm + (contentMmW - drawW) / 2;

  pdf.addImage(
    canvas.toDataURL("image/jpeg", JPEG_QUALITY),
    "JPEG",
    xOffset, marginMm,
    drawW, drawH,
  );
}

/**
 * Exporta las métricas a PDF A4 portrait.
 *
 * Cada sección [data-pdf-section] ocupa exactamente UNA hoja, sin cortes
 * por salto de página. Si una sección supera la altura útil del A4 se
 * escala proporcionalmente para que quede completa en una sola hoja.
 */
export async function exportMetricsPDF(
  container: HTMLElement,
  period:    string,
): Promise<void> {
  const { html2canvas, jsPDF } = await loadDeps();

  const marginMm   = 8;
  const contentMmW = 210 - marginMm * 2;   // 194 mm
  const contentMmH = 297 - marginMm * 2;   // 281 mm

  const sections = Array.from(
    container.querySelectorAll<HTMLElement>("[data-pdf-section]"),
  );

  const pdf         = new jsPDF({ unit: "mm", format: "a4", orientation: "portrait" });
  const isFirstPage = { value: true };

  for (const section of sections) {
    const canvas = await captureElement(section, html2canvas, METRICS_WIN_W);
    placeOnePage(pdf, canvas, contentMmW, contentMmH, marginMm, isFirstPage);
  }

  pdf.save(`Metricas_ElMorro_${period}.pdf`);
}
