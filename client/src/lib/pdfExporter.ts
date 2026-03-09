/**
 * pdfExporter.ts  —  Morro Energy S.A.
 *
 * Exportación PDF con pdfmake:
 *  – Informes: el HTML generado se post-procesa para inyectar estilos inline,
 *    luego html-to-pdfmake lo convierte a definición de documento pdfmake
 *    (texto vectorial, tablas nativas, sin html2canvas para los informes).
 *  – Métricas: html2canvas captura cada sección [data-pdf-section] como
 *    imagen JPEG que pdfmake maqueta en A4 portrait (1 hoja por sección,
 *    sin cortes).
 */

const JPEG_QUALITY  = 0.92;
const RENDER_SCALE  = 3;
const METRICS_WIN_W = 794;  // A4 portrait @96dpi

/* eslint-disable @typescript-eslint/no-explicit-any */
let _h2c: any = null;
let _pm:  any = null;
let _h2pm: any = null;

async function loadH2C(): Promise<any> {
  if (!_h2c) _h2c = (await import("html2canvas")).default;
  return _h2c;
}

async function loadPdfMake(): Promise<any> {
  if (!_pm) {
    const [pmMod, fontsMod] = await Promise.all([
      import("pdfmake/build/pdfmake"),
      import("pdfmake/build/vfs_fonts"),
    ]);
    _pm = (pmMod as any).default ?? pmMod;
    const fonts = (fontsMod as any).default ?? fontsMod;
    _pm.addVirtualFileSystem(fonts);
  }
  return _pm;
}

async function loadHtmlToPdfmake(): Promise<any> {
  if (!_h2pm) {
    const mod = await import("html-to-pdfmake");
    _h2pm = (mod as any).default ?? mod;
  }
  return _h2pm;
}

// ─── Inyección de estilos inline ──────────────────────────────────────────────

/**
 * Post-procesa el HTML del informe:
 * – Elimina <style> embebidos (innecesarios en pdfmake).
 * – Aplica estilos inline en todos los elementos clave para que
 *   html-to-pdfmake los procese sin depender de CSS externo.
 *
 * Estrategia de prioridad:
 *  • td base: estilos PREPEND (menor prioridad, último en la cadena pierde).
 *  • clases especiales (label, hi, warn, row-total…): APPEND → ganan.
 */
function injectInlineStyles(html: string): string {
  const parser = new DOMParser();
  const doc    = parser.parseFromString(`<div id="__r">${html}</div>`, "text/html");
  const root   = doc.getElementById("__r")!;

  function set(el: Element, styles: string): void {
    (el as HTMLElement).setAttribute("style", styles);
  }
  function add(el: Element, styles: string): void {
    const prev = (el as HTMLElement).getAttribute("style") || "";
    (el as HTMLElement).setAttribute("style", prev ? `${prev};${styles}` : styles);
  }

  // Eliminar <style> embebidos
  root.querySelectorAll("style").forEach(s => s.remove());

  // ── Header principal ──────────────────────────────────────────────────────
  root.querySelectorAll(".rpt-header").forEach(el =>
    set(el, "background-color:#1e3a6e;color:#ffffff;padding:14px 20px;display:block;margin-bottom:14px"));
  root.querySelectorAll(".rpt-logo-circle").forEach(el => el.remove());
  root.querySelectorAll(".rpt-header-stripe").forEach(el => el.remove());
  root.querySelectorAll(".rpt-header-body,.rpt-header-left").forEach(el =>
    set(el, "display:block"));
  root.querySelectorAll(".rpt-header-right").forEach(el =>
    set(el, "display:block;margin-top:6px;text-align:right"));
  root.querySelectorAll(".rpt-empresa").forEach(el =>
    set(el, "font-size:13px;font-weight:bold;color:#ffffff;display:block"));
  root.querySelectorAll(".rpt-tipo").forEach(el =>
    set(el, "font-size:10px;color:#b0c4e8;display:block;margin-top:2px"));
  root.querySelectorAll(".rpt-subtitulo-label").forEach(el =>
    set(el, "font-size:9px;color:#b0c4e8;display:block"));
  root.querySelectorAll(".rpt-subtitulo").forEach(el =>
    set(el, "font-size:14px;font-weight:bold;color:#ffffff;display:block"));

  // ── Títulos de sección ────────────────────────────────────────────────────
  root.querySelectorAll(".rpt-section-title").forEach(el =>
    set(el, "font-size:12px;font-weight:bold;color:#1e3a6e;display:block;margin-top:14px;margin-bottom:6px;padding:4px 0;border-bottom:1px solid #c5d9f1"));
  root.querySelectorAll(".rpt-section-num").forEach(el =>
    set(el, "background-color:#1e3a6e;color:#ffffff;font-weight:bold;padding:2px 6px;margin-right:6px;font-size:10px"));

  // ── Tablas de datos ───────────────────────────────────────────────────────
  root.querySelectorAll(".data-table th").forEach(el =>
    set(el, "background-color:#dbe9f8;color:#1f3a5f;font-weight:bold;padding:5px 7px;border:1px solid #c5d9f1;font-size:10px;text-align:left"));

  // td base — PREPEND para que clases especiales puedan sobreescribir
  root.querySelectorAll(".data-table td").forEach(el => {
    const e = el as HTMLElement;
    const existing = e.getAttribute("style") || "";
    e.setAttribute("style",
      `border:1px solid #e5ecf8;padding:4px 7px;text-align:right;color:#1f2d3d;font-size:10px${existing ? ";" + existing : ""}`);
  });

  // Clases especiales de celda — APPEND para ganar
  root.querySelectorAll(".data-table td.label").forEach(el =>
    add(el, "text-align:left;font-weight:500"));
  root.querySelectorAll(".data-table td.num").forEach(el =>
    add(el, "text-align:right"));
  root.querySelectorAll(".data-table td.hi").forEach(el =>
    add(el, "background-color:#fff8e1;color:#7c5500"));
  root.querySelectorAll(".data-table td.warn").forEach(el =>
    add(el, "background-color:#fff0f0;color:#c00000"));
  root.querySelectorAll(".data-table td.fuel-ahorro").forEach(el =>
    add(el, "color:#15803d;font-weight:bold"));

  root.querySelectorAll(".data-table tr.rpt-row-total td").forEach(el =>
    add(el, "background-color:#e8f0fb;font-weight:bold"));
  root.querySelectorAll(".data-table tr.rpt-row-grupo td").forEach(el =>
    add(el, "background-color:#f0f5fe;font-weight:bold;color:#3b5b8c"));
  root.querySelectorAll(".data-table tr.rpt-row-grand td").forEach(el =>
    add(el, "background-color:#dbe9f8;font-weight:bold;color:#1f3a5f"));

  // Filas alternas (simulación de :nth-child)
  root.querySelectorAll(".data-table tbody tr").forEach((row, i) => {
    if (i % 2 === 1) {
      const special = (row as HTMLElement).classList.contains("rpt-row-total")
        || (row as HTMLElement).classList.contains("rpt-row-grupo")
        || (row as HTMLElement).classList.contains("rpt-row-grand");
      if (!special) {
        (row as HTMLElement).querySelectorAll("td").forEach(td => {
          const s = td.getAttribute("style") || "";
          if (!s.includes("background-color")) add(td, "background-color:#f7fbff");
        });
      }
    }
  });

  // ── KPI table ─────────────────────────────────────────────────────────────
  root.querySelectorAll(".rpt-kpi-row td").forEach(el =>
    set(el, "background-color:#f0f5ff;border:1px solid #c5d9f1;padding:8px 10px;text-align:center;vertical-align:top"));
  root.querySelectorAll(".rpt-kpi-label").forEach(el =>
    set(el, "font-size:9px;color:#64748b;display:block;margin-bottom:3px"));
  root.querySelectorAll(".rpt-kpi-big").forEach(el =>
    set(el, "font-size:16px;font-weight:bold;color:#1e3a6e;display:block;line-height:1.2"));
  root.querySelectorAll(".rpt-kpi-unit").forEach(el =>
    set(el, "font-size:9px;color:#64748b"));
  root.querySelectorAll(".rpt-kpi-sub").forEach(el =>
    set(el, "font-size:9px;color:#94a3b8;display:block;margin-top:2px"));
  root.querySelectorAll(".rpt-kpi-inline").forEach(el =>
    set(el, "font-size:10px;color:#334155;margin:4px 0;display:block"));
  root.querySelectorAll(".rpt-kpi-val").forEach(el =>
    set(el, "font-weight:bold;color:#1e3a6e"));

  // ── Cajas de combustible, avisos y pie ────────────────────────────────────
  root.querySelectorAll(".rpt-fuel-box").forEach(el =>
    set(el, "border:1px solid #dbe9f8;padding:10px 12px;margin:8px 0;background-color:#f8fbff;display:block"));
  root.querySelectorAll(".rpt-fuel-header").forEach(el =>
    set(el, "margin-bottom:6px;padding-bottom:5px;border-bottom:1px solid #c5d9f1;display:block"));
  root.querySelectorAll(".rpt-fuel-title").forEach(el =>
    set(el, "font-size:11px;font-weight:bold;color:#1e3a6e;display:block"));
  root.querySelectorAll(".rpt-fuel-causa").forEach(el =>
    set(el, "font-size:10px;color:#475569;margin:4px 0;display:block"));
  root.querySelectorAll(".rpt-notice").forEach(el =>
    set(el, "background-color:#fff8e1;border:1px solid #f59e0b;padding:8px;font-size:10px;margin:6px 0;display:block"));
  root.querySelectorAll(".rpt-muted").forEach(el =>
    set(el, "color:#64748b;font-size:9px;display:block;margin-top:3px"));
  root.querySelectorAll(".rpt-footer").forEach(el =>
    set(el, "border-top:1px solid #e5ecf8;padding-top:6px;margin-top:10px;font-size:9px;color:#94a3b8;text-align:center;display:block"));

  return root.innerHTML;
}

// ─── API PÚBLICA ───────────────────────────────────────────────────────────────

/**
 * Exporta un informe (diario / mensual / facturación) a PDF A4 portrait.
 *
 * Usa pdfmake (texto vectorial) vía html-to-pdfmake. El HTML generado se
 * post-procesa con injectInlineStyles para que html-to-pdfmake aplique
 * colores, bordes y tipografía correctamente sin CSS externo.
 */
export async function exportReportPDF(
  element:  HTMLElement,
  filename: string,
): Promise<void> {
  const [pm, htmlToPdfmake] = await Promise.all([
    loadPdfMake(),
    loadHtmlToPdfmake(),
  ]);

  const styledHtml = injectInlineStyles(element.innerHTML);
  const content    = htmlToPdfmake(styledHtml, { window });

  const docDef: any = {
    pageSize:        "A4",
    pageOrientation: "portrait",
    pageMargins:     [22, 22, 22, 22],  // ≈ 8 mm
    content,
    defaultStyle: {
      font:       "Roboto",
      fontSize:   10,
      lineHeight: 1.3,
      color:      "#1f2d3d",
    },
    styles: {
      "html-p":      { margin: [0, 2, 0, 2] },
      "html-strong": { bold: true },
      "html-em":     { italics: true },
    },
  };

  pm.createPdf(docDef).download(filename);
}

/**
 * Exporta las métricas a PDF A4 portrait con pdfmake.
 *
 * Cada sección [data-pdf-section] se captura como imagen JPEG con
 * html2canvas y se coloca en UNA sola hoja A4 portrait sin cortes.
 * Si la imagen supera el área útil, fit la escala proporcionalmente.
 */
export async function exportMetricsPDF(
  container: HTMLElement,
  period:    string,
): Promise<void> {
  const [pm, h2c] = await Promise.all([loadPdfMake(), loadH2C()]);

  const sections = Array.from(
    container.querySelectorAll<HTMLElement>("[data-pdf-section]"),
  );

  // Captura secuencial (html2canvas es más fiable sin concurrencia)
  const images: string[] = [];
  for (const sec of sections) {
    const canvas = await h2c(sec, {
      scale:           RENDER_SCALE,
      useCORS:         true,
      allowTaint:      false,
      logging:         false,
      backgroundColor: "#ffffff",
      windowWidth:     METRICS_WIN_W,
    }) as HTMLCanvasElement;
    images.push(canvas.toDataURL("image/jpeg", JPEG_QUALITY));
  }

  // A4 portrait útil: 194mm × 281mm ≈ 549pt × 795pt
  const contentW = 549;
  const contentH = 795;

  const content: any[] = images.map((img, i) => ({
    image:     img,
    fit:       [contentW, contentH],
    alignment: "center",
    ...(i > 0 ? { pageBreak: "before" } : {}),
  }));

  pm.createPdf({
    pageSize:        "A4",
    pageOrientation: "portrait",
    pageMargins:     [22, 22, 22, 22],
    content,
  }).download(`Metricas_ElMorro_${period}.pdf`);
}
