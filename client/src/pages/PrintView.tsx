import { useEffect, useState } from "react";
import { Printer, Download, X, AlertCircle } from "lucide-react";

export default function PrintView() {
  const params = new URLSearchParams(window.location.search);
  const key   = params.get("key") ?? "";
  const title = params.get("title") ?? "Informe – Central El Morro";

  const [html,      setHtml]      = useState("");
  const [status,    setStatus]    = useState<"loading" | "ready" | "error">("loading");
  const [pdfStatus, setPdfStatus] = useState<"idle" | "loading" | "error">("idle");

  useEffect(() => {
    if (!key) { setStatus("error"); return; }
    const stored = sessionStorage.getItem(key);
    if (!stored) { setStatus("error"); return; }
    setHtml(stored);
    setStatus("ready");
    if (params.get("autoprint") === "1") {
      setTimeout(() => window.print(), 900);
    }
  }, []);

  const handlePrint = () => window.print();

  const handleDownloadPDF = async () => {
    const serviceUrl = (import.meta.env.VITE_PDF_SERVICE_URL ?? "").trim() as string;
    if (!serviceUrl) { handlePrint(); return; }

    setPdfStatus("loading");
    try {
      const printUrl = window.location.href.replace(/&?autoprint=1/, "");
      const res = await fetch(serviceUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ url: printUrl }),
      });
      if (!res.ok) throw new Error(`${res.status}`);
      const blob = await res.blob();
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement("a");
      a.href     = url;
      a.download = `${title.replace(/[^\w\s\-]/g, "_").trim()}.pdf`;
      a.click();
      URL.revokeObjectURL(url);
      setPdfStatus("idle");
    } catch {
      setPdfStatus("error");
      handlePrint();
    }
  };

  return (
    <div className="print-shell">
      {/* ── Barra de control (oculta al imprimir) ── */}
      <div className="print-toolbar no-print">
        <div className="print-toolbar-inner">
          <span className="print-toolbar-title">{title}</span>
          <div className="print-toolbar-actions">
            <button
              data-testid="button-print"
              onClick={handlePrint}
              className="print-btn print-btn-outline"
            >
              <Printer size={14} /> Imprimir
            </button>
            <button
              data-testid="button-download-pdf"
              onClick={handleDownloadPDF}
              disabled={pdfStatus === "loading"}
              className="print-btn print-btn-primary"
            >
              <Download size={14} />
              {pdfStatus === "loading" ? "Generando…" : "Descargar PDF"}
            </button>
            <button
              onClick={() => window.close()}
              className="print-btn print-btn-ghost"
              title="Cerrar"
            >
              <X size={14} />
            </button>
          </div>
        </div>
        {pdfStatus === "error" && (
          <p className="print-toolbar-warning">
            El servicio PDF no respondió. Usando impresión del navegador como alternativa.
          </p>
        )}
      </div>

      {/* ── Contenido ── */}
      <div className="print-page-wrap">
        {status === "loading" && (
          <div className="print-status">Cargando informe…</div>
        )}

        {status === "error" && (
          <div className="print-error">
            <AlertCircle size={40} className="mb-3 opacity-60" />
            <h2>No se encontró el informe</h2>
            <p>
              La sesión puede haber expirado. Genera el informe nuevamente
              desde el panel principal y usa el botón "Vista de Impresión".
            </p>
          </div>
        )}

        {status === "ready" && (
          <div
            className="print-report report-content"
            dangerouslySetInnerHTML={{ __html: html }}
          />
        )}
      </div>
    </div>
  );
}
