import type { Express } from "express";
import type { Server } from "http";
import { storage } from "./storage";
import { api } from "@shared/routes";
import { z } from "zod";
import puppeteer from "puppeteer-core";
import { execSync } from "child_process";

function findChromiumExecutable(): string {
  try {
    return execSync("which chromium", { timeout: 5000 }).toString().trim();
  } catch {
    const candidates = [
      "/usr/bin/chromium",
      "/usr/bin/chromium-browser",
      "/usr/bin/google-chrome",
    ];
    for (const p of candidates) {
      try {
        execSync(`test -x ${p}`, { timeout: 2000 });
        return p;
      } catch { /* skip */ }
    }
    throw new Error("Chromium not found. Install chromium via system dependencies.");
  }
}

const PRINT_CSS = `
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

@page {
  size: A4 portrait;
  margin: 14mm 16mm;
}

body {
  font-family: "Segoe UI", system-ui, -apple-system, Arial, sans-serif;
  font-size: 12px;
  line-height: 1.5;
  color: #1b2134;
  background: #fff;
  -webkit-print-color-adjust: exact;
  print-color-adjust: exact;
  color-adjust: exact;
}

.report-content {
  font-family: "Segoe UI", system-ui, -apple-system, Arial, sans-serif;
  font-size: 12px;
  line-height: 1.5;
  color: #1b2134;
}

/* HEADER */
.report-content .rpt-header {
  border-radius: 6px;
  overflow: hidden;
  margin-bottom: 18px;
  break-inside: avoid;
  page-break-inside: avoid;
}
.report-content .rpt-header-body {
  background: linear-gradient(120deg, #1b3b6f 0%, #1260a8 55%, #0e7fd4 100%);
  -webkit-print-color-adjust: exact;
  print-color-adjust: exact;
  color: #fff;
  padding: 14px 18px;
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 14px;
  flex-wrap: wrap;
}
.report-content .rpt-header-left {
  display: flex;
  align-items: center;
  gap: 10px;
}
.report-content .rpt-logo-circle {
  width: 40px;
  height: 40px;
  background: rgba(255,255,255,0.12);
  border: 1.5px solid rgba(255,255,255,0.3);
  border-radius: 50%;
  display: flex;
  align-items: center;
  justify-content: center;
  flex-shrink: 0;
  -webkit-print-color-adjust: exact;
  print-color-adjust: exact;
}
.report-content .rpt-empresa {
  font-size: 17px;
  font-weight: 700;
  letter-spacing: 0.02em;
  line-height: 1.2;
}
.report-content .rpt-tipo {
  font-size: 12px;
  font-weight: 400;
  opacity: 0.85;
  text-transform: uppercase;
  letter-spacing: 0.07em;
  margin-top: 2px;
}
.report-content .rpt-header-right { text-align: right; }
.report-content .rpt-subtitulo-label {
  font-size: 10px;
  text-transform: uppercase;
  letter-spacing: 0.07em;
  opacity: 0.7;
  margin-bottom: 2px;
}
.report-content .rpt-subtitulo {
  font-size: 18px;
  font-weight: 600;
}
.report-content .rpt-header-stripe {
  height: 4px;
  background: linear-gradient(90deg, #facc15, #fb923c, #f87171, #c084fc);
  -webkit-print-color-adjust: exact;
  print-color-adjust: exact;
}

/* SECTIONS */
.report-content .rpt-section-title {
  display: flex;
  align-items: center;
  gap: 10px;
  margin: 18px 0 7px;
  font-size: 12px;
  font-weight: 700;
  color: #1b3b6f;
  text-transform: uppercase;
  letter-spacing: 0.04em;
  border-left: 4px solid #1260a8;
  padding: 5px 0 5px 11px;
  line-height: 1.2;
  break-after: avoid;
  page-break-after: avoid;
}
.report-content .rpt-section-num {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-width: 20px;
  height: 20px;
  background: #1260a8;
  -webkit-print-color-adjust: exact;
  print-color-adjust: exact;
  color: #fff;
  font-size: 11px;
  font-weight: 700;
  border-radius: 3px;
  flex-shrink: 0;
  padding: 0 4px;
}
.report-content .rpt-section-label { display: inline; }

/* TABLES */
.report-content .data-table {
  width: 100%;
  border-collapse: collapse;
  margin: 3px 0 12px;
  font-size: 11px;
  break-inside: avoid;
  page-break-inside: avoid;
}
.report-content .data-table thead tr {
  background: linear-gradient(90deg, #dce8fa, #eff4fd);
  -webkit-print-color-adjust: exact;
  print-color-adjust: exact;
}
.report-content .data-table th {
  border: 1px solid #c8d6ee;
  padding: 7px 11px;
  text-align: left;
  font-weight: 700;
  font-size: 10px;
  color: #1b3b6f;
  white-space: normal;
  vertical-align: middle;
  line-height: 1.3;
}
.report-content .data-table td {
  border: 1px solid #e5ecf8;
  padding: 5px 11px;
  text-align: right;
  color: #1f2d3d;
  font-size: 11px;
  vertical-align: middle;
  line-height: 1.3;
}
.report-content .data-table tbody tr:nth-child(even) {
  background: #f8faff;
  -webkit-print-color-adjust: exact;
  print-color-adjust: exact;
}
.report-content .data-table td.label {
  text-align: left;
  color: #374151;
  font-weight: 500;
}
.report-content .data-table td.num {
  font-variant-numeric: tabular-nums;
  font-weight: 400;
}
.report-content .data-table td.hi { font-weight: 700; color: #1b3b6f; }
.report-content .data-table td.warn { color: #b45309; font-weight: 600; }
.report-content .data-table td.fuel-ahorro { color: #15803d; font-weight: 600; }

.report-content .data-table tr.rpt-row-total td {
  background: #e8f0fd;
  -webkit-print-color-adjust: exact;
  print-color-adjust: exact;
  font-weight: 600;
  border-top: 1.5px solid #b8cce8;
}
.report-content .data-table tr.rpt-row-grupo td {
  background: #f0f4fb;
  -webkit-print-color-adjust: exact;
  print-color-adjust: exact;
  font-size: 10px;
  font-weight: 700;
  color: #1260a8;
  text-transform: uppercase;
  letter-spacing: 0.04em;
  border-top: 1.5px solid #c8d6ee;
}
.report-content .data-table tr.rpt-row-grand td {
  background: linear-gradient(90deg, #dce8fa, #e8f4fd);
  -webkit-print-color-adjust: exact;
  print-color-adjust: exact;
  font-size: 11px;
  font-weight: 700;
  color: #1b3b6f;
  border-top: 2px solid #1260a8;
}

/* KPIs */
.report-content .rpt-kpi-row {
  display: flex;
  gap: 7px;
  margin: 5px 0 9px;
  flex-wrap: wrap;
  break-inside: avoid;
  page-break-inside: avoid;
}
.report-content .rpt-kpi-card {
  flex: 1;
  min-width: 80px;
  background: linear-gradient(135deg, #f0f6ff, #e8f0fd);
  -webkit-print-color-adjust: exact;
  print-color-adjust: exact;
  border: 1px solid #c8d6ee;
  border-radius: 6px;
  padding: 6px 9px;
}
.report-content .rpt-kpi-label {
  font-size: 9px;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 0.05em;
  color: #4b6a9b;
  margin-bottom: 3px;
}
.report-content .rpt-kpi-big {
  font-size: 24px;
  font-weight: 700;
  color: #1b3b6f;
  line-height: 1.1;
}
.report-content .rpt-kpi-unit { font-size: 12px; font-weight: 600; color: #4b6a9b; margin-left: 2px; }
.report-content .rpt-kpi-sub { font-size: 10px; color: #6b82a8; margin-top: 2px; }
.report-content .rpt-kpi-inline { font-size: 12px; color: #374151; margin: 3px 0 9px; }
.report-content .rpt-kpi-val { font-weight: 700; color: #1b3b6f; font-size: 13px; }

/* FUEL BOX */
.report-content .rpt-fuel-box {
  background: #f9fbff;
  -webkit-print-color-adjust: exact;
  print-color-adjust: exact;
  border: 1px solid #c8d6ee;
  border-left: 4px solid #1260a8;
  border-radius: 5px;
  padding: 7px 10px;
  margin: 3px 0 9px;
  break-inside: avoid;
  page-break-inside: avoid;
}
.report-content .rpt-fuel-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 5px;
  flex-wrap: wrap;
  gap: 4px;
}
.report-content .rpt-fuel-title {
  font-size: 11px;
  font-weight: 700;
  color: #1b3b6f;
  text-transform: uppercase;
  letter-spacing: 0.04em;
}
.report-content .rpt-fuel-causa {
  font-size: 11px;
  color: #374151;
  margin-bottom: 6px;
  line-height: 1.4;
}

/* IDOM / SINDESIS */
.report-content .rpt-idom-box {
  background: linear-gradient(135deg, #f0f6ff, #e8f0fd);
  -webkit-print-color-adjust: exact;
  print-color-adjust: exact;
  border: 1px solid #c8d6ee;
  border-radius: 6px;
  padding: 10px 14px;
  margin: 4px 0 10px;
  break-inside: avoid;
  page-break-inside: avoid;
}
.report-content .rpt-idom-score {
  font-size: 20px;
  font-weight: 800;
  color: #1b3b6f;
}
.report-content .rpt-idom-conclusion {
  font-size: 11px;
  color: #374151;
  border-top: 1px solid #c8d6ee;
  margin-top: 7px;
  padding-top: 6px;
}

/* MISC */
.report-content .rpt-obs {
  font-size: 12px;
  color: #374151;
  background: #f9fafb;
  border-left: 3px solid #9ca3af;
  padding: 8px 12px;
  border-radius: 4px;
  line-height: 1.55;
}
.report-content .rpt-obs-empty { color: #6b7280; font-style: italic; }
.report-content .rpt-muted { font-size: 10px; color: #6b7280; margin: 3px 0; }
.report-content .rpt-notice {
  padding: 8px 12px;
  border-radius: 5px;
  font-size: 11px;
  margin: 5px 0 10px;
}
.report-content .rpt-notice-warn { background: #fef9c3; border: 1px solid #fde047; color: #713f12; }
.report-content .rpt-notice-error { background: #fee2e2; border: 1px solid #fca5a5; color: #7f1d1d; }

/* HORÓMETROS */
.report-content .rpt-horom-row {
  display: flex;
  gap: 8px;
  margin: 5px 0 10px;
  flex-wrap: wrap;
  break-inside: avoid;
  page-break-inside: avoid;
}
.report-content .rpt-horom-card {
  flex: 1;
  min-width: 140px;
  background: linear-gradient(135deg, #f0f6ff, #e8f0fd);
  -webkit-print-color-adjust: exact;
  print-color-adjust: exact;
  border: 1px solid #c8d6ee;
  border-radius: 6px;
  padding: 6px 10px;
}
.report-content .rpt-horom-label { font-size: 9px; font-weight: 700; text-transform: uppercase; color: #4b6a9b; margin-bottom: 2px; }
.report-content .rpt-horom-big { font-size: 18px; font-weight: 700; color: #1b3b6f; }
.report-content .rpt-horom-unit { font-size: 11px; color: #4b6a9b; margin-left: 2px; }
.report-content .rpt-horom-sub { font-size: 10px; color: #6b82a8; margin-top: 1px; }
.report-content .rpt-horom-warn { color: #b45309 !important; }
.report-content .rpt-horom-ok { color: #15803d !important; }
`;

function buildReportPuppeteerHTML(reportHtml: string, title = "Reporte"): string {
  return `<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${title}</title>
  <style>${PRINT_CSS}</style>
</head>
<body>
  <div class="report-content">
    ${reportHtml}
  </div>
</body>
</html>`;
}

export async function registerRoutes(
  httpServer: Server,
  app: Express
): Promise<Server> {

  app.get(api.reports.list.path, async (req, res) => {
    try {
      const allReports = await storage.getReports();
      res.status(200).json(allReports);
    } catch (err) {
      console.error(err);
      res.status(500).json({ message: "Internal server error" });
    }
  });

  app.get(api.reports.get.path, async (req, res) => {
    try {
      const id = Number(req.params.id);
      if (isNaN(id)) {
        return res.status(400).json({ message: "Invalid ID" });
      }
      const report = await storage.getReport(id);
      if (!report) {
        return res.status(404).json({ message: "Report not found" });
      }
      res.status(200).json(report);
    } catch (err) {
      res.status(500).json({ message: "Internal server error" });
    }
  });

  app.post(api.reports.create.path, async (req, res) => {
    try {
      const input = api.reports.create.input.parse(req.body);
      const report = await storage.createReport(input);
      res.status(201).json(report);
    } catch (err) {
      if (err instanceof z.ZodError) {
        return res.status(400).json({
          message: err.errors[0].message,
          field: err.errors[0].path.join("."),
        });
      }
      res.status(500).json({ message: "Internal server error" });
    }
  });

  app.delete(api.reports.delete.path, async (req, res) => {
    try {
      const id = Number(req.params.id);
      if (isNaN(id)) {
        return res.status(400).json({ message: "Invalid ID" });
      }
      const report = await storage.getReport(id);
      if (!report) {
        return res.status(404).json({ message: "Report not found" });
      }
      await storage.deleteReport(id);
      res.status(204).send();
    } catch (err) {
      res.status(500).json({ message: "Internal server error" });
    }
  });

  app.post("/api/export/pdf", async (req, res) => {
    const { html, title } = req.body as { html?: string; title?: string };
    if (!html || typeof html !== "string") {
      return res.status(400).json({ error: "Missing html content" });
    }

    const fullHtml = buildReportPuppeteerHTML(html, title || "Reporte");

    let browser;
    try {
      const executablePath = findChromiumExecutable();
      browser = await puppeteer.launch({
        executablePath,
        headless: true,
        args: [
          "--no-sandbox",
          "--disable-setuid-sandbox",
          "--disable-dev-shm-usage",
          "--disable-gpu",
          "--no-first-run",
          "--no-zygote",
          "--single-process",
          "--disable-extensions",
        ],
      });

      const page = await browser.newPage();
      await page.setViewport({ width: 1123, height: 1587 });
      await page.setContent(fullHtml, { waitUntil: "networkidle0", timeout: 30000 });

      const pdfBuffer = await page.pdf({
        format: "A4",
        printBackground: true,
        margin: { top: "14mm", bottom: "14mm", left: "16mm", right: "16mm" },
        displayHeaderFooter: false,
        preferCSSPageSize: true,
      });

      const safeFilename = (title || "reporte.pdf")
        .replace(/[^a-zA-Z0-9_.\- ]/g, "_")
        .trim();
      res.setHeader("Content-Type", "application/pdf");
      res.setHeader(
        "Content-Disposition",
        `attachment; filename="${safeFilename}"`
      );
      res.end(pdfBuffer);
    } catch (err) {
      console.error("Puppeteer PDF error:", err);
      res.status(500).json({ error: "PDF generation failed", detail: String(err) });
    } finally {
      if (browser) await browser.close();
    }
  });

  return httpServer;
}
