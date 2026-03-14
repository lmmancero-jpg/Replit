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
      await page.setViewport({ width: 1123, height: 794 });
      await page.setContent(html, { waitUntil: "networkidle0", timeout: 30000 });

      const pdfBuffer = await page.pdf({
        format: "A4",
        printBackground: true,
        margin: { top: "14mm", bottom: "14mm", left: "16mm", right: "16mm" },
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
