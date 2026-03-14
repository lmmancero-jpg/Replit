import type { Express } from "express";
import type { Server } from "http";
import { storage } from "./storage";
import { api } from "@shared/routes";
import { z } from "zod";
import puppeteer from "puppeteer-core";
import chromium from "@sparticuz/chromium";

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
          field: err.errors[0].path.join('.'),
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
    const { html, filename } = req.body as { html?: string; filename?: string };
    if (!html || typeof html !== "string") {
      return res.status(400).json({ error: "Missing html content" });
    }
    let browser;
    try {
      browser = await puppeteer.launch({
        args: chromium.args,
        executablePath: await chromium.executablePath(),
        headless: true,
      });
      const page = await browser.newPage();
      await page.setContent(html, { waitUntil: "networkidle0", timeout: 30000 });
      const pdfBuffer = await page.pdf({
        format: "A4",
        printBackground: true,
        margin: { top: "15mm", bottom: "15mm", left: "15mm", right: "15mm" },
      });
      const safeFilename = (filename || "reporte.pdf").replace(/[^a-zA-Z0-9_.\-]/g, "_");
      res.setHeader("Content-Type", "application/pdf");
      res.setHeader("Content-Disposition", `attachment; filename="${safeFilename}"`);
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
