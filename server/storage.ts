import { db } from "./db";
import {
  reports, invoices,
  type Report, type InsertReport,
  type Invoice, type InsertInvoice,
  INVOICE_CATEGORIES,
} from "@shared/schema";
import { eq, desc, and } from "drizzle-orm";

export type InvoiceSummary = Record<string, number>;

export interface IStorage {
  getReports(): Promise<Report[]>;
  getReport(id: number): Promise<Report | undefined>;
  createReport(report: InsertReport): Promise<Report>;
  deleteReport(id: number): Promise<void>;

  getInvoices(period?: string): Promise<Invoice[]>;
  getInvoice(id: number): Promise<Invoice | undefined>;
  createInvoice(invoice: InsertInvoice): Promise<Invoice>;
  updateInvoice(id: number, invoice: Partial<InsertInvoice>): Promise<Invoice>;
  deleteInvoice(id: number): Promise<void>;
  getInvoiceSummary(period: string): Promise<InvoiceSummary>;
}

export class DatabaseStorage implements IStorage {
  // ── Reports ────────────────────────────────────────────────────────────────

  async getReports(): Promise<Report[]> {
    return await db.select().from(reports).orderBy(desc(reports.createdAt));
  }

  async getReport(id: number): Promise<Report | undefined> {
    const [report] = await db.select().from(reports).where(eq(reports.id, id));
    return report;
  }

  async createReport(insertReport: InsertReport): Promise<Report> {
    const [report] = await db.insert(reports).values(insertReport).returning();
    return report;
  }

  async deleteReport(id: number): Promise<void> {
    await db.delete(reports).where(eq(reports.id, id));
  }

  // ── Invoices ───────────────────────────────────────────────────────────────

  async getInvoices(period?: string): Promise<Invoice[]> {
    if (period) {
      return await db
        .select()
        .from(invoices)
        .where(eq(invoices.period, period))
        .orderBy(invoices.issueDate, invoices.category);
    }
    return await db.select().from(invoices).orderBy(desc(invoices.createdAt));
  }

  async getInvoice(id: number): Promise<Invoice | undefined> {
    const [inv] = await db.select().from(invoices).where(eq(invoices.id, id));
    return inv;
  }

  async createInvoice(invoice: InsertInvoice): Promise<Invoice> {
    const [inv] = await db.insert(invoices).values(invoice).returning();
    return inv;
  }

  async updateInvoice(id: number, invoice: Partial<InsertInvoice>): Promise<Invoice> {
    const [inv] = await db
      .update(invoices)
      .set(invoice)
      .where(eq(invoices.id, id))
      .returning();
    if (!inv) throw new Error(`Invoice ${id} not found`);
    return inv;
  }

  async deleteInvoice(id: number): Promise<void> {
    await db.delete(invoices).where(eq(invoices.id, id));
  }

  async getInvoiceSummary(period: string): Promise<InvoiceSummary> {
    const rows = await db
      .select()
      .from(invoices)
      .where(eq(invoices.period, period));

    const summary: InvoiceSummary = {};
    for (const cat of INVOICE_CATEGORIES) {
      summary[cat] = 0;
    }
    for (const row of rows) {
      const cat = row.category;
      if (cat in summary) {
        summary[cat] = (summary[cat] ?? 0) + parseFloat(row.total ?? "0");
      }
    }
    return summary;
  }
}

export const storage = new DatabaseStorage();
