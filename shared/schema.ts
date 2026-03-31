import { pgTable, text, serial, timestamp, numeric } from "drizzle-orm/pg-core";
import { createInsertSchema } from "drizzle-zod";
import { z } from "zod";

// ── REPORTS ──────────────────────────────────────────────────────────────────

export const reports = pgTable("reports", {
  id: serial("id").primaryKey(),
  title: text("title").notNull(),
  reportType: text("type").notNull(),
  date: text("date").notNull(),
  content: text("content").notNull(),
  createdAt: timestamp("created_at").defaultNow(),
});

export const insertReportSchema = createInsertSchema(reports).omit({ id: true, createdAt: true });

export type Report = typeof reports.$inferSelect;
export type InsertReport = z.infer<typeof insertReportSchema>;

// ── INVOICES ─────────────────────────────────────────────────────────────────

export const INVOICE_CATEGORIES = [
  "combustible_transporte",
  "lubricantes_quimicos",
  "agua_insumos",
  "repuestos_predictivo",
  "impacto_ambiental",
  "servicios_auxiliares",
] as const;

export type InvoiceCategory = typeof INVOICE_CATEGORIES[number];

export const INVOICE_CATEGORY_LABELS: Record<InvoiceCategory, string> = {
  combustible_transporte: "Combustible + Transporte",
  lubricantes_quimicos:   "Lubricantes + Químicos",
  agua_insumos:           "Agua + Insumos",
  repuestos_predictivo:   "Repuestos Mantenimiento Predictivo",
  impacto_ambiental:      "Impacto Ambiental",
  servicios_auxiliares:   "Servicios Auxiliares",
};

export const invoices = pgTable("invoices", {
  id: serial("id").primaryKey(),
  period: text("period").notNull(),
  issueDate: text("issue_date").notNull(),
  supplier: text("supplier").notNull(),
  invoiceNumber: text("invoice_number").notNull(),
  category: text("category").notNull(),
  description: text("description").notNull().default(""),
  subtotal: numeric("subtotal").notNull().default("0"),
  iva: numeric("iva").notNull().default("0"),
  total: numeric("total").notNull().default("0"),
  lineItems: text("line_items").default("[]"),
  observations: text("observations").default(""),
  createdAt: timestamp("created_at").defaultNow(),
});

export const insertInvoiceSchema = createInsertSchema(invoices).omit({ id: true, createdAt: true });

export type Invoice = typeof invoices.$inferSelect;
export type InsertInvoice = z.infer<typeof insertInvoiceSchema>;
