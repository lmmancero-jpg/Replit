# Central El Morro — Power Plant Operations (NEXUS)

## Overview

NEXUS is a full-stack web application for **Central El Morro – Morro Energy S.A.**, a power plant operations center. It automates post-operative reports (daily, monthly) and has a complete billing module with invoice CRUD and two billing report types. Users upload Excel production files, configure operational parameters, and the app computes energy/fuel/hours/billing data and generates formatted HTML reports exported as PDF via Puppeteer.

Named "NEXUS — Power Plant Ops" in the UI. Originally a standalone HTML/JS tool rebuilt as a React + Express app.

---

## User Preferences

Preferred communication style: Simple, everyday language.

---

## System Architecture

### Full-Stack Monorepo

- `client/` — React frontend (Vite, TypeScript)
- `server/` — Express backend (Node.js, TypeScript)
- `shared/` — Shared types, schema, and route definitions

### Frontend Pages (wouter routing)

| Route | Page | Description |
|---|---|---|
| `/` | Generator | Upload Excel files, configure parameters, generate daily/monthly reports |
| `/invoices` | Facturas | Full CRUD for monthly supplier invoices |
| `/billing` | Facturación | Generate client-billing or real-cost billing reports |
| `/metrics` | Métricas | Charts and KPIs from production data |

### Key Libraries

- **UI:** shadcn/ui (Radix + Tailwind), lucide-react icons
- **State:** TanStack Query v5
- **Forms:** React Hook Form + Zod
- **Routing:** wouter
- **Charts:** Recharts (metrics page)
- **Excel:** xlsx (SheetJS) — parsed entirely on the frontend
- **PDF:** Puppeteer (backend) via POST /api/export/pdf; browser print window as fallback

### Report Engine (`client/src/lib/reportEngine.ts`)

Core business logic (1700+ lines):
- Parses `.xlsx` files client-side with SheetJS
- Reads column-mapped data (energy kWh, fuel HFO+DO, tank stocks, horómetros per unit)
- Exports: `generarInformeDiario`, `generarInformeMensual`, `generarInformeFacturacion`
- Also exports constants and helpers: `CONFIG`, `COSTOS_VARIABLES`, `COSTO_FIJO_MENSUAL_POR_UNIDAD`, `CBMT_U1_MENSUAL`, `P_CONTR_LANEC`, `P_CONTR_GRACA`, `P_CONTR_TOT`, `posNum`, `fmt`, `excelDateKey`, `getDaysInMonth`, `getMesNombreES`, `getProdSheetAndRows`, `rptHeader`, `seccion`
- Fuel analysis: HFO and DO separated per unit, compared to 90-day reference
- Section 6 ("Síntesis Operativa"): replaces IDOM table with a 6-row executive table

### Billing Engine (`client/src/lib/billingEngine.ts`)

- `buildClientBillingReport`: invoice combustible_transporte total ÷ kWh as adjusted CV; other items at contractual rate
- `buildRealBillingReport`: all invoice totals as real CVs + contractual margen_variable
- `getMonthlyProductionSummary`: reads production workbook for a given YYYY-MM period

### Invoice Excel (`client/src/lib/invoiceExcel.ts`)

- `exportInvoicesExcel`: exports invoice list + summary to .xlsx via SheetJS

### Hooks

- `client/src/hooks/use-invoices.ts`: `useInvoices`, `useInvoiceSummary`, `useCreateInvoice`, `useUpdateInvoice`, `useDeleteInvoice`
- `client/src/hooks/use-reports.ts`: report list CRUD

### File Store (`client/src/lib/fileStore.tsx`)

- React context holding the loaded production workbook (`wbProd`) and aforo file
- Shared across Generator, Metrics, Invoices, and Billing pages

### PDF Export

- **Backend Puppeteer** (primary): POST `/api/export/pdf` → `buildReportPuppeteerHTML` wraps HTML with `PRINT_CSS` → puppeteer-core + chromium → A4 PDF binary
- **Browser fallback**: `openPrintWindow` in `client/src/lib/printPDF.ts` opens a print-optimized window with embedded CSS
- Metrics PDF: `exportMetricsPDF` in `client/src/lib/pdfExporter.ts` uses html2canvas + pdfmake

### Backend (`server/routes.ts`)

API endpoints:
- `GET/POST /api/reports` — report history CRUD
- `GET /api/reports/:id`, `DELETE /api/reports/:id`
- `GET /api/invoices?period=YYYY-MM` — list invoices for period
- `GET /api/invoices/summary?period=YYYY-MM` — totals by category
- `POST /api/invoices` — create invoice
- `PUT /api/invoices/:id` — update invoice
- `DELETE /api/invoices/:id` — delete invoice
- `POST /api/export/pdf` — Puppeteer PDF generation

### Database (`shared/schema.ts`)

Two tables:
1. `reports`: id, title, reportType, date, content (HTML), createdAt
2. `invoices`: id, period (YYYY-MM), issueDate, supplier, invoiceNumber, category (enum), description, subtotal, iva, total, observations, createdAt

`INVOICE_CATEGORIES`: combustible_transporte, lubricantes_quimicos, agua_insumos, repuestos_predictivo, impacto_ambiental, servicios_auxiliares

### CSS

- `client/src/index.css`: screen view report styles (.report-content) with large fonts (tables 16px, KPI big 28px, section titles 22px) + @media print styles
- `server/routes.ts` → `PRINT_CSS`: compact PDF-optimized styles (tables 11px) injected by Puppeteer
- `client/src/lib/printPDF.ts` → `REPORT_PRINT_CSS`: same compact CSS for browser print fallback

### Downtime Persistence

Generator page writes `localStorage.setItem("nexus_u1dt", ...)` / `nexus_u2dt` whenever the user changes downtime values. Billing page reads these on mount as default values.

---

## Environment Variables Required

| Variable | Description |
|---|---|
| `DATABASE_URL` | Full PostgreSQL connection string |
| `SESSION_SECRET` | Session secret (used by Express session middleware) |

---

## Build

- `npm run dev` — tsx server/index.ts + Vite middleware (port 5000)
- `npm run build` — esbuild + Vite production build
- `npm run db:push` — sync Drizzle schema to PostgreSQL
