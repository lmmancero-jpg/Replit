# NEXUS Power Plant Operations — replit.md

## Overview

NEXUS is a web application for **Central El Morro**, a power plant operations center. Its primary purpose is to automate the generation of post-operative reports (daily, monthly, and billing/facturación types) for the plant. Users upload Excel production and tank-aforo files, provide operational parameters (dates, downtime days, observations), and the app processes the data to generate formatted HTML reports. These reports are saved to a PostgreSQL database and can be viewed, searched, and exported to PDF from a history page.

The app is named "NEXUS — Power Plant Ops" in the UI and was originally a standalone HTML/JS tool (see `attached_assets/` and `client/src/lib/original.*`) that has been rebuilt as a full-stack React + Express application.

---

## User Preferences

Preferred communication style: Simple, everyday language.

---

## System Architecture

### Full-Stack Monorepo Structure

The project uses a single-repo layout with three main areas:
- `client/` — React frontend (Vite, TypeScript)
- `server/` — Express backend (Node.js, TypeScript)
- `shared/` — Shared types, schema, and route definitions used by both client and server

This avoids duplication of types and API contracts. The `@shared/*` path alias connects both sides.

### Frontend Architecture

- **Framework:** React 18 with TypeScript, bundled by Vite
- **Routing:** `wouter` (lightweight client-side router) with two main pages:
  - `/` → Generator page: upload files, set parameters, generate reports
  - `/history` → History page: list, search, view, and delete saved reports
- **UI Components:** shadcn/ui (Radix UI primitives + Tailwind CSS). The `components.json` file configures shadcn with "new-york" style. All UI components live in `client/src/components/ui/`.
- **State / Data Fetching:** TanStack Query (React Query v5) for server state. Custom hooks in `client/src/hooks/use-reports.ts` wrap all API calls with Zod validation for type safety.
- **Forms:** React Hook Form + Zod (`@hookform/resolvers/zod`) for the report generator form
- **Styling:** Tailwind CSS with CSS custom properties for theming. Dark mode supported via class strategy. Custom design tokens for sidebar, cards, primary, etc.
- **Layout:** Sidebar layout using shadcn's `Sidebar` component with a collapsible sidebar, sticky header, and a main content area.

### Report Engine (`client/src/lib/reportEngine.ts`)

This is the core business logic. It:
- Parses `.xlsx` files using the `xlsx` (SheetJS) library
- Reads column-mapped data (energy kWh, fuel consumption, tank stocks, horómetros)
- Computes operational KPIs (utilization, availability, billing factors, maintenance hours)
- Generates three report types as HTML strings:
  - `generarInformeDiario` — Daily post-operative report
  - `generarInformeMensual` — Monthly summary report
  - `generarInformeFacturacion` — Billing report with contract factor calculations
- The generated HTML strings are saved to the database and rendered in-browser for preview and PDF export

### Backend Architecture

- **Framework:** Express 5 (Node.js, TypeScript, ESM)
- **Entry point:** `server/index.ts` creates an HTTP server, registers routes, and serves static files (or Vite middleware in dev)
- **Routes:** Defined in `server/routes.ts`, typed against `shared/routes.ts` API definitions:
  - `GET /api/reports` — list all reports (newest first)
  - `GET /api/reports/:id` — get one report
  - `POST /api/reports` — create a report (body validated by Zod)
  - `DELETE /api/reports/:id` — delete a report
- **Storage layer:** `server/storage.ts` defines an `IStorage` interface and a `DatabaseStorage` class that uses Drizzle ORM. This abstraction allows swapping storage implementations.
- **Dev vs Prod:** In development, Vite middleware is mounted on the Express server for HMR. In production, the built static files are served from `dist/public/`.

### Shared API Contract (`shared/routes.ts`)

A single source of truth for all API routes, input schemas (Zod), and response schemas. Both the frontend hooks and backend route handlers import from here, ensuring the client and server never drift out of sync.

### Database

- **Database:** PostgreSQL via `pg` (node-postgres)
- **ORM:** Drizzle ORM with `drizzle-zod` for auto-generating Zod schemas from table definitions
- **Schema** (`shared/schema.ts`): Single `reports` table:
  - `id` (serial PK)
  - `title` (text)
  - `reportType` (text: `'diario'`, `'mensual'`, `'facturacion'`)
  - `date` (text)
  - `content` (text — stores the full HTML string of the generated report)
  - `createdAt` (timestamp, default now)
- **Migrations:** Drizzle Kit, config in `drizzle.config.ts`. Run with `npm run db:push`.
- **Seeding:** `seed.ts` in the root inserts a sample report if the table is empty.

### Build System

- `npm run dev` — runs the backend with `tsx` (TypeScript execute), which also serves Vite in middleware mode
- `npm run build` — runs `script/build.ts` which: (1) builds the client with Vite, (2) bundles the server with esbuild into `dist/index.cjs`
- `npm start` — runs the production build

### PDF Export

PDF export uses **pdfmake 0.2.20** (vector text, no rasterization) combined with **html-to-pdfmake** for report pages, and **html2canvas** for the metrics charts page:

- **Reports** (daily / monthly / billing): `exportReportPDF()` in `client/src/lib/pdfExporter.ts` post-processes the report HTML with `injectInlineStyles()` to embed all colors, borders, and typography as inline `style=""` attributes (so html-to-pdfmake renders correctly without external CSS), then calls pdfmake to produce an A4 portrait PDF download.
- **Metrics**: `exportMetricsPDF()` captures each `[data-pdf-section]` as a JPEG image via html2canvas at 3× scale, then embeds each image in a pdfmake document (one A4 portrait page per section).
- Fonts registered via `pdfMake.addVirtualFileSystem(vfsFonts)` using the bundled Roboto VFS from `pdfmake/build/vfs_fonts`.
- Custom type declarations for pdfmake and html-to-pdfmake are in `client/src/vendor.d.ts`.

---

## External Dependencies

| Dependency | Purpose |
|---|---|
| **PostgreSQL** | Primary data store. Required via `DATABASE_URL` environment variable. |
| **Drizzle ORM** (`drizzle-orm`, `drizzle-kit`, `drizzle-zod`) | Database access, schema management, and Zod schema generation |
| **`pg` (node-postgres)** | PostgreSQL driver for Node.js |
| **`xlsx` (SheetJS)** | Parsing `.xlsx` Excel files uploaded by the user for production and aforo data |
| **`pdfmake` 0.2.20** | Vector-text PDF generation for reports and metrics |
| **`html-to-pdfmake`** | Converts styled HTML into pdfmake content nodes |
| **`html2canvas`** | Rasterizes chart sections for metrics PDF export |
| **`date-fns`** | Date formatting and manipulation throughout the app |
| **TanStack Query v5** | Server state management and caching on the frontend |
| **React Hook Form + Zod** | Form state and validation |
| **wouter** | Lightweight client-side routing |
| **shadcn/ui + Radix UI** | Accessible headless UI primitives |
| **Tailwind CSS** | Utility-first styling |
| **Recharts** | Chart components (included via shadcn's chart component, not yet actively used in pages) |
| **`@replit/vite-plugin-runtime-error-modal`** | Replit-specific dev overlay |
| **`@replit/vite-plugin-cartographer`** | Replit-specific dev tool (dev mode only) |
| **`nanoid`** | Generating unique IDs (used in Vite dev server template cache-busting) |

### Environment Variables Required

| Variable | Description |
|---|---|
| `DATABASE_URL` | Full PostgreSQL connection string (e.g., `postgresql://user:pass@host:5432/dbname`) |