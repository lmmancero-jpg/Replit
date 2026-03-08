import { db } from "./server/db";
import { reports } from "./shared/schema";

async function seed() {
  const existing = await db.select().from(reports);
  if (existing.length === 0) {
    await db.insert(reports).values({
      title: "Informe diario - 2026-03-08",
      reportType: "diario",
      date: "2026-03-08",
      content: "<div class='report-document'><h2>Informe de Ejemplo</h2><p>Este es un informe de prueba generado automáticamente.</p></div>"
    });
    console.log("Seeded database with mock report");
  } else {
    console.log("Database already seeded");
  }
}

seed().catch(console.error).then(() => process.exit(0));
