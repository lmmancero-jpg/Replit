import * as XLSX from "xlsx";

// ---- Columnas (0-based) ----
const COL = {
  FECHA: 1,
  H_U1: 4,
  H_U2: 10,
  AUX: 24,
  LANEC: 25,
  GRACA: 30,
  ETOTAL: 35,
  E_U1: 36,
  E_U2: 37,
  REND: 38,
  HFO_TOT: 43,
  DO_TOT: 46,
} as const;

function toNumber(value: unknown): number {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return isFinite(value) ? value : 0;
  let str = String(value).trim().replace(/\s+/g, "");
  if (!str) return 0;
  if (/^\d{1,3}(\.\d{3})*(,\d+)?$/.test(str)) str = str.replace(/\./g, "").replace(",", ".");
  else if (/^\d{1,3}(,\d{3})*(\.\d+)?$/.test(str)) str = str.replace(/,/g, "");
  else if (/^\d+,\d+$/.test(str)) str = str.replace(",", ".");
  const n = Number(str);
  return isFinite(n) ? n : 0;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function parseExcelDate(value: any): Date | null {
  if (value === null || value === undefined || value === "") return null;
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return isNaN((value as Date).getTime()) ? null : (value as Date);
  }
  if (typeof value === "number") {
    const o = XLSX.SSF.parse_date_code(value);
    if (!o) return null;
    return new Date(o.y, o.m - 1, o.d);
  }
  const s = String(value).trim();
  if (!s) return null;
  const d1 = new Date(s);
  if (!isNaN(d1.getTime())) return d1;
  const m = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (m) {
    let dd = parseInt(m[1], 10), mm = parseInt(m[2], 10), yy = parseInt(m[3], 10);
    if (yy < 100) yy = 2000 + yy;
    const d2 = new Date(yy, mm - 1, dd);
    return isNaN(d2.getTime()) ? null : d2;
  }
  return null;
}

function dayLabel(date: Date): string {
  return `${String(date.getDate()).padStart(2, "0")}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function dominantMonthYear(rows: any[][]): { year: number; month: number } | null {
  const counts: Record<string, number> = {};
  for (const r of rows) {
    const d = parseExcelDate(r?.[COL.FECHA]);
    if (!d) continue;
    const key = `${d.getFullYear()}-${d.getMonth() + 1}`;
    counts[key] = (counts[key] || 0) + 1;
  }
  let bestKey: string | null = null, best = 0;
  for (const [k, c] of Object.entries(counts)) {
    if (c > best) { best = c; bestKey = k; }
  }
  if (!bestKey) return null;
  const [y, mo] = bestKey.split("-").map(Number);
  return { year: y, month: mo };
}

export interface ProdData {
  sheetName: string;
  targetMonth: number;
  targetYear: number;
  labels: string[];
  etotal: (number | null)[];
  aux: (number | null)[];
  lanec: (number | null)[];
  graca: (number | null)[];
  e_u1: (number | null)[];
  e_u2: (number | null)[];
  h_u1: (number | null)[];
  h_u2: (number | null)[];
  pot_u1: (number | null)[];
  pot_u2: (number | null)[];
  rend: (number | null)[];
  hfoTot: number[];
  doTot: number[];
  hfoG1: number[];
  hfoG2: number[];
  doG1: number[];
  doG2: number[];
  hfoG1_galH: (number | null)[];
  hfoG2_galH: (number | null)[];
  doG1_galH: (number | null)[];
  doG2_galH: (number | null)[];
}

export interface AforoData {
  labels: string[];
  t601: (number | null)[];
  t602: (number | null)[];
  t610: (number | null)[];
  t611: (number | null)[];
  cisterna2: (number | null)[];
}

export interface Resumen {
  energiaTotalMWh: number;
  energiaU1MWh: number;
  energiaU2MWh: number;
  energiaLanecMWh: number;
  energiaGracaMWh: number;
  energiaAuxMWh: number;
  horasU1: number;
  horasU2: number;
  potPromU1: number | null;
  potPromU2: number | null;
  eficProm: number | null;
  hfoGal: number;
  doGal: number;
  dias: number;
}

export function extractProduction(wb: XLSX.WorkBook, sheetName: string): ProdData | null {
  const ws = wb.Sheets[sheetName];
  if (!ws) return null;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null }) as any[][];
  if (!rows.length) return null;
  const dom = dominantMonthYear(rows);
  if (!dom) return null;

  const labels: string[] = [];
  const etotal: (number | null)[] = [];
  const aux: (number | null)[] = [];
  const lanec: (number | null)[] = [];
  const graca: (number | null)[] = [];
  const e_u1: (number | null)[] = [];
  const e_u2: (number | null)[] = [];
  const h_u1: (number | null)[] = [];
  const h_u2: (number | null)[] = [];
  const pot_u1: (number | null)[] = [];
  const pot_u2: (number | null)[] = [];
  const rend: (number | null)[] = [];
  const hfoTot: number[] = [];
  const doTot: number[] = [];
  const hfoG1: number[] = [];
  const hfoG2: number[] = [];
  const doG1: number[] = [];
  const doG2: number[] = [];
  const hfoG1_galH: (number | null)[] = [];
  const hfoG2_galH: (number | null)[] = [];
  const doG1_galH: (number | null)[] = [];
  const doG2_galH: (number | null)[] = [];

  for (const r of rows) {
    const d = parseExcelDate(r?.[COL.FECHA]);
    if (!d) continue;
    if (d.getMonth() + 1 !== dom.month || d.getFullYear() !== dom.year) continue;

    const eTot = toNumber(r[COL.ETOTAL]);
    const e1   = toNumber(r[COL.E_U1]);
    const e2   = toNumber(r[COL.E_U2]);
    const a    = toNumber(r[COL.AUX]);
    const l    = toNumber(r[COL.LANEC]);
    const g    = toNumber(r[COL.GRACA]);
    const hu1  = toNumber(r[COL.H_U1]);
    const hu2  = toNumber(r[COL.H_U2]);
    const re   = toNumber(r[COL.REND]);
    const hf   = toNumber(r[COL.HFO_TOT]);
    const dox  = toNumber(r[COL.DO_TOT]);

    if (!eTot && !e1 && !e2 && !a && !l && !g && !hu1 && !hu2 && !re && !hf && !dox) continue;

    labels.push(dayLabel(d));
    etotal.push(eTot || null);
    aux.push(a || null);
    lanec.push(l || null);
    graca.push(g || null);
    e_u1.push(e1 || null);
    e_u2.push(e2 || null);
    h_u1.push(hu1 || null);
    h_u2.push(hu2 || null);
    rend.push(re || null);
    pot_u1.push(e1 > 0 && hu1 > 0 ? e1 / hu1 : null);
    pot_u2.push(e2 > 0 && hu2 > 0 ? e2 / hu2 : null);
    hfoTot.push(hf);
    doTot.push(dox);
    const eSum = e1 + e2;
    let g1hfo = 0, g2hfo = 0, g1do = 0, g2do = 0;
    if (eSum > 0) {
      g1hfo = hf * (e1 / eSum); g2hfo = hf * (e2 / eSum);
      g1do  = dox * (e1 / eSum); g2do  = dox * (e2 / eSum);
    }
    hfoG1.push(g1hfo); hfoG2.push(g2hfo);
    doG1.push(g1do);   doG2.push(g2do);

    hfoG1_galH.push(g1hfo > 0 && hu1 > 0 ? g1hfo / hu1 : null);
    hfoG2_galH.push(g2hfo > 0 && hu2 > 0 ? g2hfo / hu2 : null);
    doG1_galH.push(g1do  > 0 && hu1 > 0 ? g1do  / hu1  : null);
    doG2_galH.push(g2do  > 0 && hu2 > 0 ? g2do  / hu2  : null);
  }

  return { sheetName, targetMonth: dom.month, targetYear: dom.year, labels, etotal, aux, lanec, graca, e_u1, e_u2, h_u1, h_u2, pot_u1, pot_u2, rend, hfoTot, doTot, hfoG1, hfoG2, doG1, doG2, hfoG1_galH, hfoG2_galH, doG1_galH, doG2_galH };
}

export function extractAforo(wb: XLSX.WorkBook, targetMonth: number, targetYear: number): AforoData | null {
  const ws = wb.Sheets["Sondas 00 00 hrs"];
  if (!ws) return null;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null }) as any[][];
  if (!rows.length) return null;

  const labels: string[] = [];
  const t601: (number | null)[] = [];
  const t602: (number | null)[] = [];
  const t610: (number | null)[] = [];
  const t611: (number | null)[] = [];
  const cisterna2: (number | null)[] = [];

  let currentDate: Date | null = null;

  for (let i = 2; i < rows.length; i++) {
    const r = rows[i];
    const maybeDate = parseExcelDate(r?.[0]);
    if (maybeDate) currentDate = maybeDate;
    if (!currentDate) continue;
    if (currentDate.getMonth() + 1 !== targetMonth || currentDate.getFullYear() !== targetYear) continue;

    const label = dayLabel(currentDate);
    const rawTag = String(r?.[1] ?? "").trim().toUpperCase().replace(/\s+/g, "");
    const tipoNorm = String(r?.[2] ?? "").trim().toUpperCase().replace(/\s+/g, "");
    const vol = toNumber(r?.[5]);
    if (!vol) continue;

    const idxForLabel = (): number => {
      let idx = labels.indexOf(label);
      if (idx === -1) {
        labels.push(label); t601.push(null); t602.push(null); t610.push(null); t611.push(null); cisterna2.push(null);
        idx = labels.length - 1;
      }
      return idx;
    };

    if      (rawTag.startsWith("T601"))    { const ix = idxForLabel(); t601[ix] = vol; }
    else if (rawTag.startsWith("T602"))    { const ix = idxForLabel(); t602[ix] = vol; }
    else if (rawTag.startsWith("T610"))    { const ix = idxForLabel(); t610[ix] = vol; }
    else if (rawTag.startsWith("T611"))    { const ix = idxForLabel(); t611[ix] = vol; }
    else if (tipoNorm === "CISTERNA2")     { const ix = idxForLabel(); cisterna2[ix] = vol; }
  }

  return { labels, t601, t602, t610, t611, cisterna2 };
}

export function buildResumen(prod: ProdData): Resumen {
  const sum = (arr: (number | null)[]): number =>
    arr.reduce<number>((a, b) => a + (b ?? 0), 0);
  const avgNoZero = (arr: (number | null)[]): number | null => {
    const f = arr.filter((v) => v != null && v !== 0) as number[];
    return f.length ? f.reduce<number>((a, b) => a + b, 0) / f.length : null;
  };
  const sumU1  = sum(prod.e_u1);
  const sumU2  = sum(prod.e_u2);
  const sumHU1 = sum(prod.h_u1);
  const sumHU2 = sum(prod.h_u2);
  return {
    energiaTotalMWh:  sum(prod.etotal) / 1000,
    energiaU1MWh:     sumU1 / 1000,
    energiaU2MWh:     sumU2 / 1000,
    energiaLanecMWh:  sum(prod.lanec) / 1000,
    energiaGracaMWh:  sum(prod.graca) / 1000,
    energiaAuxMWh:    sum(prod.aux)   / 1000,
    horasU1:          sumHU1,
    horasU2:          sumHU2,
    potPromU1:        sumU1 > 0 && sumHU1 > 0 ? sumU1 / sumHU1 : null,
    potPromU2:        sumU2 > 0 && sumHU2 > 0 ? sumU2 / sumHU2 : null,
    eficProm:         avgNoZero(prod.rend),
    hfoGal:           sum(prod.hfoTot),
    doGal:            sum(prod.doTot),
    dias:             prod.labels.length,
  };
}

export function fmt(value: number | null | undefined, decimals = 0): string {
  if (value === null || value === undefined || !isFinite(value)) return "—";
  try {
    return value.toLocaleString("es-EC", { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
  } catch {
    return Number(value).toFixed(decimals);
  }
}
