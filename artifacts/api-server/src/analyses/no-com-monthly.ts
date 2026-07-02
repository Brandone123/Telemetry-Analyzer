/**
 * Monthly NO COM consolidation + predictions.
 *
 * Reads data DIRECTLY from the Comm Issue sheet (date serials in header row,
 * 1/0 per site per column, SUBTOTAL at the bottom) — this is always up to date
 * and does not depend on the NO COM HISTORY date column or the DAYS sheet.
 *
 * Output: a standalone 3-sheet Excel:
 *   MONTHLY STATS   — KPIs, region breakdown, top recurring sites, insights
 *   PREDICTIONS     — linear regression + WMA side-by-side, 7-day forecast
 *   SITE RISK SCORES — per-site risk rating
 */

import * as XLSX from "xlsx";
import ExcelJS from "exceljs";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface MonthlyConsolidationOptions {
  month?: number;   // 1-based. Defaults to last completed month.
  year?: number;
  referenceDate?: Date;
}

export interface MonthlyConsolidationResult {
  buffer: Buffer;
  filename: string;
  summary: {
    month: string;
    avgNoComPct: number;
    avgNoComSites: number;
    peakDay: string;
    peakCount: number;
    bestDay: string;
    bestCount: number;
    totalDays: number;
    byRegion: Record<string, { avgNoCom: number; avgTotal: number; pct: number }>;
    topRecurringSites: Array<{ site: string; region: string; days: number; pct: number }>;
    linearForecastNextWeek: number;
    wmaForecastNextWeek: number;
    trend: "improving" | "stable" | "worsening";
  };
}

interface DailyEntry {
  dateSerial: number;
  date: Date;
  dateStr: string;
  dayName: string;
  total: number;      // SUBTOTAL value for that column
  onAir: number;      // number of On Air sites in that column
}

interface SiteEntry {
  rowIndex: number;
  ihsId: string;
  siteId: string;
  region: string;
  projectStatus: string;
  noComDays: number;
  totalDays: number;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const MONTH_NAMES = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];

function excelSerialToDate(serial: number): Date {
  return new Date(Date.UTC(1899, 11, 30) + serial * 86_400_000);
}

function fmtDate(d: Date): string {
  return `${String(d.getUTCMonth() + 1).padStart(2, "0")}/${String(d.getUTCDate()).padStart(2, "0")}/${d.getUTCFullYear()}`;
}

function dayNameOf(d: Date): string {
  return d.toLocaleDateString("en-US", { weekday: "long", timeZone: "UTC" }).toUpperCase();
}

// ---------------------------------------------------------------------------
// Linear regression
// ---------------------------------------------------------------------------
function linearRegression(points: number[]): { slope: number; intercept: number; r2: number; predict: (x: number) => number } {
  const n = points.length;
  if (n < 2) return { slope: 0, intercept: points[0] ?? 0, r2: 0, predict: () => points[0] ?? 0 };
  let sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;
  for (let i = 0; i < n; i++) { sumX += i; sumY += points[i]!; sumXY += i * points[i]!; sumX2 += i * i; }
  const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
  const intercept = (sumY - slope * sumX) / n;
  const yMean = sumY / n;
  let ssTot = 0, ssRes = 0;
  for (let i = 0; i < n; i++) { ssTot += (points[i]! - yMean) ** 2; ssRes += (points[i]! - (slope * i + intercept)) ** 2; }
  const r2 = ssTot === 0 ? 1 : Math.max(0, 1 - ssRes / ssTot);
  return { slope, intercept, r2, predict: (x: number) => slope * x + intercept };
}

// ---------------------------------------------------------------------------
// Weighted Moving Average
// ---------------------------------------------------------------------------
function wma(points: number[], window: number): number[] {
  const weights = Array.from({ length: window }, (_, i) => i + 1);
  const sumW = weights.reduce((a, b) => a + b, 0);
  const result: number[] = [];
  for (let i = window - 1; i < points.length; i++) {
    let v = 0;
    for (let j = 0; j < window; j++) v += points[i - window + 1 + j]! * weights[j]!;
    result.push(v / sumW);
  }
  return result;
}

// ---------------------------------------------------------------------------
// Read monthly data from Comm Issue
// ---------------------------------------------------------------------------
function readCommIssueMonthly(wb: XLSX.WorkBook, targetMonth: number, targetYear: number): {
  daily: DailyEntry[];
  sites: SiteEntry[];
  allDailyHistory: DailyEntry[]; // all columns ever, for regression
} {
  const ws = wb.Sheets["Comm Issue"];
  if (!ws) return { daily: [], sites: [], allDailyHistory: [] };

  const range = XLSX.utils.decode_range(ws["!ref"] ?? "A1:A1");

  // ---- Step 1: Scan header row for date serials (row 0, col 9+ / col J+)
  // Some trackers start dates earlier; scan from col 5 onward to be safe.
  const dateCols: Array<{ col: number; serial: number; date: Date }> = [];
  for (let c = 5; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: 0, c })];
    // Date serials: Excel date serials > 40000 correspond to dates after 2009
    if (cell && cell.t === "n" && Number(cell.v) > 40000) {
      dateCols.push({ col: c, serial: Number(cell.v), date: excelSerialToDate(Number(cell.v)) });
    }
  }

  // ---- Step 2: Identify On Air site rows and skip non-data rows.
  // A site row must have an IHS ID in col B (index 1) and project status "on air" in col E (index 4).
  // We also detect non-site rows: formula rows, header repeats, subtotal rows (these have no IHS ID).
  const siteRows = new Map<number, SiteEntry>();
  for (let r = 1; r <= range.e.r; r++) {
    const ihsCell = ws[XLSX.utils.encode_cell({ r, c: 1 })];
    if (!ihsCell?.v) continue; // blank IHS ID → not a site row (header repeat, subtotal, etc.)
    const ihsStr = String(ihsCell.v).trim();
    if (!ihsStr || /subtotal|total|sum/i.test(ihsStr)) continue; // skip label rows
    const psCell = ws[XLSX.utils.encode_cell({ r, c: 4 })];
    const ps = String(psCell?.v ?? "").trim().toLowerCase();
    if (ps !== "on air") continue;
    siteRows.set(r, {
      rowIndex: r,
      ihsId: ihsStr,
      siteId: String(ws[XLSX.utils.encode_cell({ r, c: 2 })]?.v ?? "").trim(),
      region: String(ws[XLSX.utils.encode_cell({ r, c: 0 })]?.v ?? "").trim(),
      projectStatus: ps,
      noComDays: 0,
      totalDays: 0,
    });
  }

  // ---- Step 3: For every date column, sum 1/0 values across On Air site rows.
  // This is the correct NO COM count — no dependency on SUBTOTAL formulas.
  const allDailyHistory: DailyEntry[] = [];
  const targetDateCols: typeof dateCols = [];

  for (const dc of dateCols) {
    let noComCount = 0;
    let filledCount = 0; // number of On Air sites that have any value in this col
    for (const [r] of siteRows) {
      const cell = ws[XLSX.utils.encode_cell({ r, c: dc.col })];
      if (cell?.v != null) {
        filledCount++;
        if (Number(cell.v) === 1) noComCount++;
      }
    }
    // Only include this date column if at least one site has data (avoids empty future cols)
    if (filledCount === 0) continue;

    const entry: DailyEntry = {
      dateSerial: dc.serial,
      date: dc.date,
      dateStr: fmtDate(dc.date),
      dayName: dayNameOf(dc.date),
      total: noComCount,
      onAir: filledCount,
    };
    allDailyHistory.push(entry);
    if (dc.date.getUTCMonth() + 1 === targetMonth && dc.date.getUTCFullYear() === targetYear) {
      targetDateCols.push(dc);
    }
  }

  // ---- Step 4: Per-site stats for the target month
  for (const dc of targetDateCols) {
    for (const [r, site] of siteRows) {
      const cell = ws[XLSX.utils.encode_cell({ r, c: dc.col })];
      if (cell?.v != null) {
        site.totalDays++;
        if (Number(cell.v) === 1) site.noComDays++;
      }
    }
  }

  const daily = allDailyHistory.filter(
    (e) => e.date.getUTCMonth() + 1 === targetMonth && e.date.getUTCFullYear() === targetYear,
  );
  const sites = Array.from(siteRows.values()).filter((s) => s.totalDays > 0);

  return { daily, sites, allDailyHistory };
}

// ---------------------------------------------------------------------------
// Main export
// ---------------------------------------------------------------------------
export async function consolidateMonthlyNoCom(
  trackerBuffer: Buffer,
  _commBuffer: Buffer | null,
  _ticketsBuffer: Buffer | null,
  options: MonthlyConsolidationOptions = {},
): Promise<MonthlyConsolidationResult> {
  const now = options.referenceDate ?? new Date();
  let targetMonth = options.month;
  let targetYear = options.year;
  if (!targetMonth || !targetYear) {
    const prev = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth() - 1, 1));
    targetMonth = prev.getUTCMonth() + 1;
    targetYear = prev.getUTCFullYear();
  }

  const wb = XLSX.read(trackerBuffer, {
    type: "buffer",
    cellFormula: true,
    sheets: ["Comm Issue"],
  });

  const { daily, sites, allDailyHistory } = readCommIssueMonthly(wb, targetMonth, targetYear);

  // ---- KPIs ----
  const totals = daily.map((e) => e.total);
  const avgNoComSites = totals.length > 0 ? totals.reduce((a, b) => a + b, 0) / totals.length : 0;
  const onAirCount = daily.length > 0 ? Math.max(...daily.map((e) => e.onAir), sites.length) : sites.length;
  const avgNoComPct = onAirCount > 0 ? (avgNoComSites / onAirCount) * 100 : 0;

  const empty = { dateStr: "-", total: 0, date: new Date(), dayName: "", dateSerial: 0, onAir: 0 };
  const peakEntry = daily.reduce((a, b) => (b.total > a.total ? b : a), daily[0] ?? empty);
  const bestEntry = daily.reduce((a, b) => (b.total < a.total ? b : a), daily[0] ?? empty);

  // ---- Region breakdown from site-level data ----
  const regionMap = new Map<string, { noComSum: number; totalSum: number; siteCount: number }>();
  for (const s of sites) {
    const r = s.region || "Unknown";
    const cur = regionMap.get(r) ?? { noComSum: 0, totalSum: 0, siteCount: 0 };
    cur.noComSum += s.noComDays;
    cur.totalSum += s.totalDays;
    cur.siteCount++;
    regionMap.set(r, cur);
  }
  const byRegion: Record<string, { avgNoCom: number; avgTotal: number; pct: number }> = {};
  for (const [region, d] of regionMap) {
    byRegion[region] = {
      avgNoCom: d.siteCount > 0 ? d.noComSum / d.siteCount : 0,
      avgTotal: d.siteCount > 0 ? d.totalSum / d.siteCount : 0,
      pct: d.totalSum > 0 ? Math.round((d.noComSum / d.totalSum) * 100) : 0,
    };
  }

  // ---- Top recurring sites ----
  const topRecurringSites = sites
    .filter((s) => s.noComDays > 0)
    .map((s) => ({
      site: s.siteId || s.ihsId,
      region: s.region,
      days: s.noComDays,
      pct: s.totalDays > 0 ? Math.round((s.noComDays / s.totalDays) * 100) : 0,
    }))
    .sort((a, b) => b.days - a.days)
    .slice(0, 30);

  // ---- Predictions (use all history up to end of target month) ----
  const historyUpToNow = allDailyHistory.filter(
    (e) => e.date <= new Date(Date.UTC(targetYear, targetMonth, 0)),
  );
  const series = historyUpToNow.slice(-60).map((e) => e.total);
  const lr = linearRegression(series);
  const wmaWindow = Math.min(7, series.length);
  const wmaValues = wma(series, wmaWindow);
  const lastWma = wmaValues[wmaValues.length - 1] ?? avgNoComSites;

  const linearForecastNextWeek = Math.max(0, Math.round(lr.predict(series.length + 3)));
  const wmaForecastNextWeek = Math.max(0, Math.round(lastWma));

  const last7 = series.slice(-7);
  const prev7 = series.slice(-14, -7);
  const avg7 = last7.reduce((a, b) => a + b, 0) / (last7.length || 1);
  const avgPrev7 = prev7.reduce((a, b) => a + b, 0) / (prev7.length || 1);
  const trend: "improving" | "stable" | "worsening" =
    avg7 < avgPrev7 * 0.95 ? "improving" : avg7 > avgPrev7 * 1.05 ? "worsening" : "stable";

  const monthLabel = `${MONTH_NAMES[targetMonth - 1]} ${targetYear}`;

  // ---- Build the Excel output ----
  const buffer = await buildMonthlyExcel({
    monthLabel, daily, allDailyHistory, sites,
    avgNoComSites, avgNoComPct, onAirCount,
    peakEntry, bestEntry, byRegion, topRecurringSites,
    series, lr, wmaValues, wmaWindow, lastWma,
    linearForecastNextWeek, wmaForecastNextWeek, trend,
  });

  const filename = `NoCom_Monthly_${String(targetYear)}_${String(targetMonth).padStart(2, "0")}.xlsx`;

  return {
    buffer,
    filename,
    summary: {
      month: monthLabel,
      avgNoComPct: Math.round(avgNoComPct * 10) / 10,
      avgNoComSites: Math.round(avgNoComSites),
      peakDay: peakEntry.dateStr,
      peakCount: peakEntry.total,
      bestDay: bestEntry.dateStr,
      bestCount: bestEntry.total,
      totalDays: daily.length,
      byRegion,
      topRecurringSites: topRecurringSites.slice(0, 10),
      linearForecastNextWeek,
      wmaForecastNextWeek,
      trend,
    },
  };
}

// ---------------------------------------------------------------------------
// Excel builder
// ---------------------------------------------------------------------------
interface BuildArgs {
  monthLabel: string;
  daily: DailyEntry[];
  allDailyHistory: DailyEntry[];
  sites: SiteEntry[];
  avgNoComSites: number;
  avgNoComPct: number;
  onAirCount: number;
  peakEntry: DailyEntry;
  bestEntry: DailyEntry;
  byRegion: Record<string, { avgNoCom: number; avgTotal: number; pct: number }>;
  topRecurringSites: Array<{ site: string; region: string; days: number; pct: number }>;
  series: number[];
  lr: { slope: number; intercept: number; r2: number; predict: (x: number) => number };
  wmaValues: number[];
  wmaWindow: number;
  lastWma: number;
  linearForecastNextWeek: number;
  wmaForecastNextWeek: number;
  trend: string;
}

async function buildMonthlyExcel(a: BuildArgs): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Telemetry Analyzer";

  const NAVY: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E78" } };
  const TEAL: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF246C82" } };
  const RED_FILL: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC00000" } };
  const GREEN_FILL: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF375623" } };
  const GOLD_FILL: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFC000" } };
  const LIGHT: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD6E4F0" } };
  const WHITE: Partial<ExcelJS.Font> = { color: { argb: "FFFFFFFF" }, bold: true };

  const hdr = (ws: ExcelJS.Worksheet, ref: string, text: string, fill: ExcelJS.Fill, span?: string) => {
    if (span) ws.mergeCells(span);
    const c = ws.getCell(ref);
    c.value = text; c.font = WHITE; c.fill = fill;
    c.alignment = { horizontal: "center", vertical: "middle" };
  };
  const lbl = (ws: ExcelJS.Worksheet, ref: string, text: string) => {
    ws.getCell(ref).value = text; ws.getCell(ref).font = { bold: true };
  };
  const num = (ws: ExcelJS.Worksheet, ref: string, v: number | string) => {
    ws.getCell(ref).value = v; ws.getCell(ref).alignment = { horizontal: "right" };
  };

  // ================================================================
  // Sheet 1: MONTHLY STATS
  // ================================================================
  const ws1 = wb.addWorksheet("MONTHLY STATS");
  ws1.getRow(1).height = 30;
  hdr(ws1, "A1", `NO COM — Monthly Statistics · ${a.monthLabel}`, NAVY, "A1:G1");

  // KPI block
  hdr(ws1, "A3", "KEY PERFORMANCE INDICATORS", TEAL, "A3:G3");
  const kpis: [string, string | number][] = [
    ["Month", a.monthLabel],
    ["Days recorded in tracker", a.daily.length],
    ["Average NO COM sites / day", Math.round(a.avgNoComSites)],
    ["Average NO COM rate", `${(Math.round(a.avgNoComPct * 10) / 10).toFixed(1)}%`],
    ["Peak day (worst)", `${a.peakEntry.dateStr} — ${a.peakEntry.total} sites`],
    ["Best day (lowest)", `${a.bestEntry.dateStr} — ${a.bestEntry.total} sites`],
    ["On Air sites count (approx)", a.onAirCount],
  ];
  kpis.forEach(([k, v], i) => {
    lbl(ws1, `B${4 + i}`, String(k));
    num(ws1, `D${4 + i}`, v);
  });

  // Daily detail table
  const detailStart = 4 + kpis.length + 2;
  hdr(ws1, `A${detailStart}`, "DAILY NO COM COUNT — DETAIL", TEAL, `A${detailStart}:G${detailStart}`);
  const dhRow = ws1.getRow(detailStart + 1);
  dhRow.values = ["", "Date", "Day", "NO COM Sites", "vs Average", "", ""];
  dhRow.font = { bold: true }; dhRow.fill = LIGHT;

  const avg = a.avgNoComSites;
  a.daily.forEach((e, i) => {
    const r = detailStart + 2 + i;
    const dev = Math.round(e.total - avg);
    ws1.getCell(`B${r}`).value = e.dateStr;
    ws1.getCell(`C${r}`).value = e.dayName;
    ws1.getCell(`D${r}`).value = e.total;
    ws1.getCell(`D${r}`).alignment = { horizontal: "center" };
    const dc = ws1.getCell(`E${r}`);
    dc.value = dev >= 0 ? `+${dev}` : `${dev}`;
    dc.alignment = { horizontal: "center" };
    if (dev > 5) dc.font = { color: { argb: "FFC00000" } };
    else if (dev < -5) dc.font = { color: { argb: "FF375623" } };
  });

  // Region breakdown
  const regStart = detailStart + 2 + a.daily.length + 2;
  hdr(ws1, `A${regStart}`, "BREAKDOWN BY REGION", TEAL, `A${regStart}:G${regStart}`);
  const rhRow = ws1.getRow(regStart + 1);
  rhRow.values = ["", "Region", "Avg NO COM days", "Avg total days", "NO COM %", "", ""];
  rhRow.font = { bold: true }; rhRow.fill = LIGHT;

  const sortedRegions = Object.entries(a.byRegion).sort(([, x], [, y]) => y.pct - x.pct);
  sortedRegions.forEach(([region, d], i) => {
    const r = regStart + 2 + i;
    ws1.getCell(`B${r}`).value = region;
    ws1.getCell(`C${r}`).value = Math.round(d.avgNoCom * 10) / 10;
    ws1.getCell(`D${r}`).value = Math.round(d.avgTotal * 10) / 10;
    const pc = ws1.getCell(`E${r}`);
    pc.value = `${d.pct}%`; pc.alignment = { horizontal: "center" };
    if (d.pct >= 30) pc.font = { bold: true, color: { argb: "FFC00000" } };
    else if (d.pct >= 15) pc.font = { color: { argb: "FFD68000" } };
  });

  // Top recurring sites
  const siteStart = regStart + 2 + sortedRegions.length + 2;
  hdr(ws1, `A${siteStart}`, "TOP 30 MOST RECURRING NO COM SITES (this month)", RED_FILL, `A${siteStart}:G${siteStart}`);
  const shRow = ws1.getRow(siteStart + 1);
  shRow.values = ["", "Site ID", "Region", "Days in NO COM", "% of month in NO COM", "", ""];
  shRow.font = { bold: true }; shRow.fill = LIGHT;
  a.topRecurringSites.forEach((s, i) => {
    const r = siteStart + 2 + i;
    ws1.getCell(`B${r}`).value = s.site;
    ws1.getCell(`C${r}`).value = s.region;
    ws1.getCell(`D${r}`).value = s.days; ws1.getCell(`D${r}`).alignment = { horizontal: "center" };
    const pc = ws1.getCell(`E${r}`);
    pc.value = `${s.pct}%`; pc.alignment = { horizontal: "center" };
    if (s.pct >= 80) pc.font = { bold: true, color: { argb: "FFC00000" } };
    else if (s.pct >= 50) pc.font = { color: { argb: "FFD68000" } };
  });

  // Insights
  const insStart = siteStart + 2 + a.topRecurringSites.length + 2;
  hdr(ws1, `A${insStart}`, "INSIGHTS & RECOMMENDED ACTIONS", GREEN_FILL, `A${insStart}:G${insStart}`);
  const insights: string[] = [];
  const worstReg = sortedRegions[0];
  if (worstReg && worstReg[1].pct >= 20)
    insights.push(`⚠ Region "${worstReg[0]}" has the highest NO COM rate (${worstReg[1].pct}%) — prioritize field investigation and controller replacement.`);
  const chronic = a.topRecurringSites.filter((s) => s.pct >= 80);
  if (chronic.length > 0)
    insights.push(`⚠ ${chronic.length} site(s) in NO COM for ≥ 80% of the month — chronic outages requiring escalation (hardware, fibre/radio link audit).`);
  const highDays = a.daily.filter((e) => e.total > avg * 1.3);
  if (highDays.length > 2)
    insights.push(`📌 ${highDays.length} days had NO COM > 30% above average — check for power outages or maintenance on those dates.`);
  insights.push("✅ Ensure all Comm Issue columns are filled daily to maintain accurate history for predictions.");
  insights.push("✅ Cross-check sites appearing every month with the PowerGen team — chronic sites often have RTU/power supply issues.");
  insights.push("✅ Use the Predictions sheet to anticipate high-risk weeks and pre-position field teams.");

  insights.forEach((txt, i) => {
    ws1.mergeCells(`A${insStart + 1 + i}:G${insStart + 1 + i}`);
    const c = ws1.getCell(`A${insStart + 1 + i}`);
    c.value = txt; c.alignment = { wrapText: true };
    ws1.getRow(insStart + 1 + i).height = 22;
  });

  ws1.columns = [{ width: 3 }, { width: 24 }, { width: 20 }, { width: 18 }, { width: 22 }, { width: 3 }, { width: 3 }];
  ws1.views = [{ state: "frozen", ySplit: 1 }];

  // ================================================================
  // Sheet 2: PREDICTIONS
  // ================================================================
  const ws2 = wb.addWorksheet("PREDICTIONS");
  ws2.getRow(1).height = 28;
  hdr(ws2, "A1", `NO COM — Predictions & Trend Analysis · Generated ${new Date().toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" })}`, NAVY, "A1:H1");

  const trendFill = a.trend === "improving" ? GREEN_FILL : a.trend === "worsening" ? RED_FILL : GOLD_FILL;
  const trendText = a.trend === "improving"
    ? "▼ Trend: IMPROVING — NO COM count is decreasing over the recent period"
    : a.trend === "worsening"
    ? "▲ Trend: WORSENING — NO COM count is increasing over the recent period"
    : "→ Trend: STABLE — NO COM count is roughly flat";
  hdr(ws2, "A3", trendText, trendFill, "A3:H3");

  hdr(ws2, "A5", "NEXT-WEEK FORECAST (midpoint of week)", TEAL, "A5:H5");
  const forecastData = [
    ["Method", "Forecast sites/day", "Confidence", "Description"],
    ["Linear Regression", a.linearForecastNextWeek, `R² = ${(a.lr.r2 * 100).toFixed(1)}%`, "Best for long-term monotonic trends"],
    ["Weighted Moving Average", a.wmaForecastNextWeek, `Window: ${a.wmaWindow} days`, "Reacts faster to recent changes"],
    ["Combined estimate (recommended)", Math.round((a.linearForecastNextWeek + a.wmaForecastNextWeek) / 2), "—", "Average of both methods"],
  ];
  forecastData.forEach((row, i) => {
    const r = 6 + i;
    const exRow = ws2.getRow(r);
    exRow.values = ["", ...row];
    if (i === 0 || i === forecastData.length - 1) exRow.font = { bold: true };
    if (i === 0) exRow.fill = LIGHT;
  });

  // Historical series + regression + WMA
  const recentHistory = a.allDailyHistory.slice(-60);
  const histStart = 6 + forecastData.length + 2;
  hdr(ws2, `A${histStart}`, `HISTORICAL DATA (last ${recentHistory.length} recorded days) + REGRESSION`, TEAL, `A${histStart}:H${histStart}`);
  const hhRow = ws2.getRow(histStart + 1);
  hhRow.values = ["", "Date", "Day", "Actual NO COM", "Linear Regression", "WMA", "Residual (Actual − LR)", ""];
  hhRow.font = { bold: true }; hhRow.fill = LIGHT;

  recentHistory.forEach((e, i) => {
    const r = histStart + 2 + i;
    const lrVal = Math.round(a.lr.predict(i) * 10) / 10;
    const wmaIdx = i - (a.wmaWindow - 1);
    const wmaVal = wmaIdx >= 0 ? Math.round((a.wmaValues[wmaIdx] ?? 0) * 10) / 10 : null;
    const residual = Math.round((e.total - lrVal) * 10) / 10;
    ws2.getCell(`B${r}`).value = e.dateStr;
    ws2.getCell(`C${r}`).value = e.dayName;
    ws2.getCell(`D${r}`).value = e.total;
    ws2.getCell(`E${r}`).value = lrVal;
    if (wmaVal !== null) ws2.getCell(`F${r}`).value = wmaVal;
    const res = ws2.getCell(`G${r}`);
    res.value = residual;
    if (residual > 10) res.font = { color: { argb: "FFC00000" } };
    else if (residual < -10) res.font = { color: { argb: "FF375623" } };
  });

  // 7-day forecast rows
  const forecastStart = histStart + 2 + recentHistory.length + 1;
  hdr(ws2, `A${forecastStart}`, "FORECAST — NEXT 7 DAYS", TEAL, `A${forecastStart}:H${forecastStart}`);
  const now2 = new Date();
  for (let d = 1; d <= 7; d++) {
    const fd = new Date(now2.getTime() + d * 86_400_000);
    const r = forecastStart + d;
    ws2.getCell(`B${r}`).value = fmtDate(fd);
    ws2.getCell(`C${r}`).value = fd.toLocaleDateString("en-US", { weekday: "long" }).toUpperCase();
    ws2.getCell(`E${r}`).value = Math.max(0, Math.round(a.lr.predict(recentHistory.length + d - 1)));
    ws2.getCell(`F${r}`).value = Math.max(0, Math.round(a.lastWma));
    ["B", "C", "E", "F"].forEach((col) => {
      ws2.getCell(`${col}${r}`).font = { italic: true, color: { argb: col === "E" ? "FF1F4E78" : col === "F" ? "FF246C82" : "FF595959" } };
    });
  }

  ws2.columns = [{ width: 3 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 16 }, { width: 12 }, { width: 18 }, { width: 3 }];
  ws2.views = [{ state: "frozen", ySplit: 1 }];

  // ================================================================
  // Sheet 3: SITE RISK SCORES
  // ================================================================
  const ws3 = wb.addWorksheet("SITE RISK SCORES");
  ws3.getRow(1).height = 26;
  hdr(ws3, "A1", "NO COM — Per-Site Risk Scores (On Air sites, current month)", NAVY, "A1:F1");
  ws3.mergeCells("A2:F2");
  const note = ws3.getCell("A2");
  note.value = "Risk score = % of the month the site was in NO COM. HIGH ≥ 50%  |  MEDIUM 20–49%  |  LOW < 20%";
  note.font = { italic: true, color: { argb: "FF595959" } };

  const rHdr = ws3.getRow(3);
  rHdr.values = ["Risk Level", "Site ID", "Region", "NO COM days", "Total days recorded", "Risk %"];
  rHdr.font = { bold: true, color: { argb: "FFFFFFFF" } }; rHdr.fill = TEAL;

  const allSites = a.sites
    .filter((s) => s.totalDays > 0)
    .map((s) => ({ ...s, pct: (s.noComDays / s.totalDays) * 100 }))
    .sort((x, y) => y.pct - x.pct);

  allSites.forEach((s, i) => {
    const r = 4 + i;
    const risk = s.pct >= 50 ? "HIGH" : s.pct >= 20 ? "MEDIUM" : "LOW";
    const rf = s.pct >= 50 ? RED_FILL : s.pct >= 20 ? GOLD_FILL : GREEN_FILL;
    const rc = ws3.getCell(`A${r}`);
    rc.value = risk; rc.fill = rf; rc.font = { bold: true, color: { argb: "FFFFFFFF" } };
    rc.alignment = { horizontal: "center" };
    ws3.getCell(`B${r}`).value = s.siteId || s.ihsId;
    ws3.getCell(`C${r}`).value = s.region;
    ws3.getCell(`D${r}`).value = s.noComDays; ws3.getCell(`D${r}`).alignment = { horizontal: "center" };
    ws3.getCell(`E${r}`).value = s.totalDays; ws3.getCell(`E${r}`).alignment = { horizontal: "center" };
    const pc = ws3.getCell(`F${r}`);
    pc.value = `${Math.round(s.pct)}%`; pc.alignment = { horizontal: "center" };
    if (s.pct >= 50) pc.font = { bold: true, color: { argb: "FFC00000" } };
  });

  ws3.autoFilter = { from: "A3", to: "F3" };
  ws3.views = [{ state: "frozen", ySplit: 3 }];
  ws3.columns = [{ width: 12 }, { width: 20 }, { width: 18 }, { width: 14 }, { width: 18 }, { width: 10 }];

  return Buffer.from(await wb.xlsx.writeBuffer());
}
