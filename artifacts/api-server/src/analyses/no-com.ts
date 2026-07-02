import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import { parseSheet, isSlaProject, parseDate, type Row } from "../lib/excel-utils";
import { XlsxTemplate, colLetter, dateToExcelSerial } from "../lib/xlsx-template";
import { renderLineChart } from "../lib/chart";

export interface NoComOptions {
  cutoffDays: number; // default 2 (>48h)
  referenceDate: Date;
}

export interface NoComResult {
  buffer: Buffer;
  filename: string;
  ticketListBuffer: Buffer;
  ticketListFilename: string;
  summary: {
    totalSites: number;
    onAirSites: number;
    noComSites: number;
    ticketsToCreate: number;
    ticketsToResolve: number;
    moreThan24h: number;
    within48h: number;
    todayColumn: string;
    todayDateSerial: number;
    historyEntries: number;
    byCategory: Record<string, { com: number; noCom: number }>;
    ticketsToCreateList: Array<{ region: string; site: string; ihsId: string; vip: boolean; lastComm: string; daysSinceComm: number | null }>;
    ticketsToResolveList: Array<{ ttid: string; region: string; site: string; status: string }>;
  };
}

interface CommRow {
  region: string;
  site: string;
  projectStatus: string;
  status: string;
  lastCommunication: Date | null;
  lastCommRaw: string;
}

interface TrackerSite {
  rowIndex: number;        // 1-based row number in Comm Issue sheet
  state: string;
  ihsId: string;           // IHS_YND_244 (no trailing letter)
  siteId: string;          // IHS_YND_244O / M
  vip: boolean;
  noso: boolean;
  notNoso: boolean;
  aiot: boolean;
  projectStatus: string;
  siteAddress: string;
}

interface TicketRow {
  ttid: string;
  region: string;
  site: string;
  status: string;
  type: string;
  item: string;
  category: string;
  createTime: string;
  projectStatus: string;
}

function pickField(row: Row, candidates: string[]): string {
  for (const c of candidates) {
    for (const k of Object.keys(row)) {
      if (k.toLowerCase().trim() === c.toLowerCase().trim()) {
        const v = row[k];
        if (v !== undefined && v !== null && String(v).trim() !== "") return String(v);
      }
    }
  }
  return "";
}

function parseCommReport(rows: Row[]): CommRow[] {
  return rows.map((r) => {
    const lastRaw = pickField(r, ["Last Communication", "LastCommunication"]);
    const site = pickField(r, ["Site", "SiteName"]);
    return {
      region: pickField(r, ["Region", "RegionName"]),
      site,
      projectStatus: pickField(r, ["projectStatus", "Project Status", "ProjectStatus"]),
      status: pickField(r, ["Status"]),
      lastCommunication: parseDate(lastRaw),
      lastCommRaw: lastRaw,
    };
  }).filter((r) => r.site);
}

function parseTickets(rows: Row[]): TicketRow[] {
  return rows.map((r) => ({
    ttid: pickField(r, ["TTID", "TT ID"]),
    region: pickField(r, ["Region"]),
    site: pickField(r, ["Site"]),
    status: pickField(r, ["Status"]),
    type: pickField(r, ["Type"]),
    item: pickField(r, ["Item"]),
    category: pickField(r, ["Category"]),
    createTime: pickField(r, ["Create Time", "CreateTime"]),
    projectStatus: pickField(r, ["Project Status", "ProjectStatus", "projectStatus"]),
  })).filter((t) => t.site);
}

/**
 * Read the Comm Issue sheet from the SLA Tracker (using xlsx for fast read of formula
 * results). Returns the list of sites with their classification, plus the row index
 * of the SUBTOTAL aggregate row at the bottom of each date column.
 */
function parseTrackerCommIssue(wb: XLSX.WorkBook): {
  sites: TrackerSite[];
  totalRow: number;          // row containing SUBTOTAL (e.g. 2384)
  lastDataRow: number;       // last data row included in subtotal range (e.g. 2348)
} {
  const ws = wb.Sheets["Comm Issue"];
  if (!ws) throw new Error("Sheet 'Comm Issue' not found in SLA Tracker");
  const range = XLSX.utils.decode_range(ws["!ref"] ?? "A1:A1");
  const sites: TrackerSite[] = [];
  let lastDataRow = 1;

  // We need to detect where actual data ends. Iterate rows starting at 2,
  // stop when both column A (state) and column B (IHS ID) are empty for many rows in a row.
  let consecutiveEmpty = 0;
  for (let r = 1; r <= range.e.r; r++) {
    const stateCell = ws[XLSX.utils.encode_cell({ r, c: 0 })];
    const ihsCell = ws[XLSX.utils.encode_cell({ r, c: 1 })];
    const siteIdCell = ws[XLSX.utils.encode_cell({ r, c: 2 })];
    const psCell = ws[XLSX.utils.encode_cell({ r, c: 4 })];
    const c6 = ws[XLSX.utils.encode_cell({ r, c: 5 })]; // NOT NOSO
    const c7 = ws[XLSX.utils.encode_cell({ r, c: 6 })]; // NOSO
    const c8 = ws[XLSX.utils.encode_cell({ r, c: 7 })]; // AIOT
    const c4 = ws[XLSX.utils.encode_cell({ r, c: 3 })]; // VIP
    const addrCell = ws[XLSX.utils.encode_cell({ r, c: 8 })];

    const ihsId = String(ihsCell?.v ?? "").trim();
    if (!ihsId) {
      consecutiveEmpty++;
      // If we've found data and now we see >5 empty rows, stop
      if (sites.length > 0 && consecutiveEmpty > 5) break;
      continue;
    }
    consecutiveEmpty = 0;
    lastDataRow = r + 1;

    const norm = (cell: XLSX.CellObject | undefined) => String(cell?.v ?? "").trim().toUpperCase();
    const c6v = norm(c6);
    const c7v = norm(c7);
    const notNoso = c6v.includes("NOT NOSO") || c7v === "NOT NOSO";
    const noso = !notNoso && (c6v === "NOSO" || c7v === "NOSO");

    sites.push({
      rowIndex: r + 1, // 1-based
      state: String(stateCell?.v ?? "").trim(),
      ihsId,
      siteId: String(siteIdCell?.v ?? "").trim(),
      vip: norm(c4) === "VIP",
      noso,
      notNoso,
      aiot: norm(c8) === "AIOT",
      projectStatus: String(psCell?.v ?? "").trim(),
      siteAddress: String(addrCell?.v ?? "").trim(),
    });
  }

  // SUBTOTAL row sits a few rows below the last data row (typically lastDataRow + 36)
  // The user's formula example was AD2384 with SUBTOTAL(9,AD2:AD2348). So total row = 2384
  // We'll use a heuristic: scan for an existing SUBTOTAL formula in column J..AM at any row
  let totalRow = lastDataRow + 36;
  for (let r = lastDataRow; r <= Math.min(range.e.r, lastDataRow + 100); r++) {
    for (let c = 9; c <= 38; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell?.f && /SUBTOTAL/i.test(cell.f)) {
        totalRow = r + 1;
        break;
      }
    }
    if (totalRow !== lastDataRow + 36) break;
  }

  return { sites, totalRow, lastDataRow };
}

/**
 * Find today's date column letter from the header row of a sheet (J..AM contain
 * date serials in both Comm Issue and ON AIR). Returns null if not present.
 */
function findDateColumn(wb: XLSX.WorkBook, sheetName: string, todaySerial: number): string | null {
  const ws = wb.Sheets[sheetName];
  if (!ws) return null;
  const range = XLSX.utils.decode_range(ws["!ref"] ?? "A1:A1");
  for (let c = 0; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: 0, c })];
    if (cell && cell.t === "n" && Number(cell.v) === todaySerial) {
      return XLSX.utils.encode_col(c);
    }
  }
  return null;
}

/**
 * Read the ON AIR sheet's Site ID column and return a map from siteId → row index.
 * Used to mirror today's column from Comm Issue into ON AIR (the pivot's source).
 */
function parseOnAirSiteIndex(wb: XLSX.WorkBook): Map<string, number> {
  const ws = wb.Sheets["ON AIR"];
  const out = new Map<string, number>();
  if (!ws) return out;
  const range = XLSX.utils.decode_range(ws["!ref"] ?? "A1:A1");
  for (let r = 1; r <= range.e.r; r++) {
    const cell = ws[XLSX.utils.encode_cell({ r, c: 2 })]; // col C = Site ID
    const id = String(cell?.v ?? "").trim();
    if (id) out.set(id.toUpperCase(), r + 1); // 1-based row
  }
  return out;
}

/**
 * Read the ON AIR header row date serials between cols J and AM. Returns the
 * map column-letter → date serial, used to build cache field names when we
 * extend the pivot cache to today's column.
 */
function readOnAirDateHeaders(wb: XLSX.WorkBook): Array<{ col: string; serial: number }> {
  const ws = wb.Sheets["ON AIR"];
  const out: Array<{ col: string; serial: number }> = [];
  if (!ws) return out;
  const range = XLSX.utils.decode_range(ws["!ref"] ?? "A1:A1");
  for (let c = 9; c <= Math.min(range.e.c, 38); c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: 0, c })];
    if (cell && cell.t === "n") {
      out.push({ col: XLSX.utils.encode_col(c), serial: Number(cell.v) });
    }
  }
  return out;
}

function colLetterToIndex(letter: string): number {
  let n = 0;
  for (const ch of letter) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

function excelSerialToDate(serial: number): Date {
  const epoch = Date.UTC(1899, 11, 30);
  return new Date(epoch + serial * 86_400_000);
}

function fmtMDY(d: Date): string {
  const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
  const dd = String(d.getUTCDate()).padStart(2, "0");
  return `${mm}/${dd}/${d.getUTCFullYear()}`;
}

function fmtDMY(d: Date): string {
  const dd = String(d.getUTCDate()).padStart(2, "0");
  const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
  return `${dd}/${mm}/${d.getUTCFullYear()}`;
}

/**
 * Normalize a NO COM HISTORY day label to the full uppercase weekday name.
 * The user requires every entry to use the same format (e.g. "FRIDAY"), never
 * a mixed/abbreviated form like "FRI 29/05". We map by the leading 3 letters and
 * fall back to the trimmed/uppercased original if it can't be matched.
 */
function normalizeDayName(raw: string): string {
  const days: Record<string, string> = {
    MON: "MONDAY", TUE: "TUESDAY", WED: "WEDNESDAY", THU: "THURSDAY",
    FRI: "FRIDAY", SAT: "SATURDAY", SUN: "SUNDAY",
  };
  const key = raw.trim().toUpperCase().slice(0, 3);
  return days[key] ?? raw.trim().toUpperCase();
}

/**
 * A sub-table inside the NO COM HISTORY sheet: a fixed window of rows in two
 * adjacent columns (DAY, TOTAL). Older entries are shifted up when the window
 * is full, so the sub-table always shows the most recent N days and any chart
 * with a fixed range keeps working.
 */
interface HistoryTable {
  label: string;
  dateCol?: string;   // column just before dayCol containing dates, if present
  dayCol: string;     // e.g. "B"
  totalCol: string;   // e.g. "C"
  startRow: number;   // first data row (inclusive)
  endRow: number;     // last data row (inclusive)
  entries: Array<{ row: number; date?: string; day: string; total: number }>;
}

/**
 * Detect each sub-table inside NO COM HISTORY:
 *   - Total NO COM            (cols B/C — text "TOTAL")
 *   - NO COM more than 24hrs  (cols E/F — text contains "24")
 *   - NO COM within 48 hrs    (cols B/C — text contains "48")
 *
 * A sub-table is bounded by:
 *   - its header row (where col contains "DAY")
 *   - the next sub-table header in the same column, or the sheet end
 */
function readHistoryTables(wb: XLSX.WorkBook): HistoryTable[] {
  const ws = wb.Sheets["NO COM HISTORY"];
  if (!ws) return [];
  const range = XLSX.utils.decode_range(ws["!ref"] ?? "A1:A1");
  const tables: HistoryTable[] = [];

  type Found = { col: number; headerRow: number; label: string };
  const headers: Found[] = [];

  // Look for "DAY" in cols B and E (and the labels above them)
  for (const col of [1, 4]) {
    for (let r = 0; r <= range.e.r; r++) {
      const v = String(ws[XLSX.utils.encode_cell({ r, c: col })]?.v ?? "").trim().toUpperCase();
      if (v === "DAY") {
        // Look up to 3 rows above for a label
        let label = "";
        for (let up = 1; up <= 3 && r - up >= 0; up++) {
          const lv = String(ws[XLSX.utils.encode_cell({ r: r - up, c: col })]?.v ?? "").trim();
          if (lv && lv.toUpperCase() !== "DAY") { label = lv; break; }
        }
        // First sub-table in col B has its label as a sibling (the header row itself)
        if (!label) {
          const tot = String(ws[XLSX.utils.encode_cell({ r, c: col + 1 })]?.v ?? "").trim();
          label = tot || "TOTAL";
        }
        headers.push({ col, headerRow: r + 1, label });
      }
    }
  }

  // Sort by col, headerRow → determine endRow as next header in same col (-2)
  headers.sort((a, b) => a.col - b.col || a.headerRow - b.headerRow);
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i]!;
    const sameCol = headers.slice(i + 1).find((x) => x.col === h.col);
    const endRow = sameCol ? sameCol.headerRow - 2 : range.e.r + 1;
    const startRow = h.headerRow + 1;

    // Detect an optional date column in the column immediately before DAY.
    // We check: (a) the header cell of that col says "DATE", OR
    //           (b) at least one data row already has a value there.
    let dateColIdx: number | undefined;
    if (h.col > 0) {
      const dateHdrVal = String(ws[XLSX.utils.encode_cell({ r: h.headerRow - 1, c: h.col - 1 })]?.v ?? "").trim().toUpperCase();
      const hasDateHeader = dateHdrVal.includes("DATE");
      let hasDateData = false;
      for (let r = startRow; r <= endRow && !hasDateData; r++) {
        const cell = ws[XLSX.utils.encode_cell({ r: r - 1, c: h.col - 1 })];
        if (cell?.v != null && String(cell.v).trim() !== "") hasDateData = true;
      }
      if (hasDateHeader || hasDateData) dateColIdx = h.col - 1;
    }

    const entries: Array<{ row: number; date?: string; day: string; total: number }> = [];
    for (let r = startRow; r <= endRow; r++) {
      const day = String(ws[XLSX.utils.encode_cell({ r: r - 1, c: h.col })]?.v ?? "").trim();
      const tot = ws[XLSX.utils.encode_cell({ r: r - 1, c: h.col + 1 })];
      if (day && tot?.v != null) {
        const dateCell = dateColIdx != null ? ws[XLSX.utils.encode_cell({ r: r - 1, c: dateColIdx })] : undefined;
        const date = dateCell?.v != null ? String(dateCell.v).trim() : undefined;
        entries.push({ row: r, date, day, total: Number(tot.v) });
      }
    }
    tables.push({
      label: h.label,
      dateCol: dateColIdx != null ? XLSX.utils.encode_col(dateColIdx) : undefined,
      dayCol: XLSX.utils.encode_col(h.col),
      totalCol: XLSX.utils.encode_col(h.col + 1),
      startRow,
      endRow,
      entries,
    });
  }
  return tables;
}

/** Find a sub-table by a label keyword (case-insensitive, "contains"). */
function findTable(tables: HistoryTable[], keywords: string[], excludeKeywords: string[] = []): HistoryTable | undefined {
  return tables.find((t) => {
    const lbl = t.label.toUpperCase();
    if (excludeKeywords.some((k) => lbl.includes(k.toUpperCase()))) return false;
    return keywords.some((k) => lbl.includes(k.toUpperCase()));
  });
}

/**
 * Roll a new (date, day, total) into a NO COM HISTORY sub-table.
 *
 * Per the user's spec, each table is a fixed-size *rolling window* that always
 * spans "the same weekday last week → the same weekday this week". So whenever a
 * NEW day is added we drop the oldest (first) recorded entry, shift every other
 * entry up by one, and write today's values in the last slot. The number of
 * entries therefore stays constant — the table never grows.
 *
 * Exceptions:
 *   - Empty table → write today at the first row.
 *   - Re-running the analysis for the SAME date → update that entry in place
 *     (no roll), so repeated runs on one day don't shrink the window.
 *
 * Day names are written full-uppercase (e.g. "THURSDAY"); dates are MM/DD/YYYY
 * and only written when the table has a date column.
 */
async function appendOrRoll(
  tpl: XlsxTemplate,
  sheetName: string,
  table: HistoryTable,
  day: string,
  total: number,
  date: string,
  windowRows?: number,
  rolling = false,
): Promise<Array<{ day: string; total: number }>> {
  const today = normalizeDayName(day);
  const oldFootprint = table.entries.length; // rows previously populated

  // Build the sequence of entries (normalized) after applying today's value.
  let seq = table.entries.map((e) => ({
    day: normalizeDayName(e.day),
    total: e.total,
    date: e.date as string | undefined,
  }));

  if (seq.length === 0) {
    seq = [{ day: today, total, date }];
  } else {
    const last = seq[seq.length - 1]!;
    const sameDay =
      table.dateCol && last.date ? last.date === date : last.day === today;
    if (sameDay) {
      last.day = today;
      last.total = total;
      last.date = date;
    } else {
      seq.push({ day: today, total, date });
    }
  }

  // Window size: when the caller pins it (e.g. to a chart's plotted range) use
  // that; otherwise keep the table's existing size. Keep only the most recent
  // `W` entries so the window never grows.
  const W = windowRows ?? Math.max(oldFootprint, 1);
  if (rolling && seq.length > W) seq = seq.slice(seq.length - W);

  // Write the window into rows startRow .. startRow + seq.length - 1. Day labels
  // are normalized so legacy "FRI 29/05"-style values clean up as they roll.
  for (let i = 0; i < seq.length; i++) {
    const e = seq[i]!;
    const row = table.startRow + i;
    if (table.dateCol) await tpl.setCellString(sheetName, `${table.dateCol}${row}`, e.date ?? "");
    await tpl.setCellString(sheetName, `${table.dayCol}${row}`, e.day);
    await tpl.setCellNumber(sheetName, `${table.totalCol}${row}`, e.total);
  }

  // Clear any rows that were populated before but now fall outside the window
  // (e.g. shrinking a 6-row table to a 5-row charted window).
  for (let row = table.startRow + seq.length; row <= table.startRow + oldFootprint - 1; row++) {
    if (table.dateCol) await tpl.clearCell(sheetName, `${table.dateCol}${row}`);
    await tpl.clearCell(sheetName, `${table.dayCol}${row}`);
    await tpl.clearCell(sheetName, `${table.totalCol}${row}`);
  }

  return seq.map((e) => ({ day: e.day, total: e.total }));
}

type MarkedForList = {
  state: string; siteId: string; ihsId: string; vip: boolean;
  lastCommRaw: string; daysSinceComm: number | null;
};

async function generateTicketListExcel(
  toCreate: MarkedForList[],
  toResolve: TicketRow[],
  refDate: Date,
): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Telemetry Analyzer";

  const dateLabel = `${String(refDate.getUTCDate()).padStart(2,"0")}/${String(refDate.getUTCMonth()+1).padStart(2,"0")}/${refDate.getUTCFullYear()}`;
  const HDR_FILL: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E78" } };
  const CREATE_FILL: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC00000" } };
  const RESOLVE_FILL: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF375623" } };

  // ---- Sheet 1: Tickets to CREATE ----
  const ws1 = wb.addWorksheet("Tickets to Create");
  ws1.mergeCells("A1:F1");
  ws1.getCell("A1").value = `NO COM — Tickets to CREATE · ${dateLabel}`;
  ws1.getCell("A1").font = { bold: true, size: 13, color: { argb: "FFFFFFFF" } };
  ws1.getCell("A1").fill = CREATE_FILL;
  ws1.getCell("A1").alignment = { horizontal: "center" };

  ws1.getRow(2).values = ["Region", "Site", "IHS ID", "VIP", "Last Communication", "Days without comm"];
  ws1.getRow(2).font = { bold: true, color: { argb: "FFFFFFFF" } };
  ws1.getRow(2).fill = HDR_FILL;

  toCreate.forEach((s, i) => {
    const r = i + 3;
    ws1.getCell(`A${r}`).value = s.state;
    ws1.getCell(`B${r}`).value = s.siteId;
    ws1.getCell(`C${r}`).value = s.ihsId;
    ws1.getCell(`D${r}`).value = s.vip ? "YES" : "";
    ws1.getCell(`E${r}`).value = s.lastCommRaw;
    ws1.getCell(`F${r}`).value = s.daysSinceComm ?? "";
    if (s.vip) ws1.getCell(`D${r}`).font = { bold: true, color: { argb: "FFED7D31" } };
  });

  ws1.columns = [
    { width: 16 }, { width: 18 }, { width: 18 }, { width: 8 }, { width: 22 }, { width: 20 },
  ];
  ws1.autoFilter = { from: "A2", to: "F2" };
  ws1.views = [{ state: "frozen", ySplit: 2 }];
  if (toCreate.length > 0) {
    ws1.addConditionalFormatting({
      ref: `F3:F${toCreate.length + 2}`,
      rules: [
        { type: "cellIs", operator: "greaterThan", formulae: ["7"], priority: 1,
          style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFC00000" } }, font: { color: { argb: "FFFFFFFF" } } } },
        { type: "cellIs", operator: "greaterThan", formulae: ["2"], priority: 2,
          style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFFC000" } } } },
      ],
    });
  }

  // ---- Sheet 2: Tickets to RESOLVE ----
  const ws2 = wb.addWorksheet("Tickets to Resolve");
  ws2.mergeCells("A1:E1");
  ws2.getCell("A1").value = `NO COM — Tickets to RESOLVE (site OK, ticket still Open) · ${dateLabel}`;
  ws2.getCell("A1").font = { bold: true, size: 13, color: { argb: "FFFFFFFF" } };
  ws2.getCell("A1").fill = RESOLVE_FILL;
  ws2.getCell("A1").alignment = { horizontal: "center" };

  ws2.getRow(2).values = ["Region", "Site", "TTID", "Ticket Status", "Created"];
  ws2.getRow(2).font = { bold: true, color: { argb: "FFFFFFFF" } };
  ws2.getRow(2).fill = HDR_FILL;

  toResolve.forEach((t, i) => {
    const r = i + 3;
    ws2.getCell(`A${r}`).value = t.region;
    ws2.getCell(`B${r}`).value = t.site;
    ws2.getCell(`C${r}`).value = t.ttid;
    ws2.getCell(`D${r}`).value = t.status;
    ws2.getCell(`E${r}`).value = t.createTime;
  });

  ws2.columns = [{ width: 16 }, { width: 18 }, { width: 18 }, { width: 16 }, { width: 22 }];
  ws2.autoFilter = { from: "A2", to: "E2" };
  ws2.views = [{ state: "frozen", ySplit: 2 }];

  return Buffer.from(await wb.xlsx.writeBuffer());
}

const CHART_BG = "#246C82";
const CHART_BG_TREND = "#595959";

// ---------------------------------------------------------------------------
// Month-change helpers
// ---------------------------------------------------------------------------

/**
 * Returns true when the first date column in Comm Issue belongs to a month
 * different from today's month, signalling that the tracker is from a previous
 * month and needs to be reset for the new month.
 */
function detectMonthChange(trackerWb: XLSX.WorkBook, today: Date): boolean {
  const ws = trackerWb.Sheets["Comm Issue"];
  if (!ws) return false;
  const range = XLSX.utils.decode_range(ws["!ref"] ?? "A1");
  const todayMonth = today.getUTCMonth();
  const todayYear  = today.getUTCFullYear();
  let foundAnyDate = false;
  for (let c = 0; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: 0, c })];
    if (cell && cell.t === "n" && Number(cell.v) > 40000) {
      foundAnyDate = true;
      const d = new Date(Date.UTC(1899, 11, 30) + Number(cell.v) * 86_400_000);
      // If ANY date column already belongs to the current month, no reset needed
      if (d.getUTCMonth() === todayMonth && d.getUTCFullYear() === todayYear) return false;
    }
  }
  // Reset only when date columns exist and ALL are from a previous month
  return foundAnyDate;
}

/**
 * Resets the Comm Issue (and ON AIR) date columns for a new month:
 *   1. Zeroes all 1/0 data cells in every date column (single XML pass).
 *   2. Overwrites header serials with day-1, day-2, ... of the new month.
 *   3. Zeroes the cached SUBTOTAL value in each column.
 * Returns the column letter that corresponds to today's date (or the last
 * column if today's day number exceeds the number of pre-formatted columns).
 */
async function handleMonthReset(
  tpl: XlsxTemplate,
  trackerWb: XLSX.WorkBook,
  today: Date,
  todaySerial: number,
  lastDataRow: number,
  totalRow: number,
): Promise<{ todayCol: string; onAirTodayCol: string }> {
  const year = today.getUTCFullYear();
  const month = today.getUTCMonth(); // 0-based

  // -- Collect existing date columns from Comm Issue --
  const ciWs = trackerWb.Sheets["Comm Issue"]!;
  const ciRange = XLSX.utils.decode_range(ciWs["!ref"] ?? "A1");
  const ciDateCols: Array<{ col: string; colIdx: number }> = [];
  for (let c = 0; c <= ciRange.e.c; c++) {
    const cell = ciWs[XLSX.utils.encode_cell({ r: 0, c })];
    if (cell && cell.t === "n" && Number(cell.v) > 40000)
      ciDateCols.push({ col: XLSX.utils.encode_col(c), colIdx: c });
  }

  // -- Zero all 1/0 data cells in one pass --
  const ciColSet = new Set(ciDateCols.map((d) => d.col));
  await tpl.clearColumnsDataRows("Comm Issue", ciColSet, 2, lastDataRow);

  // -- Rewrite headers and SUBTOTAL cached values --
  let todayCol = ciDateCols[ciDateCols.length - 1]?.col ?? "J"; // fallback to last col
  for (let i = 0; i < ciDateCols.length; i++) {
    const dayNum = i + 1;
    const newDate = new Date(Date.UTC(year, month, dayNum));
    const newSerial = dateToExcelSerial(newDate);
    const col = ciDateCols[i]!.col;
    await tpl.setCellNumber("Comm Issue", `${col}1`, newSerial);
    const subtotalFormula = `SUBTOTAL(9,${col}2:${col}${lastDataRow})`;
    await tpl.updateCellFormula("Comm Issue", `${col}${totalRow}`, subtotalFormula, 0);
    if (newSerial === todaySerial) todayCol = col;
  }

  // -- Reset ON AIR sheet --
  let onAirTodayCol = todayCol;
  const oaWs = trackerWb.Sheets["ON AIR"];
  if (oaWs) {
    const oaRange = XLSX.utils.decode_range(oaWs["!ref"] ?? "A1");
    const oaDateCols: Array<{ col: string }> = [];
    let oaLastDataRow = 1;
    for (let r = 1; r <= oaRange.e.r; r++) {
      const cell = oaWs[XLSX.utils.encode_cell({ r, c: 1 })];
      if (cell?.v) oaLastDataRow = r + 1;
    }
    for (let c = 0; c <= oaRange.e.c; c++) {
      const cell = oaWs[XLSX.utils.encode_cell({ r: 0, c })];
      if (cell && cell.t === "n" && Number(cell.v) > 40000)
        oaDateCols.push({ col: XLSX.utils.encode_col(c) });
    }
    if (oaDateCols.length > 0) {
      const oaColSet = new Set(oaDateCols.map((d) => d.col));
      await tpl.clearColumnsDataRows("ON AIR", oaColSet, 2, oaLastDataRow);
      for (let i = 0; i < oaDateCols.length; i++) {
        const dayNum = i + 1;
        const newDate = new Date(Date.UTC(year, month, dayNum));
        const newSerial = dateToExcelSerial(newDate);
        await tpl.setCellNumber("ON AIR", `${oaDateCols[i]!.col}1`, newSerial);
        if (newSerial === todaySerial) onAirTodayCol = oaDateCols[i]!.col;
      }
    }
  }

  // -- Reset DAYS sheet date columns (blank, not zero) --
  const daysWs = trackerWb.Sheets["DAYS"];
  if (daysWs) {
    const daysRange = XLSX.utils.decode_range(daysWs["!ref"] ?? "A1");
    const daysDateCols: string[] = [];
    let daysLastDataRow = 1;
    for (let r = 1; r <= daysRange.e.r; r++) {
      if (daysWs[XLSX.utils.encode_cell({ r, c: 1 })]?.v) daysLastDataRow = r + 1;
    }
    for (let c = 0; c <= daysRange.e.c; c++) {
      const cell = daysWs[XLSX.utils.encode_cell({ r: 0, c })];
      if (cell && cell.t === "n" && Number(cell.v) > 40000)
        daysDateCols.push(XLSX.utils.encode_col(c));
    }
    if (daysDateCols.length > 0) {
      await tpl.clearColumnsDataRows("DAYS", new Set(daysDateCols), 2, daysLastDataRow);
      for (let i = 0; i < daysDateCols.length; i++) {
        const newSerial = dateToExcelSerial(new Date(Date.UTC(year, month, i + 1)));
        await tpl.setCellNumber("DAYS", `${daysDateCols[i]!}1`, newSerial);
      }
    }
  }

  // -- Reset NO COM HISTORY tables EXCEPT "more than 24hrs" --
  // The 24hrs table accumulates across months; the others (TOTAL, WITHIN 48hrs) reset.
  const histTables = readHistoryTables(trackerWb);
  const more24Table = findTable(histTables, ["24"]);
  for (const table of histTables) {
    if (more24Table && table.startRow === more24Table.startRow) continue;
    for (let row = table.startRow; row <= table.endRow; row++) {
      if (table.dateCol) await tpl.clearCell("NO COM HISTORY", `${table.dateCol}${row}`);
      await tpl.clearCell("NO COM HISTORY", `${table.dayCol}${row}`);
      await tpl.clearCell("NO COM HISTORY", `${table.totalCol}${row}`);
    }
  }

  return { todayCol, onAirTodayCol };
}

export async function analyzeNoCom(
  trackerBuffer: Buffer,
  commBuffer: Buffer,
  ticketsBuffer: Buffer,
  options: NoComOptions,
): Promise<NoComResult> {
  // 1) Parse comm report and tickets
  const allComm = parseCommReport(parseSheet(commBuffer));
  const allTickets = parseTickets(parseSheet(ticketsBuffer));

  // Filter Communication Report: drop NON-SLA rows (per the user procedure)
  const slaComm = allComm.filter((c) => isSlaProject(c.projectStatus));

  // 2) Parse the tracker ONCE (13MB+ file): decompressing the zip per-sheet was
  // costing ~600-850ms each and we used to do it 5+ times. One combined read of
  // the three sheets we touch (~1s) replaces all of that.
  const trackerWb = XLSX.read(trackerBuffer, {
    type: "buffer",
    cellFormula: true,
    sheets: ["Comm Issue", "ON AIR", "NO COM HISTORY", "DAYS"],
  });
  const { sites, totalRow, lastDataRow } = parseTrackerCommIssue(trackerWb);

  // Build VLOOKUP map: site IHS id (with trailing letter stripped) → status
  // Per user procedure: =GAUCHE(B2,NBCAR(B2)-1) on the site ID column
  const commByKey = new Map<string, CommRow>();
  for (const c of slaComm) {
    const key1 = c.site.replace(/[A-Z]$/i, ""); // drop trailing O/M
    commByKey.set(key1.toUpperCase(), c);
    commByKey.set(c.site.toUpperCase(), c);
  }

  // 3) For each tracker site, look up status
  type Marked = TrackerSite & { isNoCom: boolean; status: string; daysSinceComm: number | null; lastCommRaw: string };
  const marked: Marked[] = sites.map((s) => {
    const c = commByKey.get(s.ihsId.toUpperCase()) ?? commByKey.get(s.siteId.toUpperCase());
    const status = c?.status ?? "";
    const upper = status.toUpperCase().trim();
    // Comm report uses "No Comm" (two m's); SLA tracker historically used "NO COM".
    // Sites NOT found in the comm report (VLOOKUP → N/A) are treated as NO COM = 1.
    const isNoCom = !c
      || upper === "NO COM" || upper === "NO COMM"
      || upper === "NOCOM"  || upper === "NOCOMM"
      || upper === "NO_COM" || upper === "NO_COMM";
    const days = c?.lastCommunication
      ? Math.floor((options.referenceDate.getTime() - c.lastCommunication.getTime()) / 86_400_000)
      : null;
    return { ...s, isNoCom, status, daysSinceComm: days, lastCommRaw: c?.lastCommRaw ?? "" };
  });

  const onAir = marked.filter((s) => s.projectStatus.trim().toLowerCase() === "on air");
  const noComOnAir = onAir.filter((s) => s.isNoCom);

  // The NO COM HISTORY total table reflects the SUBTOTAL of today's column in
  // Comm Issue. The column gets a 1 for every site whose comm-report status is
  // "No Comm" (regardless of project status), so we count the same way here.
  const totalNoComCount = marked.filter((s) => s.isNoCom).length;

  // 4) Pivot Region × {NOSO, NOT NOSO, AIOT, VIP} × {COM, NO COM}
  const REGIONS = [
    "Adamawa", "Centre", "East", "Extreme North", "Littoral",
    "North", "NorthWest", "South", "SouthWest", "West",
  ];
  type CategoryKey = "NOSO" | "NOT_NOSO" | "AIOT" | "VIP";
  const pivot: Record<string, Record<CategoryKey, { com: number; noCom: number }>> = {};
  for (const reg of REGIONS) {
    pivot[reg] = {
      NOSO: { com: 0, noCom: 0 }, NOT_NOSO: { com: 0, noCom: 0 },
      AIOT: { com: 0, noCom: 0 }, VIP: { com: 0, noCom: 0 },
    };
  }
  for (const s of onAir) {
    const reg = s.state;
    if (!pivot[reg]) {
      pivot[reg] = {
        NOSO: { com: 0, noCom: 0 }, NOT_NOSO: { com: 0, noCom: 0 },
        AIOT: { com: 0, noCom: 0 }, VIP: { com: 0, noCom: 0 },
      };
    }
    const inc = (k: CategoryKey) => {
      if (s.isNoCom) pivot[reg]![k].noCom++; else pivot[reg]![k].com++;
    };
    if (s.noso) inc("NOSO");
    if (s.notNoso) inc("NOT_NOSO");
    if (s.aiot) inc("AIOT");
    if (s.vip) inc("VIP");
  }

  // 5) Tickets to create / resolve
  const cutoffMs = options.referenceDate.getTime() - options.cutoffDays * 86_400_000;
  // Normalize site keys the same way `commByKey` does (uppercase + drop the
  // trailing O/M variant letter) so ticket VLOOKUPs are immune to casing,
  // whitespace, or trailing-letter mismatches in the CMS exports.
  const ticketKey = (s: string) => s.trim().replace(/[A-Z]$/i, "").toUpperCase();
  const openCommTicketsBySite = new Map<string, TicketRow>();
  for (const t of allTickets) {
    // Per user procedure, an existing No-Com ticket only counts when:
    //   1. NOT SLA rows are dropped (keep SLA only)
    //   2. Category is "Controller" or "AIOT Controller"
    //   3. Type is "Communication"
    //   4. Status is Open
    // Only then do we VLOOKUP it against the sites.
    if (t.projectStatus && !isSlaProject(t.projectStatus)) continue;
    const cat = t.category.trim().toLowerCase();
    const type = t.type.trim().toLowerCase();
    const status = t.status.trim().toLowerCase();
    const isController = cat === "controller" || cat === "aiot controller";
    const isCommunication = type === "communication";
    if (status === "open" && isController && isCommunication) {
      const key = ticketKey(t.site);
      if (!openCommTicketsBySite.has(key)) openCommTicketsBySite.set(key, t);
    }
  }
  const ticketsToCreate: Marked[] = [];
  for (const s of noComOnAir) {
    const c = commByKey.get(s.ihsId.toUpperCase()) ?? commByKey.get(s.siteId.toUpperCase());
    if (!c?.lastCommunication) continue;
    if (c.lastCommunication.getTime() > cutoffMs) continue;
    if (openCommTicketsBySite.has(ticketKey(s.siteId)) || openCommTicketsBySite.has(ticketKey(s.ihsId))) continue;
    ticketsToCreate.push(s);
  }
  ticketsToCreate.sort((a, b) => a.state.localeCompare(b.state) || a.ihsId.localeCompare(b.ihsId));

  const onAirByKey = new Map<string, Marked>();
  for (const s of onAir) {
    onAirByKey.set(ticketKey(s.siteId), s);
    onAirByKey.set(ticketKey(s.ihsId), s);
  }
  const ticketsToResolve: TicketRow[] = [];
  for (const [key, t] of openCommTicketsBySite) {
    const s = onAirByKey.get(key);
    if (s && !s.isNoCom) ticketsToResolve.push(t);
  }
  ticketsToResolve.sort((a, b) => a.region.localeCompare(b.region) || a.site.localeCompare(b.site));

  // 6) Counts for the two NO COM HISTORY trend tables.
  //
  //   • "NO COM SITES MORE THAN 24hrs" — straight from the Communication Report:
  //      filter to NO COM, then drop everything whose last_comm is within the last
  //      2 days (i.e. keep sites whose outage is older than the cutoff). The result
  //      is the count of remaining sites.
  //   • "NO COM WITHIN 48 hrs" — symmetric: NO COM with last_comm in the last 2 days.
  const cutoffMs2 = options.referenceDate.getTime() - options.cutoffDays * 86_400_000;
  const noComInReport = slaComm.filter((c) => {
    const u = c.status.toUpperCase().trim();
    return u === "NO COM" || u === "NO COMM" || u === "NOCOM" || u === "NOCOMM";
  });
  // Use <= so a site whose lastCommunication falls exactly on the cutoff day
  // (midnight) is counted in moreThan24h, matching the user's manual count.
  const moreThan24h = noComInReport.filter(
    (c) => c.lastCommunication && c.lastCommunication.getTime() <= cutoffMs2,
  ).length;
  const within48h = noComInReport.filter(
    (c) => !c.lastCommunication || c.lastCommunication.getTime() > cutoffMs2,
  ).length;

  // ---------------------------------------------------------------------------
  // 7) SURGICAL EDIT of the SLA Tracker template
  // ---------------------------------------------------------------------------
  const tpl = await XlsxTemplate.load(trackerBuffer);

  // 7a) Find today's date column in Comm Issue (and mirror to ON AIR — same headers).
  //
  // Three cases handled automatically:
  //   A) Month changed — reset all date columns to the new month, zero all data.
  //   B) Same month, column missing — auto-insert today's column after the last one.
  //   C) Same month, column present — normal update.
  const todaySerial = dateToExcelSerial(options.referenceDate);
  let todayCol: string;
  let onAirTodayCol: string;
  let todayColIsNew = false;

  let monthReset = false;
  if (detectMonthChange(trackerWb, options.referenceDate)) {
    monthReset = true;
    // Case A: new month — reset the whole date grid and get today's column
    const reset = await handleMonthReset(
      tpl, trackerWb, options.referenceDate, todaySerial, lastDataRow, totalRow,
    );
    todayCol = reset.todayCol;
    onAirTodayCol = reset.onAirTodayCol;
    // After reset the cells exist (zeroed) — use updateColumnValues, not insert
    todayColIsNew = false;
  } else {
    const existing = findDateColumn(trackerWb, "Comm Issue", todaySerial);
    if (existing) {
      // Case C: normal update
      todayCol = existing;
      onAirTodayCol = findDateColumn(trackerWb, "ON AIR", todaySerial) ?? todayCol;
    } else {
      // Case B: new day in same month — auto-insert after last date column
      const ciWs = trackerWb.Sheets["Comm Issue"];
      if (!ciWs) throw new Error('Sheet "Comm Issue" not found in tracker.');
      const ciRange = XLSX.utils.decode_range(ciWs["!ref"] ?? "A1");
      let lastDateColIdx = -1;
      for (let c = 0; c <= ciRange.e.c; c++) {
        const cell = ciWs[XLSX.utils.encode_cell({ r: 0, c })];
        if (cell && cell.t === "n" && Number(cell.v) > 40000) lastDateColIdx = c;
      }
      if (lastDateColIdx < 0) throw new Error('No date columns found in "Comm Issue" header. Please check your tracker template.');
      todayCol = XLSX.utils.encode_col(lastDateColIdx + 1);
      todayColIsNew = true;
      await tpl.setCellNumber("Comm Issue", `${todayCol}1`, todaySerial);
      // Mirror to ON AIR
      const oaWs = trackerWb.Sheets["ON AIR"];
      onAirTodayCol = todayCol;
      if (oaWs) {
        const oaRange = XLSX.utils.decode_range(oaWs["!ref"] ?? "A1");
        let lastOaDateColIdx = -1;
        for (let c = 0; c <= oaRange.e.c; c++) {
          const cell = oaWs[XLSX.utils.encode_cell({ r: 0, c })];
          if (cell && cell.t === "n" && Number(cell.v) > 40000) lastOaDateColIdx = c;
        }
        if (lastOaDateColIdx >= 0) {
          onAirTodayCol = XLSX.utils.encode_col(lastOaDateColIdx + 1);
          await tpl.setCellNumber("ON AIR", `${onAirTodayCol}1`, todaySerial);
        }
      }
    }
  }

  // 7b) Write 1/0 in today's column for each site row in Comm Issue
  const valuesByRow = new Map<number, number>();
  for (const s of marked) {
    valuesByRow.set(s.rowIndex, s.isNoCom ? 1 : 0);
  }
  if (todayColIsNew) {
    await tpl.insertColumnValues("Comm Issue", todayCol, valuesByRow);
  } else {
    await tpl.updateColumnValues("Comm Issue", todayCol, valuesByRow);
  }

  // 7b-bis) Mirror today's column into ON AIR (the source of the NO COM STATUS pivot).
  // ON AIR is the On-Air subset of Comm Issue and uses the same headers; we match
  // by Site ID so row numbers don't have to align.
  const onAirIndex = parseOnAirSiteIndex(trackerWb);
  const onAirValuesByRow = new Map<number, number>();
  for (const s of marked) {
    const rowInOnAir = onAirIndex.get(s.siteId.toUpperCase());
    if (rowInOnAir != null) onAirValuesByRow.set(rowInOnAir, s.isNoCom ? 1 : 0);
  }
  if (todayColIsNew) {
    await tpl.insertColumnValues("ON AIR", onAirTodayCol, onAirValuesByRow);
  } else {
    await tpl.updateColumnValues("ON AIR", onAirTodayCol, onAirValuesByRow);
  }

  // 7c) Write SUBTOTAL formula at the total row of today's column. The cached
  // result is the count of all 1s we just wrote (= totalNoComCount).
  const subtotalFormula = `SUBTOTAL(9,${todayCol}2:${todayCol}${lastDataRow})`;
  await tpl.updateCellFormula("Comm Issue", `${todayCol}${totalRow}`, subtotalFormula, totalNoComCount);

  // 7d) Append today's data to NO COM HISTORY (3 sub-tables). The total table
  // takes the SUBTOTAL value (totalNoComCount), not the on-air filtered count.
  // DAY label is the full English weekday name (e.g. "THURSDAY") — matching the
  // existing entries. The date (MM/DD/YYYY) is written to the date column if one
  // is detected in the template.
  const dayName = options.referenceDate.toLocaleDateString("en-US", { weekday: "long" }).toUpperCase();
  const todayDate = fmtMDY(options.referenceDate);
  const tables = readHistoryTables(trackerWb);
  const totalTable    = findTable(tables, ["TOTAL", "NO COM"], ["24", "48"]);
  const more24Table   = findTable(tables, ["24"]);
  const within48Table = findTable(tables, ["48"]);

  if (totalTable)    await appendOrRoll(tpl, "NO COM HISTORY", totalTable, dayName, totalNoComCount, todayDate);
  if (more24Table)   await appendOrRoll(tpl, "NO COM HISTORY", more24Table, dayName, moreThan24h, todayDate, undefined, true);
  if (within48Table) await appendOrRoll(tpl, "NO COM HISTORY", within48Table, dayName, within48h, todayDate);

  // 7d-bis) Update DAYS sheet: write today's NO COM flag (1=NO COM, 0=COM) for
  // each site row. The DAYS table already has formula columns:
  //   • "Total no com days" = SUM of daily flags across the month
  //   • "ToutNoCom"         = (Total = NbJoursTravailles)  ← TRUE if NO COM all month
  // "Prolonged No Com" sheet uses COUNTIFS(T_DAYS[ToutNoCom],TRUE) per region,
  // so it auto-refreshes in Excel once the daily column is populated here.
  const daysTodayCol = findDateColumn(trackerWb, "DAYS", todaySerial);
  if (daysTodayCol) {
    const daysWs = trackerWb.Sheets["DAYS"];
    if (daysWs) {
      const daysRange = XLSX.utils.decode_range(daysWs["!ref"] ?? "A1");
      const dayRowValues = new Map<number, number>();
      for (let r = 1; r <= daysRange.e.r; r++) {
        const ihsCell = daysWs[XLSX.utils.encode_cell({ r, c: 0 })];
        if (!ihsCell?.v) continue;
        const ihsId = String(ihsCell.v).trim();
        const comm =
          commByKey.get(ihsId.toUpperCase()) ??
          commByKey.get(ihsId.replace(/[A-Z]$/i, "").toUpperCase());
        const u = comm ? comm.status.toUpperCase().trim() : "";
        const isNoComToday =
          u === "NO COM" || u === "NO COMM" || u === "NOCOM" || u === "NOCOMM";
        dayRowValues.set(r + 1, isNoComToday ? 1 : 0); // r+1: XLSX r is 0-based, Excel rows are 1-based
      }
      await tpl.updateColumnValues("DAYS", daysTodayCol, dayRowValues);
    }
  }

  // 7e) Extend the NO COM STATUS pivot cache (cacheDef #2) so today's column is
  // included, then re-target the pivot's column-axis field to today's date.
  // This makes the pivot in the NOCOM STATUS sheet show today's COM/NO COM split
  // by region, refreshed automatically when Excel opens the file.
  const dateHeaders = readOnAirDateHeaders(trackerWb);
  // Cache currently spans cols A..(end of source range). We figure out the new
  // last column letter (today's) and which date columns are *new* (between the
  // previous end and today, inclusive).
  const cacheEndOldLetter = await tpl.getPivotCacheEndColumn(2);
  const todayColIdx = colLetterToIndex(onAirTodayCol);
  const cacheEndOldIdx = cacheEndOldLetter ? colLetterToIndex(cacheEndOldLetter) : -1;
  if (todayColIdx > cacheEndOldIdx) {
    const newDates: string[] = [];
    for (let c = cacheEndOldIdx + 1; c <= todayColIdx; c++) {
      const colL = colLetter(c);
      const hdr = dateHeaders.find((h) => h.col === colL);
      const name = hdr ? fmtMDY(excelSerialToDate(hdr.serial)) : `col_${colL}`;
      newDates.push(name);
    }
    await tpl.extendPivotCacheToColumn(2, onAirTodayCol, newDates);
  }
  // Pivot field index for today = (cacheFields count - 1) = colIdx in source ref.
  // Since we extended the source ref to start at A (col 0), the field index = todayColIdx.
  await tpl.retargetPivotColumnField(1, todayColIdx);

  // 7f) Update chart titles with today's date ranges:
  //   • chart1 + chart4: "NO COM SITES MORE THAN 24hrs" → 7-day window ending today
  //   • chart3: "NO COM STATUS SLA" → current month (1st → last day)
  const sevenAgo = new Date(options.referenceDate.getTime() - 7 * 86_400_000);
  const dayWindow = `${fmtMDY(sevenAgo)}-${fmtMDY(options.referenceDate)}`;
  const dateInline = /<a:t>\d{1,2}\/\d{1,2}\/\d{4}-\d{1,2}\/\d{1,2}\/\d{4}<\/a:t>/g;
  await tpl.updateChartTitle(1, [{ from: dateInline, to: `<a:t>${dayWindow}</a:t>` }]);
  await tpl.updateChartTitle(4, [{ from: dateInline, to: `<a:t>${dayWindow}</a:t>` }]);

  const monthStart = new Date(Date.UTC(
    options.referenceDate.getUTCFullYear(),
    options.referenceDate.getUTCMonth(), 1,
  ));
  const monthEnd = new Date(Date.UTC(
    options.referenceDate.getUTCFullYear(),
    options.referenceDate.getUTCMonth() + 1, 0,
  ));
  await tpl.updateChartTitle(3, [
    { from: /<a:t>\d{1,2}\/\d{1,2}\/\d{4}<\/a:t>/, to: `<a:t>${fmtMDY(monthStart)}</a:t>` },
    { from: /<a:t> - \d{1,2}\/\d{1,2}\/\d{4}<\/a:t>/, to: `<a:t> - ${fmtMDY(monthEnd)}</a:t>` },
  ]);

  // 7g) Update Sheet9: write today's date to B3 and write the full pivot data
  // (region × NOSO/NOT NOSO/AIOT/VIP × COM/NO COM) so the table stays in sync
  // with the ON AIR pivot without requiring manual updates.
  // Write as an explicit MM/DD/YYYY string so the 4-digit year always shows
  // (the cell's stored number format was m/d/yy, which truncated to "5/28/26").
  await tpl.setCellString("Sheet9", "B3", fmtMDY(options.referenceDate));
  const SHEET9_ROWS: Record<string, number> = {
    Adamawa: 6, Centre: 7, East: 8, "Extreme North": 9, Littoral: 10,
    North: 11, NorthWest: 12, South: 13, SouthWest: 14, West: 15,
  };
  for (const reg of REGIONS) {
    const row = SHEET9_ROWS[reg];
    if (!row) continue;
    const p = pivot[reg];
    if (!p) continue;
    // Template layout: C=NOSO COM, D=NOSO NO COM, E=NOT NOSO COM, F=NOT NOSO NO COM,
    //                  G=AIOT COM, H=AIOT NO COM, I=VIP COM, J=VIP NO COM
    if (p.NOSO.com > 0)     await tpl.setCellNumber("Sheet9", `C${row}`, p.NOSO.com);
    if (p.NOSO.noCom > 0)   await tpl.setCellNumber("Sheet9", `D${row}`, p.NOSO.noCom);
    if (p.NOT_NOSO.com > 0) await tpl.setCellNumber("Sheet9", `E${row}`, p.NOT_NOSO.com);
    if (p.NOT_NOSO.noCom > 0) await tpl.setCellNumber("Sheet9", `F${row}`, p.NOT_NOSO.noCom);
    if (p.AIOT.com > 0)     await tpl.setCellNumber("Sheet9", `G${row}`, p.AIOT.com);
    if (p.AIOT.noCom > 0)   await tpl.setCellNumber("Sheet9", `H${row}`, p.AIOT.noCom);
    if (p.VIP.com > 0)      await tpl.setCellNumber("Sheet9", `I${row}`, p.VIP.com);
    if (p.VIP.noCom > 0)    await tpl.setCellNumber("Sheet9", `J${row}`, p.VIP.noCom);
  }
  // Total row (row 16)
  await tpl.setCellNumber("Sheet9", "C16", sumCol(pivot, "NOSO", "com"));
  await tpl.setCellNumber("Sheet9", "D16", sumCol(pivot, "NOSO", "noCom"));
  await tpl.setCellNumber("Sheet9", "E16", sumCol(pivot, "NOT_NOSO", "com"));
  await tpl.setCellNumber("Sheet9", "F16", sumCol(pivot, "NOT_NOSO", "noCom"));
  await tpl.setCellNumber("Sheet9", "G16", sumCol(pivot, "AIOT", "com"));
  await tpl.setCellNumber("Sheet9", "H16", sumCol(pivot, "AIOT", "noCom"));
  await tpl.setCellNumber("Sheet9", "I16", sumCol(pivot, "VIP", "com"));
  await tpl.setCellNumber("Sheet9", "J16", sumCol(pivot, "VIP", "noCom"));

  // 7h) Make pivot tables refresh on open + recompute formulas
  await tpl.setPivotCachesRefreshOnLoad();
  await tpl.setFullCalcOnLoad();

  const buffer = await tpl.toBuffer();

  const dateStr = formatDate(options.referenceDate).replace(/\//g, "-");
  const filename = `SLA_TRACKER_UPDATES_F1_${dateStr}.xlsx`;

  // 8) Generate the standalone ticket-list Excel (two sheets: to create + to resolve)
  const ticketListBuffer = await generateTicketListExcel(ticketsToCreate, ticketsToResolve, options.referenceDate);
  const ticketListFilename = `NoCom_TicketList_${dateStr}.xlsx`;

  return {
    buffer,
    filename,
    ticketListBuffer,
    ticketListFilename,
    summary: {
      totalSites: sites.length,
      onAirSites: onAir.length,
      noComSites: noComOnAir.length,
      ticketsToCreate: ticketsToCreate.length,
      ticketsToResolve: ticketsToResolve.length,
      moreThan24h,
      within48h,
      todayColumn: todayCol,
      todayDateSerial: todaySerial,
      historyEntries: (totalTable?.entries.length ?? 0) + 1,
      monthReset,
      byCategory: {
        NOSO: { com: sumCol(pivot, "NOSO", "com"), noCom: sumCol(pivot, "NOSO", "noCom") },
        "NOT NOSO": { com: sumCol(pivot, "NOT_NOSO", "com"), noCom: sumCol(pivot, "NOT_NOSO", "noCom") },
        AIOT: { com: sumCol(pivot, "AIOT", "com"), noCom: sumCol(pivot, "AIOT", "noCom") },
        VIP: { com: sumCol(pivot, "VIP", "com"), noCom: sumCol(pivot, "VIP", "noCom") },
      },
      ticketsToCreateList: ticketsToCreate.slice(0, 200).map((s) => ({
        region: s.state, site: s.siteId, ihsId: s.ihsId, vip: s.vip,
        lastComm: s.lastCommRaw, daysSinceComm: s.daysSinceComm,
      })),
      ticketsToResolveList: ticketsToResolve.slice(0, 200).map((t) => ({
        ttid: t.ttid, region: t.region, site: t.site, status: t.status,
      })),
    },
  };
}

function sumCol(
  pivot: Record<string, Record<string, { com: number; noCom: number }>>,
  cat: string,
  field: "com" | "noCom",
): number {
  let sum = 0;
  for (const reg of Object.keys(pivot)) sum += pivot[reg]![cat]![field];
  return sum;
}

function formatDate(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const da = String(d.getDate()).padStart(2, "0");
  return `${da}/${m}/${y}`;
}

// Re-exports kept for compatibility (renderLineChart no longer used but we keep
// the import path stable if other modules reference it later)
export { renderLineChart };
