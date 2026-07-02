import ExcelJS from "exceljs";
import { parseSheet, isSlaProject, type Row } from "../lib/excel-utils";

export interface FuelSensorResult {
  buffer: Buffer;
  filename: string;
  summary: {
    totalComponents: number;
    uniqueSites: number;
    sitesWithOpenTicket: number;
    sitesNeedingTicket: number;
  };
}

/** Find the first key in `row` whose name (case-insensitive) matches any candidate. */
function findKey(row: Row, candidates: string[]): string | undefined {
  for (const c of candidates) {
    if (row[c] !== undefined) return c;
    const found = Object.keys(row).find((k) => k.toLowerCase() === c.toLowerCase());
    if (found !== undefined) return found;
  }
  return undefined;
}

/** Resolve column name by scanning the first data row (whose keys = header names). */
function colOf(rows: Row[], candidates: string[]): string | undefined {
  if (!rows.length) return undefined;
  return findKey(rows[0]!, candidates);
}

export async function analyzeFuelSensor(
  componentsBuffer: Buffer,
  ticketsBuffer: Buffer,
): Promise<FuelSensorResult> {
  // ── 1. Parse and filter the Components file ──────────────────────────────
  const compRows = parseSheet(componentsBuffer);

  const siteCol      = colOf(compRows, ["Site", "SiteName", "Site Name", "SiteId", "Site ID", "SITE"]);
  const locCol       = colOf(compRows, ["Location", "LOCATION"]);
  const valCol       = colOf(compRows, ["Value", "VALUE", "Quantity", "Quant"]);
  const projStatCol  = colOf(compRows, ["ProjectStatus", "Project Status", "PROJECT STATUS"]);

  // Rows where SLA + Location = "Quant" and Value ∈ {0, -1}
  const filtered = compRows.filter((row) => {
    const projStatus = String(row[projStatCol ?? "ProjectStatus"] ?? "").trim();
    if (!isSlaProject(projStatus)) return false;
    const loc = String(row[locCol ?? "Location"] ?? "").trim().toLowerCase();
    const raw = row[valCol ?? "Value"];
    const val = Number(raw);
    return loc === "quant" && (val === 0 || val === -1);
  });

  // ── 2. Parse the Fuel Sensor Tickets file ────────────────────────────────
  const tickRows = parseSheet(ticketsBuffer);

  const tSiteCol      = colOf(tickRows, ["Site", "SiteName", "Site Name", "SiteId", "Site ID"]);
  const tStatusCol    = colOf(tickRows, ["Status", "Ticket Status", "STATUS"]);
  const tTtidCol      = colOf(tickRows, ["TTID", "ttid", "Ticket ID", "TicketID", "Ticket_ID", "ID"]);
  const tTypeCol      = colOf(tickRows, ["Type", "Ticket Type", "Category", "Problem Type"]);
  const tProjStatCol  = colOf(tickRows, ["Project Status", "ProjectStatus", "PROJECT STATUS"]);

  // Map: siteKey → { ttid, status } — keep only open SLA tickets
  const openTickets = new Map<string, { ttid: string; status: string; type: string }>();
  for (const row of tickRows) {
    const site      = String(row[tSiteCol      ?? "Site"]           ?? "").trim();
    const status    = String(row[tStatusCol    ?? "Status"]         ?? "").trim();
    const ttid      = String(row[tTtidCol      ?? "TTID"]           ?? "").trim();
    const type      = String(row[tTypeCol      ?? "Type"]           ?? "").trim();
    const projStat  = String(row[tProjStatCol  ?? "Project Status"] ?? "").trim();
    if (!site) continue;
    // Skip non-SLA tickets
    if (!isSlaProject(projStat)) continue;
    // Consider a ticket closed if status contains "close", "resolv", "done", "cancel"
    const isClose = /close|resolv|done|cancel/i.test(status);
    if (!isClose) {
      openTickets.set(site.toLowerCase(), { ttid, status, type });
    }
  }

  // ── 3. VLOOKUP result ────────────────────────────────────────────────────
  const allOrigCols = filtered.length > 0
    ? Object.keys(filtered[0]!).filter((k) => !k.startsWith("_"))
    : [];

  const vlookupRows = filtered.map((row) => {
    const site = String(row[siteCol ?? "Site"] ?? "").trim();
    const ticket = openTickets.get(site.toLowerCase()) ?? null;
    return { row, site, hasTicket: !!ticket, ticket };
  });

  const allSites = [...new Set(vlookupRows.map((r) => r.site).filter(Boolean))];
  const sitesWithOpen  = allSites.filter((s) => openTickets.has(s.toLowerCase()));
  const sitesNeedingTicket = allSites.filter((s) => !openTickets.has(s.toLowerCase()));

  // ── 4. Build output Excel ─────────────────────────────────────────────────
  const wb = new ExcelJS.Workbook();
  wb.creator = "Telemetry Analyzer";

  const HDR_FILL:    ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E78" } };
  const TITLE_FILL:  ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF833C00" } };
  const OK_FILL:     ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF375623" } };
  const WARN_FILL:   ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC00000" } };
  const WHITE_FONT: Partial<ExcelJS.Font> = { bold: true, color: { argb: "FFFFFFFF" } };

  // Sheet 1 — full VLOOKUP table ──────────────────────────────────────────
  const ws1 = wb.addWorksheet("VLOOKUP Table");
  const extraCols = ["Has Open Ticket", "Ticket ID", "Ticket Status", "Ticket Type"];
  const hdrTotal = allOrigCols.length + extraCols.length;

  ws1.mergeCells(1, 1, 1, hdrTotal);
  const title1 = ws1.getCell("A1");
  title1.value = `FUEL SENSOR — Components at 0 or −1 (Location = Quant) · ${filtered.length} rows`;
  title1.font = { ...WHITE_FONT, size: 13 };
  title1.fill = TITLE_FILL;
  title1.alignment = { horizontal: "center" };

  ws1.getRow(2).values = [...allOrigCols, ...extraCols];
  ws1.getRow(2).font = WHITE_FONT;
  ws1.getRow(2).fill = HDR_FILL;

  vlookupRows.forEach(({ row, hasTicket, ticket }, i) => {
    const r = i + 3;
    allOrigCols.forEach((h, ci) => { ws1.getCell(r, ci + 1).value = row[h] ?? ""; });
    const base = allOrigCols.length;
    const hasCell = ws1.getCell(r, base + 1);
    hasCell.value = hasTicket ? "YES" : "NO";
    hasCell.fill  = hasTicket ? OK_FILL : WARN_FILL;
    hasCell.font  = WHITE_FONT;
    ws1.getCell(r, base + 2).value = ticket?.ttid   ?? "";
    ws1.getCell(r, base + 3).value = ticket?.status ?? "";
    ws1.getCell(r, base + 4).value = ticket?.type   ?? "";
  });

  ws1.columns = [
    ...allOrigCols.map(() => ({ width: 18 as number })),
    { width: 16 }, { width: 18 }, { width: 18 }, { width: 18 },
  ];
  if (filtered.length > 0) {
    ws1.autoFilter = {
      from: "A2",
      to: ws1.getCell(2, hdrTotal).address,
    };
  }
  ws1.views = [{ state: "frozen", ySplit: 2 }];

  // Sheet 2 — sites needing new tickets ───────────────────────────────────
  const ws2 = wb.addWorksheet("Sites Needing Tickets");
  ws2.mergeCells("A1:B1");
  const title2 = ws2.getCell("A1");
  title2.value = `FUEL SENSOR — Sites with NO open ticket · ${sitesNeedingTicket.length} sites`;
  title2.font = { ...WHITE_FONT, size: 13 };
  title2.fill = WARN_FILL;
  title2.alignment = { horizontal: "center" };

  ws2.getRow(2).values = ["#", "Site"];
  ws2.getRow(2).font = WHITE_FONT;
  ws2.getRow(2).fill = HDR_FILL;

  sitesNeedingTicket.forEach((site, i) => {
    ws2.getCell(i + 3, 1).value = i + 1;
    ws2.getCell(i + 3, 2).value = site;
  });
  ws2.columns = [{ width: 6 }, { width: 28 }];

  // Sheet 3 — sites already covered ───────────────────────────────────────
  const ws3 = wb.addWorksheet("Sites Already Covered");
  ws3.mergeCells("A1:C1");
  const title3 = ws3.getCell("A1");
  title3.value = `FUEL SENSOR — Sites with open ticket · ${sitesWithOpen.length} sites`;
  title3.font = { ...WHITE_FONT, size: 13 };
  title3.fill = OK_FILL;
  title3.alignment = { horizontal: "center" };

  ws3.getRow(2).values = ["#", "Site", "Ticket ID"];
  ws3.getRow(2).font = WHITE_FONT;
  ws3.getRow(2).fill = HDR_FILL;

  sitesWithOpen.forEach((site, i) => {
    const tk = openTickets.get(site.toLowerCase());
    ws3.getCell(i + 3, 1).value = i + 1;
    ws3.getCell(i + 3, 2).value = site;
    ws3.getCell(i + 3, 3).value = tk?.ttid ?? "";
  });
  ws3.columns = [{ width: 6 }, { width: 28 }, { width: 18 }];

  const d = new Date();
  const dateStr = `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,"0")}${String(d.getDate()).padStart(2,"0")}`;
  const filename = `FuelSensor_Analysis_${dateStr}.xlsx`;
  const buffer = Buffer.from(await wb.xlsx.writeBuffer());

  return {
    buffer,
    filename,
    summary: {
      totalComponents: filtered.length,
      uniqueSites: allSites.length,
      sitesWithOpenTicket: sitesWithOpen.length,
      sitesNeedingTicket: sitesNeedingTicket.length,
    },
  };
}
