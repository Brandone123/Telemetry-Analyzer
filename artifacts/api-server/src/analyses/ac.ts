import ExcelJS from "exceljs";
import { parseSheet, isSlaProject, type Row } from "../lib/excel-utils";

const AC_ALERT_CODES = new Set(["AL690-0055", "EM690-0001", "E_ACM-0001", "USER-0016"]);

export interface AcAnalysisResult {
  buffer: Buffer;
  filename: string;
  summary: {
    m1TotalAlerts: number;
    m1UniqueSites: number;
    m1SitesNeedingTicket: number;
    m1ByCode: Record<string, number>;
    m1ByRegion: Record<string, number>;
    m2TotalSites: number;
    m2SitesNeedingTicket: number;
    m2ByRegion: Record<string, number>;
    falseTickets: number;
    dateRange: string[];
  };
}

function findKey(row: Row, candidates: string[]): string | undefined {
  for (const c of candidates) {
    const found = Object.keys(row).find(
      (k) => k.toLowerCase().trim() === c.toLowerCase().trim(),
    );
    if (found !== undefined) return found;
  }
  return undefined;
}

function colOf(rows: Row[], candidates: string[]): string | undefined {
  if (!rows.length) return undefined;
  return findKey(rows[0]!, candidates);
}

function str(row: Row, col: string | undefined, fallback: string): string {
  return String(row[col ?? fallback] ?? "").trim();
}

export async function analyzeAC(
  alertsBuffer: Buffer,
  ersBigDataBuffer: Buffer,
  ticketsBuffer: Buffer | null,
): Promise<AcAnalysisResult> {
  const alertRows = parseSheet(alertsBuffer);
  const ersRows   = parseSheet(ersBigDataBuffer);
  const ticketRows = ticketsBuffer ? parseSheet(ticketsBuffer) : [];

  // ── 1. Parse tickets: open ticket set + full info for false-ticket detection ─
  interface OpenTicket {
    site: string;
    region: string;
    ticketId: string;
    status: string;
    category: string;
    projectStatus: string;
  }
  const openTicketsBySite = new Map<string, OpenTicket[]>();

  if (ticketRows.length > 0) {
    const tSiteCol   = colOf(ticketRows, ["Site", "SiteName", "Site Name"]);
    const tStatusCol = colOf(ticketRows, ["Status", "Ticket Status"]);
    const tProjCol   = colOf(ticketRows, ["Project Status", "ProjectStatus"]);
    const tIdCol     = colOf(ticketRows, ["Ticket ID", "TicketID", "ID", "Reference"]);
    const tCatCol    = colOf(ticketRows, ["Category", "Type", "Alarm Type"]);
    const tRegCol    = colOf(ticketRows, ["Region", "RegionName"]);

    for (const row of ticketRows) {
      const site    = str(row, tSiteCol, "Site");
      const status  = str(row, tStatusCol, "Status");
      const projStat = str(row, tProjCol, "Project Status");
      if (!site) continue;
      if (!isSlaProject(projStat)) continue;
      const isClosed = /close|resolv|done|cancel/i.test(status);
      if (isClosed) continue;
      const key = site.toLowerCase();
      if (!openTicketsBySite.has(key)) openTicketsBySite.set(key, []);
      openTicketsBySite.get(key)!.push({
        site,
        region:        str(row, tRegCol, "Region"),
        ticketId:      str(row, tIdCol, "Ticket ID"),
        status,
        category:      str(row, tCatCol, "Category"),
        projectStatus: projStat,
      });
    }
  }

  // Convenience: simple set for O(1) "has open ticket" checks
  const openTickets = new Set(openTicketsBySite.keys());

  // ── 2. Method 1 — Alert code filtering ────────────────────────────────────
  const m1CodeCol    = colOf(alertRows, ["Alert Code", "AlertCode"]);
  const m1SiteCol    = colOf(alertRows, ["Site", "SiteName", "Site Name"]);
  const m1RegionCol  = colOf(alertRows, ["Region", "RegionName"]);
  const m1AddrCol    = colOf(alertRows, ["Address"]);
  const m1PriorityCol = colOf(alertRows, ["Site Priority", "Priority"]);
  const m1ProjStatCol = colOf(alertRows, ["Project Status", "ProjectStatus"]);
  const m1TopoCol    = colOf(alertRows, ["Power Topology", "PowerTopology"]);
  const m1DescCol    = colOf(alertRows, ["Description"]);

  const m1Filtered = alertRows.filter((r) => {
    const code = str(r, m1CodeCol, "Alert Code");
    const ps   = str(r, m1ProjStatCol, "Project Status");
    return AC_ALERT_CODES.has(code) && isSlaProject(ps);
  });

  const m1ByCode: Record<string, number> = {};
  const m1BySite = new Map<string, {
    site: string; region: string; address: string; priority: string;
    projectStatus: string; topology: string; codes: Set<string>;
    descriptions: Set<string>; count: number;
  }>();

  for (const row of m1Filtered) {
    const site = str(row, m1SiteCol, "Site");
    const code = str(row, m1CodeCol, "Alert Code");
    m1ByCode[code] = (m1ByCode[code] ?? 0) + 1;
    if (!m1BySite.has(site)) {
      m1BySite.set(site, {
        site,
        region:        str(row, m1RegionCol, "Region"),
        address:       str(row, m1AddrCol, "Address"),
        priority:      str(row, m1PriorityCol, "Site Priority"),
        projectStatus: str(row, m1ProjStatCol, "Project Status"),
        topology:      str(row, m1TopoCol, "Power Topology"),
        codes: new Set(),
        descriptions: new Set(),
        count: 0,
      });
    }
    const s = m1BySite.get(site)!;
    s.codes.add(code);
    s.descriptions.add(str(row, m1DescCol, "Description"));
    s.count++;
  }

  const m1SiteList         = [...m1BySite.values()];
  const m1SitesNeedingTicket = m1SiteList.filter((s) => !openTickets.has(s.site.toLowerCase()));
  const m1ByRegion: Record<string, number> = {};
  m1SitesNeedingTicket.forEach((s) => { m1ByRegion[s.region] = (m1ByRegion[s.region] ?? 0) + 1; });

  // ── 3. Method 2 — ERS Big Data (5 days) ───────────────────────────────────
  const ersRegionCol  = colOf(ersRows, ["RegionName", "Region"]);
  const ersSiteCol    = colOf(ersRows, ["SiteName", "Site"]);
  const ersProjStatCol = colOf(ersRows, ["ProjectStatus", "Project Status"]);
  const ersDayCol     = colOf(ersRows, ["Day", "Date"]);
  const ersTopoCol    = colOf(ersRows, ["PowerTopology", "Power Topology"]);
  const ersTotalRHCol = colOf(ersRows, ["TotalPower RH", "Total Power RH", "TotalPowerRH"]);
  const ersTotalKWhCol = colOf(ersRows, ["TotalPower KWh", "Total Power KWh", "TotalPowerKWh"]);
  const ersNoCommCol  = colOf(ersRows, ["No Comm. H", "No Comm H", "NoCommH", "No Comm"]);

  const ersSlaRows = ersRows.filter((r) =>
    isSlaProject(str(r, ersProjStatCol, "ProjectStatus")),
  );

  const allDates = [...new Set(
    ersSlaRows.map((r) => str(r, ersDayCol, "Day")).filter(Boolean),
  )].sort();
  const last5Dates = allDates.slice(-5);

  const m2BySite = new Map<string, {
    site: string; region: string; projectStatus: string; topology: string;
    totalRH: number; totalKWh: number; noCommH: number; days: number;
  }>();

  for (const row of ersSlaRows) {
    const day = str(row, ersDayCol, "Day");
    if (!last5Dates.includes(day)) continue;
    const site = str(row, ersSiteCol, "SiteName");
    if (!site) continue;
    if (!m2BySite.has(site)) {
      m2BySite.set(site, {
        site,
        region:        str(row, ersRegionCol, "RegionName"),
        projectStatus: str(row, ersProjStatCol, "ProjectStatus"),
        topology:      str(row, ersTopoCol, "PowerTopology"),
        totalRH: 0, totalKWh: 0, noCommH: 0, days: 0,
      });
    }
    const s = m2BySite.get(site)!;
    s.totalRH  += parseFloat(String(row[ersTotalRHCol  ?? "TotalPower RH"]  ?? "0")) || 0;
    s.totalKWh += parseFloat(String(row[ersTotalKWhCol ?? "TotalPower KWh"] ?? "0")) || 0;
    s.noCommH  += parseFloat(String(row[ersNoCommCol   ?? "No Comm. H"]     ?? "0")) || 0;
    s.days++;
  }

  // M2: communicating (noCommH=0) + continuous power (≥119.5h) + zero AC consumption
  const m2Matching          = [...m2BySite.values()].filter(
    (s) => s.noCommH === 0 && s.totalRH >= 119.5 && s.totalKWh === 0,
  );
  const m2SitesNeedingTicket = m2Matching.filter((s) => !openTickets.has(s.site.toLowerCase()));
  const m2ByRegion: Record<string, number> = {};
  m2SitesNeedingTicket.forEach((s) => { m2ByRegion[s.region] = (m2ByRegion[s.region] ?? 0) + 1; });

  // ── 4. False tickets: AC working but open ticket exists ───────────────────
  // Site is communicating (noCommH=0), continuously powered (≥119.5h), AND consuming AC (KWh > 0)
  // → AC is healthy, ticket should be closed
  const falseTicketSites: { site: string; region: string; projectStatus: string; topology: string;
    totalRH: number; totalKWh: number; noCommH: number;
    ticketId: string; ticketStatus: string; ticketCategory: string; }[] = [];

  for (const s of m2BySite.values()) {
    if (s.noCommH === 0 && s.totalRH >= 119.5 && s.totalKWh > 0) {
      const tickets = openTicketsBySite.get(s.site.toLowerCase());
      if (tickets && tickets.length > 0) {
        // One row per open ticket for this site
        for (const t of tickets) {
          falseTicketSites.push({
            site:            s.site,
            region:          s.region,
            projectStatus:   s.projectStatus,
            topology:        s.topology,
            totalRH:         Math.round(s.totalRH * 100) / 100,
            totalKWh:        Math.round(s.totalKWh * 100) / 100,
            noCommH:         s.noCommH,
            ticketId:        t.ticketId,
            ticketStatus:    t.status,
            ticketCategory:  t.category,
          });
        }
      }
    }
  }
  falseTicketSites.sort((a, b) => a.region.localeCompare(b.region) || a.site.localeCompare(b.site));

  // ── 5. Build Excel workbook ────────────────────────────────────────────────
  const wb = new ExcelJS.Workbook();
  wb.creator = "Telemetry Analyzer";

  const HDR_BLUE:   ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E78" } };
  const HDR_PURPLE: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4C1D95" } };
  const HDR_RED:    ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC00000" } };
  const HDR_GREEN:  ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF375623" } };
  const WHITE_BOLD: Partial<ExcelJS.Font> = { bold: true, color: { argb: "FFFFFFFF" } };

  // ── Sheet 1: M1 — Alert codes ──────────────────────────────────────────────
  {
    const ws = wb.addWorksheet("M1 - Tickets à créer");
    ws.mergeCells("A1:H1");
    ws.getCell("A1").value = "M1 — Alertes AC (AL690-0055 · EM690-0001 · E_ACM-0001 · USER-0016) — Tickets à créer";
    ws.getCell("A1").font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
    ws.getCell("A1").fill = HDR_BLUE;
    ws.getCell("A1").alignment = { horizontal: "center" };
    ws.columns = [
      { header: "Region",         key: "region",        width: 16 },
      { header: "Site",           key: "site",          width: 20 },
      { header: "Address",        key: "address",       width: 35 },
      { header: "Priority",       key: "priority",      width: 10 },
      { header: "Project Status", key: "projectStatus", width: 14 },
      { header: "Power Topology", key: "topology",      width: 28 },
      { header: "Alert Codes",    key: "codes",         width: 40 },
      { header: "Alert Count",    key: "count",         width: 12 },
    ];
    const h = ws.getRow(2);
    ws.columns.forEach((col, i) => { h.getCell(i + 1).value = col.header as string; });
    h.eachCell((c) => { c.fill = HDR_BLUE; c.font = WHITE_BOLD; c.alignment = { horizontal: "center", vertical: "middle" }; });
    h.height = 20;
    for (const s of m1SitesNeedingTicket) {
      ws.addRow([s.region, s.site, s.address, s.priority, s.projectStatus, s.topology, [...s.codes].join(" | "), s.count]);
    }
    ws.autoFilter = { from: "A2", to: "H2" };
    ws.views = [{ state: "frozen", ySplit: 2 }];
  }

  // ── Sheet 2: M2 — ERS Big Data ─────────────────────────────────────────────
  {
    const ws = wb.addWorksheet("M2 - Tickets à créer");
    ws.mergeCells("A1:G1");
    ws.getCell("A1").value = `M2 — ERS Big Data (5 jours) : communicant · alimentation continue · consommation AC = 0 — Tickets à créer`;
    ws.getCell("A1").font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
    ws.getCell("A1").fill = HDR_PURPLE;
    ws.getCell("A1").alignment = { horizontal: "center" };
    const headers = ["Region", "Site", "Project Status", "Power Topology", "Total Power RH (5j)", "Total Power KWh (5j)", "No Comm H (5j)"];
    ws.getRow(2).values = headers;
    ws.getRow(2).eachCell((c) => { c.fill = HDR_PURPLE; c.font = WHITE_BOLD; c.alignment = { horizontal: "center", vertical: "middle" }; });
    ws.getRow(2).height = 20;
    ws.columns = [
      { width: 16 }, { width: 20 }, { width: 14 }, { width: 28 }, { width: 20 }, { width: 20 }, { width: 16 },
    ];
    for (const s of m2SitesNeedingTicket) {
      ws.addRow([s.region, s.site, s.projectStatus, s.topology,
        Math.round(s.totalRH * 100) / 100, s.totalKWh, s.noCommH]);
    }
    ws.autoFilter = { from: "A2", to: "G2" };
    ws.views = [{ state: "frozen", ySplit: 2 }];
  }

  // ── Sheet 3: False tickets — AC is working ─────────────────────────────────
  {
    const ws = wb.addWorksheet("Faux Tickets - À fermer");
    ws.mergeCells("A1:J1");
    ws.getCell("A1").value = "Faux Tickets AC — AC fonctionne (KWh > 0) mais ticket ouvert → À fermer";
    ws.getCell("A1").font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
    ws.getCell("A1").fill = HDR_GREEN;
    ws.getCell("A1").alignment = { horizontal: "center" };
    const headers = ["Region", "Site", "Project Status", "Power Topology",
      "Total Power RH (5j)", "Total Power KWh (5j)", "No Comm H",
      "Ticket ID", "Ticket Status", "Category"];
    ws.getRow(2).values = headers;
    ws.getRow(2).eachCell((c) => { c.fill = HDR_GREEN; c.font = WHITE_BOLD; c.alignment = { horizontal: "center", vertical: "middle" }; });
    ws.getRow(2).height = 20;
    ws.columns = [
      { width: 16 }, { width: 20 }, { width: 14 }, { width: 28 },
      { width: 18 }, { width: 18 }, { width: 12 },
      { width: 20 }, { width: 16 }, { width: 22 },
    ];
    for (const s of falseTicketSites) {
      ws.addRow([
        s.region, s.site, s.projectStatus, s.topology,
        s.totalRH, s.totalKWh, s.noCommH,
        s.ticketId, s.ticketStatus, s.ticketCategory,
      ]);
    }
    ws.autoFilter = { from: "A2", to: "J2" };
    ws.views = [{ state: "frozen", ySplit: 2 }];
    // Highlight the ticket columns
    if (falseTicketSites.length > 0) {
      ws.addConditionalFormatting({
        ref: `H3:H${falseTicketSites.length + 2}`,
        rules: [{ type: "expression", formulae: [`H3<>""`], priority: 1,
          style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFFE699" } } } }],
      });
    }
  }

  // ── Sheet 4: Summary ───────────────────────────────────────────────────────
  {
    const ws = wb.addWorksheet("Summary");
    ws.mergeCells("A1:B1");
    ws.getCell("A1").value = "AC Analysis — Summary";
    ws.getCell("A1").font = { bold: true, size: 14, color: { argb: "FFFFFFFF" } };
    ws.getCell("A1").fill = HDR_BLUE;
    ws.getCell("A1").alignment = { horizontal: "center" };
    ws.getColumn(1).width = 40; ws.getColumn(2).width = 16;

    const rows: [string, string | number][] = [
      ["Period (ERS)", last5Dates.length > 0 ? `${last5Dates[0]} → ${last5Dates[last5Dates.length - 1]}` : "—"],
      ["", ""],
      ["M1 — Total alerts (AC codes)", m1Filtered.length],
      ["M1 — Unique sites with alert", m1SiteList.length],
      ["M1 — Tickets to create (no open ticket)", m1SitesNeedingTicket.length],
      ["", ""],
      ["M2 — Sites: communicant + alimentation + KWh=0", m2Matching.length],
      ["M2 — Tickets to create (no open ticket)", m2SitesNeedingTicket.length],
      ["", ""],
      ["Faux tickets à fermer", falseTicketSites.length],
    ];
    rows.forEach((row, i) => {
      const r = i + 3;
      ws.getCell(`A${r}`).value = row[0];
      ws.getCell(`B${r}`).value = row[1];
      if (typeof row[1] === "number" && row[1] > 0) ws.getCell(`B${r}`).font = { bold: true };
    });

    ws.getCell("D3").value = "Region"; ws.getCell("D3").font = { bold: true };
    ws.getCell("E3").value = "M1 tickets"; ws.getCell("E3").font = { bold: true };
    ws.getCell("F3").value = "M2 tickets"; ws.getCell("F3").font = { bold: true };
    ws.getCell("G3").value = "Faux tickets"; ws.getCell("G3").font = { bold: true };
    const regions = [...new Set([
      ...m1SitesNeedingTicket.map((s) => s.region),
      ...m2SitesNeedingTicket.map((s) => s.region),
      ...falseTicketSites.map((s) => s.region),
    ])].filter(Boolean).sort();
    regions.forEach((reg, i) => {
      const r = i + 4;
      ws.getCell(`D${r}`).value = reg;
      ws.getCell(`E${r}`).value = m1SitesNeedingTicket.filter((s) => s.region === reg).length;
      ws.getCell(`F${r}`).value = m2SitesNeedingTicket.filter((s) => s.region === reg).length;
      ws.getCell(`G${r}`).value = falseTicketSites.filter((s) => s.region === reg).length;
    });
    [4, 5, 6, 7].forEach((c) => { ws.getColumn(c).width = 16; });
  }

  const buf  = await wb.xlsx.writeBuffer();
  const date = new Date().toISOString().slice(0, 10);

  return {
    buffer: Buffer.from(buf),
    filename: `AC_Analysis_${date}.xlsx`,
    summary: {
      m1TotalAlerts:        m1Filtered.length,
      m1UniqueSites:        m1SiteList.length,
      m1SitesNeedingTicket: m1SitesNeedingTicket.length,
      m1ByCode,
      m1ByRegion,
      m2TotalSites:         m2Matching.length,
      m2SitesNeedingTicket: m2SitesNeedingTicket.length,
      m2ByRegion,
      falseTickets:         falseTicketSites.length,
      dateRange:            last5Dates,
    },
  };
}
