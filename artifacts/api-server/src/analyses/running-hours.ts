import ExcelJS from "exceljs";
import { parseAllSheets, parseSheet, isSlaProject, type Row } from "../lib/excel-utils";

export interface RHOptions {
  periodDays: number;
  lowThresholdHours: number;
  highThresholdHours: number;
  batteryAbnormalThreshold: number;
  noCommDailyThreshold: number; // exclude site if avg No Comm H/day >= this (default 20)
}

export interface RHResult {
  buffer: Buffer;
  filename: string;
  summary: {
    totalRows: number;
    slaSites: number;
    excludedNoComm: number;
    excludedOpenPower: number;
    excludedOpenNoCom: number;
    excludedOpenGen: number;
    candidates: number;
    lowRH: number;
    highRH: number;
    genOnLoad: number;
  };
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

function num(v: string): number {
  if (!v) return 0;
  const n = parseFloat(String(v).replace(/,/g, ""));
  return Number.isNaN(n) ? 0 : n;
}

interface ERSAggregate {
  site: string;
  region: string;
  projectStatus: string;
  powerTopology: string;
  days: Set<string>;
  totalRH: number;
  totalGenRH: number;
  totalDownH: number;
  totalNoCommH: number;
  rhPerDay: Map<string, number>;
}

function hasGenerator(topology: string): boolean {
  return /\bgen\b/i.test(topology);
}

export async function analyzeRunningHours(
  ersBuffer: Buffer,
  noComTicketsBuffer: Buffer | null,
  powerTicketsBuffer: Buffer | null,
  genTicketsBuffer: Buffer | null,
  options: RHOptions,
): Promise<RHResult> {
  // ── 1. Parse ERS Big Data ────────────────────────────────────────────────
  const allSheets = parseAllSheets(ersBuffer);
  let ersRows: Row[] = [];
  for (const [name, rows] of Object.entries(allSheets)) {
    if (rows.length === 0) continue;
    const first = rows[0]!;
    const keys = Object.keys(first).map((k) => k.toLowerCase());
    if (keys.some((k) => k.includes("totalpower")) || keys.some((k) => k.includes("gen rh") || k.includes("genrh"))) {
      ersRows = rows;
      break;
    }
    if (name.toLowerCase().includes("ers big data")) {
      ersRows = rows;
      break;
    }
  }
  if (ersRows.length === 0) {
    const names = Object.keys(allSheets);
    if (names.length > 0) ersRows = allSheets[names[names.length - 1]!]!;
  }

  // ── 2. Aggregate per site ────────────────────────────────────────────────
  const agg = new Map<string, ERSAggregate>();
  for (const r of ersRows) {
    const site = pickField(r, ["SiteName", "Site", "Site Name"]);
    if (!site) continue;
    const projectStatus = pickField(r, ["ProjectStatus", "Project Status", "projectStatus"]);
    if (!isSlaProject(projectStatus)) continue;
    const day = pickField(r, ["Day", "Date"]);
    let a = agg.get(site);
    if (!a) {
      a = {
        site,
        region: pickField(r, ["RegionName", "Region"]),
        projectStatus,
        powerTopology: pickField(r, ["PowerTopology", "Power Topology"]),
        days: new Set(),
        totalRH: 0,
        totalGenRH: 0,
        totalDownH: 0,
        totalNoCommH: 0,
        rhPerDay: new Map(),
      };
      agg.set(site, a);
    }
    if (day) a.days.add(day);
    const genRH   = num(pickField(r, ["Gen RH", "GenRH", "Generator Working H"]));
    const totalRH = num(pickField(r, ["TotalPower RH", "Total Power RH", "TotalPowerRH"]));
    const downH   = num(pickField(r, ["SiteDown H", "SiteDown", "SiteDownH"]));
    const noComm  = num(pickField(r, ["No Comm. H", "No Comm H", "NoComm H", "NoCommH"]));
    a.totalRH     += totalRH || genRH;
    a.totalGenRH  += genRH;
    a.totalDownH  += downH;
    a.totalNoCommH += noComm;
    a.rhPerDay.set(day, (a.rhPerDay.get(day) ?? 0) + (totalRH || genRH));
  }

  // ── 3. NO COM tickets → exclude sites with open NO COM ticket ───────────
  const openNoComSet = new Set<string>();
  if (noComTicketsBuffer) {
    const tRows = parseSheet(noComTicketsBuffer);
    for (const r of tRows) {
      const site = pickField(r, ["Site"]);
      const status = pickField(r, ["Status"]);
      if (!site) continue;
      if (status.toLowerCase() === "open") openNoComSet.add(site.trim().toUpperCase());
    }
  }

  // ── 4. Open Power tickets ────────────────────────────────────────────────
  const openPowerSet = new Set<string>();
  if (powerTicketsBuffer) {
    const tRows = parseSheet(powerTicketsBuffer);
    for (const r of tRows) {
      const site = pickField(r, ["Site"]);
      const status = pickField(r, ["Status"]);
      if (!site) continue;
      if (status.toLowerCase() === "open") openPowerSet.add(site.trim().toUpperCase());
    }
  }

  // ── 5. Open Generator / Gen On Load tickets ──────────────────────────────
  const openGenSet = new Set<string>();
  if (genTicketsBuffer) {
    const tRows = parseSheet(genTicketsBuffer);
    for (const r of tRows) {
      const site = pickField(r, ["Site"]);
      const status = pickField(r, ["Status"]);
      if (!site) continue;
      if (status.toLowerCase() === "open") openGenSet.add(site.trim().toUpperCase());
    }
  }

  // ── 6. Classify ──────────────────────────────────────────────────────────
  type Classified = {
    site: string;
    region: string;
    projectStatus: string;
    powerTopology: string;
    daysCount: number;
    totalRH: number;
    avgRH: number;
    totalGenRH: number;
    siteDownH: number;
    noCommH: number;
    avgDailyNoCommH: number;
    excludedReason: string;
    flag: string;
  };

  const list: Classified[] = [];
  let excludedNoComm = 0, excludedOpenPower = 0, excludedOpenNoCom = 0, excludedOpenGen = 0;

  const low  = options.lowThresholdHours;
  const high = options.highThresholdHours;
  const noCommThr = options.noCommDailyThreshold;

  for (const a of agg.values()) {
    const daysCount = a.days.size || options.periodDays;
    const avgDailyNoCommH = a.totalNoCommH / Math.max(1, daysCount);
    const siteKey = a.site.trim().toUpperCase();

    let excluded = "";
    // Priority: no-comm first (ERS-based), then ticket exclusions
    if (avgDailyNoCommH >= noCommThr) {
      excluded = "Not communicating (No Comm H ≥ 20/day)";
      excludedNoComm++;
    } else if (openPowerSet.has(siteKey)) {
      excluded = "Open Power ticket";
      excludedOpenPower++;
    } else if (openNoComSet.has(siteKey)) {
      excluded = "Open NO COM ticket";
      excludedOpenNoCom++;
    } else if (openGenSet.has(siteKey)) {
      excluded = "Open Generator ticket";
      excludedOpenGen++;
    }

    const totalRH = Math.round(a.totalRH * 100) / 100;
    const avgRH   = Math.round((totalRH / Math.max(1, daysCount)) * 100) / 100;

    let flag = "OK";
    if (!excluded) {
      if (totalRH < low)  flag = "LOW";
      else if (totalRH > high) flag = "HIGH";
    }

    list.push({
      site: a.site,
      region: a.region,
      projectStatus: a.projectStatus,
      powerTopology: a.powerTopology,
      daysCount,
      totalRH,
      avgRH,
      totalGenRH: Math.round(a.totalGenRH * 100) / 100,
      siteDownH: Math.round(a.totalDownH * 100) / 100,
      noCommH: Math.round(a.totalNoCommH * 100) / 100,
      avgDailyNoCommH: Math.round(avgDailyNoCommH * 100) / 100,
      excludedReason: excluded,
      flag,
    });
  }

  const candidates = list.filter((c) => !c.excludedReason);
  const lowSites   = candidates.filter((c) => c.flag === "LOW").sort((a, b) => a.totalRH - b.totalRH);
  const highSites  = candidates.filter((c) => c.flag === "HIGH").sort((a, b) => b.totalRH - a.totalRH);

  // Gen On Load: site has Gen in topology, Gen RH = 0 over the period, communicating, no open gen ticket
  const genOnLoadSites = list.filter((c) => {
    if (c.avgDailyNoCommH >= noCommThr) return false; // not communicating
    if (openGenSet.has(c.site.trim().toUpperCase())) return false; // already has gen ticket
    if (!hasGenerator(c.powerTopology)) return false;
    return c.totalGenRH === 0;
  }).sort((a, b) => a.region.localeCompare(b.region) || a.site.localeCompare(b.site));

  // ── 7. Build workbook ────────────────────────────────────────────────────
  const wb = new ExcelJS.Workbook();
  wb.creator = "Telemetry Analyzer";
  wb.created = new Date();

  const HDR_BLUE:   ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E78" } };
  const HDR_RED:    ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC00000" } };
  const HDR_ORANGE: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFED7D31" } };
  const HDR_GREEN:  ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF375623" } };
  const WHITE_BOLD: Partial<ExcelJS.Font> = { bold: true, color: { argb: "FFFFFFFF" } };

  // ---- Sheet 1: Low RH — Sites to Investigate ----
  {
    const ws = wb.addWorksheet("Low RH - Investigate");
    ws.mergeCells("A1:G1");
    ws.getCell("A1").value = `LOW Running Hours — Investigate (Total RH < ${low}h over ${options.periodDays} days)`;
    ws.getCell("A1").font = { bold: true, size: 13, color: { argb: "FFFFFFFF" } };
    ws.getCell("A1").fill = HDR_RED;
    ws.getCell("A1").alignment = { horizontal: "center" };
    ws.getRow(2).values = ["Region", "Site", "Project Status", "Power Topology", "Total RH (h)", "Avg RH/day", "Site Down H"];
    ws.getRow(2).font = WHITE_BOLD;
    ws.getRow(2).fill = HDR_BLUE;
    lowSites.forEach((c, i) => {
      const r = i + 3;
      ws.getCell(`A${r}`).value = c.region;
      ws.getCell(`B${r}`).value = c.site;
      ws.getCell(`C${r}`).value = c.projectStatus;
      ws.getCell(`D${r}`).value = c.powerTopology;
      ws.getCell(`E${r}`).value = c.totalRH;     ws.getCell(`E${r}`).numFmt = "0.00";
      ws.getCell(`F${r}`).value = c.avgRH;       ws.getCell(`F${r}`).numFmt = "0.00";
      ws.getCell(`G${r}`).value = c.siteDownH;   ws.getCell(`G${r}`).numFmt = "0.00";
    });
    ws.columns = [{ width: 16 }, { width: 20 }, { width: 14 }, { width: 30 }, { width: 14 }, { width: 12 }, { width: 12 }];
    ws.autoFilter = { from: "A2", to: "G2" };
    ws.views = [{ state: "frozen", ySplit: 2 }];
    if (lowSites.length > 0) {
      ws.addConditionalFormatting({
        ref: `E3:E${lowSites.length + 2}`,
        rules: [{ type: "dataBar", priority: 1, cfvo: [{ type: "min" }, { type: "max" }], color: { argb: "FFC00000" } } as never],
      });
    }
  }

  // ---- Sheet 2: High RH — Sites to Investigate ----
  {
    const ws = wb.addWorksheet("High RH - Investigate");
    ws.mergeCells("A1:G1");
    ws.getCell("A1").value = `HIGH Running Hours — Investigate (Total RH > ${high}h over ${options.periodDays} days)`;
    ws.getCell("A1").font = { bold: true, size: 13, color: { argb: "FFFFFFFF" } };
    ws.getCell("A1").fill = HDR_ORANGE;
    ws.getCell("A1").alignment = { horizontal: "center" };
    ws.getRow(2).values = ["Region", "Site", "Project Status", "Power Topology", "Total RH (h)", "Avg RH/day", "Site Down H"];
    ws.getRow(2).font = WHITE_BOLD;
    ws.getRow(2).fill = HDR_BLUE;
    highSites.forEach((c, i) => {
      const r = i + 3;
      ws.getCell(`A${r}`).value = c.region;
      ws.getCell(`B${r}`).value = c.site;
      ws.getCell(`C${r}`).value = c.projectStatus;
      ws.getCell(`D${r}`).value = c.powerTopology;
      ws.getCell(`E${r}`).value = c.totalRH;   ws.getCell(`E${r}`).numFmt = "0.00";
      ws.getCell(`F${r}`).value = c.avgRH;     ws.getCell(`F${r}`).numFmt = "0.00";
      ws.getCell(`G${r}`).value = c.siteDownH; ws.getCell(`G${r}`).numFmt = "0.00";
    });
    ws.columns = [{ width: 16 }, { width: 20 }, { width: 14 }, { width: 30 }, { width: 14 }, { width: 12 }, { width: 12 }];
    ws.autoFilter = { from: "A2", to: "G2" };
    ws.views = [{ state: "frozen", ySplit: 2 }];
    if (highSites.length > 0) {
      ws.addConditionalFormatting({
        ref: `E3:E${highSites.length + 2}`,
        rules: [{ type: "dataBar", priority: 1, cfvo: [{ type: "min" }, { type: "max" }], color: { argb: "FFED7D31" } } as never],
      });
    }
  }

  // ---- Sheet 3: Gen on Load — Tickets to Create ----
  {
    const ws = wb.addWorksheet("Gen on Load - Tickets");
    ws.mergeCells("A1:F1");
    ws.getCell("A1").value = `Gen On Load — Tickets to Create (Gen in topology · Gen RH = 0 over ${options.periodDays} days)`;
    ws.getCell("A1").font = { bold: true, size: 13, color: { argb: "FFFFFFFF" } };
    ws.getCell("A1").fill = HDR_GREEN;
    ws.getCell("A1").alignment = { horizontal: "center" };
    ws.getRow(2).values = ["Region", "Site", "Project Status", "Power Topology", "Total RH (h)", "No Comm H (total)"];
    ws.getRow(2).font = WHITE_BOLD;
    ws.getRow(2).fill = HDR_BLUE;
    genOnLoadSites.forEach((c, i) => {
      const r = i + 3;
      ws.getCell(`A${r}`).value = c.region;
      ws.getCell(`B${r}`).value = c.site;
      ws.getCell(`C${r}`).value = c.projectStatus;
      ws.getCell(`D${r}`).value = c.powerTopology;
      ws.getCell(`E${r}`).value = c.totalRH;   ws.getCell(`E${r}`).numFmt = "0.00";
      ws.getCell(`F${r}`).value = c.noCommH;   ws.getCell(`F${r}`).numFmt = "0.00";
    });
    ws.columns = [{ width: 16 }, { width: 20 }, { width: 14 }, { width: 34 }, { width: 14 }, { width: 18 }];
    ws.autoFilter = { from: "A2", to: "F2" };
    ws.views = [{ state: "frozen", ySplit: 2 }];
  }

  // ---- Sheet 4: All Sites (full data) ----
  {
    const ws = wb.addWorksheet("All Sites");
    ws.columns = [
      { header: "Region",          key: "region",          width: 16 },
      { header: "Site",            key: "site",            width: 20 },
      { header: "Project Status",  key: "projectStatus",   width: 14 },
      { header: "Power Topology",  key: "powerTopology",   width: 30 },
      { header: "Days",            key: "daysCount",       width: 8  },
      { header: "Total RH (h)",    key: "totalRH",         width: 13 },
      { header: "Avg RH/day",      key: "avgRH",           width: 11 },
      { header: "Gen RH (h)",      key: "totalGenRH",      width: 12 },
      { header: "Site Down H",     key: "siteDownH",       width: 12 },
      { header: "No Comm H",       key: "noCommH",         width: 12 },
      { header: "Avg No Comm/day", key: "avgDailyNoCommH", width: 16 },
      { header: "Excluded Reason", key: "excludedReason",  width: 28 },
      { header: "Flag",            key: "flag",            width: 10 },
    ];
    ws.getRow(1).font = WHITE_BOLD;
    ws.getRow(1).fill = HDR_BLUE;
    list.forEach((c) => {
      const row = ws.addRow(c);
      ["F","G","H","I","J","K"].forEach((col) => { row.getCell(col).numFmt = "0.00"; });
    });
    ws.autoFilter = { from: "A1", to: "M1" };
    ws.views = [{ state: "frozen", ySplit: 1 }];
    ws.addConditionalFormatting({
      ref: `M2:M${list.length + 1}`,
      rules: [
        { type: "containsText", operator: "containsText", text: "LOW",  priority: 1, style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFF8CBAD" } } } },
        { type: "containsText", operator: "containsText", text: "HIGH", priority: 2, style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFFC000" } } } },
      ],
    });
  }

  // ---- Sheet 5: Summary ----
  {
    const ws = wb.addWorksheet("Summary");
    ws.mergeCells("A1:B1");
    ws.getCell("A1").value = "Running Hours Analysis — Summary";
    ws.getCell("A1").font = { bold: true, size: 14, color: { argb: "FFFFFFFF" } };
    ws.getCell("A1").fill = HDR_BLUE;
    ws.getCell("A1").alignment = { horizontal: "center" };
    const meta: [string, number | string][] = [
      ["Period (days)",                     options.periodDays],
      ["Low threshold (h)",                 options.lowThresholdHours],
      ["High threshold (h)",                options.highThresholdHours],
      ["No Comm exclusion threshold (h/day)", options.noCommDailyThreshold],
      ["", ""],
      ["Total SLA sites in ERS",            agg.size],
      ["Excluded — Not communicating",       excludedNoComm],
      ["Excluded — Open Power ticket",       excludedOpenPower],
      ["Excluded — Open NO COM ticket",      excludedOpenNoCom],
      ["Excluded — Open Generator ticket",   excludedOpenGen],
      ["Candidates after exclusions",        candidates.length],
      ["LOW RH (< " + low + "h)",            lowSites.length],
      ["HIGH RH (> " + high + "h)",          highSites.length],
      ["Gen On Load — no gen run in period", genOnLoadSites.length],
    ];
    meta.forEach((m, i) => {
      const r = i + 3;
      ws.getCell(`A${r}`).value = m[0];
      ws.getCell(`B${r}`).value = m[1];
      if (i < 4) ws.getCell(`A${r}`).font = { bold: true };
    });
    ws.getColumn(1).width = 42;
    ws.getColumn(2).width = 18;

    // Region breakdown
    ws.getCell("D3").value = "Region"; ws.getCell("D3").font = { bold: true };
    ws.getCell("E3").value = "LOW";    ws.getCell("E3").font = { bold: true };
    ws.getCell("F3").value = "HIGH";   ws.getCell("F3").font = { bold: true };
    ws.getCell("G3").value = "Gen OL"; ws.getCell("G3").font = { bold: true };
    const regions = Array.from(new Set(list.map((l) => l.region))).filter(Boolean).sort();
    regions.forEach((reg, i) => {
      const r = i + 4;
      ws.getCell(`D${r}`).value = reg;
      ws.getCell(`E${r}`).value = lowSites.filter((c) => c.region === reg).length;
      ws.getCell(`F${r}`).value = highSites.filter((c) => c.region === reg).length;
      ws.getCell(`G${r}`).value = genOnLoadSites.filter((c) => c.region === reg).length;
    });
    [4, 5, 6, 7].forEach((c) => { ws.getColumn(c).width = 16; });
  }

  const buffer = Buffer.from(await wb.xlsx.writeBuffer());
  const date = new Date().toISOString().slice(0, 10);
  const filename = `RunningHours_Analysis_${options.periodDays}d_${date}.xlsx`;

  return {
    buffer,
    filename,
    summary: {
      totalRows: ersRows.length,
      slaSites: agg.size,
      excludedNoComm,
      excludedOpenPower,
      excludedOpenNoCom,
      excludedOpenGen,
      candidates: candidates.length,
      lowRH: lowSites.length,
      highRH: highSites.length,
      genOnLoad: genOnLoadSites.length,
    },
  };
}
