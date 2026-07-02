const API_BASE = "/api";

export interface AnalysisSummary {
  [key: string]: unknown;
}

export interface AnalysisResponse {
  blob: Blob;
  filename: string;
  summary: AnalysisSummary | null;
}

export interface NoComAnalysisResponse extends AnalysisResponse {
  ticketListBlob?: Blob;
  ticketListFilename?: string;
}

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

function base64ToBlob(b64: string, mimeType = XLSX_MIME): Blob {
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return new Blob([bytes], { type: mimeType });
}

// All analysis endpoints now return JSON: { filename, file (base64), mimeType, summary, ...extras }
async function postAnalysis(path: string, fd: FormData): Promise<AnalysisResponse> {
  const res = await fetch(`${API_BASE}${path}`, { method: "POST", body: fd });
  if (!res.ok) {
    let msg = `Request failed (${res.status})`;
    try {
      const j = await res.json();
      if (j?.error) msg = j.error;
    } catch {}
    throw new Error(msg);
  }
  const j = await res.json();
  const blob = base64ToBlob(j.file, j.mimeType ?? XLSX_MIME);
  let summary: AnalysisSummary | null = null;
  if (j.summary) { try { summary = j.summary; } catch {} }
  return { blob, filename: j.filename ?? "analysis.xlsx", summary };
}

export async function runNoComAnalysis(args: {
  slaTracker?: File | null;
  commReport: File;
  noComTickets: File;
  cutoffDays: number;
  referenceDate: string;
}): Promise<NoComAnalysisResponse> {
  const fd = new FormData();
  if (args.slaTracker) fd.append("slaTracker", args.slaTracker);
  fd.append("commReport", args.commReport);
  fd.append("noComTickets", args.noComTickets);
  fd.append("cutoffDays", String(args.cutoffDays));
  fd.append("referenceDate", args.referenceDate);

  const res = await fetch(`${API_BASE}/analyses/no-com`, { method: "POST", body: fd });
  if (!res.ok) {
    let msg = `Request failed (${res.status})`;
    try {
      const j = await res.json();
      if (j?.error) msg = j.error;
    } catch {}
    throw new Error(msg);
  }
  const j = await res.json();
  const blob = base64ToBlob(j.file, j.mimeType ?? XLSX_MIME);
  let summary: AnalysisSummary | null = null;
  if (j.summary) { try { summary = j.summary; } catch {} }
  let ticketListBlob: Blob | undefined;
  if (j.ticketListFile) {
    try { ticketListBlob = base64ToBlob(j.ticketListFile, XLSX_MIME); } catch {}
  }
  return {
    blob,
    filename: j.filename ?? "NoCom_Analysis.xlsx",
    summary,
    ticketListBlob,
    ticketListFilename: j.ticketListFilename,
  };
}

export interface TemplateInfo {
  exists: boolean;
  filename?: string;
  sizeBytes?: number;
  uploadedAt?: string;
}

export async function getTemplateInfo(): Promise<TemplateInfo> {
  const res = await fetch(`${API_BASE}/template/sla-tracker`);
  return res.json();
}

export async function uploadTemplate(file: File): Promise<TemplateInfo> {
  const fd = new FormData();
  fd.append("file", file);
  const res = await fetch(`${API_BASE}/template/sla-tracker`, { method: "POST", body: fd });
  if (!res.ok) throw new Error(`Upload failed (${res.status})`);
  return res.json();
}

export async function deleteTemplate(): Promise<void> {
  await fetch(`${API_BASE}/template/sla-tracker`, { method: "DELETE" });
}

export async function runRunningHoursAnalysis(args: {
  ersBigData: File;
  noComTickets?: File | null;
  powerTickets?: File | null;
  genTickets?: File | null;
  periodDays: number;
  lowThresholdHours: number;
  highThresholdHours: number;
  noCommDailyThreshold?: number;
}): Promise<AnalysisResponse> {
  const fd = new FormData();
  fd.append("ersBigData", args.ersBigData);
  if (args.noComTickets) fd.append("noComTickets", args.noComTickets);
  if (args.powerTickets) fd.append("powerTickets", args.powerTickets);
  if (args.genTickets) fd.append("genTickets", args.genTickets);
  fd.append("periodDays", String(args.periodDays));
  fd.append("lowThresholdHours", String(args.lowThresholdHours));
  fd.append("highThresholdHours", String(args.highThresholdHours));
  fd.append("noCommDailyThreshold", String(args.noCommDailyThreshold ?? 20));
  return postAnalysis("/analyses/running-hours", fd);
}

export async function runFuelSensorAnalysis(args: {
  componentsFile: File;
  ticketsFile: File;
}): Promise<AnalysisResponse> {
  const fd = new FormData();
  fd.append("componentsFile", args.componentsFile);
  fd.append("ticketsFile", args.ticketsFile);
  return postAnalysis("/analyses/fuel-sensor", fd);
}

export async function runACAnalysis(args: {
  alertsFile: File;
  ersBigData: File;
  ticketsFile?: File | null;
}): Promise<AnalysisResponse> {
  const fd = new FormData();
  fd.append("alertsFile", args.alertsFile);
  fd.append("ersBigData", args.ersBigData);
  if (args.ticketsFile) fd.append("ticketsFile", args.ticketsFile);
  return postAnalysis("/analyses/ac", fd);
}

export interface MonthlyConsolidationResponse {
  blob: Blob;
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
  } | null;
}

export async function runMonthlyConsolidation(args: {
  slaTracker?: File | null;
  month?: number;
  year?: number;
  referenceDate?: string;
}): Promise<MonthlyConsolidationResponse> {
  const fd = new FormData();
  if (args.slaTracker) fd.append("slaTracker", args.slaTracker);
  if (args.month) fd.append("month", String(args.month));
  if (args.year) fd.append("year", String(args.year));
  if (args.referenceDate) fd.append("referenceDate", args.referenceDate);

  const res = await fetch(`${API_BASE}/analyses/no-com/monthly`, { method: "POST", body: fd });
  if (!res.ok) {
    let msg = `Request failed (${res.status})`;
    try { const j = await res.json(); if (j?.error) msg = j.error; } catch {}
    throw new Error(msg);
  }
  const j = await res.json();
  const blob = base64ToBlob(j.file, j.mimeType ?? XLSX_MIME);
  return { blob, filename: j.filename ?? "NoCom_Monthly.xlsx", summary: j.summary ?? null };
}

export function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 200);
}
