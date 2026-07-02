import * as XLSX from "xlsx";

export type Row = Record<string, string>;

export function parseSheet(buffer: Buffer, sheetName?: string): Row[] {
  const wb = XLSX.read(buffer, { type: "buffer", cellDates: false });
  const sn = sheetName && wb.SheetNames.includes(sheetName) ? sheetName : wb.SheetNames[0];
  if (!sn) return [];
  const ws = wb.Sheets[sn];
  if (!ws) return [];
  return XLSX.utils.sheet_to_json<Row>(ws, { defval: "", raw: false });
}

export function parseAllSheets(buffer: Buffer): Record<string, Row[]> {
  const wb = XLSX.read(buffer, { type: "buffer", cellDates: false });
  const out: Record<string, Row[]> = {};
  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    if (!ws) continue;
    out[sn] = XLSX.utils.sheet_to_json<Row>(ws, { defval: "", raw: false });
  }
  return out;
}

export function isSlaProject(projectStatus: string | undefined): boolean {
  if (!projectStatus) return false;
  const v = String(projectStatus).trim().toUpperCase();
  if (!v) return false;
  if (v.startsWith("NOT")) return false;
  if (v.includes("NOT_SLA") || v.includes("NOT SLA")) return false;
  return v.startsWith("SLA");
}

export function parseDate(value: string | undefined): Date | null {
  if (!value) return null;
  const s = String(value).trim();
  if (!s) return null;
  // try ISO/yyyy-mm-dd hh:mm:ss
  let d = new Date(s);
  if (!Number.isNaN(d.getTime())) return d;
  // try m/d/yy or m/d/yyyy [h:mm]
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (m) {
    let yy = parseInt(m[3]!, 10);
    if (yy < 100) yy += 2000;
    const mo = parseInt(m[1]!, 10) - 1;
    const da = parseInt(m[2]!, 10);
    const hh = m[4] ? parseInt(m[4], 10) : 0;
    const mi = m[5] ? parseInt(m[5], 10) : 0;
    d = new Date(yy, mo, da, hh, mi);
    if (!Number.isNaN(d.getTime())) return d;
  }
  return null;
}

export function diffDays(later: Date, earlier: Date): number {
  return Math.floor((later.getTime() - earlier.getTime()) / 86_400_000);
}

export function extractCity(siteId: string): string {
  const m = siteId.match(/^IHS_([A-Z]+)_/);
  return m ? m[1]! : "";
}
