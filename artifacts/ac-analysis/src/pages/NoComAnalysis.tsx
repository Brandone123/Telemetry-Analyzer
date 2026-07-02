import { useEffect, useState } from "react";
import { Link } from "wouter";
import {
  ArrowLeft, Download, Loader2, Radio, AlertTriangle, CheckCircle2,
  Upload, Trash2, FileSpreadsheet, ListChecks, CalendarDays,
  TrendingDown, TrendingUp, Minus,
} from "lucide-react";
import { FileDrop } from "@/components/FileDrop";
import {
  runNoComAnalysis,
  runMonthlyConsolidation,
  downloadBlob,
  getTemplateInfo,
  uploadTemplate,
  deleteTemplate,
  type NoComAnalysisResponse,
  type MonthlyConsolidationResponse,
  type TemplateInfo,
} from "@/lib/api";

function todayStr() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

const MONTH_NAMES = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];

interface Summary {
  totalSites: number;
  onAirSites: number;
  noComSites: number;
  ticketsToCreate: number;
  ticketsToResolve: number;
  moreThan24h: number;
  within48h: number;
  todayColumn?: string;
  monthReset?: boolean;
  byCategory: Record<string, { com: number; noCom: number }>;
}

export default function NoComAnalysis() {
  const [template, setTemplate] = useState<TemplateInfo>({ exists: false });
  const [tracker, setTracker] = useState<File | null>(null);
  const [commReport, setCommReport] = useState<File | null>(null);
  const [tickets, setTickets] = useState<File | null>(null);
  const [cutoffDays, setCutoffDays] = useState(2);
  const [refDate, setRefDate] = useState(todayStr());
  const [loading, setLoading] = useState(false);
  const [tplLoading, setTplLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [result, setResult] = useState<NoComAnalysisResponse | null>(null);

  // Monthly consolidation state
  const now = new Date();
  const prevMonth = now.getMonth() === 0 ? 12 : now.getMonth();
  const prevYear = now.getMonth() === 0 ? now.getFullYear() - 1 : now.getFullYear();
  const [monthlyMonth, setMonthlyMonth] = useState(prevMonth);
  const [monthlyYear, setMonthlyYear] = useState(prevYear);
  const [monthlyLoading, setMonthlyLoading] = useState(false);
  const [monthlyError, setMonthlyError] = useState<string | null>(null);
  const [monthlyResult, setMonthlyResult] = useState<MonthlyConsolidationResponse | null>(null);

  const summary = result?.summary as Summary | null;
  const trackerProvided = !!tracker || template.exists;
  const canRun = trackerProvided && !!commReport && !!tickets && !loading;
  const canRunMonthly = (template.exists || !!tracker) && !monthlyLoading;

  useEffect(() => { void refreshTemplate(); }, []);

  async function refreshTemplate() {
    try { setTemplate(await getTemplateInfo()); } catch {}
  }

  async function uploadAsTemplate() {
    if (!tracker) return;
    setTplLoading(true); setError(null);
    try {
      const info = await uploadTemplate(tracker);
      setTemplate(info);
      setTracker(null);
    } catch (e) { setError((e as Error).message); }
    finally { setTplLoading(false); }
  }

  async function clearTemplate() {
    setTplLoading(true);
    try { await deleteTemplate(); await refreshTemplate(); }
    finally { setTplLoading(false); }
  }

  async function run() {
    if (!commReport || !tickets) return;
    if (!trackerProvided) return;
    setLoading(true); setError(null); setResult(null);
    try {
      const res = await runNoComAnalysis({
        slaTracker: tracker, commReport, noComTickets: tickets,
        cutoffDays, referenceDate: refDate,
      });
      setResult(res);
      void refreshTemplate();
      setTracker(null);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setLoading(false);
    }
  }

  async function runMonthly() {
    setMonthlyLoading(true); setMonthlyError(null); setMonthlyResult(null);
    try {
      const res = await runMonthlyConsolidation({
        month: monthlyMonth,
        year: monthlyYear,
        referenceDate: todayStr(),
      });
      setMonthlyResult(res);
    } catch (e) {
      setMonthlyError((e as Error).message);
    } finally {
      setMonthlyLoading(false);
    }
  }

  const monthlySummary = monthlyResult?.summary ?? null;
  const TrendIcon = monthlySummary?.trend === "improving" ? TrendingDown
    : monthlySummary?.trend === "worsening" ? TrendingUp : Minus;
  const trendColor = monthlySummary?.trend === "improving"
    ? "text-emerald-600 dark:text-emerald-400"
    : monthlySummary?.trend === "worsening"
    ? "text-red-500"
    : "text-amber-500";

  return (
    <div className="min-h-screen bg-gradient-to-b from-background to-muted/30">
      <header className="border-b border-border/50 bg-card/50 backdrop-blur">
        <div className="max-w-5xl mx-auto px-6 py-4 flex items-center justify-between">
          <Link href="/" className="text-sm text-muted-foreground hover:text-foreground inline-flex items-center gap-1">
            <ArrowLeft className="h-4 w-4" /> Back
          </Link>
          <div className="flex items-center gap-2">
            <div className="h-8 w-8 rounded-lg bg-gradient-to-br from-blue-500 to-cyan-500 grid place-items-center text-white">
              <Radio className="h-4 w-4" />
            </div>
            <span className="font-semibold">NO COM Analysis</span>
          </div>
        </div>
      </header>

      <main className="max-w-5xl mx-auto px-6 py-8 grid lg:grid-cols-[1fr_320px] gap-6">
        <section className="space-y-6">
          {/* Template card */}
          <div className="rounded-xl border bg-card p-6 space-y-4">
            <div className="flex items-start justify-between gap-3">
              <div>
                <h2 className="font-semibold">SLA Tracker template</h2>
                <p className="text-xs text-muted-foreground mt-1 leading-relaxed">
                  Your SLA Tracker file is the template. The server keeps the latest copy and
                  returns it updated each day — pivots, charts and formulas all preserved.
                </p>
              </div>
              {template.exists && (
                <div className="inline-flex items-center gap-1.5 rounded-full border border-emerald-500/30 bg-emerald-500/10 px-2.5 py-1 text-xs text-emerald-600 dark:text-emerald-400">
                  <CheckCircle2 className="h-3.5 w-3.5" /> Registered
                </div>
              )}
            </div>

            {template.exists ? (
              <div className="rounded-md bg-muted/40 p-3 text-sm flex items-start justify-between gap-3">
                <div className="flex items-start gap-2 min-w-0">
                  <FileSpreadsheet className="h-4 w-4 mt-0.5 text-emerald-600 dark:text-emerald-400 shrink-0" />
                  <div className="min-w-0">
                    <div className="truncate font-medium">{template.filename}</div>
                    <div className="text-xs text-muted-foreground">
                      Stored {template.uploadedAt ? new Date(template.uploadedAt).toLocaleString() : ""}
                      {template.sizeBytes ? ` · ${(template.sizeBytes / 1024 / 1024).toFixed(1)} MB` : ""}
                    </div>
                  </div>
                </div>
                <button
                  onClick={clearTemplate}
                  disabled={tplLoading}
                  className="inline-flex items-center gap-1 rounded-md border px-2 py-1 text-xs text-muted-foreground hover:text-destructive hover:border-destructive/40"
                >
                  <Trash2 className="h-3.5 w-3.5" /> Replace
                </button>
              </div>
            ) : (
              <>
                <FileDrop
                  label="SLA Tracker (one-time upload)"
                  hint="The full SLA_TRACKER.xlsx with all sheets, charts, pivots and formulas"
                  required
                  file={tracker}
                  onChange={setTracker}
                />
                <button
                  onClick={uploadAsTemplate}
                  disabled={!tracker || tplLoading}
                  className="inline-flex items-center gap-2 rounded-md bg-secondary px-4 py-2 text-sm font-medium hover:opacity-90 disabled:opacity-50"
                >
                  {tplLoading ? <Loader2 className="h-4 w-4 animate-spin" /> : <Upload className="h-4 w-4" />}
                  Register as template
                </button>
              </>
            )}
          </div>

          {/* Daily inputs */}
          <div className="rounded-xl border bg-card p-6 space-y-5">
            <h2 className="font-semibold">1 · Daily inputs</h2>
            {!template.exists && (
              <div className="rounded-md border border-amber-500/30 bg-amber-500/10 p-3 text-xs text-amber-700 dark:text-amber-400 flex items-start gap-2">
                <AlertTriangle className="h-4 w-4 shrink-0 mt-0.5" />
                Register your SLA Tracker as a template first (above) — or attach a fresh one below for this run.
              </div>
            )}
            {template.exists && (
              <FileDrop
                label="SLA Tracker (optional override)"
                hint="Leave empty to use the stored template. Attach a file only to replace it."
                file={tracker}
                onChange={setTracker}
              />
            )}
            <FileDrop
              label="Communication Report (today)"
              hint="Excel/HTML export — Site, Status, Last Communication"
              required
              file={commReport}
              onChange={setCommReport}
            />
            <FileDrop
              label="Existing No Com tickets"
              hint="CSV/Excel export of the open No Communication ticket list"
              required
              file={tickets}
              onChange={setTickets}
            />
          </div>

          <div className="rounded-xl border bg-card p-6 space-y-4">
            <h2 className="font-semibold">2 · Parameters</h2>
            <div className="grid sm:grid-cols-2 gap-4">
              <label className="block">
                <span className="text-sm font-medium">Reference date</span>
                <input
                  type="date"
                  value={refDate}
                  onChange={(e) => setRefDate(e.target.value)}
                  className="mt-1 w-full rounded-md border bg-background px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ring"
                />
              </label>
              <label className="block">
                <span className="text-sm font-medium">Ticket cutoff (days)</span>
                <input
                  type="number"
                  min={1}
                  value={cutoffDays}
                  onChange={(e) => setCutoffDays(Math.max(1, Number(e.target.value)))}
                  className="mt-1 w-full rounded-md border bg-background px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ring"
                />
              </label>
            </div>
            <p className="text-xs text-muted-foreground leading-relaxed">
              Today's column in <b>Comm Issue</b> is auto-located by date. Each site is
              looked up against the Communication Report (VLOOKUP-style) and written as 1 / 0,
              the SUBTOTAL formula is set at the bottom, and today's totals are appended to
              <b> NO COM HISTORY</b>. Native pivots refresh on file open.
            </p>
          </div>

          <button
            onClick={run}
            disabled={!canRun}
            className="w-full sm:w-auto inline-flex items-center justify-center gap-2 rounded-md bg-primary px-6 py-3 text-primary-foreground font-medium shadow-sm hover:opacity-90 disabled:opacity-50 disabled:cursor-not-allowed transition-opacity"
          >
            {loading ? <><Loader2 className="h-4 w-4 animate-spin" /> Updating tracker…</> : <>Update SLA Tracker</>}
          </button>

          {error && (
            <div className="rounded-md border border-destructive/30 bg-destructive/5 p-4 text-sm text-destructive flex items-start gap-2">
              <AlertTriangle className="h-4 w-4 shrink-0 mt-0.5" />
              <div>{error}</div>
            </div>
          )}

          {/* ---- End of Month Consolidation ---- */}
          <div className="rounded-xl border-2 border-dashed border-blue-300/60 bg-blue-500/5 p-6 space-y-4">
            <div className="flex items-center gap-2">
              <CalendarDays className="h-5 w-5 text-blue-500" />
              <h2 className="font-semibold">End of Month — Statistics & Predictions</h2>
            </div>
            <p className="text-sm text-muted-foreground leading-relaxed">
              Run at the end of each month to generate a standalone Excel with:
              monthly KPIs, region breakdown, top recurring NO COM sites,
              month-over-month comparison, and next-week forecasts using
              <b> linear regression</b> and <b>weighted moving average</b>.
            </p>

            <div className="grid sm:grid-cols-2 gap-4">
              <label className="block">
                <span className="text-sm font-medium">Month</span>
                <select
                  value={monthlyMonth}
                  onChange={(e) => setMonthlyMonth(Number(e.target.value))}
                  className="mt-1 w-full rounded-md border bg-background px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ring"
                >
                  {MONTH_NAMES.map((name, i) => (
                    <option key={i} value={i + 1}>{name}</option>
                  ))}
                </select>
              </label>
              <label className="block">
                <span className="text-sm font-medium">Year</span>
                <input
                  type="number"
                  value={monthlyYear}
                  min={2023}
                  max={2030}
                  onChange={(e) => setMonthlyYear(Number(e.target.value))}
                  className="mt-1 w-full rounded-md border bg-background px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ring"
                />
              </label>
            </div>

            <button
              onClick={runMonthly}
              disabled={!canRunMonthly}
              className="inline-flex items-center gap-2 rounded-md bg-blue-600 px-5 py-2.5 text-white text-sm font-medium hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
            >
              {monthlyLoading
                ? <><Loader2 className="h-4 w-4 animate-spin" /> Consolidating…</>
                : <><CalendarDays className="h-4 w-4" /> Generate Monthly Report</>}
            </button>

            {monthlyError && (
              <div className="rounded-md border border-destructive/30 bg-destructive/5 p-3 text-sm text-destructive flex items-start gap-2">
                <AlertTriangle className="h-4 w-4 shrink-0 mt-0.5" />
                <div>{monthlyError}</div>
              </div>
            )}

            {monthlyResult && monthlySummary && (
              <div className="rounded-lg border bg-card p-4 space-y-3">
                <div className="flex items-center gap-2 text-sm font-semibold">
                  <CheckCircle2 className="h-4 w-4 text-emerald-500" />
                  {monthlySummary.month} — report ready
                </div>
                <div className="grid grid-cols-2 gap-x-6 gap-y-1 text-sm">
                  <Row label="Days recorded" value={monthlySummary.totalDays} />
                  <Row label="Avg NO COM sites/day" value={monthlySummary.avgNoComSites} />
                  <Row label="Avg NO COM %" value={`${monthlySummary.avgNoComPct}%`} />
                  <Row label="Peak day" value={`${monthlySummary.peakDay} (${monthlySummary.peakCount})`} />
                  <Row label="Best day" value={`${monthlySummary.bestDay} (${monthlySummary.bestCount})`} />
                </div>
                <div className={`flex items-center gap-1.5 text-sm font-medium ${trendColor}`}>
                  <TrendIcon className="h-4 w-4" />
                  Trend: {monthlySummary.trend.toUpperCase()}
                </div>
                <div className="grid grid-cols-2 gap-x-6 text-sm">
                  <Row label="LR forecast next week" value={monthlySummary.linearForecastNextWeek} />
                  <Row label="WMA forecast next week" value={monthlySummary.wmaForecastNextWeek} />
                </div>
                {monthlySummary.topRecurringSites.length > 0 && (
                  <div className="text-xs text-muted-foreground">
                    <span className="font-medium">Top recurring sites: </span>
                    {monthlySummary.topRecurringSites.slice(0, 5).map((s) => `${s.site} (${s.pct}%)`).join(" · ")}
                    {monthlySummary.topRecurringSites.length > 5 && " · …"}
                  </div>
                )}
                <button
                  onClick={() => downloadBlob(monthlyResult.blob, monthlyResult.filename)}
                  className="w-full inline-flex items-center justify-center gap-2 rounded-md bg-blue-600 px-4 py-2 text-white text-sm font-medium hover:bg-blue-700"
                >
                  <Download className="h-4 w-4" /> Download Monthly Report
                </button>
                <p className="text-xs text-muted-foreground">
                  3 sheets: MONTHLY STATS · PREDICTIONS (LR + WMA) · SITE RISK SCORES
                </p>
              </div>
            )}
          </div>
        </section>

        <aside className="space-y-4">
          <div className="rounded-xl border bg-card p-5 sticky top-6">
            <h3 className="font-semibold">Result</h3>
            {!result ? (
              <p className="mt-2 text-sm text-muted-foreground">
                Run the analysis to see the summary here.
              </p>
            ) : (
              <div className="mt-3 space-y-3">
                <div className="flex items-center gap-2 text-sm text-emerald-600 dark:text-emerald-400">
                  <CheckCircle2 className="h-4 w-4" /> Tracker updated
                </div>
                {summary && (
                  <div className="space-y-1.5 text-sm">
                    <Row label="On Air sites" value={summary.onAirSites} muted />
                    <Row label="Sites in NO COM" value={summary.noComSites} accent />
                    <Row label="More than 24h" value={summary.moreThan24h} />
                    <Row label="Within 48h" value={summary.within48h} />
                    {summary.todayColumn && (
                      <div className="text-xs text-muted-foreground pt-1">
                        Wrote column <b>{summary.todayColumn}</b> in Comm Issue
                      </div>
                    )}
                    {summary.monthReset === true && (
                      <div className="flex items-center gap-1.5 rounded-md bg-amber-500/10 border border-amber-400/40 px-3 py-1.5 text-xs font-medium text-amber-700 dark:text-amber-400">
                        🔄 Month reset — date columns cleared and rewritten for the new month
                      </div>
                    )}
                    {summary.monthReset === false && (
                      <div className="text-xs text-muted-foreground pt-0.5">
                        Month detection: same month, no reset
                      </div>
                    )}
                    <div className="h-px bg-border my-2" />
                    <Row label="Tickets to create" value={summary.ticketsToCreate} accent />
                    <Row label="Tickets to resolve" value={summary.ticketsToResolve} accent />
                  </div>
                )}
                <button
                  onClick={() => downloadBlob(result.blob, result.filename)}
                  className="w-full inline-flex items-center justify-center gap-2 rounded-md bg-accent px-4 py-2 text-accent-foreground text-sm font-medium hover:opacity-90"
                >
                  <Download className="h-4 w-4" /> Download SLA Tracker
                </button>
                {result.ticketListBlob && (
                  <button
                    onClick={() => {
                      downloadBlob(
                        result.ticketListBlob!,
                        result.ticketListFilename ?? "NoCom_TicketList.xlsx",
                      );
                    }}
                    className="w-full inline-flex items-center justify-center gap-2 rounded-md border border-primary/40 bg-primary/10 px-4 py-2 text-primary text-sm font-medium hover:bg-primary/20"
                  >
                    <ListChecks className="h-4 w-4" /> Download Ticket List
                  </button>
                )}
                <p className="text-xs text-muted-foreground leading-snug">
                  The SLA Tracker is updated in place (today's column, SUBTOTAL, NO COM HISTORY, Sheet9 pivot). The ticket list is a separate Excel with two sheets: tickets to create and tickets to resolve.
                </p>
              </div>
            )}
          </div>
        </aside>
      </main>
    </div>
  );
}

function Row({
  label, value, accent, muted,
}: {
  label: string; value: number | string; accent?: boolean; muted?: boolean;
}) {
  return (
    <div className="flex items-center justify-between">
      <span className={muted ? "text-muted-foreground" : ""}>{label}</span>
      <span className={`tabular-nums font-semibold ${accent ? "text-primary" : ""}`}>{value}</span>
    </div>
  );
}
