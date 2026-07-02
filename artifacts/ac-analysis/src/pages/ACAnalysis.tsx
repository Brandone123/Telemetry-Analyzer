import { useState } from "react";
import { Link } from "wouter";
import {
  ArrowLeft,
  Download,
  Loader2,
  Activity,
  AlertTriangle,
  CheckCircle2,
} from "lucide-react";
import { FileDrop } from "@/components/FileDrop";
import { runACAnalysis, downloadBlob, type AnalysisResponse } from "@/lib/api";

interface AcSummary {
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
}

export default function ACAnalysis() {
  const [alertsFile, setAlertsFile] = useState<File | null>(null);
  const [ersBigData, setErsBigData] = useState<File | null>(null);
  const [ticketsFile, setTicketsFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [result, setResult] = useState<AnalysisResponse | null>(null);

  const canRun = !!alertsFile && !!ersBigData && !loading;
  const summary = result?.summary as AcSummary | null;

  async function run() {
    if (!alertsFile || !ersBigData) return;
    setLoading(true);
    setError(null);
    setResult(null);
    try {
      const res = await runACAnalysis({ alertsFile, ersBigData, ticketsFile });
      setResult(res);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="min-h-screen bg-gradient-to-b from-background to-muted/30">
      <header className="border-b border-border/50 bg-card/50 backdrop-blur">
        <div className="max-w-5xl mx-auto px-6 py-4 flex items-center justify-between">
          <Link
            href="/"
            className="text-sm text-muted-foreground hover:text-foreground inline-flex items-center gap-1"
          >
            <ArrowLeft className="h-4 w-4" /> Back
          </Link>
          <div className="flex items-center gap-2">
            <div className="h-8 w-8 rounded-lg bg-gradient-to-br from-purple-500 to-fuchsia-500 grid place-items-center text-white">
              <Activity className="h-4 w-4" />
            </div>
            <span className="font-semibold">AC Issue Analysis</span>
          </div>
        </div>
      </header>

      <main className="max-w-5xl mx-auto px-6 py-8 grid lg:grid-cols-[1fr_300px] gap-6">
        <section className="space-y-6">
          <div className="rounded-xl border bg-card p-6 space-y-5">
            <div>
              <h2 className="font-semibold">Upload files</h2>
              <p className="text-sm text-muted-foreground mt-1">
                Two methods: alert code filtering (M1) and ERS Big Data power
                anomaly detection (M2). Output is a 2-sheet Excel.
              </p>
            </div>
            <FileDrop
              label="Alerts file"
              hint="CMS alerts export — AC powercom codes (AL690-0055, EM690-0001, E_ACM-0001, USER-0016)"
              required
              file={alertsFile}
              onChange={setAlertsFile}
            />
            <FileDrop
              label="ERS Big Data"
              hint="5-day ERS export — detects sites with no AC power consumption despite being on grid"
              required
              file={ersBigData}
              onChange={setErsBigData}
            />
            <FileDrop
              label="AC Tickets"
              hint="Existing AC tickets — sites with an open ticket will be excluded from results"
              file={ticketsFile}
              onChange={setTicketsFile}
            />
          </div>

          <button
            onClick={run}
            disabled={!canRun}
            className="w-full sm:w-auto inline-flex items-center justify-center gap-2 rounded-md bg-primary px-6 py-3 text-primary-foreground font-medium shadow-sm hover:opacity-90 disabled:opacity-50 disabled:cursor-not-allowed transition-opacity"
          >
            {loading ? (
              <>
                <Loader2 className="h-4 w-4 animate-spin" /> Running analysis…
              </>
            ) : (
              <>Run AC Analysis</>
            )}
          </button>

          {error && (
            <div className="rounded-md border border-destructive/30 bg-destructive/5 p-4 text-sm text-destructive flex items-start gap-2">
              <AlertTriangle className="h-4 w-4 shrink-0 mt-0.5" />
              <div>{error}</div>
            </div>
          )}
        </section>

        {/* Sidebar */}
        <aside className="space-y-4">
          <div className="rounded-xl border bg-card p-5 sticky top-6 space-y-4">
            <h3 className="font-semibold">Result</h3>

            {!result ? (
              <p className="text-sm text-muted-foreground">
                Run the analysis to see the summary here.
              </p>
            ) : (
              <>
                <div className="flex items-center gap-2 text-sm text-emerald-600 dark:text-emerald-400">
                  <CheckCircle2 className="h-4 w-4" /> Analysis complete
                </div>

                {summary && (
                  <div className="space-y-4 text-sm">
                    <div>
                      <p className="font-semibold text-purple-600 dark:text-purple-400 mb-1.5">
                        M1 — Alert codes
                      </p>
                      <div className="space-y-1">
                        <SummaryRow
                          label="Total alerts"
                          value={summary.m1TotalAlerts}
                          muted
                        />
                        <SummaryRow
                          label="Unique sites"
                          value={summary.m1UniqueSites}
                          muted
                        />
                        <div className="h-px bg-border my-1" />
                        <SummaryRow
                          label="Tickets to create"
                          value={summary.m1SitesNeedingTicket}
                          accent
                        />
                      </div>
                      {Object.keys(summary.m1ByCode).length > 0 && (
                        <div className="mt-2 space-y-1">
                          {Object.entries(summary.m1ByCode).map(
                            ([code, count]) => (
                              <div
                                key={code}
                                className="flex justify-between text-xs text-muted-foreground"
                              >
                                <code className="bg-muted px-1 rounded">
                                  {code}
                                </code>
                                <span className="font-medium">{count}</span>
                              </div>
                            ),
                          )}
                        </div>
                      )}
                    </div>

                    <div className="h-px bg-border" />

                    <div>
                      <p className="font-semibold text-fuchsia-600 dark:text-fuchsia-400 mb-1.5">
                        M2 — ERS Big Data
                      </p>
                      {summary.dateRange?.length > 0 && (
                        <p className="text-xs text-muted-foreground mb-1.5">
                          Period: {summary.dateRange[0]} →{" "}
                          {summary.dateRange[summary.dateRange.length - 1]}
                        </p>
                      )}
                      <div className="space-y-1">
                        <SummaryRow
                          label="Matching sites"
                          value={summary.m2TotalSites}
                          muted
                        />
                        <div className="h-px bg-border my-1" />
                        <SummaryRow
                          label="Tickets to create"
                          value={summary.m2SitesNeedingTicket}
                          accent
                        />
                      </div>
                    </div>

                    <div className="h-px bg-border" />

                    <div>
                      <p className="font-semibold text-emerald-700 dark:text-emerald-400 mb-1.5">
                        False Tickets
                      </p>
                      <SummaryRow
                        label="AC ok but open ticket → to close"
                        value={summary.falseTickets ?? 0}
                        accent
                      />
                    </div>
                  </div>
                )}

                <button
                  onClick={() => downloadBlob(result.blob, result.filename)}
                  className="w-full inline-flex items-center justify-center gap-2 rounded-md bg-accent px-4 py-2 text-accent-foreground text-sm font-medium hover:opacity-90"
                >
                  <Download className="h-4 w-4" /> Download Excel
                </button>
                <p className="text-xs text-muted-foreground leading-snug">
                  Sheets: M1 tickets · M2 tickets · False tickets to close ·
                  Summary
                </p>
              </>
            )}
          </div>
        </aside>
      </main>
    </div>
  );
}

function SummaryRow({
  label,
  value,
  accent,
  muted,
}: {
  label: string;
  value: number;
  accent?: boolean;
  muted?: boolean;
}) {
  return (
    <div className="flex items-center justify-between">
      <span className={muted ? "text-muted-foreground" : ""}>{label}</span>
      <span
        className={`tabular-nums font-semibold ${accent ? "text-primary" : ""}`}
      >
        {value}
      </span>
    </div>
  );
}
