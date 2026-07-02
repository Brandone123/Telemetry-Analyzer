import { useState, useEffect } from "react";
import { Link } from "wouter";
import {
  ArrowLeft,
  Download,
  Loader2,
  Gauge,
  AlertTriangle,
  CheckCircle2,
} from "lucide-react";
import { FileDrop } from "@/components/FileDrop";
import {
  runRunningHoursAnalysis,
  downloadBlob,
  type AnalysisResponse,
} from "@/lib/api";

interface Summary {
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
}

export default function RunningHoursAnalysis() {
  const [ers, setErs] = useState<File | null>(null);
  const [noComTickets, setNoComTickets] = useState<File | null>(null);
  const [power, setPower] = useState<File | null>(null);
  const [gen, setGen] = useState<File | null>(null);
  const [periodDays, setPeriodDays] = useState(4);
  const [low, setLow] = useState(80);
  const [high, setHigh] = useState(88);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [result, setResult] = useState<AnalysisResponse | null>(null);

  useEffect(() => {
    setHigh(periodDays * 22);
    setLow(periodDays * 20);
  }, [periodDays]);

  const summary = result?.summary as Summary | null;
  const canRun = !!ers && !loading;

  async function run() {
    if (!ers) return;
    setLoading(true);
    setError(null);
    setResult(null);
    try {
      const res = await runRunningHoursAnalysis({
        ersBigData: ers,
        noComTickets,
        powerTickets: power,
        genTickets: gen,
        periodDays,
        lowThresholdHours: low,
        highThresholdHours: high,
      });
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
            <div className="h-8 w-8 rounded-lg bg-gradient-to-br from-orange-500 to-amber-500 grid place-items-center text-white">
              <Gauge className="h-4 w-4" />
            </div>
            <span className="font-semibold">Running Hours Analysis</span>
          </div>
        </div>
      </header>

      <main className="max-w-5xl mx-auto px-6 py-8 grid lg:grid-cols-[1fr_300px] gap-6">
        <section className="space-y-6">
          <div className="rounded-xl border bg-card p-6 space-y-5">
            <div>
              <h2 className="font-semibold">Upload files</h2>
              <p className="text-sm text-muted-foreground mt-1">
                ERS Big Data (4 days, Daily summary). Optional ticket files are
                used to exclude sites with existing open tickets.
              </p>
            </div>
            <FileDrop
              label="ERS Big Data"
              hint="4-day daily ERS export — SiteName, Day, TotalPower RH, Gen RH, No Comm. H"
              required
              file={ers}
              onChange={setErs}
            />
            <FileDrop
              label="NO COM Tickets"
              hint="Exclude sites with an open NO COM ticket"
              file={noComTickets}
              onChange={setNoComTickets}
            />
            <FileDrop
              label="Power Tickets"
              hint="Exclude sites with an open Power ticket"
              file={power}
              onChange={setPower}
            />
            <FileDrop
              label="Gen on Load Tickets"
              hint="Exclude sites already ticketed for Gen on Load"
              file={gen}
              onChange={setGen}
            />
          </div>

          <div className="rounded-xl border bg-card p-6 space-y-3">
            <h2 className="font-semibold">2 · What this analysis does</h2>
            <ol className="list-decimal pl-5 space-y-1.5 text-sm text-muted-foreground">
              <li>Loads the ERS Big Data export and keeps only <b>SLA / On Air</b> sites.</li>
              <li>Excludes sites with <b>average No Comm ≥ 20 h/day</b> (unreliable data) and sites that already have an open Power, Gen, or NO COM ticket.</li>
              <li>Flags sites whose cumulative generator running hours fall <b>below the low threshold</b> (gen underused or not started) or <b>above the high threshold</b> (gen running almost continuously).</li>
              <li>Detects <b>Gen on Load</b> cases — gen present in the topology but 0 h recorded, indicating a counting or start-up fault.</li>
              <li>
                Outputs a 5-sheet Excel:
                <ul className="list-disc pl-5 mt-1 space-y-0.5">
                  <li><b>Low RH</b> — sites below threshold.</li>
                  <li><b>High RH</b> — sites above threshold.</li>
                  <li><b>Gen on Load</b> — sites with gen in topology but 0 h.</li>
                  <li><b>All Sites</b> — full filtered dataset.</li>
                  <li><b>Summary</b> — counts and exclusion breakdown.</li>
                </ul>
              </li>
            </ol>
          </div>

          <div className="rounded-xl border bg-card p-6 space-y-4">
            <h2 className="font-semibold">3 · Parameters</h2>
            <div className="grid sm:grid-cols-3 gap-4">
              <label className="block">
                <span className="text-sm font-medium">Period (days)</span>
                <input
                  type="number"
                  min={1}
                  value={periodDays}
                  onChange={(e) =>
                    setPeriodDays(Math.max(1, Number(e.target.value)))
                  }
                  className="mt-1 w-full rounded-md border bg-background px-3 py-2 text-sm"
                />
              </label>
              <label className="block">
                <span className="text-sm font-medium">Low threshold (h)</span>
                <input
                  type="number"
                  value={low}
                  onChange={(e) => setLow(Number(e.target.value))}
                  className="mt-1 w-full rounded-md border bg-background px-3 py-2 text-sm"
                />
                <span className="text-xs text-muted-foreground">
                  Total RH below → investigate
                </span>
              </label>
              <label className="block">
                <span className="text-sm font-medium">High threshold (h)</span>
                <input
                  type="number"
                  value={high}
                  onChange={(e) => setHigh(Number(e.target.value))}
                  className="mt-1 w-full rounded-md border bg-background px-3 py-2 text-sm"
                />
                <span className="text-xs text-muted-foreground">
                  Total RH above → investigate
                </span>
              </label>
            </div>
          </div>

          <button
            onClick={run}
            disabled={!canRun}
            className="w-full sm:w-auto inline-flex items-center justify-center gap-2 rounded-md bg-primary px-6 py-3 text-primary-foreground font-medium shadow-sm hover:opacity-90 disabled:opacity-50 disabled:cursor-not-allowed"
          >
            {loading ? (
              <>
                <Loader2 className="h-4 w-4 animate-spin" /> Analyzing…
              </>
            ) : (
              <>Run analysis</>
            )}
          </button>

          {error && (
            <div className="rounded-md border border-destructive/30 bg-destructive/5 p-4 text-sm text-destructive flex items-start gap-2">
              <AlertTriangle className="h-4 w-4 shrink-0 mt-0.5" />
              <div>{error}</div>
            </div>
          )}
        </section>

        <aside>
          <div className="rounded-xl border bg-card p-5 sticky top-6 space-y-4">
            <h3 className="font-semibold">Result</h3>
            {!result ? (
              <p className="text-sm text-muted-foreground">
                Run the analysis to see the summary here.
              </p>
            ) : (
              <>
                <div className="flex items-center gap-2 text-sm text-emerald-600 dark:text-emerald-400">
                  <CheckCircle2 className="h-4 w-4" /> Analysis ready
                </div>
                {summary && (
                  <div className="space-y-1.5 text-sm">
                    <Row
                      label="SLA sites in ERS"
                      value={summary.slaSites}
                      muted
                    />
                    <Row
                      label="Excl. — not communicating"
                      value={summary.excludedNoComm}
                      muted
                    />
                    <Row
                      label="Excl. — open Power"
                      value={summary.excludedOpenPower}
                      muted
                    />
                    <Row
                      label="Excl. — open NO COM"
                      value={summary.excludedOpenNoCom}
                      muted
                    />
                    <Row
                      label="Excl. — open Gen"
                      value={summary.excludedOpenGen}
                      muted
                    />
                    <Row label="Candidates" value={summary.candidates} />
                    <div className="h-px bg-border my-1.5" />
                    <Row
                      label="LOW RH (investigate)"
                      value={summary.lowRH}
                      accent
                    />
                    <Row
                      label="HIGH RH (investigate)"
                      value={summary.highRH}
                      accent
                    />
                    <Row
                      label="Gen on Load — tickets"
                      value={summary.genOnLoad}
                      accent
                    />
                  </div>
                )}
                <button
                  onClick={() => downloadBlob(result.blob, result.filename)}
                  className="w-full inline-flex items-center justify-center gap-2 rounded-md bg-accent px-4 py-2 text-accent-foreground text-sm font-medium hover:opacity-90"
                >
                  <Download className="h-4 w-4" /> Download Excel
                </button>
                <p className="text-xs text-muted-foreground leading-snug">
                  Sheets: Low RH · High RH · Gen on Load tickets · All Sites ·
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

function Row({
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
