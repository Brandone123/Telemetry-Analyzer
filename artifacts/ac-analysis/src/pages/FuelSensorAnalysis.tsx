import { useState } from "react";
import { Link } from "wouter";
import {
  ArrowLeft, Download, Loader2, Droplet, AlertTriangle, CheckCircle2,
} from "lucide-react";
import { FileDrop } from "@/components/FileDrop";
import { runFuelSensorAnalysis, downloadBlob, type AnalysisResponse } from "@/lib/api";

interface Summary {
  totalComponents: number;
  uniqueSites: number;
  sitesWithOpenTicket: number;
  sitesNeedingTicket: number;
}

export default function FuelSensorAnalysis() {
  const [componentsFile, setComponentsFile] = useState<File | null>(null);
  const [ticketsFile, setTicketsFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [result, setResult] = useState<AnalysisResponse | null>(null);

  const canRun = !!componentsFile && !!ticketsFile && !loading;
  const summary = result?.summary as Summary | null;

  async function run() {
    if (!componentsFile || !ticketsFile) return;
    setLoading(true);
    setError(null);
    setResult(null);
    try {
      const res = await runFuelSensorAnalysis({ componentsFile, ticketsFile });
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
          <Link href="/" className="text-sm text-muted-foreground hover:text-foreground inline-flex items-center gap-1">
            <ArrowLeft className="h-4 w-4" /> Back
          </Link>
          <div className="flex items-center gap-2">
            <div className="h-8 w-8 rounded-lg bg-gradient-to-br from-rose-500 to-pink-500 grid place-items-center text-white">
              <Droplet className="h-4 w-4" />
            </div>
            <span className="font-semibold">Fuel Sensor Analysis</span>
          </div>
        </div>
      </header>

      <main className="max-w-5xl mx-auto px-6 py-8 grid lg:grid-cols-[1fr_300px] gap-6">
        <section className="space-y-6">
          {/* Inputs */}
          <div className="rounded-xl border bg-card p-6 space-y-5">
            <h2 className="font-semibold">1 · Upload files</h2>
            <FileDrop
              label="Components file"
              hint="Excel/CSV with Location and Value columns — rows where Location=Quant and Value=0 or -1 will be extracted"
              required
              file={componentsFile}
              onChange={setComponentsFile}
            />
            <FileDrop
              label="Fuel Sensor Tickets"
              hint="Excel/CSV export of general fuel sensor tickets — used to check which sites already have an open ticket"
              required
              file={ticketsFile}
              onChange={setTicketsFile}
            />
          </div>

          {/* Info box */}
          <div className="rounded-xl border bg-card p-6 space-y-3">
            <h2 className="font-semibold">2 · What this analysis does</h2>
            <ol className="list-decimal pl-5 space-y-1.5 text-sm text-muted-foreground">
              <li>Filters the Components file where <b>Location = Quant</b> and <b>Value = 0 or −1</b>.</li>
              <li>Extracts the site name from each matching row.</li>
              <li>Cross-references against the Fuel Sensor ticket list to identify open tickets per site.</li>
              <li>
                Outputs a 3-sheet Excel:
                <ul className="list-disc pl-5 mt-1 space-y-0.5">
                  <li><b>VLOOKUP Table</b> — every filtered component row with a "Has Open Ticket" and "Ticket ID" column.</li>
                  <li><b>Sites Needing Tickets</b> — sites with no open fuel sensor ticket.</li>
                  <li><b>Sites Already Covered</b> — sites that already have an open ticket.</li>
                </ul>
              </li>
            </ol>
          </div>

          <button
            onClick={run}
            disabled={!canRun}
            className="w-full sm:w-auto inline-flex items-center justify-center gap-2 rounded-md bg-primary px-6 py-3 text-primary-foreground font-medium shadow-sm hover:opacity-90 disabled:opacity-50 disabled:cursor-not-allowed transition-opacity"
          >
            {loading
              ? <><Loader2 className="h-4 w-4 animate-spin" /> Running analysis…</>
              : <>Run Fuel Sensor Analysis</>}
          </button>

          {error && (
            <div className="rounded-md border border-destructive/30 bg-destructive/5 p-4 text-sm text-destructive flex items-start gap-2">
              <AlertTriangle className="h-4 w-4 shrink-0 mt-0.5" />
              <div>{error}</div>
            </div>
          )}
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
                  <CheckCircle2 className="h-4 w-4" /> Analysis complete
                </div>
                {summary && (
                  <div className="space-y-1.5 text-sm">
                    <Row label="Components filtered" value={summary.totalComponents} muted />
                    <Row label="Unique sites" value={summary.uniqueSites} />
                    <div className="h-px bg-border my-1" />
                    <Row label="Sites with open ticket" value={summary.sitesWithOpenTicket} />
                    <Row label="Sites needing ticket" value={summary.sitesNeedingTicket} accent />
                  </div>
                )}
                <button
                  onClick={() => downloadBlob(result.blob, result.filename)}
                  className="w-full inline-flex items-center justify-center gap-2 rounded-md bg-accent px-4 py-2 text-accent-foreground text-sm font-medium hover:opacity-90"
                >
                  <Download className="h-4 w-4" /> Download Excel
                </button>
                <p className="text-xs text-muted-foreground leading-snug">
                  3-sheet Excel: full VLOOKUP table, sites to ticket, and sites already covered.
                </p>
              </div>
            )}
          </div>
        </aside>
      </main>
    </div>
  );
}

function Row({ label, value, accent, muted }: { label: string; value: number; accent?: boolean; muted?: boolean }) {
  return (
    <div className="flex items-center justify-between">
      <span className={muted ? "text-muted-foreground" : ""}>{label}</span>
      <span className={`tabular-nums font-semibold ${accent ? "text-primary" : ""}`}>{value}</span>
    </div>
  );
}
