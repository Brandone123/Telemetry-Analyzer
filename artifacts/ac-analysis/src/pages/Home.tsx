import { useState } from "react";
import { Link } from "wouter";
import {
  ArrowRight,
  Radio,
  Gauge,
  Activity,
  Battery,
  Droplet,
} from "lucide-react";

const analyses = [
  {
    id: "no-com",
    title: "NO COM Analysis",
    description:
      "Identify sites that need No-Communication tickets to be created or resolved, reproducing the SLA Tracker pivot and trend charts.",
    icon: Radio,
    color: "from-blue-500 to-cyan-500",
    href: "/analysis/no-com",
    available: true,
  },
  {
    id: "running-hours",
    title: "Running Hours Analysis",
    description:
      "Investigate sites with abnormally low or high generator running hours, after filtering out non-comm sites and open Power/Gen tickets.",
    icon: Gauge,
    color: "from-orange-500 to-amber-500",
    href: "/analysis/running-hours",
    available: true,
  },
  {
    id: "ac",
    title: "AC Issue Analysis",
    description:
      "Detect grid availability issues and AC alarm patterns using two methods: alert code filtering (M1) and ERS Big Data continuous-power anomaly detection (M2).",
    icon: Activity,
    color: "from-purple-500 to-fuchsia-500",
    href: "/analysis/ac",
    available: true,
  },
  {
    id: "dc",
    title: "DC Issue Analysis",
    description: "Analyze DC voltage and battery health alarms across sites.",
    icon: Battery,
    color: "from-emerald-500 to-teal-500",
    href: "#",
    available: false,
  },
  {
    id: "fuel",
    title: "Fuel Sensor Analysis",
    description:
      "Filter components at 0 / −1 (Location = Quant), cross-check against fuel sensor tickets, and output a VLOOKUP table with sites needing new tickets.",
    icon: Droplet,
    color: "from-rose-500 to-pink-500",
    href: "/analysis/fuel-sensor",
    available: true,
  },
];

export default function Home() {
  const [now] = useState(() => new Date());

  return (
    <div className="min-h-screen bg-gradient-to-b from-background via-background to-muted/30">
      <header className="border-b border-border/50 bg-card/50 backdrop-blur sticky top-0 z-10">
        <div className="max-w-6xl mx-auto px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="h-9 w-9 rounded-lg bg-gradient-to-br from-primary to-accent grid place-items-center text-primary-foreground font-bold">
              T
            </div>
            <div>
              <div className="font-semibold tracking-tight">
                Telemetry Analyzer
              </div>
              <div className="text-xs text-muted-foreground">
                CMS analyses · Cameroon BTS network
              </div>
            </div>
          </div>
          <div className="text-sm text-muted-foreground tabular-nums">
            {now.toLocaleDateString("en-US", {
              weekday: "long",
              year: "numeric",
              month: "long",
              day: "numeric",
            })}
          </div>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-6 py-10">
        <div className="mb-10">
          <h1 className="text-3xl font-semibold tracking-tight">
            Choose an analysis
          </h1>
          <p className="mt-2 text-muted-foreground max-w-2xl">
            Upload the CMS exports, set the parameters, and download an Excel
            report that reproduces your SLA Tracker pivot tables and trend
            charts — with formulas you can audit.
          </p>
        </div>

        <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-4">
          {analyses.map((a) => {
            const Icon = a.icon;
            const card = (
              <div
                className={`group relative h-full rounded-xl border bg-card p-5 transition-all ${a.available ? "hover:border-primary/50 hover:shadow-lg cursor-pointer" : "opacity-60"}`}
              >
                <div
                  className={`h-10 w-10 rounded-lg bg-gradient-to-br ${a.color} grid place-items-center text-white shadow`}
                >
                  <Icon className="h-5 w-5" />
                </div>
                <h3 className="mt-4 font-semibold">{a.title}</h3>
                <p className="mt-1 text-sm text-muted-foreground line-clamp-3">
                  {a.description}
                </p>
                <div className="mt-4 flex items-center justify-between">
                  {a.available ? (
                    <span className="text-sm font-medium text-primary inline-flex items-center gap-1 group-hover:gap-2 transition-all">
                      Open <ArrowRight className="h-4 w-4" />
                    </span>
                  ) : (
                    <span className="text-xs px-2 py-0.5 rounded-full bg-muted text-muted-foreground">
                      Coming soon
                    </span>
                  )}
                </div>
              </div>
            );
            return a.available ? (
              <Link key={a.id} href={a.href}>
                {card}
              </Link>
            ) : (
              <div key={a.id}>{card}</div>
            );
          })}
        </div>

        <div className="mt-12 rounded-xl border bg-card p-6">
          <h2 className="font-semibold">How it works</h2>
          <ol className="mt-3 space-y-2 text-sm text-muted-foreground list-decimal pl-5">
            <li>Pick the analysis you want to run.</li>
            <li>
              Upload the required CMS files (SLA Tracker, Communication Report,
              ticket exports…).
            </li>
            <li>
              Tune the parameters — cutoff days, thresholds — for your week.
            </li>
            <li>
              The server VLOOKUPs the comm report against the tracker, filters
              On Air sites, builds the Region × NOSO/NOT NOSO/AIOT/VIP pivot,
              appends today to the history, and renders the trend charts as
              images embedded in the Excel.
            </li>
            <li>
              Open it in Excel — formulas, pivot, charts and ticket lists are
              ready.
            </li>
          </ol>
        </div>
      </main>

      <footer className="max-w-6xl mx-auto px-6 py-8 text-xs text-muted-foreground">
        Built for GAIO Cameroon Limited — only SLA / On Air sites are included
        in every analysis.
      </footer>
    </div>
  );
}
