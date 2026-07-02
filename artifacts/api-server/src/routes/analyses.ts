import { Router, type IRouter, type Request, type Response } from "express";
import multer from "multer";
import { analyzeNoCom } from "../analyses/no-com";
import { analyzeRunningHours } from "../analyses/running-hours";
import { analyzeFuelSensor } from "../analyses/fuel-sensor";
import { analyzeAC } from "../analyses/ac";
import { consolidateMonthlyNoCom } from "../analyses/no-com-monthly";
import { saveTemplate, loadTemplate, getTemplateInfo, deleteTemplate } from "../lib/template-storage";

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 80 * 1024 * 1024 } });
const router: IRouter = Router();

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

function getFile(req: Request, name: string): Buffer | null {
  const files = req.files as { [field: string]: Express.Multer.File[] } | undefined;
  if (!files) return null;
  const f = files[name]?.[0];
  return f ? f.buffer : null;
}
function getFileWithName(req: Request, name: string): { buffer: Buffer; originalname: string } | null {
  const files = req.files as { [field: string]: Express.Multer.File[] } | undefined;
  if (!files) return null;
  const f = files[name]?.[0];
  return f ? { buffer: f.buffer, originalname: f.originalname } : null;
}

// Helper: send analysis result as JSON (file as base64) so it survives any proxy
function sendAnalysisJson(res: Response, data: {
  buffer: Buffer;
  filename: string;
  summary: unknown;
  extra?: Record<string, unknown>;
}) {
  res.json({
    filename: data.filename,
    file: data.buffer.toString("base64"),
    mimeType: XLSX_MIME,
    summary: data.summary,
    ...data.extra,
  });
}

// ----- SLA Tracker template management -----
router.get("/template/sla-tracker", async (_req, res) => {
  res.json(await getTemplateInfo("slaTracker"));
});

router.post(
  "/template/sla-tracker",
  upload.single("file"),
  async (req: Request, res: Response) => {
    const file = req.file;
    if (!file) { res.status(400).json({ error: "Missing 'file' field" }); return; }
    const info = await saveTemplate("slaTracker", file.buffer, file.originalname);
    res.json(info);
  },
);

router.delete("/template/sla-tracker", async (_req, res) => {
  await deleteTemplate("slaTracker");
  res.json({ ok: true });
});

// ----- NO COM analysis -----
router.post(
  "/analyses/no-com",
  upload.fields([
    { name: "slaTracker", maxCount: 1 },
    { name: "commReport", maxCount: 1 },
    { name: "noComTickets", maxCount: 1 },
  ]),
  async (req: Request, res: Response) => {
    try {
      let tracker = getFileWithName(req, "slaTracker");
      const comm = getFile(req, "commReport");
      const tickets = getFile(req, "noComTickets");
      if (!comm || !tickets) {
        res.status(400).json({ error: "commReport and noComTickets files are required" });
        return;
      }
      if (tracker) {
        await saveTemplate("slaTracker", tracker.buffer, tracker.originalname);
      } else {
        const stored = await loadTemplate("slaTracker");
        if (!stored) {
          res.status(400).json({ error: "No SLA Tracker template stored. Upload one first." });
          return;
        }
        tracker = { buffer: stored, originalname: "sla_tracker_template.xlsx" };
      }
      const cutoffDays = Number(req.body.cutoffDays ?? 2);
      const referenceDate = req.body.referenceDate ? new Date(String(req.body.referenceDate)) : new Date();
      const result = await analyzeNoCom(tracker.buffer, comm, tickets, { cutoffDays, referenceDate });
      await saveTemplate("slaTracker", result.buffer, result.filename);
      sendAnalysisJson(res, {
        buffer: result.buffer,
        filename: result.filename,
        summary: result.summary,
        extra: {
          ticketListFile: result.ticketListBuffer.toString("base64"),
          ticketListFilename: result.ticketListFilename,
        },
      });
    } catch (err) {
      req.log.error({ err }, "no-com analysis failed");
      res.status(500).json({ error: (err as Error).message });
    }
  },
);

// ----- Running Hours analysis -----
router.post(
  "/analyses/running-hours",
  upload.fields([
    { name: "ersBigData", maxCount: 1 },
    { name: "noComTickets", maxCount: 1 },
    { name: "powerTickets", maxCount: 1 },
    { name: "genTickets", maxCount: 1 },
  ]),
  async (req: Request, res: Response) => {
    try {
      const ers = getFile(req, "ersBigData");
      if (!ers) {
        res.status(400).json({ error: "ersBigData file is required" });
        return;
      }
      const periodDays = Number(req.body.periodDays ?? 4);
      const lowThresholdHours = Number(req.body.lowThresholdHours ?? 80);
      const highThresholdHours = Number(req.body.highThresholdHours ?? periodDays * 22);
      const batteryAbnormalThreshold = Number(req.body.batteryAbnormalThreshold ?? 47.5);
      const noCommDailyThreshold = Number(req.body.noCommDailyThreshold ?? 20);
      const result = await analyzeRunningHours(
        ers,
        getFile(req, "noComTickets"),
        getFile(req, "powerTickets"),
        getFile(req, "genTickets"),
        { periodDays, lowThresholdHours, highThresholdHours, batteryAbnormalThreshold, noCommDailyThreshold },
      );
      sendAnalysisJson(res, { buffer: result.buffer, filename: result.filename, summary: result.summary });
    } catch (err) {
      req.log.error({ err }, "running-hours analysis failed");
      res.status(500).json({ error: (err as Error).message });
    }
  },
);

// ----- Fuel Sensor analysis -----
router.post(
  "/analyses/fuel-sensor",
  upload.fields([
    { name: "componentsFile", maxCount: 1 },
    { name: "ticketsFile", maxCount: 1 },
  ]),
  async (req: Request, res: Response) => {
    try {
      const components = getFile(req, "componentsFile");
      const tickets = getFile(req, "ticketsFile");
      if (!components || !tickets) {
        res.status(400).json({ error: "componentsFile and ticketsFile are both required" });
        return;
      }
      const result = await analyzeFuelSensor(components, tickets);
      sendAnalysisJson(res, { buffer: result.buffer, filename: result.filename, summary: result.summary });
    } catch (err) {
      req.log.error({ err }, "fuel-sensor analysis failed");
      res.status(500).json({ error: (err as Error).message });
    }
  },
);

// ----- AC analysis -----
router.post(
  "/analyses/ac",
  upload.fields([
    { name: "alertsFile", maxCount: 1 },
    { name: "ersBigData", maxCount: 1 },
    { name: "ticketsFile", maxCount: 1 },
  ]),
  async (req: Request, res: Response) => {
    try {
      const alerts = getFile(req, "alertsFile");
      const ers = getFile(req, "ersBigData");
      if (!alerts || !ers) {
        res.status(400).json({ error: "alertsFile and ersBigData files are required" });
        return;
      }
      const tickets = getFile(req, "ticketsFile");
      const result = await analyzeAC(alerts, ers, tickets);
      sendAnalysisJson(res, { buffer: result.buffer, filename: result.filename, summary: result.summary });
    } catch (err) {
      req.log.error({ err }, "ac analysis failed");
      res.status(500).json({ error: (err as Error).message });
    }
  },
);

// ----- NO COM Monthly consolidation -----
router.post(
  "/analyses/no-com/monthly",
  upload.fields([
    { name: "slaTracker", maxCount: 1 },
    { name: "commReport", maxCount: 1 },
    { name: "noComTickets", maxCount: 1 },
  ]),
  async (req: Request, res: Response) => {
    try {
      let trackerBuf: Buffer | null = getFile(req, "slaTracker");
      if (!trackerBuf) {
        const stored = await loadTemplate("slaTracker");
        if (!stored) {
          res.status(400).json({ error: "No SLA Tracker available. Upload one first or register a template." });
          return;
        }
        trackerBuf = stored;
      }
      const month = req.body.month ? Number(req.body.month) : undefined;
      const year = req.body.year ? Number(req.body.year) : undefined;
      const referenceDate = req.body.referenceDate ? new Date(String(req.body.referenceDate)) : new Date();

      const result = await consolidateMonthlyNoCom(
        trackerBuf,
        getFile(req, "commReport"),
        getFile(req, "noComTickets"),
        { month, year, referenceDate },
      );
      sendAnalysisJson(res, { buffer: result.buffer, filename: result.filename, summary: result.summary });
    } catch (err) {
      req.log.error({ err }, "no-com monthly consolidation failed");
      res.status(500).json({ error: (err as Error).message });
    }
  },
);

export default router;
