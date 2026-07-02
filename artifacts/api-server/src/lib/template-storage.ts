import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const DATA_DIR = path.resolve(__dirname, "..", "..", "data");

export interface TemplateInfo {
  exists: boolean;
  filename?: string;
  sizeBytes?: number;
  uploadedAt?: string; // ISO date
}

const TEMPLATES = {
  slaTracker: { file: "sla_tracker_template.xlsx", meta: "sla_tracker_template.json" },
} as const;

export type TemplateKey = keyof typeof TEMPLATES;

async function ensureDir() {
  await fs.mkdir(DATA_DIR, { recursive: true });
}

export async function saveTemplate(key: TemplateKey, buffer: Buffer, originalName: string): Promise<TemplateInfo> {
  await ensureDir();
  const t = TEMPLATES[key];
  const filePath = path.join(DATA_DIR, t.file);
  const metaPath = path.join(DATA_DIR, t.meta);
  await fs.writeFile(filePath, buffer);
  const info: TemplateInfo = {
    exists: true,
    filename: originalName,
    sizeBytes: buffer.length,
    uploadedAt: new Date().toISOString(),
  };
  await fs.writeFile(metaPath, JSON.stringify(info, null, 2));
  return info;
}

export async function loadTemplate(key: TemplateKey): Promise<Buffer | null> {
  const t = TEMPLATES[key];
  const filePath = path.join(DATA_DIR, t.file);
  try {
    return await fs.readFile(filePath);
  } catch {
    return null;
  }
}

export async function getTemplateInfo(key: TemplateKey): Promise<TemplateInfo> {
  const t = TEMPLATES[key];
  const metaPath = path.join(DATA_DIR, t.meta);
  try {
    const raw = await fs.readFile(metaPath, "utf-8");
    return JSON.parse(raw);
  } catch {
    return { exists: false };
  }
}

export async function deleteTemplate(key: TemplateKey): Promise<void> {
  const t = TEMPLATES[key];
  await Promise.allSettled([
    fs.rm(path.join(DATA_DIR, t.file), { force: true }),
    fs.rm(path.join(DATA_DIR, t.meta), { force: true }),
  ]);
}
