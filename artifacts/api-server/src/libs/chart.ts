import { createCanvas } from "@napi-rs/canvas";

export interface LinePoint {
  label: string;
  value: number;
}

export interface LineChartOptions {
  width: number;
  height: number;
  title: string;
  subtitle?: string;
  data: LinePoint[];
  background: string; // background color
  titleColor: string;
  lineColor: string;
  pointColor: string;
  labelColor: string;
  showGrid?: boolean;
  showAxis?: boolean;
  yMin?: number;
  yMax?: number;
  yStep?: number;
}

export function renderLineChart(opt: LineChartOptions): Buffer {
  const W = opt.width;
  const H = opt.height;
  const canvas = createCanvas(W, H);
  const ctx = canvas.getContext("2d");

  // Background
  ctx.fillStyle = opt.background;
  ctx.fillRect(0, 0, W, H);

  // Title
  ctx.fillStyle = opt.titleColor;
  ctx.textAlign = "center";
  ctx.font = "bold 22px sans-serif";
  const titleLines = opt.title.split("\n");
  let ty = 36;
  for (const tl of titleLines) {
    ctx.fillText(tl, W / 2, ty);
    ty += 26;
  }
  if (opt.subtitle) {
    ctx.font = "bold 18px sans-serif";
    ctx.fillText(opt.subtitle, W / 2, ty);
    ty += 22;
  }

  // Compute layout
  const padTop = ty + 30;
  const padBottom = 70;
  const padLeft = opt.showAxis ? 60 : 40;
  const padRight = 40;
  const plotW = W - padLeft - padRight;
  const plotH = H - padTop - padBottom;

  if (opt.data.length === 0) return canvas.toBuffer("image/png");

  const values = opt.data.map((d) => d.value);
  const dataMin = Math.min(...values);
  const dataMax = Math.max(...values);
  const yMin = opt.yMin ?? Math.max(0, Math.floor((dataMin - 5) / 5) * 5);
  const yMax = opt.yMax ?? Math.ceil((dataMax + 5) / 5) * 5;
  const yRange = Math.max(1, yMax - yMin);

  // Grid + Y axis
  if (opt.showGrid) {
    ctx.strokeStyle = "rgba(255,255,255,0.15)";
    ctx.lineWidth = 1;
    const step = opt.yStep ?? 5;
    ctx.font = "12px sans-serif";
    ctx.fillStyle = opt.labelColor;
    ctx.textAlign = "right";
    for (let v = yMin; v <= yMax; v += step) {
      const y = padTop + plotH - ((v - yMin) / yRange) * plotH;
      ctx.beginPath();
      ctx.moveTo(padLeft, y);
      ctx.lineTo(padLeft + plotW, y);
      ctx.stroke();
      if (opt.showAxis) ctx.fillText(String(v), padLeft - 8, y + 4);
    }
  }

  // X positions
  const n = opt.data.length;
  const xStep = n > 1 ? plotW / (n - 1) : 0;
  const points = opt.data.map((d, i) => ({
    x: padLeft + (n === 1 ? plotW / 2 : i * xStep),
    y: padTop + plotH - ((d.value - yMin) / yRange) * plotH,
    value: d.value,
    label: d.label,
  }));

  // Vertical drop lines from each point to baseline (the SLA tracker style)
  ctx.strokeStyle = "rgba(255,255,255,0.35)";
  ctx.lineWidth = 1;
  for (const p of points) {
    ctx.beginPath();
    ctx.moveTo(p.x, p.y + 6);
    ctx.lineTo(p.x, padTop + plotH);
    ctx.stroke();
  }

  // Line
  ctx.strokeStyle = opt.lineColor;
  ctx.lineWidth = 3;
  ctx.lineJoin = "round";
  ctx.beginPath();
  points.forEach((p, i) => {
    if (i === 0) ctx.moveTo(p.x, p.y);
    else ctx.lineTo(p.x, p.y);
  });
  ctx.stroke();

  // Points + value labels above
  ctx.fillStyle = opt.pointColor;
  ctx.font = "bold 14px sans-serif";
  ctx.textAlign = "center";
  for (const p of points) {
    ctx.beginPath();
    ctx.arc(p.x, p.y, 4, 0, Math.PI * 2);
    ctx.fill();
    ctx.fillStyle = opt.labelColor;
    ctx.fillText(String(p.value), p.x, p.y - 10);
    ctx.fillStyle = opt.pointColor;
  }

  // X labels (rotated if many)
  ctx.fillStyle = opt.labelColor;
  ctx.font = "11px sans-serif";
  const rotate = n > 8;
  for (const p of points) {
    if (rotate) {
      ctx.save();
      ctx.translate(p.x, padTop + plotH + 14);
      ctx.rotate(-Math.PI / 2.2);
      ctx.textAlign = "right";
      ctx.fillText(p.label.toUpperCase(), 0, 0);
      ctx.restore();
    } else {
      ctx.textAlign = "center";
      ctx.fillText(p.label.toUpperCase(), p.x, padTop + plotH + 22);
    }
  }

  return canvas.toBuffer("image/png");
}
