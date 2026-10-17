import type { ChartSpec, ResolvedChart } from '@sheets/core';

export interface ChartTheme {
  dark: boolean;
  background: string;
  text: string;
  muted: string;
  grid: string;
  fontFamily: string;
}

/** Categorical palette (validated for color-vision deficiencies), light and dark steps. */
export const PALETTE_LIGHT = ['#2a78d6', '#eb6834', '#1baf7a', '#eda100', '#e87ba4', '#008300', '#6250d6', '#e34948'];
export const PALETTE_DARK = ['#3987e5', '#d95926', '#199e70', '#c98500', '#d55181', '#008300', '#9085e9', '#e66767'];

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

function niceStep(range: number, ticks: number): number {
  const raw = range / Math.max(1, ticks);
  const mag = Math.pow(10, Math.floor(Math.log10(raw || 1)));
  const n = raw / mag;
  return (n <= 1 ? 1 : n <= 2 ? 2 : n <= 2.5 ? 2.5 : n <= 5 ? 5 : 10) * mag;
}

export function niceScale(min: number, max: number, ticks = 5): { min: number; max: number; step: number } {
  if (!Number.isFinite(min) || !Number.isFinite(max)) return { min: 0, max: 1, step: 0.2 };
  if (min === max) {
    if (min === 0) return { min: 0, max: 1, step: 0.2 };
    min = Math.min(0, min);
    max = Math.max(0, max);
  }
  const step = niceStep(max - min, ticks);
  return { min: Math.floor(min / step) * step, max: Math.ceil(max / step) * step, step };
}

export function compactNumber(v: number): string {
  const a = Math.abs(v);
  const fmt = (x: number, s: string) => `${Number(x.toFixed(1))}${s}`;
  if (a >= 1e9) return fmt(v / 1e9, 'B');
  if (a >= 1e6) return fmt(v / 1e6, 'M');
  if (a >= 1e4) return fmt(v / 1e3, 'k');
  return String(Number(v.toPrecision(6)));
}

function textWidth(s: string, size: number): number {
  return s.length * size * 0.56;
}

function truncate(s: string, maxWidth: number, size: number): string {
  if (textWidth(s, size) <= maxWidth) return s;
  const n = Math.max(1, Math.floor(maxWidth / (size * 0.56)) - 1);
  return s.slice(0, n) + '…';
}

/** Renders a chart as an SVG string. */
export function renderChartSvg(data: ResolvedChart, spec: ChartSpec, width: number, height: number, theme: ChartTheme): string {
  const palette = theme.dark ? PALETTE_DARK : PALETTE_LIGHT;
  const color = (i: number) => data.series[i]?.color ?? palette[i % palette.length]!;
  const parts: string[] = [];
  const font = `font-family="${esc(theme.fontFamily)}"`;
  let top = 10;
  let bottom = height - 8;
  let left = 10;
  let right = width - 12;

  if (spec.title) {
    parts.push(`<text x="${width / 2}" y="24" text-anchor="middle" font-size="15" font-weight="600" fill="${theme.text}">${esc(truncate(spec.title, width - 20, 15))}</text>`);
    top = 38;
  }

  // Legend (always for >= 2 series; pies list categories).
  const isPie = spec.type === 'pie' || spec.type === 'doughnut';
  const legendItems = isPie ? data.categories.map((c, i) => ({ label: c, color: palette[i % palette.length]! })) : data.series.map((s, i) => ({ label: s.name, color: color(i) }));
  const legendPos = spec.legend ?? (legendItems.length > 1 ? 'right' : 'none');
  if (legendPos !== 'none' && legendItems.length > (isPie ? 0 : 1)) {
    const items = legendItems.slice(0, 20);
    if (legendPos === 'right' || legendPos === 'left') {
      const w = Math.min(160, Math.max(...items.map((it) => textWidth(it.label, 12))) + 26);
      const x0 = legendPos === 'right' ? width - w - 6 : 8;
      const y0 = Math.max(top, (height - items.length * 18) / 2);
      items.forEach((it, i) => {
        parts.push(`<rect x="${x0}" y="${y0 + i * 18}" width="10" height="10" rx="2" fill="${it.color}"/>`);
        parts.push(`<text x="${x0 + 16}" y="${y0 + i * 18 + 9}" font-size="12" fill="${theme.text}">${esc(truncate(it.label, w - 22, 12))}</text>`);
      });
      if (legendPos === 'right') right -= w + 8;
      else left += w + 8;
    } else {
      let x = 0;
      const row: string[] = [];
      for (const it of items) {
        const label = truncate(it.label, 140, 12);
        row.push(`<rect x="${x}" y="0" width="10" height="10" rx="2" fill="${it.color}"/><text x="${x + 14}" y="9" font-size="12" fill="${theme.text}">${esc(label)}</text>`);
        x += textWidth(label, 12) + 30;
      }
      const y = legendPos === 'top' ? top : height - 20;
      parts.push(`<g transform="translate(${Math.max(8, (width - x) / 2)},${y})">${row.join('')}</g>`);
      if (legendPos === 'top') top += 22;
      else bottom -= 24;
    }
  }

  if (!data.series.length || data.series.every((s) => s.values.every((v) => v === null))) {
    parts.push(`<text x="${width / 2}" y="${height / 2}" text-anchor="middle" font-size="12" fill="${theme.muted}">—</text>`);
    return wrap(parts, width, height, font, theme);
  }

  if (isPie) {
    renderPie(parts, data, spec, { left, top, right, bottom }, palette, theme);
    return wrap(parts, width, height, font, theme);
  }

  // Axis titles
  if (spec.yAxisTitle) {
    parts.push(`<text transform="translate(${left + 10},${(top + bottom) / 2}) rotate(-90)" text-anchor="middle" font-size="12" fill="${theme.muted}">${esc(spec.yAxisTitle)}</text>`);
    left += 18;
  }
  if (spec.xAxisTitle) {
    parts.push(`<text x="${(left + right) / 2}" y="${bottom - 2}" text-anchor="middle" font-size="12" fill="${theme.muted}">${esc(spec.xAxisTitle)}</text>`);
    bottom -= 18;
  }

  if (spec.type === 'scatter') renderScatter(parts, data, { left, top, right, bottom }, color, theme);
  else renderCategorical(parts, data, spec, { left, top, right, bottom }, color, theme);
  return wrap(parts, width, height, font, theme);
}

function wrap(parts: string[], width: number, height: number, font: string, theme: ChartTheme): string {
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${width} ${height}" width="${width}" height="${height}" ${font} role="img"><rect width="100%" height="100%" fill="${theme.background}"/>${parts.join('')}</svg>`;
}

interface Box {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

function valueAxis(parts: string[], scale: { min: number; max: number; step: number }, box: Box, horizontal: boolean, theme: ChartTheme): (v: number) => number {
  const span = scale.max - scale.min || 1;
  const pos = horizontal
    ? (v: number) => box.left + ((v - scale.min) / span) * (box.right - box.left)
    : (v: number) => box.bottom - ((v - scale.min) / span) * (box.bottom - box.top);
  for (let v = scale.min, i = 0; v <= scale.max + scale.step / 2 && i < 50; v += scale.step, i++) {
    const p = pos(v);
    const label = compactNumber(Number(v.toPrecision(12)));
    if (horizontal) {
      parts.push(`<line x1="${p}" x2="${p}" y1="${box.top}" y2="${box.bottom}" stroke="${theme.grid}" stroke-width="1"/>`);
      parts.push(`<text x="${p}" y="${box.bottom + 14}" text-anchor="middle" font-size="11" fill="${theme.muted}">${label}</text>`);
    } else {
      parts.push(`<line x1="${box.left}" x2="${box.right}" y1="${p}" y2="${p}" stroke="${theme.grid}" stroke-width="1"/>`);
      parts.push(`<text x="${box.left - 6}" y="${p + 4}" text-anchor="end" font-size="11" fill="${theme.muted}">${label}</text>`);
    }
  }
  return pos;
}

function renderCategorical(parts: string[], data: ResolvedChart, spec: ChartSpec, outer: Box, color: (i: number) => string, theme: ChartTheme): void {
  const horizontal = spec.type === 'bar';
  const n = Math.max(data.categories.length, ...data.series.map((s) => s.values.length));
  const stacked = !!spec.stacked;
  let min = 0;
  let max = 0;
  for (let i = 0; i < n; i++) {
    let pos = 0;
    let neg = 0;
    for (const s of data.series) {
      const v = s.values[i] ?? 0;
      if (stacked) {
        if (v >= 0) pos += v;
        else neg += v;
      } else {
        max = Math.max(max, v);
        min = Math.min(min, v);
      }
    }
    if (stacked) {
      max = Math.max(max, pos);
      min = Math.min(min, neg);
    }
  }
  const scale = niceScale(min, max);
  const labelWidth = horizontal ? Math.min(120, Math.max(...data.categories.map((c) => textWidth(c, 11)), 20) + 8) : Math.max(...[scale.min, scale.max].map((v) => textWidth(compactNumber(v), 11))) + 10;
  const box: Box = { left: outer.left + labelWidth, top: outer.top + 4, right: outer.right, bottom: outer.bottom - 20 };
  const vpos = valueAxis(parts, scale, box, horizontal, theme);
  const bandLen = ((horizontal ? box.bottom - box.top : box.right - box.left) || 1) / Math.max(1, n);
  const band = (i: number) => (horizontal ? box.top + i * bandLen : box.left + i * bandLen);

  // Category labels (thinned to avoid collisions)
  const every = Math.max(1, Math.ceil((n * (horizontal ? 14 : 48)) / ((horizontal ? box.bottom - box.top : box.right - box.left) || 1)));
  for (let i = 0; i < n; i += every) {
    const label = esc(truncate(data.categories[i] ?? String(i + 1), horizontal ? labelWidth - 8 : bandLen * every - 4, 11));
    if (horizontal) parts.push(`<text x="${box.left - 6}" y="${band(i) + bandLen / 2 + 4}" text-anchor="end" font-size="11" fill="${theme.muted}">${label}</text>`);
    else parts.push(`<text x="${band(i) + bandLen / 2}" y="${box.bottom + 14}" text-anchor="middle" font-size="11" fill="${theme.muted}">${label}</text>`);
  }
  const zero = vpos(Math.max(scale.min, Math.min(0, scale.max)));
  // Baseline
  parts.push(horizontal
    ? `<line x1="${zero}" x2="${zero}" y1="${box.top}" y2="${box.bottom}" stroke="${theme.muted}" stroke-width="1"/>`
    : `<line x1="${box.left}" x2="${box.right}" y1="${zero}" y2="${zero}" stroke="${theme.muted}" stroke-width="1"/>`);

  if (spec.type === 'column' || spec.type === 'bar') {
    const groups = stacked ? 1 : data.series.length;
    const inner = bandLen * 0.72;
    const barLen = Math.max(1, inner / groups - (groups > 1 ? 2 : 0));
    const posAcc = new Array(n).fill(0);
    const negAcc = new Array(n).fill(0);
    data.series.forEach((s, si) => {
      s.values.forEach((v, i) => {
        if (v === null) return;
        const base = stacked ? (v >= 0 ? posAcc[i] : negAcc[i]) : 0;
        const end = base + v;
        if (stacked) {
          if (v >= 0) posAcc[i] = end;
          else negAcc[i] = end;
        }
        const a = vpos(base);
        const b = vpos(end);
        const offset = band(i) + (bandLen - inner) / 2 + (stacked ? 0 : si * (barLen + 2));
        const tip = `<title>${esc(s.name)} · ${esc(data.categories[i] ?? '')}: ${compactNumber(v)}</title>`;
        const gap = stacked ? 1 : 0; // surface gap between stacked segments
        if (horizontal) {
          const x = Math.min(a, b);
          const w = Math.max(0, Math.abs(b - a) - gap);
          parts.push(`<rect x="${x}" y="${offset}" width="${w}" height="${barLen}" rx="${Math.min(2, barLen / 2)}" fill="${color(si)}">${tip}</rect>`);
        } else {
          const y = Math.min(a, b);
          const h = Math.max(0, Math.abs(b - a) - gap);
          parts.push(`<rect x="${offset}" y="${y + (v >= 0 ? gap : 0)}" width="${barLen}" height="${h}" rx="${Math.min(2, barLen / 2)}" fill="${color(si)}">${tip}</rect>`);
        }
      });
    });
    return;
  }

  // Line & area
  const acc = new Array(n).fill(0);
  const layers = data.series.map((s) =>
    s.values.map((v, i) => {
      if (v === null) return null;
      const base = stacked ? acc[i] : 0;
      if (stacked) acc[i] += v;
      return { x: band(i) + bandLen / 2, y: vpos(base + v), y0: vpos(base), v, i };
    }),
  );
  layers.forEach((pts, si) => {
    const valid = pts.filter((p): p is NonNullable<typeof p> => p !== null);
    if (!valid.length) return;
    const line = valid.map((p, k) => `${k ? 'L' : 'M'}${p.x.toFixed(1)},${p.y.toFixed(1)}`).join('');
    if (spec.type === 'area') {
      const back = [...valid].reverse().map((p) => `L${p.x.toFixed(1)},${p.y0.toFixed(1)}`).join('');
      parts.push(`<path d="${line}${back}Z" fill="${color(si)}" fill-opacity="${stacked ? 0.85 : 0.35}" stroke="none"/>`);
    }
    parts.push(`<path d="${line}" fill="none" stroke="${color(si)}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>`);
    // Invisible enlarged hit targets carry the tooltips; visible dots only when sparse.
    for (const p of valid) {
      if (n <= 24) parts.push(`<circle cx="${p.x}" cy="${p.y}" r="3" fill="${color(si)}" stroke="${theme.background}" stroke-width="1.5"/>`);
      parts.push(`<circle cx="${p.x}" cy="${p.y}" r="8" fill="transparent"><title>${esc(data.series[si]!.name)} · ${esc(data.categories[p.i] ?? '')}: ${compactNumber(p.v)}</title></circle>`);
    }
  });
}

function renderScatter(parts: string[], data: ResolvedChart, outer: Box, color: (i: number) => string, theme: ChartTheme): void {
  const xs = data.series.flatMap((s) => (s.x ?? s.values.map((_, i) => i + 1)).filter((v): v is number => v !== null));
  const ys = data.series.flatMap((s) => s.values.filter((v): v is number => v !== null));
  const sx = niceScale(Math.min(...xs), Math.max(...xs));
  const sy = niceScale(Math.min(0, ...ys), Math.max(...ys));
  const labelWidth = Math.max(...[sy.min, sy.max].map((v) => textWidth(compactNumber(v), 11))) + 10;
  const box: Box = { left: outer.left + labelWidth, top: outer.top + 4, right: outer.right, bottom: outer.bottom - 20 };
  const ypos = valueAxis(parts, sy, box, false, theme);
  const xpos = valueAxis(parts, sx, box, true, theme);
  data.series.forEach((s, si) => {
    s.values.forEach((v, i) => {
      const x = s.x ? s.x[i] : i + 1;
      if (v === null || x === null || x === undefined) return;
      parts.push(
        `<circle cx="${xpos(x)}" cy="${ypos(v)}" r="4.5" fill="${color(si)}" fill-opacity="0.85" stroke="${theme.background}" stroke-width="1.5"><title>${esc(s.name)}: (${compactNumber(x)}, ${compactNumber(v)})</title></circle>`,
      );
    });
  });
}

function renderPie(parts: string[], data: ResolvedChart, spec: ChartSpec, box: Box, palette: string[], theme: ChartTheme): void {
  const values = (data.series[0]?.values ?? []).map((v) => (v !== null && v > 0 ? v : 0));
  const total = values.reduce((a, b) => a + b, 0);
  if (total <= 0) return;
  const cx = (box.left + box.right) / 2;
  const cy = (box.top + box.bottom) / 2;
  const r = Math.max(10, Math.min(box.right - box.left, box.bottom - box.top) / 2 - 6);
  const inner = spec.type === 'doughnut' ? r * 0.55 : 0;
  let angle = -Math.PI / 2;
  values.forEach((v, i) => {
    if (v <= 0) return;
    const a2 = angle + (v / total) * Math.PI * 2;
    const large = a2 - angle > Math.PI ? 1 : 0;
    const p = (a: number, rr: number) => `${(cx + rr * Math.cos(a)).toFixed(2)},${(cy + rr * Math.sin(a)).toFixed(2)}`;
    const full = v === total;
    let d: string;
    if (full) {
      d = `M${cx - r},${cy}a${r},${r} 0 1,0 ${2 * r},0a${r},${r} 0 1,0 ${-2 * r},0` + (inner ? `M${cx - inner},${cy}a${inner},${inner} 0 1,1 ${2 * inner},0a${inner},${inner} 0 1,1 ${-2 * inner},0` : '');
    } else if (inner) {
      d = `M${p(angle, r)}A${r},${r} 0 ${large} 1 ${p(a2, r)}L${p(a2, inner)}A${inner},${inner} 0 ${large} 0 ${p(angle, inner)}Z`;
    } else {
      d = `M${cx},${cy}L${p(angle, r)}A${r},${r} 0 ${large} 1 ${p(a2, r)}Z`;
    }
    const pct = (v / total) * 100;
    parts.push(`<path d="${d}" fill="${palette[i % palette.length]}" fill-rule="evenodd" stroke="${theme.background}" stroke-width="2"><title>${esc(data.categories[i] ?? '')}: ${compactNumber(v)} (${pct.toFixed(1)}%)</title></path>`);
    if (pct >= 6) {
      const mid = (angle + a2) / 2;
      const lr = inner ? (r + inner) / 2 : r * 0.65;
      parts.push(`<text x="${cx + lr * Math.cos(mid)}" y="${cy + lr * Math.sin(mid) + 4}" text-anchor="middle" font-size="11" font-weight="600" fill="#ffffff">${Math.round(pct)}%</text>`);
    }
    angle = a2;
  });
}
