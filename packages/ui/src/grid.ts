import { cellKey, colToName, MAX_COLS, formatGeneral, formatValueEx, parseRange, tableAt, type CellStyle, type FormatLocale, type RangeAddress, type Sheet } from '@sheets/core';
import type { Geometry } from './geometry.js';

export interface GridTheme {
  bg: string;
  fg: string;
  muted: string;
  grid: string;
  headerBg: string;
  headerFg: string;
  headerActiveBg: string;
  headerActiveFg: string;
  accent: string;
  selection: string;
  fontFamily: string;
  dark: boolean;
}

export const DEFAULT_CELL_FONT = 'Calibri, Carlito, "Segoe UI", system-ui, -apple-system, sans-serif';
export const DEFAULT_FONT_SIZE_PT = 11;

export interface Highlight {
  range: RangeAddress;
  color: string;
  dashed?: boolean;
}

/** Scroll/frozen-pane aware mapping between sheet coordinates and screen pixels. */
export class Viewport {
  sheet!: Sheet;
  geo!: Geometry;
  scrollX = 0;
  scrollY = 0;
  width = 0;
  height = 0;
  headerW = 46;
  headerH = 24;

  get frozenRows(): number {
    return this.sheet.frozen.rows;
  }
  get frozenCols(): number {
    return this.sheet.frozen.cols;
  }
  get frozenW(): number {
    return this.geo.cols.offset(this.frozenCols);
  }
  get frozenH(): number {
    return this.geo.rows.offset(this.frozenRows);
  }

  x(col: number): number {
    return this.headerW + this.geo.cols.offset(col) - (col >= this.frozenCols ? this.scrollX : 0);
  }
  y(row: number): number {
    return this.headerH + this.geo.rows.offset(row) - (row >= this.frozenRows ? this.scrollY : 0);
  }

  colAt(x: number): number {
    const xr = x - this.headerW;
    return xr < this.frozenW ? this.geo.cols.indexAt(xr) : this.geo.cols.indexAt(xr + this.scrollX);
  }
  rowAt(y: number): number {
    const yr = y - this.headerH;
    return yr < this.frozenH ? this.geo.rows.indexAt(yr) : this.geo.rows.indexAt(yr + this.scrollY);
  }

  /** Screen rect of a cell (merged ranges should be passed as a range). */
  rect(range: RangeAddress): { x: number; y: number; w: number; h: number } {
    // Start and end are mapped independently so ranges spanning frozen and
    // scrolled panes stay attached to both.
    const x = this.x(range.start.col);
    const y = this.y(range.start.row);
    const right = this.x(range.end.col) + this.geo.cols.size(range.end.col);
    const bottom = this.y(range.end.row) + this.geo.rows.size(range.end.row);
    return { x, y, w: Math.max(0, right - x), h: Math.max(0, bottom - y) };
  }

  /** Range of rows visible in the scrolling pane. */
  visibleRows(): [number, number] {
    const start = Math.max(this.frozenRows, this.geo.rows.indexAt(this.frozenH + this.scrollY));
    const end = this.geo.rows.indexAt(this.frozenH + this.scrollY + (this.height - this.headerH - this.frozenH));
    return [start, end];
  }
  visibleCols(): [number, number] {
    const start = Math.max(this.frozenCols, this.geo.cols.indexAt(this.frozenW + this.scrollX));
    const end = this.geo.cols.indexAt(this.frozenW + this.scrollX + (this.width - this.headerW - this.frozenW));
    return [start, end];
  }
}

export interface RenderState {
  cellStyle?: (row: number, col: number) => CellStyle | undefined;
  selection: RangeAddress;
  active: { row: number; col: number };
  highlights: Highlight[];
  copyRange: RangeAddress | null;
  fillPreview: RangeAddress | null;
  locale: FormatLocale;
  showFillHandle: boolean;
  /** Cells that hold a hyperlink (drawn underlined in the link color). */
  isLink?: (row: number, col: number) => boolean;
}

function luminance(hex: string): number {
  const h = hex.replace('#', '');
  if (h.length < 6) return 1;
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16) / 255) as [number, number, number];
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

export function cssFont(style: CellStyle | undefined, theme: GridTheme, zoom = 1): string {
  const size = ((style?.fontSize ?? DEFAULT_FONT_SIZE_PT) * 96) / 72 * zoom;
  const family = style?.fontName ? `"${style.fontName}", ${theme.fontFamily}` : theme.fontFamily;
  return `${style?.italic ? 'italic ' : ''}${style?.bold ? 'bold ' : ''}${size.toFixed(1)}px ${family}`;
}

interface Pane {
  rows: [number, number];
  cols: [number, number];
  clip: { x: number; y: number; w: number; h: number };
}

export class GridRenderer {
  readonly canvas: HTMLCanvasElement;
  private ctx: CanvasRenderingContext2D;
  private dpr = 1;
  theme: GridTheme;

  constructor(canvas: HTMLCanvasElement, theme: GridTheme) {
    this.canvas = canvas;
    this.ctx = canvas.getContext('2d')!;
    this.theme = theme;
  }

  resize(width: number, height: number): void {
    this.dpr = window.devicePixelRatio || 1;
    this.canvas.width = Math.max(1, Math.floor(width * this.dpr));
    this.canvas.height = Math.max(1, Math.floor(height * this.dpr));
    this.canvas.style.width = `${width}px`;
    this.canvas.style.height = `${height}px`;
  }

  draw(vp: Viewport, state: RenderState): void {
    const ctx = this.ctx;
    const t = this.theme;
    ctx.setTransform(this.dpr, 0, 0, this.dpr, 0, 0);
    ctx.fillStyle = t.bg;
    ctx.fillRect(0, 0, vp.width, vp.height);

    const fr = vp.frozenRows;
    const fc = vp.frozenCols;
    const mainX = vp.headerW + vp.frozenW;
    const mainY = vp.headerH + vp.frozenH;
    const rows = vp.visibleRows();
    const cols = vp.visibleCols();
    const panes: Pane[] = [{ rows, cols, clip: { x: mainX, y: mainY, w: vp.width - mainX, h: vp.height - mainY } }];
    if (fr) panes.push({ rows: [0, fr - 1], cols, clip: { x: mainX, y: vp.headerH, w: vp.width - mainX, h: vp.frozenH } });
    if (fc) panes.push({ rows, cols: [0, fc - 1], clip: { x: vp.headerW, y: mainY, w: vp.frozenW, h: vp.height - mainY } });
    if (fr && fc) panes.push({ rows: [0, fr - 1], cols: [0, fc - 1], clip: { x: vp.headerW, y: vp.headerH, w: vp.frozenW, h: vp.frozenH } });
    for (const p of panes) this.drawPane(vp, state, p);
    this.drawHeaders(vp, state, rows, cols);

    // Frozen pane dividers
    ctx.strokeStyle = t.muted;
    ctx.lineWidth = 1.5;
    if (fr) this.line(vp.headerW, mainY - 0.5, vp.width, mainY - 0.5);
    if (fc) this.line(mainX - 0.5, vp.headerH, mainX - 0.5, vp.height);
  }

  private line(x1: number, y1: number, x2: number, y2: number): void {
    const ctx = this.ctx;
    ctx.beginPath();
    ctx.moveTo(x1, y1);
    ctx.lineTo(x2, y2);
    ctx.stroke();
  }

  private drawPane(vp: Viewport, state: RenderState, pane: Pane): void {
    const ctx = this.ctx;
    const t = this.theme;
    const sheet = vp.sheet;
    const styleAt = state.cellStyle ?? ((r: number, c: number) => sheet.getCell(r, c)?.style);
    const { clip } = pane;
    if (clip.w <= 0 || clip.h <= 0) return;
    ctx.save();
    ctx.beginPath();
    ctx.rect(clip.x, clip.y, clip.w, clip.h);
    ctx.clip();

    const [r0, r1] = pane.rows;
    const [c0, c1] = pane.cols;
    // Positions within one pane share the same scroll offset; the trailing
    // edge is derived from sizes so it never jumps across a frozen boundary.
    const colX: number[] = [];
    for (let c = c0; c <= c1; c++) colX.push(vp.x(c));
    colX.push(colX[colX.length - 1]! + vp.geo.cols.size(c1));
    const rowY: number[] = [];
    for (let r = r0; r <= r1; r++) rowY.push(vp.y(r));
    rowY.push(rowY[rowY.length - 1]! + vp.geo.rows.size(r1));
    const cx = (c: number) => colX[c - c0]!;
    const ry = (r: number) => rowY[r - r0]!;

    // Merges touching this pane
    const merges = sheet.merges.filter((m) => m.end.row >= r0 && m.start.row <= r1 && m.end.col >= c0 && m.start.col <= c1);
    const covered = new Set<number>();
    for (const m of merges) for (let r = m.start.row; r <= m.end.row; r++) for (let c = m.start.col; c <= m.end.col; c++) covered.add(cellKey(r, c));

    // 1. Fills
    for (let r = r0; r <= r1; r++) {
      const h = ry(r + 1) - ry(r);
      if (h <= 0) continue;
      for (let c = c0; c <= c1; c++) {
        const fill = styleAt(r, c)?.fill;
        if (!fill || covered.has(cellKey(r, c))) continue;
        ctx.fillStyle = fill;
        ctx.fillRect(cx(c), ry(r), cx(c + 1) - cx(c), h);
      }
    }

    // 2. Grid lines
    ctx.strokeStyle = t.grid;
    ctx.lineWidth = 1;
    ctx.beginPath();
    for (let c = c0; c <= c1 + 1; c++) {
      const x = Math.round(cx(c)) - 0.5;
      ctx.moveTo(x, clip.y);
      ctx.lineTo(x, clip.y + clip.h);
    }
    for (let r = r0; r <= r1 + 1; r++) {
      const y = Math.round(ry(r)) - 0.5;
      ctx.moveTo(clip.x, y);
      ctx.lineTo(clip.x + clip.w, y);
    }
    ctx.stroke();

    // 3. Merged cell backgrounds
    for (const m of merges) {
      const rc = vp.rect(m);
      ctx.fillStyle = styleAt(m.start.row, m.start.col)?.fill ?? t.bg;
      ctx.fillRect(rc.x, rc.y, rc.w - 1, rc.h - 1);
    }

    // 4. Text
    for (let r = r0; r <= r1; r++) {
      const y = ry(r);
      const h = ry(r + 1) - y;
      if (h <= 0) continue;
      for (let c = c0; c <= c1; c++) {
        const key = cellKey(r, c);
        if (covered.has(key)) continue;
        const cell = sheet.getCell(r, c);
        const style = styleAt(r, c);
        if (!cell || cell.value === null) continue;
        let w = cx(c + 1) - cx(c);
        if (w <= 0) continue;
        const table = tableAt(sheet, r, c);
        const tableHeader = table?.headerRow && parseRange(table.range)!.range.start.row === r;
        if (tableHeader) w = Math.max(0, w - 18);
        // Left-aligned text overflows into empty neighbours.
        let spill = 0;
        const isText = typeof cell.value === 'string';
        const align = style?.hAlign ?? (typeof cell.value === 'number' ? 'right' : typeof cell.value === 'string' ? 'left' : 'center');
        if (isText && align === 'left' && !style?.wrap && !tableHeader) {
          let nc = c + 1;
          while (nc <= c1 + 20 && sheet.getValue(r, nc) === null && !covered.has(cellKey(r, nc)) && !styleAt(r, nc)?.fill) {
            spill += vp.geo.cols.size(nc);
            nc++;
          }
        }
        this.drawCellText(vp, state, cell.value, this.linkStyle(state, r, c, style), cx(c), y, w, h, spill, align);
      }
    }
    for (const m of merges) {
      const cell = sheet.getCell(m.start.row, m.start.col);
      const style = styleAt(m.start.row, m.start.col);
      if (!cell || cell.value === null) continue;
      const rc = vp.rect(m);
      const align = style?.hAlign ?? (typeof cell.value === 'number' ? 'right' : typeof cell.value === 'string' ? 'left' : 'center');
      this.drawCellText(vp, state, cell.value, this.linkStyle(state, m.start.row, m.start.col, style), rc.x, rc.y, rc.w, rc.h, 0, align);
    }

    // 5. Borders
    for (let r = r0; r <= r1; r++) {
      for (let c = c0; c <= c1; c++) {
        const b = styleAt(r, c)?.border;
        if (!b) continue;
        const m = merges.find((mm) => mm.start.row === r && mm.start.col === c);
        const rc = m ? vp.rect(m) : { x: cx(c), y: ry(r), w: cx(c + 1) - cx(c), h: ry(r + 1) - ry(r) };
        const sides: [keyof typeof b, number, number, number, number][] = [
          ['top', rc.x, rc.y, rc.x + rc.w, rc.y],
          ['bottom', rc.x, rc.y + rc.h, rc.x + rc.w, rc.y + rc.h],
          ['left', rc.x, rc.y, rc.x, rc.y + rc.h],
          ['right', rc.x + rc.w, rc.y, rc.x + rc.w, rc.y + rc.h],
        ];
        for (const [side, x1, y1, x2, y2] of sides) {
          const s = b[side];
          if (!s) continue;
          // Black borders would vanish on dark backgrounds; draw them in the text color instead.
          ctx.strokeStyle = !s.color || (t.dark && /^#0{6}$/i.test(s.color)) ? t.fg : s.color;
          ctx.lineWidth = s.style === 'thick' ? 3 : s.style === 'medium' || s.style === 'double' ? 2 : 1;
          ctx.setLineDash(s.style === 'dashed' ? [4, 2] : s.style === 'dotted' || s.style === 'hair' ? [1, 2] : []);
          const off = ctx.lineWidth % 2 ? 0.5 : 0;
          this.line(Math.round(x1) - off, Math.round(y1) - off, Math.round(x2) - off, Math.round(y2) - off);
        }
        ctx.setLineDash([]);
      }
    }

    // 6. Note indicators (red triangle in the top-right corner)
    if (sheet.notes.size) {
      ctx.fillStyle = '#C00000';
      for (const key of sheet.notes.keys()) {
        const row = Math.floor(key / MAX_COLS), col = key % MAX_COLS;
        if (row < r0 || row > r1 || col < c0 || col > c1) continue;
        const m = merges.find((mm) => mm.start.row === row && mm.start.col === col);
        const rc = m ? vp.rect(m) : { x: cx(col), y: ry(row), w: cx(col + 1) - cx(col), h: ry(row + 1) - ry(row) };
        if (rc.w <= 0 || rc.h <= 0) continue;
        const size = Math.min(6 * vp.geo.zoom, rc.w / 2, rc.h / 2);
        ctx.beginPath();
        ctx.moveTo(rc.x + rc.w - size - 1, rc.y + 1);
        ctx.lineTo(rc.x + rc.w - 1, rc.y + 1);
        ctx.lineTo(rc.x + rc.w - 1, rc.y + size + 1);
        ctx.fill();
      }
    }

    // 7. Overlays
    this.drawOverlays(vp, state);
    ctx.restore();
  }

  private linkStyle(state: RenderState, row: number, col: number, style: CellStyle | undefined): CellStyle | undefined {
    if (!state.isLink?.(row, col)) return style;
    return { ...style, underline: true, color: style?.color ?? (this.theme.dark ? '#4DA3FF' : '#0563C1') };
  }

  private drawCellText(vp: Viewport, state: RenderState, value: NonNullable<import('@sheets/core').Scalar>, style: CellStyle | undefined, x: number, y: number, w: number, h: number, spill: number, align: string): void {
    const ctx = this.ctx;
    const t = this.theme;
    const formatted = formatValueEx(value, style?.numFmt, state.locale);
    let text = formatted.text;
    if (!text) return;
    const zoom = vp.geo.zoom;
    ctx.font = cssFont(style, t, zoom);
    let color = style?.color ?? formatted.color;
    if (!color) color = style?.fill ? (luminance(style.fill) > 0.5 ? '#1f1f1f' : '#f5f5f5') : t.fg;
    if (typeof value === 'object') color = style?.color ?? (t.dark ? '#ff8a80' : '#c62828');
    ctx.fillStyle = color;
    const pad = 3 * zoom;
    const fontPx = (((style?.fontSize ?? DEFAULT_FONT_SIZE_PT) * 96) / 72) * zoom;
    const lineH = Math.round(fontPx * 1.2);
    const avail = w - pad * 2;

    let lines: string[];
    if (style?.wrap) {
      lines = [];
      for (const para of text.split('\n')) {
        let cur = '';
        for (const word of para.split(/(\s+)/)) {
          const next = cur + word;
          if (cur && ctx.measureText(next).width > avail) {
            lines.push(cur.trimEnd());
            cur = word.trimStart();
          } else cur = next;
        }
        lines.push(cur);
      }
    } else {
      lines = [text.replace(/\n/g, ' ')];
      if (typeof value === 'number' && ctx.measureText(text).width > avail) {
        text = this.fitGeneralNumber(value, style, state.locale, avail) ?? '#'.repeat(Math.max(1, Math.floor(avail / (ctx.measureText('#').width || 7))));
        lines = [text];
      }
    }

    const totalH = lines.length * lineH;
    const vAlign = style?.vAlign ?? 'bottom';
    let ty = vAlign === 'top' ? y + pad : vAlign === 'middle' ? y + (h - totalH) / 2 : y + h - totalH - pad + 1;
    ctx.textBaseline = 'top';
    ctx.save();
    ctx.beginPath();
    ctx.rect(x + 1, y, w + spill - 2, h);
    ctx.clip();
    if (spill) {
      // Hide grid lines below overflowing text
      const tw = ctx.measureText(lines[0]!).width;
      if (tw > avail) {
        ctx.fillStyle = t.bg;
        ctx.fillRect(x + w, y, Math.min(spill, tw - avail + pad) , h - 1);
        ctx.fillStyle = color;
      }
    }
    for (const ln of lines) {
      const tw = ctx.measureText(ln).width;
      const tx = align === 'right' ? x + w - pad - tw : align === 'center' ? x + (w - tw) / 2 : x + pad;
      ctx.fillText(ln, tx, ty + (lineH - fontPx) / 2);
      if (style?.underline || style?.strike) {
        ctx.strokeStyle = color;
        ctx.lineWidth = 1;
        const uy = style.underline ? ty + lineH - 2 : ty + lineH / 2;
        this.line(tx, Math.round(uy) + 0.5, tx + tw, Math.round(uy) + 0.5);
      }
      ty += lineH;
    }
    ctx.restore();
  }

  /** Like Excel's General format: drop decimals, then use scientific notation to fit the column. */
  private fitGeneralNumber(value: number, style: CellStyle | undefined, locale: FormatLocale, avail: number): string | null {
    if (style?.numFmt && !/^general$/i.test(style.numFmt)) return null;
    const fits = (t: string) => this.ctx.measureText(t).width <= avail;
    const intDigits = Math.max(1, Math.floor(Math.log10(Math.abs(value) || 1)) + 1);
    for (let d = 10; d >= intDigits; d--) {
      const t = formatGeneral(Number(value.toPrecision(d)), locale);
      if (fits(t)) return t;
    }
    for (let k = 4; k >= 0; k--) {
      const [m, e] = value.toExponential(k).split('e') as [string, string];
      const t = `${m.replace('.', locale.decimal)}E${Number(e) < 0 ? '-' : '+'}${String(Math.abs(Number(e))).padStart(2, '0')}`;
      if (fits(t)) return t;
    }
    return null;
  }

  private expandForMerges(vp: Viewport, r: RangeAddress): RangeAddress {
    const out = { start: { ...r.start }, end: { ...r.end } };
    for (const m of vp.sheet.merges) {
      if (m.end.row < out.start.row || m.start.row > out.end.row || m.end.col < out.start.col || m.start.col > out.end.col) continue;
      out.start.row = Math.min(out.start.row, m.start.row);
      out.start.col = Math.min(out.start.col, m.start.col);
      out.end.row = Math.max(out.end.row, m.end.row);
      out.end.col = Math.max(out.end.col, m.end.col);
    }
    return out;
  }

  private drawOverlays(vp: Viewport, state: RenderState): void {
    const ctx = this.ctx;
    const t = this.theme;
    // Selection
    const sel = this.expandForMerges(vp, state.selection);
    const rc = vp.rect(sel);
    const activeMerge = vp.sheet.mergeAt(state.active.row, state.active.col);
    const ac = vp.rect(activeMerge ?? { start: state.active, end: state.active });
    const single = sel.start.row === sel.end.row && sel.start.col === sel.end.col;
    if (!single && !(activeMerge && activeMerge.start.row === sel.start.row && activeMerge.end.row === sel.end.row && activeMerge.start.col === sel.start.col && activeMerge.end.col === sel.end.col)) {
      ctx.fillStyle = t.selection;
      ctx.save();
      ctx.beginPath();
      ctx.rect(rc.x, rc.y, rc.w, rc.h);
      ctx.rect(ac.x + ac.w, ac.y, -ac.w, ac.h); // cut out the active cell
      ctx.fill('evenodd');
      ctx.restore();
    }
    ctx.strokeStyle = t.accent;
    ctx.lineWidth = 2;
    ctx.strokeRect(Math.round(rc.x) - 0.5, Math.round(rc.y) - 0.5, Math.round(rc.w), Math.round(rc.h));
    if (!single) {
      ctx.lineWidth = 1;
      ctx.strokeRect(Math.round(ac.x) + 0.5, Math.round(ac.y) + 0.5, Math.round(ac.w) - 2, Math.round(ac.h) - 2);
    }
    if (state.showFillHandle) {
      ctx.fillStyle = t.accent;
      ctx.fillRect(rc.x + rc.w - 4, rc.y + rc.h - 4, 7, 7);
      ctx.strokeStyle = t.bg;
      ctx.lineWidth = 1;
      ctx.strokeRect(rc.x + rc.w - 4.5, rc.y + rc.h - 4.5, 8, 8);
    }
    // Formula reference highlights
    for (const hl of state.highlights) {
      const r = vp.rect(hl.range);
      ctx.strokeStyle = hl.color;
      ctx.lineWidth = 2;
      ctx.setLineDash(hl.dashed ? [5, 3] : []);
      ctx.strokeRect(Math.round(r.x) + 0.5, Math.round(r.y) + 0.5, Math.round(r.w) - 1, Math.round(r.h) - 1);
      ctx.globalAlpha = 0.08;
      ctx.fillStyle = hl.color;
      ctx.fillRect(r.x, r.y, r.w, r.h);
      ctx.globalAlpha = 1;
    }
    ctx.setLineDash([]);
    if (state.copyRange) {
      const r = vp.rect(state.copyRange);
      ctx.strokeStyle = t.accent;
      ctx.lineWidth = 2;
      ctx.setLineDash([6, 4]);
      ctx.strokeRect(Math.round(r.x) + 0.5, Math.round(r.y) + 0.5, Math.round(r.w) - 1, Math.round(r.h) - 1);
      ctx.setLineDash([]);
    }
    if (state.fillPreview) {
      const r = vp.rect(state.fillPreview);
      ctx.strokeStyle = t.muted;
      ctx.lineWidth = 1;
      ctx.setLineDash([4, 3]);
      ctx.strokeRect(Math.round(r.x) + 0.5, Math.round(r.y) + 0.5, Math.round(r.w) - 1, Math.round(r.h) - 1);
      ctx.setLineDash([]);
    }
  }

  private drawHeaders(vp: Viewport, state: RenderState, rows: [number, number], cols: [number, number]): void {
    const ctx = this.ctx;
    const t = this.theme;
    const sel = state.selection;
    ctx.font = `12px ${t.fontFamily}`;
    ctx.textBaseline = 'middle';
    ctx.textAlign = 'center';

    const colRanges: [number, number][] = [];
    if (vp.frozenCols) colRanges.push([0, vp.frozenCols - 1]);
    colRanges.push(cols);
    const rowRanges: [number, number][] = [];
    if (vp.frozenRows) rowRanges.push([0, vp.frozenRows - 1]);
    rowRanges.push(rows);

    // Column headers
    ctx.fillStyle = t.headerBg;
    ctx.fillRect(0, 0, vp.width, vp.headerH);
    ctx.save();
    ctx.beginPath();
    ctx.rect(vp.headerW, 0, vp.width - vp.headerW, vp.headerH);
    ctx.clip();
    for (const [i, [c0, c1]] of colRanges.entries()) {
      ctx.save();
      if (colRanges.length > 1) {
        ctx.beginPath();
        if (i === 0) ctx.rect(vp.headerW, 0, vp.frozenW, vp.headerH);
        else ctx.rect(vp.headerW + vp.frozenW, 0, vp.width, vp.headerH);
        ctx.clip();
      }
      for (let c = c0; c <= c1; c++) {
        const x = vp.x(c);
        const w = vp.geo.cols.size(c);
        if (w <= 0) {
          // Mark the position of hidden columns.
          ctx.fillStyle = t.muted;
          ctx.fillRect(Math.round(x) - 2, 3, 3, vp.headerH - 6);
          continue;
        }
        const active = c >= sel.start.col && c <= sel.end.col;
        if (active) {
          ctx.fillStyle = t.headerActiveBg;
          ctx.fillRect(x, 0, w, vp.headerH);
          ctx.fillStyle = t.accent;
          ctx.fillRect(x, vp.headerH - 2, w, 2);
        }
        ctx.fillStyle = active ? t.headerActiveFg : t.headerFg;
        ctx.fillText(colToName(c), x + w / 2, vp.headerH / 2 + 1);
        ctx.strokeStyle = t.grid;
        ctx.lineWidth = 1;
        this.line(Math.round(x + w) - 0.5, 0, Math.round(x + w) - 0.5, vp.headerH);
      }
      ctx.restore();
    }
    ctx.restore();

    // Row headers
    ctx.fillStyle = t.headerBg;
    ctx.fillRect(0, vp.headerH, vp.headerW, vp.height);
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, vp.headerH, vp.headerW, vp.height - vp.headerH);
    ctx.clip();
    for (const [i, [r0, r1]] of rowRanges.entries()) {
      ctx.save();
      if (rowRanges.length > 1) {
        ctx.beginPath();
        if (i === 0) ctx.rect(0, vp.headerH, vp.headerW, vp.frozenH);
        else ctx.rect(0, vp.headerH + vp.frozenH, vp.headerW, vp.height);
        ctx.clip();
      }
      for (let r = r0; r <= r1; r++) {
        const y = vp.y(r);
        const h = vp.geo.rows.size(r);
        if (h <= 0) {
          if (vp.sheet.rowHeights.get(r) === 0) {
            // Mark manually hidden rows (filtered rows are shown by the filter button).
            ctx.fillStyle = t.muted;
            ctx.fillRect(3, Math.round(y) - 2, vp.headerW - 6, 3);
          }
          continue;
        }
        const active = r >= sel.start.row && r <= sel.end.row;
        if (active) {
          ctx.fillStyle = t.headerActiveBg;
          ctx.fillRect(0, y, vp.headerW, h);
          ctx.fillStyle = t.accent;
          ctx.fillRect(vp.headerW - 2, y, 2, h);
        }
        ctx.fillStyle = active ? t.headerActiveFg : t.headerFg;
        ctx.fillText(String(r + 1), vp.headerW / 2, y + h / 2 + 1);
        ctx.strokeStyle = t.grid;
        this.line(0, Math.round(y + h) - 0.5, vp.headerW, Math.round(y + h) - 0.5);
      }
      ctx.restore();
    }
    ctx.restore();

    // Corner + header borders
    ctx.fillStyle = t.headerBg;
    ctx.fillRect(0, 0, vp.headerW, vp.headerH);
    ctx.strokeStyle = t.grid;
    ctx.lineWidth = 1;
    this.line(0, vp.headerH - 0.5, vp.width, vp.headerH - 0.5);
    this.line(vp.headerW - 0.5, 0, vp.headerW - 0.5, vp.height);
    ctx.fillStyle = t.muted;
    ctx.beginPath();
    ctx.moveTo(vp.headerW - 4, vp.headerH - 4);
    ctx.lineTo(vp.headerW - 14, vp.headerH - 4);
    ctx.lineTo(vp.headerW - 4, vp.headerH - 14);
    ctx.fill();
    ctx.textAlign = 'start';
  }

  /** Measures the width a cell's content needs (for auto-fit). */
  measure(text: string, style: CellStyle | undefined): number {
    this.ctx.font = cssFont(style, this.theme);
    return this.ctx.measureText(text).width;
  }
}
