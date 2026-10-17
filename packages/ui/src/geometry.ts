import { MAX_COLS, MAX_ROWS, tableFilteredRows, type Sheet } from '@sheets/core';

/** One axis (rows or columns) with variable sizes; O(log n) offset lookups. */
export class Axis {
  private custom: number[] = [];
  private prefix: number[] = [0];
  readonly def: number;
  readonly max: number;
  private sizes: Map<number, number>;

  constructor(sizes: Map<number, number>, def: number, max: number) {
    this.sizes = sizes;
    this.def = def;
    this.max = max;
    this.custom = [...sizes.keys()].sort((a, b) => a - b);
    for (const i of this.custom) this.prefix.push(this.prefix.at(-1)! + (sizes.get(i)! - def));
  }

  size(i: number): number {
    return this.sizes.get(i) ?? this.def;
  }

  /** Offset of the start of item i. */
  offset(i: number): number {
    let lo = 0;
    let hi = this.custom.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (this.custom[mid]! < i) lo = mid + 1;
      else hi = mid;
    }
    return i * this.def + this.prefix[lo]!;
  }

  /** Index of the item containing position p. */
  indexAt(p: number): number {
    if (p <= 0) return 0;
    let lo = 0;
    let hi = this.max - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (this.offset(mid) <= p) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  }
}

/** Screen geometry of a sheet; `zoom` scales all row heights and column widths. */
export class Geometry {
  readonly cols: Axis;
  readonly rows: Axis;
  readonly zoom: number;
  constructor(sheet: Sheet, zoom = 1) {
    this.zoom = zoom;
    const scale = (sizes: Map<number, number>) => (zoom === 1 ? sizes : new Map([...sizes].map(([i, v]) => [i, v * zoom])));
    this.cols = new Axis(scale(sheet.colWidths), sheet.defaultColWidth * zoom, MAX_COLS);
    const sizes = new Map(sheet.rowHeights);
    for (const row of tableFilteredRows(sheet)) sizes.set(row, 0);
    this.rows = new Axis(scale(sizes), sheet.defaultRowHeight * zoom, MAX_ROWS);
  }
}
