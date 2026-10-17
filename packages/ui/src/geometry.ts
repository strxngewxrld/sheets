import { MAX_COLS, MAX_ROWS, type Sheet } from '@sheets/core';

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

export class Geometry {
  readonly cols: Axis;
  readonly rows: Axis;
  constructor(sheet: Sheet) {
    this.cols = new Axis(sheet.colWidths, sheet.defaultColWidth, MAX_COLS);
    this.rows = new Axis(sheet.rowHeights, sheet.defaultRowHeight, MAX_ROWS);
  }
}
