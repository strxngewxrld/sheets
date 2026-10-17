/** Zero-based cell coordinates. */
export interface CellAddress {
  row: number;
  col: number;
}

/** Inclusive, normalized (start <= end) rectangular range. */
export interface RangeAddress {
  start: CellAddress;
  end: CellAddress;
}

export const MAX_ROWS = 1_048_576;
export const MAX_COLS = 16_384;

export function colToName(col: number): string {
  let name = '';
  let n = col + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

export function nameToCol(name: string): number {
  let col = 0;
  for (const ch of name.toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return col - 1;
}

export function formatAddress(row: number, col: number, absRow = false, absCol = false): string {
  return `${absCol ? '$' : ''}${colToName(col)}${absRow ? '$' : ''}${row + 1}`;
}

const CELL_RE = /^\$?([A-Za-z]{1,3})\$?(\d+)$/;

export function parseAddress(a1: string): CellAddress | null {
  const m = CELL_RE.exec(a1.trim());
  if (!m) return null;
  const col = nameToCol(m[1]!);
  const row = Number(m[2]) - 1;
  if (row < 0 || row >= MAX_ROWS || col >= MAX_COLS) return null;
  return { row, col };
}

export function normalizeRange(a: CellAddress, b: CellAddress): RangeAddress {
  return {
    start: { row: Math.min(a.row, b.row), col: Math.min(a.col, b.col) },
    end: { row: Math.max(a.row, b.row), col: Math.max(a.col, b.col) },
  };
}

export function quoteSheetName(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name)
    ? name
    : `'${name.replace(/'/g, "''")}'`;
}

export function splitSheetPrefix(ref: string): { sheet?: string; ref: string } {
  const bang = ref.lastIndexOf('!');
  if (bang < 0) return { ref };
  let sheet = ref.slice(0, bang);
  if (sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
  return { sheet, ref: ref.slice(bang + 1) };
}

/**
 * Parses "A1", "A1:B5", "A:C", "2:4", optionally sheet-qualified ("Sheet1!A1:B2", "'My Sheet'!A:A").
 */
export function parseRange(input: string): { sheet?: string; range: RangeAddress } | null {
  const { sheet, ref } = splitSheetPrefix(input.trim());
  const parts = ref.split(':');
  if (parts.length > 2) return null;
  const [a, b = a] = parts as [string, string?];
  const cellA = parseAddress(a);
  const cellB = parseAddress(b!);
  let range: RangeAddress | null = null;
  if (cellA && cellB) range = normalizeRange(cellA, cellB);
  else if (/^\$?[A-Za-z]{1,3}$/.test(a) && /^\$?[A-Za-z]{1,3}$/.test(b!)) {
    const c1 = nameToCol(a.replace('$', ''));
    const c2 = nameToCol(b!.replace('$', ''));
    range = normalizeRange({ row: 0, col: c1 }, { row: MAX_ROWS - 1, col: c2 });
  } else if (/^\$?\d+$/.test(a) && /^\$?\d+$/.test(b!)) {
    const r1 = Number(a.replace('$', '')) - 1;
    const r2 = Number(b!.replace('$', '')) - 1;
    if (r1 < 0 || r2 < 0) return null;
    range = normalizeRange({ row: r1, col: 0 }, { row: r2, col: MAX_COLS - 1 });
  }
  if (!range) return null;
  return sheet !== undefined ? { sheet, range } : { range };
}

export function formatRange(range: RangeAddress, sheet?: string): string {
  const prefix = sheet !== undefined ? `${quoteSheetName(sheet)}!` : '';
  const { start, end } = range;
  let body: string;
  if (start.row === 0 && end.row === MAX_ROWS - 1) body = `${colToName(start.col)}:${colToName(end.col)}`;
  else if (start.col === 0 && end.col === MAX_COLS - 1) body = `${start.row + 1}:${end.row + 1}`;
  else if (start.row === end.row && start.col === end.col) body = formatAddress(start.row, start.col);
  else body = `${formatAddress(start.row, start.col)}:${formatAddress(end.row, end.col)}`;
  return prefix + body;
}

/** Numeric key for a cell inside one sheet. */
export function cellKey(row: number, col: number): number {
  return row * MAX_COLS + col;
}

export function keyToAddress(key: number): CellAddress {
  return { row: Math.floor(key / MAX_COLS), col: key % MAX_COLS };
}

export function rangeContains(range: RangeAddress, row: number, col: number): boolean {
  return row >= range.start.row && row <= range.end.row && col >= range.start.col && col <= range.end.col;
}

export function rangesIntersect(a: RangeAddress, b: RangeAddress): boolean {
  return a.start.row <= b.end.row && b.start.row <= a.end.row && a.start.col <= b.end.col && b.start.col <= a.end.col;
}

export function intersectRanges(a: RangeAddress, b: RangeAddress): RangeAddress | null {
  if (!rangesIntersect(a, b)) return null;
  return {
    start: { row: Math.max(a.start.row, b.start.row), col: Math.max(a.start.col, b.start.col) },
    end: { row: Math.min(a.end.row, b.end.row), col: Math.min(a.end.col, b.end.col) },
  };
}

export function rangeRows(range: RangeAddress): number {
  return range.end.row - range.start.row + 1;
}

export function rangeCols(range: RangeAddress): number {
  return range.end.col - range.start.col + 1;
}

export function singleCellRange(row: number, col: number): RangeAddress {
  return { start: { row, col }, end: { row, col } };
}

/** Accepts either an A1 string or an address object. */
export type RangeLike = string | RangeAddress;
export type CellLike = string | CellAddress;

export function toRange(r: RangeLike): RangeAddress {
  if (typeof r !== 'string') return r;
  const parsed = parseRange(r);
  if (!parsed) throw new Error(`Invalid range: ${r}`);
  return parsed.range;
}

export function toCell(c: CellLike): CellAddress {
  if (typeof c !== 'string') return c;
  const parsed = parseAddress(c);
  if (!parsed) throw new Error(`Invalid cell address: ${c}`);
  return parsed;
}
