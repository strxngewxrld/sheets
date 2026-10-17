import { formatRange, keyToAddress, parseRange, rangesIntersect, type RangeAddress } from './address.js';
import type { Sheet } from './sheet.js';
import type { CellStyle } from './style.js';
import type { ApplyOptions, Workbook } from './workbook.js';
import type { CellPatch } from './ops.js';
import type { StructuredRef } from './formula/ast.js';
import type { RefValue } from './formula/evaluator.js';
import { Err, type CellError } from './values.js';

export interface TableTheme {
  body: CellStyle;
  header: CellStyle;
  stripe: CellStyle;
}

export interface SheetTable {
  name: string;
  range: string;
  headerRow: boolean;
  bandedRows: boolean;
  theme: TableTheme;
  /** Value filters, with zero-based columns relative to the table. Empty values hide all rows. */
  filters?: { column: number; values: string[] }[];
}

export function tableFilterValue(value: unknown): string {
  return value === null || value === undefined ? '' : String(value);
}

/** Derived visibility: filtering never overwrites manually assigned row heights. */
export function tableFilteredRows(sheet: Sheet): Set<number> {
  const hidden = new Set<number>();
  for (const table of sheet.tables) {
    if (!table.headerRow || !table.filters?.length) continue;
    const range = parseRange(table.range)!.range;
    const filters = table.filters.map((f) => ({ col: range.start.col + f.column, values: new Set(f.values) }));
    for (let row = range.start.row + 1; row <= range.end.row; row++) {
      if (filters.some((f) => !f.values.has(tableFilterValue(sheet.getValue(row, f.col))))) hidden.add(row);
    }
  }
  return hidden;
}

export function setTableFilter(wb: Workbook, sheet: string | number, name: string, column: number, values: string[] | null): void {
  const s = wb.requireSheet(sheet);
  if (!s.tables.some((t) => t.name === name)) throw new Error('Unknown table');
  const tables = s.tables.map((t) => {
    if (t.name !== name) return t;
    const filters = (t.filters ?? []).filter((f) => f.column !== column);
    if (values !== null) filters.push({ column, values: [...new Set(values)] });
    const result = { ...t };
    if (filters.length) result.filters = filters;
    else delete result.filters;
    return result;
  });
  wb.apply({ type: 'setTables', sheet: s.name, tables });
}

const theme = (header: string, stripe: string): TableTheme => ({
  body: { fill: '#FFFFFF', color: '#1F1F1F' },
  header: { fill: header, color: '#FFFFFF', bold: true },
  stripe: { fill: stripe },
});

export const TABLE_THEMES: Record<string, TableTheme> = {
  blue: theme('#4472C4', '#E9EFF7'),
  green: theme('#548235', '#EAF1E3'),
  orange: theme('#C65911', '#FCEBDD'),
  purple: theme('#7030A0', '#F0E8F5'),
  gray: theme('#595959', '#EEEEEE'),
};

export function tableAt(sheet: Sheet, row: number, col: number): SheetTable | undefined {
  return sheet.tables.find((table) => {
    const r = parseRange(table.range)!.range;
    return row >= r.start.row && row <= r.end.row && col >= r.start.col && col <= r.end.col;
  });
}

export function tableCellStyle(sheet: Sheet, row: number, col: number): CellStyle | undefined {
  const table = tableAt(sheet, row, col);
  if (!table) return undefined;
  const range = parseRange(table.range)!.range;
  if (table.headerRow && row === range.start.row) return { ...table.theme.body, ...table.theme.header };
  const offset = row - range.start.row - (table.headerRow ? 1 : 0);
  return { ...table.theme.body, ...(table.bandedRows && offset % 2 === 0 ? table.theme.stripe : {}) };
}

export function validateTables(wb: Workbook, sheet: Sheet, tables: SheetTable[]): void {
  const names = new Set(wb.sheets.filter((s) => s !== sheet).flatMap((s) => s.tables.map((t) => t.name.toUpperCase())));
  const ranges: RangeAddress[] = [];
  for (const table of tables) {
    const parsed = parseRange(table.range);
    if (!/^[A-Za-z_][\w.]*$/.test(table.name) || parseRange(table.name) || names.has(table.name.toUpperCase())) throw new Error('Table names must be valid and unique across the workbook');
    names.add(table.name.toUpperCase());
    if (!parsed || parsed.sheet !== undefined) throw new Error('Table range must be an unqualified A1 range');
    const columns = new Set<number>();
    for (const f of table.filters ?? []) {
      if (!table.headerRow || !Number.isInteger(f.column) || f.column < 0 || f.column > parsed.range.end.col - parsed.range.start.col || columns.has(f.column) || !Array.isArray(f.values) || f.values.some((v) => typeof v !== 'string')) throw new Error('Invalid table filter');
      columns.add(f.column);
    }
    if (ranges.some((r) => rangesIntersect(r, parsed.range))) throw new Error('Tables cannot overlap');
    if (sheet.merges.some((r) => rangesIntersect(r, parsed.range))) throw new Error('Tables cannot contain merged cells');
    for (const [key, cell] of sheet.cells) {
      const p = keyToAddress(key);
      if (p.row >= parsed.range.start.row && p.row <= parsed.range.end.row && p.col >= parsed.range.start.col && p.col <= parsed.range.end.col && (cell.spillParent !== undefined || cell.spillRange)) throw new Error('Tables cannot contain spilled arrays');
    }
    if (!table.theme?.body || !table.theme.header || !table.theme.stripe) throw new Error('Table requires body, header and stripe styles');
    ranges.push(parsed.range);
  }
}

/** Applies a table style without baking stripes into the underlying cell styles. */
export function formatTable(
  wb: Workbook, sheet: string | number, range: RangeAddress,
  options: ApplyOptions & { name?: string; headerRow?: boolean; bandedRows?: boolean; theme?: TableTheme } = {},
): SheetTable {
  const s = wb.requireSheet(sheet);
  const ref = formatRange(range);
  const existing = s.tables.find((t) => t.name === options.name || t.range === ref);
  let name = options.name ?? existing?.name;
  if (!name) {
    const used = new Set(wb.sheets.flatMap((s) => s.tables.map((t) => t.name.toUpperCase())));
    let index = 1;
    while (used.has(`TABLE${index}`)) index++;
    name = `Table${index}`;
  }
  const table: SheetTable = {
    name, range: ref,
    headerRow: options.headerRow ?? existing?.headerRow ?? true,
    bandedRows: options.bandedRows ?? existing?.bandedRows ?? true,
    theme: structuredClone(options.theme ?? existing?.theme ?? TABLE_THEMES.blue!),
    ...(existing?.filters && (options.headerRow ?? existing.headerRow) ? { filters: structuredClone(existing.filters).filter((f) => f.column <= range.end.col - range.start.col) } : {}),
  };
  const tables = [...s.tables.filter((t) => t !== existing), table];
  validateTables(wb, s, tables);
  const headers: CellPatch[] = [];
  if (table.headerRow) tableColumnNames(s, table).forEach((name, i) => {
    const col = range.start.col + i;
    const cell = s.getCell(range.start.row, col);
    if (cell?.value !== name || cell.formula !== undefined) headers.push({ row: range.start.row, col, data: { v: name, ...(cell?.style ? { s: cell.style } : {}) } });
  });
  wb.apply([
    ...(headers.length ? [{ type: 'setCells' as const, sheet: s.name, cells: headers }] : []),
    { type: 'setTables', sheet: s.name, tables },
  ], options);
  return table;
}

export function removeTable(wb: Workbook, sheet: string | number, name: string, options: ApplyOptions = {}): void {
  const s = wb.requireSheet(sheet);
  wb.apply({ type: 'setTables', sheet: s.name, tables: s.tables.filter((t) => t.name !== name) }, options);
}

/** Excel table columns require unique, nonempty labels, even for headerless tables. */
export function tableColumnNames(sheet: Sheet, table: SheetTable): string[] {
  const range = parseRange(table.range)!.range;
  const used = new Set<string>();
  const result: string[] = [];
  for (let col = range.start.col; col <= range.end.col; col++) {
    const value = table.headerRow ? sheet.getValue(range.start.row, col) : null;
    const base = value === null || value === '' ? `Column${col - range.start.col + 1}` : String(value);
    let name = base, suffix = 2;
    while (used.has(name.toUpperCase())) name = `${base}${suffix++}`;
    used.add(name.toUpperCase()); result.push(name);
  }
  return result;
}

/** Finds a table by name on any sheet (case-insensitive). */
export function findTable(wb: Workbook, name: string): { sheet: Sheet; table: SheetTable } | undefined {
  const upper = name.toUpperCase();
  for (const sheet of wb.sheets) {
    const table = sheet.tables.find((t) => t.name.toUpperCase() === upper);
    if (table) return { sheet, table };
  }
  return undefined;
}

/** Name of the table containing a cell, given a sheet's table list. */
export function tableNameAt(tables: readonly SheetTable[], row: number, col: number): string | undefined {
  return tables.find((t) => {
    const r = parseRange(t.range)?.range;
    return r && row >= r.start.row && row <= r.end.row && col >= r.start.col && col <= r.end.col;
  })?.name;
}

/**
 * Resolves a structured reference to a cell range. Unqualified references
 * ([@Col]) use the table that contains the formula cell.
 */
export function resolveStructuredRef(wb: Workbook, ref: StructuredRef, sheetName: string, row: number, col: number): RefValue | CellError {
  let found: { sheet: Sheet; table: SheetTable } | undefined;
  if (ref.table) found = findTable(wb, ref.table);
  else {
    const sheet = wb.getSheet(sheetName);
    const table = sheet && tableAt(sheet, row, col);
    if (sheet && table) found = { sheet, table };
  }
  if (!found) return Err.REF;
  const { sheet, table } = found;
  const range = parseRange(table.range)!.range;
  let c1 = range.start.col;
  let c2 = range.end.col;
  if (ref.columns) {
    const names = tableColumnNames(sheet, table).map((n) => n.toUpperCase());
    const a = names.indexOf(ref.columns[0].toUpperCase());
    const b = names.indexOf(ref.columns[1].toUpperCase());
    if (a < 0 || b < 0) return Err.REF;
    c1 = range.start.col + Math.min(a, b);
    c2 = range.start.col + Math.max(a, b);
  }
  const dataStart = table.headerRow ? range.start.row + 1 : range.start.row;
  const rows: [number, number][] = [];
  for (const area of ref.areas.length ? ref.areas : ['#Data' as const]) {
    switch (area) {
      case '#All': rows.push([range.start.row, range.end.row]); break;
      case '#Data': if (dataStart <= range.end.row) rows.push([dataStart, range.end.row]); break;
      case '#Headers': if (!table.headerRow) return Err.REF; rows.push([range.start.row, range.start.row]); break;
      case '#Totals': return Err.REF;
      case '#This Row':
        if (sheet.name.toUpperCase() !== sheetName.toUpperCase() || row < dataStart || row > range.end.row) return Err.VALUE;
        rows.push([row, row]);
        break;
    }
  }
  if (!rows.length) return Err.REF;
  return {
    kind: 'ref',
    sheet: sheet.name,
    range: { start: { row: Math.min(...rows.map((r) => r[0])), col: c1 }, end: { row: Math.max(...rows.map((r) => r[1])), col: c2 } },
  };
}
