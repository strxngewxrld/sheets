import { cellKey, keyToAddress, MAX_COLS, MAX_ROWS, parseRange, rangesIntersect, type RangeAddress } from './address.js';
import type { Node } from './formula/ast.js';
import { deref, evaluate, evaluateCell, isRef, type EvalContext, type EvalValue, type Matrix } from './formula/evaluator.js';
import { parseFormula, walk } from './formula/parser.js';
import { getFunction } from './formula/registry.js';
import { boundNames } from './formula/scope.js';
import type { Sheet } from './sheet.js';
import { resolveStructuredRef, tableFilteredRows } from './table.js';
import { CellError, Err, type Scalar } from './values.js';
import type { Workbook } from './workbook.js';

const SHEET_FACTOR = 2 ** 35;

/** Global cell id: unique across sheets. */
export function gid(sheetId: number, row: number, col: number): number {
  return sheetId * SHEET_FACTOR + cellKey(row, col);
}

function splitGid(g: number): { sheetId: number; row: number; col: number } {
  const sheetId = Math.floor(g / SHEET_FACTOR);
  return { sheetId, ...keyToAddress(g - sheetId * SHEET_FACTOR) };
}

interface FormulaInfo {
  sheetId: number;
  row: number;
  col: number;
  ast: Node | null;
  parseError?: CellError;
  cellRefs: number[];
  rangeRefs: { sheetId: number; range: RangeAddress }[];
  volatile: boolean;
  unsupported: boolean;
}

/** Above this many changed cells a full recalculation is cheaper than propagation. */
const FULL_RECALC_THRESHOLD = 2000;

/**
 * Dependency-tracking calculation engine. Keeps a graph of formula
 * precedents and recalculates only the affected cells after a change.
 */
export class CalcEngine {
  private readonly wb: Workbook;
  private formulas = new Map<number, FormulaInfo>();
  private dependents = new Map<number, Set<number>>();
  private rangeDeps = new Map<number, Map<number, RangeAddress[]>>();
  private pending = new Set<number>();
  private computing = new Set<number>();
  private astCache = new Map<string, Node | CellError>();
  private spillOwners = new Map<number, number>();
  private previousSpillOwners = new Map<number, number>();
  private arrays = new Set<number>();

  constructor(wb: Workbook) {
    this.wb = wb;
  }

  parse(formula: string): Node | CellError {
    let ast = this.astCache.get(formula);
    if (!ast) {
      try {
        ast = parseFormula(formula);
      } catch {
        ast = new CellError('#NAME?', 'Invalid formula');
      }
      if (this.astCache.size > 20_000) this.astCache.clear();
      this.astCache.set(formula, ast);
    }
    return ast;
  }

  // -------------------------------------------------------------------------
  // Graph maintenance
  // -------------------------------------------------------------------------

  /** Rebuilds the full dependency graph from the workbook contents. */
  rebuild(): void {
    this.spillOwners.clear();
    this.previousSpillOwners.clear();
    this.arrays.clear();
    this.formulas.clear();
    this.dependents.clear();
    this.rangeDeps.clear();
    for (const sheet of this.wb.sheets) {
      for (const [key, cell] of sheet.cells) {
        const { row, col } = keyToAddress(key);
        if (cell.spillParent !== undefined) {
          const parent = keyToAddress(cell.spillParent);
          this.spillOwners.set(gid(sheet.id, row, col), gid(sheet.id, parent.row, parent.col));
        }
        if (cell.spillRange) this.arrays.add(gid(sheet.id, row, col));
        if (cell.formula === undefined) continue;
        this.register(sheet, row, col, cell.formula);
      }
    }
  }

  /** Updates the graph for one cell (formula = undefined removes it). */
  update(sheet: Sheet, row: number, col: number, formula: string | undefined): void {
    this.unregister(gid(sheet.id, row, col));
    if (formula !== undefined) this.register(sheet, row, col, formula);
  }

  private register(sheet: Sheet, row: number, col: number, formula: string): void {
    const id = gid(sheet.id, row, col);
    const parsed = this.parse(formula);
    const info: FormulaInfo = { sheetId: sheet.id, row, col, ast: null, cellRefs: [], rangeRefs: [], volatile: false, unsupported: false };
    if (parsed instanceof CellError) info.parseError = parsed;
    else {
      info.ast = parsed;
      const bound = boundNames(parsed);
      walk(parsed, (n) => {
        if (n.type === 'cell' || n.type === 'range') {
          const target = n.sheet !== undefined ? this.wb.getSheet(n.sheet) : sheet;
          if (!target) return;
          if (n.type === 'cell') info.cellRefs.push(gid(target.id, n.ref.row, n.ref.col));
          else {
            const range = {
              start: { row: Math.min(n.start.row, n.end.row), col: Math.min(n.start.col, n.end.col) },
              end: { row: Math.max(n.start.row, n.end.row), col: Math.max(n.start.col, n.end.col) },
            };
            info.rangeRefs.push({ sheetId: target.id, range });
          }
        } else if (n.type === 'structured') {
          // Table ranges change with table edits; re-evaluate on every recalculation.
          info.volatile = true;
        } else if (n.type === 'name') {
          if (!bound.has(n.name.toUpperCase())) info.volatile = true;
        } else if (n.type === 'func') {
          const def = getFunction(n.name);
          if (!def && bound.has(n.name.toUpperCase())) return;
          if (!def && this.wb.names.has(n.name.toUpperCase())) info.volatile = true;
          else if (!def) info.unsupported = true;
          else if (def.volatile) info.volatile = true;
        }
      });
    }
    this.formulas.set(id, info);
    for (const ref of info.cellRefs) {
      let set = this.dependents.get(ref);
      if (!set) this.dependents.set(ref, (set = new Set()));
      set.add(id);
    }
    for (const { sheetId, range } of info.rangeRefs) {
      let map = this.rangeDeps.get(sheetId);
      if (!map) this.rangeDeps.set(sheetId, (map = new Map()));
      const list = map.get(id);
      if (list) list.push(range);
      else map.set(id, [range]);
    }
  }

  private unregister(id: number): void {
    this.arrays.delete(id);
    const info = this.formulas.get(id);
    if (!info) return;
    this.formulas.delete(id);
    for (const ref of info.cellRefs) this.dependents.get(ref)?.delete(id);
    for (const { sheetId } of info.rangeRefs) this.rangeDeps.get(sheetId)?.delete(id);
  }

  // -------------------------------------------------------------------------
  // Recalculation
  // -------------------------------------------------------------------------

  recalcAll(): void {
    this.evaluate(new Set(this.formulas.keys()));
  }

  /** Recalculates everything that depends on the given changed cells. */
  recalc(changed: Iterable<number>): void {
    const changedList = [...changed];
    if (this.arrays.size || this.spillOwners.size) return this.recalcAll();
    if (changedList.length > FULL_RECALC_THRESHOLD) return this.recalcAll();
    const dirty = new Set<number>();
    const queue: number[] = [];
    const mark = (id: number) => {
      if (dirty.has(id)) return;
      dirty.add(id);
      queue.push(id);
    };
    for (const id of changedList) {
      if (this.formulas.has(id)) dirty.add(id);
      queue.push(id);
    }
    for (const [id, info] of this.formulas) if (info.volatile) mark(id);
    while (queue.length) {
      const id = queue.pop()!;
      const deps = this.dependents.get(id);
      if (deps) for (const d of deps) mark(d);
      const { sheetId, row, col } = splitGid(id);
      const rdeps = this.rangeDeps.get(sheetId);
      if (rdeps) {
        for (const [fid, ranges] of rdeps) {
          if (dirty.has(fid)) continue;
          if (ranges.some((r) => row >= r.start.row && row <= r.end.row && col >= r.start.col && col <= r.end.col)) mark(fid);
        }
      }
    }
    this.evaluate(dirty);
  }

  private evaluate(dirty: Set<number>): void {
    // Discover spill ownership before trusting dependencies on former blank cells.
    // Further passes are needed only when output rectangles change, not values.
    for (let pass = 0; pass < Math.min(this.formulas.size + 2, 100); pass++) {
      this.previousSpillOwners = this.spillOwners;
      this.spillOwners = new Map();
      this.clearSpills();
      this.pending = new Set(dirty);
      this.computing.clear();
      for (const id of [...dirty].sort((a, b) => a - b)) this.compute(id);
      this.pending.clear();
      const stable = this.spillOwners.size === this.previousSpillOwners.size &&
        [...this.spillOwners].every(([key, owner]) => this.previousSpillOwners.get(key) === owner);
      if (stable) {
        this.previousSpillOwners.clear();
        return;
      }
      dirty = new Set(this.formulas.keys());
    }
    this.clearSpills();
    for (const id of this.arrays) {
      const { sheetId, row, col } = splitGid(id);
      const cell = this.wb.getSheetById(sheetId)?.getCell(row, col);
      if (cell) cell.value = new CellError('#CALC!', 'Array dimensions did not stabilize');
    }
    this.spillOwners.clear();
    this.previousSpillOwners.clear();
  }

  /** Discards derived output, keeping user formatting and formula anchors. */
  clearSpills(): void {
    for (const sheet of this.wb.sheets) {
      for (const [key, cell] of sheet.cells) {
        delete cell.spillRange;
        if (cell.spillParent === undefined) continue;
        if (cell.style) sheet.cells.set(key, { value: null, style: cell.style });
        else sheet.cells.delete(key);
      }
      sheet.invalidate();
    }
  }

  private spill(sheet: Sheet, row: number, col: number, matrix: Matrix): Scalar {
    const rows = matrix.length;
    const cols = matrix[0]?.length ?? 0;
    if (!rows || !cols || matrix.some((r) => r.length !== cols)) return new CellError('#CALC!', 'Invalid array shape');
    const range = { start: { row, col }, end: { row: row + rows - 1, col: col + cols - 1 } };
    if (range.end.row >= MAX_ROWS || range.end.col >= MAX_COLS) return new CellError('#SPILL!', 'Array extends beyond the sheet');
    if (sheet.tables.some((t) => rangesIntersect(range, parseRange(t.range)!.range))) return new CellError('#SPILL!', 'Arrays cannot spill into tables');
    for (let r = row; r <= range.end.row; r++) {
      for (let c = col; c <= range.end.col; c++) {
        const existing = sheet.getCell(r, c);
        if (sheet.mergeAt(r, c) || ((r !== row || c !== col) && existing && (existing.formula !== undefined || existing.value !== null || existing.spillParent !== undefined))) {
          return new CellError('#SPILL!', 'Array output is blocked by a cell or merge');
        }
      }
    }
    sheet.getCell(row, col)!.spillRange = range;
    const anchor = gid(sheet.id, row, col);
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        if (r === 0 && c === 0) continue;
        const key = cellKey(row + r, col + c);
        const style = sheet.cells.get(key)?.style;
        sheet.cells.set(key, { value: matrix[r]![c] ?? 0, style, spillParent: cellKey(row, col) });
        this.spillOwners.set(gid(sheet.id, row + r, col + c), anchor);
      }
    }
    sheet.invalidate();
    return matrix[0]![0] ?? 0;
  }

  /** Computes a pending formula after its pending single-cell precedents (iterative DFS). */
  private compute(root: number): void {
    if (!this.pending.has(root) || this.computing.has(root)) return;
    const stack = [root];
    const onStack = new Set(stack);
    while (stack.length) {
      const top = stack[stack.length - 1]!;
      if (!this.pending.has(top)) {
        stack.pop();
        onStack.delete(top);
        continue;
      }
      const info = this.formulas.get(top);
      const next = info?.cellRefs.map((p) => this.spillOwners.get(p) ?? this.previousSpillOwners.get(p) ?? p)
        .find((p) => this.pending.has(p) && !onStack.has(p) && !this.computing.has(p));
      if (next !== undefined) {
        stack.push(next);
        onStack.add(next);
        continue;
      }
      stack.pop();
      onStack.delete(top);
      this.evaluateOne(top);
    }
  }

  private evaluateOne(id: number): void {
    const info = this.formulas.get(id);
    this.pending.delete(id);
    if (!info) return;
    const sheet = this.wb.getSheetById(info.sheetId);
    const cell = sheet?.getCell(info.row, info.col);
    if (!sheet || !cell) return;
    if (info.parseError) {
      cell.value = info.parseError;
      return;
    }
    if (info.unsupported && cell.cached) return; // keep the value Excel computed
    this.computing.add(id);
    try {
      const ctx = this.contextFor(sheet, info.row, info.col);
      const result = evaluate(info.ast!, ctx);
      const output = isRef(result) ? deref(result, ctx) : result;
      if (Array.isArray(output)) {
        this.arrays.add(id);
        cell.value = this.spill(sheet, info.row, info.col, output);
      } else {
        this.arrays.delete(id);
        cell.value = output ?? 0;
      }
    } catch (e) {
      cell.value = e instanceof RangeError ? new CellError('#CALC!', 'Recursion too deep') : new CellError('#VALUE!', String(e));
    } finally {
      this.computing.delete(id);
    }
    cell.cached = false;
  }

  /** Value of a cell, computing it first if it is pending. */
  private valueOf(sheet: Sheet, row: number, col: number): Scalar {
    const id = gid(sheet.id, row, col);
    const owner = this.spillOwners.get(id) ?? this.previousSpillOwners.get(id);
    if (owner !== undefined) {
      if (this.computing.has(owner)) return Err.CIRC;
      if (this.pending.has(owner)) this.compute(owner);
    }
    if (this.computing.has(id)) return Err.CIRC;
    if (this.pending.has(id)) this.compute(id);
    if (this.computing.has(id)) return Err.CIRC;
    return sheet.getValue(row, col);
  }

  contextFor(sheet: Sheet, row: number, col: number): EvalContext & { hasFormula(s: string, r: number, c: number): boolean } {
    const wb = this.wb;
    const filtered = new Map<Sheet, Set<number>>();
    return {
      resolveStructured: (ref, name, r, c) => resolveStructuredRef(wb, ref, name, r, c),
      rowHidden: (name, r) => {
        const s = wb.getSheet(name);
        if (!s) return false;
        if (s.rowHeights.get(r) === 0) return 'manual';
        let rows = filtered.get(s);
        if (!rows) filtered.set(s, (rows = tableFilteredRows(s)));
        return rows.has(r) ? 'filtered' : false;
      },
      sheet: sheet.name,
      row,
      col,
      getCell: (name, r, c) => {
        const s = wb.getSheet(name);
        return s ? this.valueOf(s, r, c) : Err.REF;
      },
      getRange: (name, range) => {
        const s = wb.getSheet(name);
        if (!s) return [[Err.REF]];
        return this.rangeValues(s, range);
      },
      sheetExists: (name) => !!wb.getSheet(name),
      sheetNames: () => wb.sheets.map((s) => s.name),
      resolveName: (name) => {
        const ref = wb.names.get(name.toUpperCase());
        if (!ref) return undefined;
        const ast = this.parse(ref.ref);
        return ast instanceof CellError ? undefined : ast;
      },
      parseRef: (text) => {
        const p = parseRange(text);
        if (!p) return undefined;
        const ast = this.parse(text);
        return ast instanceof CellError ? undefined : ast;
      },
      hasFormula: (name, r, c) => wb.getSheet(name)?.getCell(r, c)?.formula !== undefined,
      getSpillRange: (name, r, c) => {
        const target = wb.getSheet(name);
        if (!target) return Err.REF;
        this.valueOf(target, r, c);
        if (this.computing.has(gid(target.id, r, c))) return Err.CIRC;
        return target.getCell(r, c)?.spillRange ?? Err.REF;
      },
    };
  }

  private rangeValues(sheet: Sheet, range: RangeAddress): Matrix {
    const used = sheet.usedRange();
    const { start } = range;
    // Whole-column/row references are clipped to the used area; only the far
    // edge is clipped so relative positions (INDEX/MATCH) stay correct.
    const endRow = range.end.row === MAX_ROWS - 1 ? Math.max(start.row, Math.min(range.end.row, used?.end.row ?? 0)) : range.end.row;
    const endCol = range.end.col === MAX_COLS - 1 ? Math.max(start.col, Math.min(range.end.col, used?.end.col ?? 0)) : range.end.col;
    const out: Matrix = [];
    for (let r = start.row; r <= endRow; r++) {
      const row: Scalar[] = [];
      for (let c = start.col; c <= endCol; c++) row.push(this.valueOf(sheet, r, c));
      out.push(row);
    }
    return out;
  }

  /**
   * Evaluates a formula without reducing it to one cell value (ranges stay
   * references). `assume` evaluates as if a cell of this sheet held a value.
   */
  evaluateRaw(formula: string, sheet: Sheet, row = 0, col = 0, assume?: { row: number; col: number; value: Scalar }): EvalValue {
    const ast = this.parse(formula.startsWith('=') ? formula.slice(1) : formula);
    if (ast instanceof CellError) return ast;
    const ctx = this.contextFor(sheet, row, col);
    if (assume) {
      const same = (name: string) => name.toUpperCase() === sheet.name.toUpperCase();
      const { getCell, getRange } = ctx;
      ctx.getCell = (name, r, c) => (same(name) && r === assume.row && c === assume.col ? assume.value : getCell(name, r, c));
      ctx.getRange = (name, range) => {
        const m = getRange(name, range);
        const { row: r, col: c } = assume;
        if (same(name) && r >= range.start.row && r <= range.end.row && c >= range.start.col && c <= range.end.col) {
          const line = m[r - range.start.row];
          if (line) line[c - range.start.col] = assume.value;
        }
        return m;
      };
    }
    try {
      return evaluate(ast, ctx);
    } catch (e) {
      return e instanceof RangeError ? new CellError('#CALC!', 'Recursion too deep') : new CellError('#VALUE!', String(e));
    }
  }

  /** Evaluates an ad-hoc formula in the context of a sheet (not stored). */
  evaluateFormula(formula: string, sheet: Sheet, row = 0, col = 0): Scalar {
    const ast = this.parse(formula.startsWith('=') ? formula.slice(1) : formula);
    if (ast instanceof CellError) return ast;
    try {
      return evaluateCell(ast, this.contextFor(sheet, row, col));
    } catch (e) {
      // Same mapping as for stored cells, so a pathological formula yields an error value instead of throwing.
      return e instanceof RangeError ? new CellError('#CALC!', 'Recursion too deep') : new CellError('#VALUE!', String(e));
    }
  }

  /** Cells (as sheet/row/col) that directly depend on the given cell. */
  dependentsOf(sheet: Sheet, row: number, col: number): { sheetId: number; row: number; col: number }[] {
    const id = gid(sheet.id, row, col);
    const out = new Set<number>(this.dependents.get(id) ?? []);
    for (const [fid, ranges] of this.rangeDeps.get(sheet.id) ?? []) {
      if (ranges.some((r) => row >= r.start.row && row <= r.end.row && col >= r.start.col && col <= r.end.col)) out.add(fid);
    }
    return [...out].map(splitGid);
  }
}
