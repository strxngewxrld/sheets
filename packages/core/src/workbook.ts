import { cellKey, formatRange, keyToAddress, parseRange, MAX_COLS, MAX_ROWS, type RangeAddress } from './address.js';
import { CalcEngine, gid } from './calc.js';
import type { ChartSpec } from './chart.js';
import { ConditionalFormatEvaluator, validateConditionalFormats } from './conditional.js';
import { Emitter } from './events.js';
import { registerBuiltinFunctions } from './formula/functions/index.js';
import {
  adjustFormulaForStructure,
  invalidateSheetRefs,
  renameSheetInFormula,
  shiftFormula,
  type StructuralChange,
} from './formula/rewrite.js';
import type { Operation, WorkbookData } from './ops.js';
import { Sheet, type Cell, type SheetData } from './sheet.js';
import type { Scalar } from './values.js';
import type { CellStyle } from './style.js';
import { adjustPivotForStructure, type PivotTableSpec } from './pivot.js';
import { validateTables } from './table.js';
import { validateValidations, validationRanges, type DataValidation } from './validation.js';

export interface ChangeEvent {
  ops: Operation[];
  /** Who caused the change: "user", "undo", "redo", "agent", "remote", ... */
  origin: string;
  /** True if sheets, rows or columns were added/removed (UI should fully refresh). */
  structural: boolean;
}

export type WorkbookEvents = {
  change: ChangeEvent;
};

export interface ApplyOptions {
  origin?: string;
  /** Record in undo history (default true). */
  undoable?: boolean;
  /** Human readable label for the history entry. */
  label?: string;
}

interface HistoryEntry {
  ops: Operation[];
  inverse: Operation[];
  label?: string;
}

interface Transaction {
  ops: Operation[];
  inverse: Operation[];
  changed: Set<number>;
  structural: boolean;
  options: ApplyOptions;
}

export class Workbook {
  sheets: Sheet[] = [];
  readonly names = new Map<string, { name: string; ref: string }>();
  activeSheetIndex = 0;
  readonly events = new Emitter<WorkbookEvents>();
  readonly calc: CalcEngine;
  historyLimit = 500;
  pivotTables: PivotTableSpec[] = [];
  xlsxSource?: WorkbookData['xlsxSource'];

  private undoStack: HistoryEntry[] = [];
  private redoStack: HistoryEntry[] = [];
  private tx: Transaction | null = null;
  private txDepth = 0;

  constructor(options: { sheetName?: string; empty?: boolean } = {}) {
    registerBuiltinFunctions();
    this.calc = new CalcEngine(this);
    if (!options.empty) this.sheets.push(new Sheet(options.sheetName ?? 'Sheet1'));
  }

  static fromJSON(data: WorkbookData): Workbook {
    const wb = new Workbook({ empty: true });
    wb.load(data);
    return wb;
  }

  toJSON(): WorkbookData {
    const data: WorkbookData = { sheets: this.sheets.map((s) => s.toJSON()) };
    if (this.names.size) data.names = Object.fromEntries([...this.names.values()].map((n) => [n.name, n.ref]));
    if (this.activeSheetIndex) data.activeSheet = this.activeSheetIndex;
    if (this.pivotTables.length) data.pivotTables = structuredClone(this.pivotTables);
    if (this.xlsxSource) data.xlsxSource = structuredClone(this.xlsxSource);
    return data;
  }

  /** Replaces the whole content without history (used by importers). */
  load(data: WorkbookData, options: { recalculate?: boolean } = {}): void {
    this.xlsxSource = structuredClone(data.xlsxSource);
    this.pivotTables = structuredClone(data.pivotTables ?? []);
    this.sheets = data.sheets.map((s) => Sheet.fromJSON(s));
    if (!this.sheets.length) this.sheets.push(new Sheet('Sheet1'));
    this.names.clear();
    for (const [name, ref] of Object.entries(data.names ?? {})) this.names.set(name.toUpperCase(), { name, ref });
    this.activeSheetIndex = Math.min(data.activeSheet ?? 0, this.sheets.length - 1);
    this.calc.rebuild();
    if (options.recalculate !== false) this.calc.recalcAll();
    if (this.xlsxSource && !this.xlsxSource.baseline) {
      const baseline = this.toJSON();
      delete baseline.xlsxSource;
      this.xlsxSource.baseline = baseline;
    }
  }

  // -------------------------------------------------------------------------
  // Lookup
  // -------------------------------------------------------------------------

  getSheet(nameOrIndex: string | number): Sheet | undefined {
    if (typeof nameOrIndex === 'number') return this.sheets[nameOrIndex];
    const upper = nameOrIndex.toUpperCase();
    return this.sheets.find((s) => s.name.toUpperCase() === upper);
  }

  requireSheet(nameOrIndex: string | number): Sheet {
    const s = this.getSheet(nameOrIndex);
    if (!s) throw new Error(`Sheet not found: ${nameOrIndex}`);
    return s;
  }

  getSheetById(id: number): Sheet | undefined {
    return this.sheets.find((s) => s.id === id);
  }

  get activeSheet(): Sheet {
    return this.sheets[this.activeSheetIndex] ?? this.sheets[0]!;
  }

  uniqueSheetName(base = 'Sheet'): string {
    for (let i = this.sheets.length + 1; ; i++) {
      const name = `${base}${i}`;
      if (!this.getSheet(name)) return name;
    }
  }

  /** Evaluates a formula without storing it. */
  evaluate(formula: string, sheet: string | number = this.activeSheetIndex): Scalar {
    return this.calc.evaluateFormula(formula, this.requireSheet(sheet));
  }

  // -------------------------------------------------------------------------
  // Mutation
  // -------------------------------------------------------------------------

  /** Groups several apply() calls into one undo step and one change event. */
  transaction<T>(fn: () => T, options: ApplyOptions = {}): T {
    this.begin(options);
    try {
      return fn();
    } finally {
      this.commit();
    }
  }

  apply(ops: Operation | Operation[], options: ApplyOptions = {}): void {
    const list = Array.isArray(ops) ? ops : [ops];
    if (!list.length) return;
    this.begin(options);
    const tx = this.tx!;
    const wasStructural = tx.structural;
    tx.structural = false;
    tx.changed.clear();
    try {
      for (const op of list) {
        const inverse = this.applyOp(op, tx);
        if (this.xlsxSource && (['insertRows', 'deleteRows', 'insertCols', 'deleteCols', 'addSheet', 'removeSheet', 'renameSheet', 'moveSheet'].includes(op.type) || (op.type === 'setCells' && op.reordersCells))) {
          this.xlsxSource.structuralEdits = Math.max(0, (this.xlsxSource.structuralEdits ?? 0) + (options.origin === 'undo' ? -1 : 1));
        }
        tx.ops.push(op);
        tx.inverse.unshift(...inverse);
      }
    } finally {
      // Recalculate immediately so values are current inside transactions.
      if (tx.structural) {
        this.calc.rebuild();
        this.calc.recalcAll();
      } else {
        this.calc.recalc(tx.changed);
      }
      tx.structural ||= wasStructural;
      this.commit();
    }
  }

  private begin(options: ApplyOptions): void {
    if (this.txDepth++ === 0) {
      this.tx = { ops: [], inverse: [], changed: new Set(), structural: false, options };
    }
  }

  private commit(): void {
    if (--this.txDepth > 0) return;
    const tx = this.tx!;
    this.tx = null;
    if (!tx.ops.length) return;
    if (tx.options.undoable !== false) {
      const entry: HistoryEntry = { ops: tx.ops, inverse: tx.inverse };
      if (tx.options.label) entry.label = tx.options.label;
      this.undoStack.push(entry);
      if (this.undoStack.length > this.historyLimit) this.undoStack.shift();
      this.redoStack = [];
    }
    this.events.emit('change', { ops: tx.ops, origin: tx.options.origin ?? 'user', structural: tx.structural });
  }

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  undo(): boolean {
    const entry = this.undoStack.pop();
    if (!entry) return false;
    this.apply(entry.inverse, { origin: 'undo', undoable: false });
    this.redoStack.push(entry);
    return true;
  }

  redo(): boolean {
    const entry = this.redoStack.pop();
    if (!entry) return false;
    this.apply(entry.ops, { origin: 'redo', undoable: false });
    this.undoStack.push(entry);
    return true;
  }

  clearHistory(): void {
    this.undoStack = [];
    this.redoStack = [];
  }

  // -------------------------------------------------------------------------
  // Operation handlers (each returns its inverse)
  // -------------------------------------------------------------------------

  private applyOp(op: Operation, tx: Transaction): Operation[] {
    switch (op.type) {
      case 'setCells': {
        const sheet = this.requireSheet(op.sheet);
        const replacing = new Set(op.cells.filter((p) => !p.styleOnly).map((p) => cellKey(p.row, p.col)));
        for (const patch of op.cells) {
          const parent = sheet.getCell(patch.row, patch.col)?.spillParent;
          if (parent !== undefined && !patch.styleOnly && !replacing.has(parent)) throw new Error('Cannot change part of a spilled array. Edit its formula anchor instead.');
        }
        const prev = op.cells.map(({ row, col, data, styleOnly }) => {
          const existing = sheet.getCell(row, col);
          const before = existing?.spillParent !== undefined ? (existing.style ? { s: existing.style } : null) : Sheet.cellToData(existing);
          const cell: Cell | undefined = styleOnly
            ? { ...(existing ?? { value: null }), style: data?.s }
            : data ? Sheet.dataToCell(data) : undefined;
          sheet.putCell(row, col, cell);
          this.calc.update(sheet, row, col, cell?.formula);
          tx.changed.add(gid(sheet.id, row, col));
          return { row, col, data: before, styleOnly };
        });
        return [{ type: 'setCells', sheet: sheet.name, cells: prev, ...(op.reordersCells ? { reordersCells: true } : {}) }];
      }
      case 'insertRows':
      case 'insertCols': {
        const axis = op.type === 'insertRows' ? 'row' : 'col';
        this.structuralChange({ sheet: this.requireSheet(op.sheet).name, axis, index: op.index, count: op.count });
        tx.structural = true;
        return [{ type: axis === 'row' ? 'deleteRows' : 'deleteCols', sheet: op.sheet, index: op.index, count: op.count }];
      }
      case 'deleteRows':
      case 'deleteCols': {
        const snapshot = this.toJSON();
        const axis = op.type === 'deleteRows' ? 'row' : 'col';
        this.structuralChange({ sheet: this.requireSheet(op.sheet).name, axis, index: op.index, count: -op.count });
        tx.structural = true;
        return [{ type: 'restore', data: snapshot }];
      }
      case 'addSheet': {
        if (this.getSheet(op.name)) throw new Error(`Sheet already exists: ${op.name}`);
        const sheet = op.data ? Sheet.fromJSON({ ...op.data, name: op.name }) : new Sheet(op.name);
        const index = Math.max(0, Math.min(op.index ?? this.sheets.length, this.sheets.length));
        this.sheets.splice(index, 0, sheet);
        if (index <= this.activeSheetIndex && this.sheets.length > 1) this.activeSheetIndex++;
        tx.structural = true;
        return [{ type: 'removeSheet', name: op.name }];
      }
      case 'removeSheet': {
        const sheet = this.requireSheet(op.name);
        if (this.sheets.length === 1) throw new Error('Cannot remove the last sheet');
        const snapshot = this.toJSON();
        const index = this.sheets.indexOf(sheet);
        this.sheets.splice(index, 1);
        this.pivotTables = this.pivotTables.filter((p) => p.source.sheet.toUpperCase() !== sheet.name.toUpperCase() && p.target.sheet.toUpperCase() !== sheet.name.toUpperCase());
        this.rewriteAllRefs((f) => invalidateSheetRefs(f, sheet.name));
        if (this.activeSheetIndex >= this.sheets.length || this.activeSheetIndex > index) this.activeSheetIndex = Math.max(0, this.activeSheetIndex - 1);
        tx.structural = true;
        return [{ type: 'restore', data: snapshot }];
      }
      case 'renameSheet': {
        const sheet = this.requireSheet(op.from);
        const to = op.to.trim();
        if (!to || /[\\/?*[\]:]/.test(to) || to.length > 31) throw new Error(`Invalid sheet name: ${op.to}`);
        const existing = this.getSheet(to);
        if (existing && existing !== sheet) throw new Error(`Sheet already exists: ${to}`);
        const from = sheet.name;
        sheet.name = to;
        for (const pivot of this.pivotTables) {
          if (pivot.source.sheet.toUpperCase() === from.toUpperCase()) pivot.source.sheet = to;
          if (pivot.target.sheet.toUpperCase() === from.toUpperCase()) pivot.target.sheet = to;
        }
        this.rewriteAllRefs((f) => renameSheetInFormula(f, from, to));
        tx.structural = true;
        return [{ type: 'renameSheet', from: to, to: from }];
      }
      case 'moveSheet': {
        const sheet = this.requireSheet(op.name);
        const from = this.sheets.indexOf(sheet);
        const active = this.activeSheet;
        this.sheets.splice(from, 1);
        this.sheets.splice(Math.max(0, Math.min(op.index, this.sheets.length)), 0, sheet);
        this.activeSheetIndex = this.sheets.indexOf(active);
        tx.structural = true;
        return [{ type: 'moveSheet', name: sheet.name, index: from }];
      }
      case 'setColWidths':
      case 'setRowHeights': {
        const sheet = this.requireSheet(op.sheet);
        const map = op.type === 'setColWidths' ? sheet.colWidths : sheet.rowHeights;
        const prev = op.sizes.map(({ index, size }) => {
          const before = map.get(index) ?? null;
          if (size === null) map.delete(index);
          else map.set(index, Math.max(0, size));
          return { index, size: before };
        });
        return [{ type: op.type, sheet: sheet.name, sizes: prev }];
      }
      case 'setMerges': {
        const sheet = this.requireSheet(op.sheet);
        const prev = sheet.merges.map((m) => formatRange(m));
        sheet.merges = op.merges.map((m) => parseRange(m)?.range).filter((r): r is RangeAddress => !!r);
        tx.structural = true;
        return [{ type: 'setMerges', sheet: sheet.name, merges: prev }];
      }
      case 'setFrozen': {
        const sheet = this.requireSheet(op.sheet);
        const prev = { ...sheet.frozen };
        sheet.frozen = { rows: Math.max(0, op.rows), cols: Math.max(0, op.cols) };
        return [{ type: 'setFrozen', sheet: sheet.name, ...prev }];
      }
      case 'setCharts': {
        const sheet = this.requireSheet(op.sheet);
        const prev = structuredClone(sheet.charts);
        sheet.charts = structuredClone(op.charts);
        return [{ type: 'setCharts', sheet: sheet.name, charts: prev }];
      }
      case 'setConditionalFormats': {
        validateConditionalFormats(op.rules);
        const sheet = this.requireSheet(op.sheet);
        const prev = structuredClone(sheet.conditionalFormats);
        sheet.conditionalFormats = structuredClone(op.rules);
        return [{ type: 'setConditionalFormats', sheet: sheet.name, rules: prev }];
      }
      case 'setValidations': {
        validateValidations(op.rules);
        const sheet = this.requireSheet(op.sheet);
        const prev = structuredClone(sheet.validations);
        sheet.validations = structuredClone(op.rules);
        return [{ type: 'setValidations', sheet: sheet.name, rules: prev }];
      }
      case 'setPivotTables': {
        const ids = new Set<string>();
        for (const table of op.tables) {
          if (!table.id || ids.has(table.id)) throw new Error('Pivot IDs must be nonempty and unique');
          ids.add(table.id);
        }
        const prev = structuredClone(this.pivotTables);
        this.pivotTables = structuredClone(op.tables);
        return [{ type: 'setPivotTables', tables: prev }];
      }
      case 'setTables': {
        const sheet = this.requireSheet(op.sheet);
        validateTables(this, sheet, op.tables);
        const previous = structuredClone(sheet.tables);
        sheet.tables = structuredClone(op.tables);
        sheet.invalidate();
        return [{ type: 'setTables', sheet: sheet.name, tables: previous }];
      }
      case 'setName': {
        const key = op.name.toUpperCase();
        const prev = this.names.get(key)?.ref ?? null;
        if (op.ref === null) this.names.delete(key);
        else this.names.set(key, { name: op.name, ref: op.ref.replace(/^=/, '') });
        tx.structural = true;
        return [{ type: 'setName', name: op.name, ref: prev }];
      }
      case 'restore': {
        const snapshot = this.toJSON();
        const active = this.activeSheetIndex;
        this.sheets = op.data.sheets.map((s) => Sheet.fromJSON(s));
        this.pivotTables = structuredClone(op.data.pivotTables ?? []);
        this.names.clear();
        for (const [name, ref] of Object.entries(op.data.names ?? {})) this.names.set(name.toUpperCase(), { name, ref });
        this.activeSheetIndex = Math.min(active, this.sheets.length - 1);
        tx.structural = true;
        return [{ type: 'restore', data: snapshot }];
      }
    }
  }

  /** Applies a function to every formula, chart reference and defined name. */
  private rewriteAllRefs(fn: (formula: string, sheetName: string) => string): void {
    for (const sheet of this.sheets) {
      for (const cell of sheet.cells.values()) {
        if (cell.formula !== undefined) cell.formula = fn(cell.formula, sheet.name);
      }
      for (const chart of sheet.charts) rewriteChartRefs(chart, (ref) => fn(ref, sheet.name));
      for (const rule of sheet.conditionalFormats) {
        if (rule.type !== 'colorScale') rule.formulas = rule.formulas.map((f) => fn(f, sheet.name));
      }
      for (const rule of sheet.validations) {
        if (rule.formula1 !== undefined && !/^".*"$/s.test(rule.formula1)) rule.formula1 = fn(rule.formula1, sheet.name);
        if (rule.formula2 !== undefined) rule.formula2 = fn(rule.formula2, sheet.name);
      }
    }
    for (const n of this.names.values()) n.ref = fn(n.ref, this.sheets[0]?.name ?? '');
  }

  private structuralChange(change: StructuralChange): void {
    this.calc.clearSpills();
    const sheet = this.requireSheet(change.sheet);
    const isRow = change.axis === 'row';
    const { index, count } = change;
    sheet.tables = sheet.tables.flatMap((table) => {
      const range = adjustFormulaForStructure(table.range, sheet.name, change);
      if (range === '#REF!') return [];
      const previous = parseRange(table.range)!.range;
      const headerDeleted = isRow && count < 0 && previous.start.row >= index && previous.start.row < index - count;
      const updated = { ...table, range, headerRow: table.headerRow && !headerDeleted };
      if (headerDeleted) delete updated.filters;
      else if (!isRow && table.filters) {
        const start = parseRange(range)!.range.start.col;
        updated.filters = table.filters.flatMap((f) => {
          const col = previous.start.col + f.column;
          if (count < 0 && col >= index && col < index - count) return [];
          const moved = col >= index ? col + count : col;
          return [{ ...f, column: moved - start }];
        });
      }
      return [updated];
    });
    this.pivotTables = this.pivotTables.flatMap((pivot) => {
      const updated = adjustPivotForStructure(this, pivot, change);
      return updated ? [updated] : [];
    });
    sheet.conditionalFormats = sheet.conditionalFormats.flatMap((rule) => {
      const ref = adjustFormulaForStructure(rule.range, sheet.name, change);
      if (ref === '#REF!') return [];
      const range = parseRange(rule.range)!.range;
      const key = isRow ? 'row' : 'col';
      const offset = count < 0 && range.start[key] >= index && range.start[key] < index - count ? index - count - range.start[key] : 0;
      return [{ ...rule, range: ref, ...(rule.type !== 'colorScale' && offset ? { formulas: rule.formulas.map((f) => shiftFormula(f, isRow ? offset : 0, isRow ? 0 : offset)) } : {}) }];
    });
    sheet.validations = sheet.validations.flatMap((rule) => {
      const before = validationRanges(rule);
      const parts = before.map((r) => adjustFormulaForStructure(formatRange(r), sheet.name, change)).filter((r) => r !== '#REF!');
      if (!parts.length) return [];
      // Formulas are relative to the first range; keep them anchored when its start moves by deletion.
      const first = before[0]!, key = isRow ? 'row' : 'col';
      const offset = count < 0 && first.start[key] >= index && first.start[key] < index - count ? index - count - first.start[key] : 0;
      const shift = (f: string | undefined) => (f === undefined || !offset || /^".*"$/s.test(f) ? f : shiftFormula(f, isRow ? offset : 0, isRow ? 0 : offset));
      const updated: DataValidation = { ...rule, range: parts.join(' ') };
      if (rule.formula1 !== undefined) updated.formula1 = shift(rule.formula1)!;
      if (rule.formula2 !== undefined) updated.formula2 = shift(rule.formula2)!;
      return [updated];
    });
    const limit = isRow ? MAX_ROWS : MAX_COLS;
    const move = (pos: number): number | null => {
      if (count > 0) return pos >= index ? (pos + count < limit ? pos + count : null) : pos;
      const n = -count;
      if (pos < index) return pos;
      if (pos < index + n) return null;
      return pos - n;
    };

    // Cells
    const entries = [...sheet.cells.entries()];
    sheet.cells.clear();
    for (const [key, cell] of entries) {
      const { row, col } = keyToAddress(key);
      const p = move(isRow ? row : col);
      if (p === null) continue;
      sheet.cells.set(isRow ? cellKey(p, col) : cellKey(row, p), cell);
    }
    sheet.invalidate();

    // Notes and links
    for (const map of [sheet.notes, sheet.links] as Map<number, unknown>[]) {
      const items = [...map.entries()];
      map.clear();
      for (const [key, v] of items) {
        const { row, col } = keyToAddress(key);
        const p = move(isRow ? row : col);
        if (p !== null) map.set(isRow ? cellKey(p, col) : cellKey(row, p), v);
      }
    }

    // Sizes
    const sizes = isRow ? sheet.rowHeights : sheet.colWidths;
    const sizeEntries = [...sizes.entries()];
    sizes.clear();
    for (const [i, v] of sizeEntries) {
      const p = move(i);
      if (p !== null) sizes.set(p, v);
    }

    // Merges
    sheet.merges = sheet.merges.flatMap((m) => {
      const a = isRow ? m.start.row : m.start.col;
      const b = isRow ? m.end.row : m.end.col;
      const na = move(a) ?? (count < 0 ? index : null);
      let nb = move(b);
      if (nb === null && count < 0) nb = index - 1;
      if (na === null || nb === null || nb < na) return [];
      const r = structuredClone(m);
      if (isRow) [r.start.row, r.end.row] = [na, nb];
      else [r.start.col, r.end.col] = [na, nb];
      return r.start.row === r.end.row && r.start.col === r.end.col ? [] : [r];
    });

    // Chart anchors on this sheet
    for (const chart of sheet.charts) {
      for (const pt of [chart.anchor.from, chart.anchor.to]) {
        const v = isRow ? pt.row : pt.col;
        const p = move(v) ?? index;
        if (isRow) pt.row = p;
        else pt.col = p;
      }
    }

    this.rewriteAllRefs((f, formulaSheet) => adjustFormulaForStructure(f, formulaSheet, change));
  }

  // -------------------------------------------------------------------------
  // Convenience accessors
  // -------------------------------------------------------------------------

  getValue(sheet: string | number, row: number, col: number): Scalar {
    return this.requireSheet(sheet).getValue(row, col);
  }

  getCellStyle(sheet: string | number, row: number, col: number): CellStyle | undefined {
    return new ConditionalFormatEvaluator(this).style(this.requireSheet(sheet), row, col);
  }

  getChart(id: string): { sheet: Sheet; chart: ChartSpec } | undefined {
    for (const sheet of this.sheets) {
      const chart = sheet.charts.find((c) => c.id === id);
      if (chart) return { sheet, chart };
    }
    return undefined;
  }

  addSheet(name = this.uniqueSheetName(), data?: SheetData): Sheet {
    this.apply(data ? { type: 'addSheet', name, data } : { type: 'addSheet', name });
    return this.requireSheet(name);
  }
}

function rewriteChartRefs(chart: ChartSpec, fn: (ref: string) => string): void {
  for (const s of chart.series) {
    s.values = fn(s.values);
    if (s.categories) s.categories = fn(s.categories);
    if (s.nameRef) s.nameRef = fn(s.nameRef);
  }
}
