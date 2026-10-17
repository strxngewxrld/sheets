import {
  applyBorders,
  applyDataValidation,
  cellKey,
  applyStyle,
  CHART_TYPES,
  clearRange,
  CONDITIONAL_OPERATORS,
  ConditionalFormatEvaluator,
  ICON_SETS,
  createPivotTable,
  colToName,
  copyRange,
  findCells,
  formatAddress,
  formatRange,
  formatValue,
  formatTable,
  hideCols,
  hideRows,
  insertChartFromRange,
  isError,
  keyToAddress,
  listFunctions,
  mergeCells,
  nameToCol,
  parseFormula,
  parseRange,
  pasteRange,
  PIVOT_AGGREGATES,
  refreshPivotTable,
  removePivotTable,
  removeChart,
  removeTable,
  replaceCells,
  resolveChart,
  setValues,
  setConditionalFormats,
  shiftFormula,
  sortRange,
  unhideCols,
  unhideRows,
  unmergeCells,
  updateChart,
  invalidCells,
  validationAt,
  VALIDATION_OPERATORS,
  VALIDATION_TYPES,
  walk,
  chartFromRange,
  MAX_COLS,
  MAX_ROWS,
  TABLE_THEMES,
  type CellPatch,
  type DataValidation,
  type ValidationOperator,
  type ValidationType,
  type ConditionalFormatRule,
  type ChartSpec,
  type ChartType,
  type ClearMode,
  type InputValue,
  type LegendPosition,
  type RangeAddress,
  type PivotTableSpec,
  type Scalar,
  type Sheet,
  type StylePatch,
  type Workbook,
} from '@sheets/core';

/** JSON Schema subset used for tool inputs. */
export type JsonSchema = {
  type?: string | string[];
  description?: string;
  properties?: Record<string, JsonSchema>;
  required?: string[];
  items?: JsonSchema;
  enum?: readonly (string | number | boolean)[];
  additionalProperties?: boolean | JsonSchema;
  minimum?: number;
  maximum?: number;
  default?: unknown;
};

const THRESHOLD_SCHEMA: JsonSchema = {
  type: 'object',
  properties: { type: { type: 'string', enum: ['min', 'max', 'num', 'percent', 'percentile'] }, value: { type: 'number' } },
  required: ['type'],
};

export interface ToolDefinition {
  name: string;
  description: string;
  inputSchema: JsonSchema & { type: 'object' };
  /** True if the tool never modifies the workbook. */
  readOnly?: boolean;
}

/** Optional UI integration (selection, navigation) provided by a host such as the spreadsheet view. */
export interface AgentHost {
  getSelection?(): { sheet: string; range: string } | undefined;
  select?(sheet: string, range: RangeAddress): void;
  activateSheet?(sheet: string): void;
}

export interface ToolContext {
  wb: Workbook;
  host: AgentHost | undefined;
  maxCells: number;
}

export class ToolInputError extends Error {}

type Input = Record<string, unknown>;

interface ToolImpl extends ToolDefinition {
  run(ctx: ToolContext, input: Input): unknown;
}

// ---------------------------------------------------------------------------
// Input helpers
// ---------------------------------------------------------------------------

function str(input: Input, key: string, required = true): string {
  const v = input[key];
  if (v === undefined || v === null || v === '') {
    if (required) throw new ToolInputError(`Missing required parameter "${key}"`);
    return '';
  }
  return String(v);
}

function optStr(input: Input, key: string): string | undefined {
  const v = input[key];
  return v === undefined || v === null || v === '' ? undefined : String(v);
}

function optNum(input: Input, key: string): number | undefined {
  const v = input[key];
  if (v === undefined || v === null || v === '') return undefined;
  const n = Number(v);
  if (!Number.isFinite(n)) throw new ToolInputError(`Parameter "${key}" must be a number`);
  return n;
}

function sheetOf(ctx: ToolContext, input: Input, rangeText?: string): Sheet {
  const qualified = rangeText ? parseRange(rangeText)?.sheet : undefined;
  const name = qualified ?? optStr(input, 'sheet');
  if (!name) return ctx.wb.activeSheet;
  const sheet = ctx.wb.getSheet(name);
  if (!sheet) throw new ToolInputError(`Sheet "${name}" not found. Available: ${ctx.wb.sheets.map((s) => s.name).join(', ')}`);
  return sheet;
}

function rangeOf(text: string): RangeAddress {
  const parsed = parseRange(text);
  if (!parsed) throw new ToolInputError(`Invalid range "${text}". Use A1 notation like "A1:C10", "B:B" or "Sheet1!A1".`);
  return parsed.range;
}

function cellOf(text: string): { row: number; col: number } {
  const r = rangeOf(text);
  return r.start;
}

function clip(sheet: Sheet, range: RangeAddress): RangeAddress {
  const used = sheet.usedRange();
  const end = { ...range.end };
  if (range.end.row === MAX_ROWS - 1) end.row = Math.max(range.start.row, used?.end.row ?? range.start.row);
  if (range.end.col === MAX_COLS - 1) end.col = Math.max(range.start.col, used?.end.col ?? range.start.col);
  return { start: range.start, end };
}

/** JSON-friendly scalar: errors become their code string ("#DIV/0!"). */
function jsonValue(v: Scalar): string | number | boolean | null {
  return isError(v) ? v.code : v;
}

function columnIndex(input: string): number {
  const t = input.trim().toUpperCase();
  if (/^\d+$/.test(t)) return Number(t) - 1;
  if (!/^[A-Z]{1,3}$/.test(t)) throw new ToolInputError(`Invalid column "${input}". Use a letter like "B".`);
  return nameToCol(t);
}

function rowIndex(v: unknown): number {
  const n = Number(v);
  if (!Number.isInteger(n) || n < 1) throw new ToolInputError(`Invalid row number "${v}". Rows are 1-based.`);
  return n - 1;
}

function chartSummary(wb: Workbook, sheet: string, c: ChartSpec) {
  return {
    id: c.id,
    sheet,
    type: c.type,
    title: c.title ?? null,
    series: c.series.map((s) => ({ name: s.name ?? s.nameRef ?? null, values: s.values, categories: s.categories ?? null })),
    position: `${formatAddress(c.anchor.from.row, c.anchor.from.col)}:${formatAddress(c.anchor.to.row, c.anchor.to.col)}`,
    stacked: !!c.stacked,
    legend: c.legend ?? 'right',
    preview: (() => {
      const d = resolveChart(wb, c);
      return { categories: d.categories.slice(0, 20), series: d.series.map((s) => ({ name: s.name, values: s.values.slice(0, 20) })) };
    })(),
  };
}

const STYLE_SCHEMA: JsonSchema = {
  type: 'object',
  description: 'Style properties to set. Use null to remove a property.',
  properties: {
    bold: { type: 'boolean' },
    italic: { type: 'boolean' },
    underline: { type: 'boolean' },
    strike: { type: 'boolean' },
    fontName: { type: 'string' },
    fontSize: { type: 'number', description: 'Font size in points' },
    color: { type: 'string', description: 'Font color as hex, e.g. "#C00000"' },
    fill: { type: 'string', description: 'Background color as hex, e.g. "#FFF2CC"' },
    hAlign: { type: 'string', enum: ['left', 'center', 'right', 'justify'] },
    vAlign: { type: 'string', enum: ['top', 'middle', 'bottom'] },
    wrap: { type: 'boolean' },
    numFmt: { type: 'string', description: 'Excel number format, e.g. "#,##0.00", "0%", "dd.mm.yyyy", "#,##0.00 \\"€\\""' },
  },
  additionalProperties: false,
};

const SHEET_PARAM: JsonSchema = { type: 'string', description: 'Sheet name. Defaults to the active sheet.' };

// ---------------------------------------------------------------------------
// Tools
// ---------------------------------------------------------------------------

export const TOOLS: ToolImpl[] = [
  {
    name: 'get_workbook_info',
    description:
      'Overview of the workbook: sheets with their used ranges and charts, the active sheet, defined names and the current user selection. Call this first to orient yourself.',
    readOnly: true,
    inputSchema: { type: 'object', properties: {} },
    run: ({ wb, host }) => ({
      activeSheet: wb.activeSheet.name,
      selection: host?.getSelection?.() ?? null,
      sheets: wb.sheets.map((s) => {
        const used = s.usedRange();
        return {
          name: s.name,
          usedRange: used ? formatRange({ start: { row: 0, col: 0 }, end: used.end }) : null,
          rows: used ? used.end.row + 1 : 0,
          columns: used ? used.end.col + 1 : 0,
          cellCount: s.cells.size,
          charts: s.charts.map((c) => ({ id: c.id, type: c.type, title: c.title ?? null })),
          frozen: s.frozen,
          merges: s.merges.map((m) => formatRange(m)),
          conditionalFormats: s.conditionalFormats.length,
          tables: s.tables.map((t) => ({ name: t.name, range: t.range, headerRow: t.headerRow, bandedRows: t.bandedRows })),
        };
      }),
      names: Object.fromEntries([...wb.names.values()].map((n) => [n.name, n.ref])),
      pivotTables: structuredClone(wb.pivotTables),
      canUndo: wb.canUndo,
      canRedo: wb.canRedo,
    }),
  },
  {
    name: 'read_range',
    description:
      'Reads cell values from a range. Returns a 2D array of computed values (numbers, strings, booleans, null for empty, "#ERR!" strings for errors) plus formulas for cells that contain them. Use format "markdown" for a readable table, "display" for formatted text as the user sees it.',
    readOnly: true,
    inputSchema: {
      type: 'object',
      properties: {
        sheet: SHEET_PARAM,
        range: { type: 'string', description: 'A1 range, e.g. "A1:D20", "B:B" or "Sheet2!A1:C5". Whole columns/rows are clipped to the used area.' },
        format: { type: 'string', enum: ['values', 'display', 'markdown'], default: 'values' },
        include_formulas: { type: 'boolean', default: true },
      },
      required: ['range'],
    },
    run: (ctx, input) => {
      const text = str(input, 'range');
      const sheet = sheetOf(ctx, input, text);
      const range = clip(sheet, rangeOf(text));
      const rows = range.end.row - range.start.row + 1;
      const cols = range.end.col - range.start.col + 1;
      let lastRow = range.end.row;
      let truncated = false;
      if (rows * cols > ctx.maxCells) {
        lastRow = range.start.row + Math.max(1, Math.floor(ctx.maxCells / cols)) - 1;
        truncated = true;
      }
      const format = optStr(input, 'format') ?? 'values';
      const values: unknown[][] = [];
      const formulas: Record<string, string> = {};
      const conditional = new ConditionalFormatEvaluator(ctx.wb);
      for (let r = range.start.row; r <= lastRow; r++) {
        const line: unknown[] = [];
        for (let c = range.start.col; c <= range.end.col; c++) {
          const cell = sheet.getCell(r, c);
          if (cell?.formula !== undefined && input.include_formulas !== false) formulas[formatAddress(r, c)] = `=${cell.formula}`;
          if (format === 'values') line.push(jsonValue(cell?.value ?? null));
          else line.push(formatValue(cell?.value ?? null, conditional.style(sheet, r, c)?.numFmt));
        }
        values.push(line);
      }
      const actual = formatRange({ start: range.start, end: { row: lastRow, col: range.end.col } });
      const result: Record<string, unknown> = { sheet: sheet.name, range: actual };
      if (format === 'markdown') {
        const header = ['', ...Array.from({ length: cols }, (_, i) => colToName(range.start.col + i))];
        const lines = [`| ${header.join(' | ')} |`, `|${header.map(() => '---').join('|')}|`];
        values.forEach((line, i) => lines.push(`| ${range.start.row + i + 1} | ${line.map((v) => String(v).replace(/\|/g, '\\|')).join(' | ')} |`));
        result.table = lines.join('\n');
      } else result.values = values;
      if (Object.keys(formulas).length) result.formulas = formulas;
      if (truncated) result.truncated = `Only the first ${lastRow - range.start.row + 1} rows were returned. Read the rest in smaller chunks.`;
      return result;
    },
  },
  {
    name: 'write_range',
    description:
      'Writes a 2D array of values starting at a cell. Strings starting with "=" are formulas (e.g. "=SUM(B2:B10)"); other strings are parsed like typed input ("12%", "2024-01-31", "1,234.5"). null clears a cell. Existing formatting is kept. Returns the computed values of written formulas.',
    inputSchema: {
      type: 'object',
      properties: {
        sheet: SHEET_PARAM,
        start: { type: 'string', description: 'Top-left cell, e.g. "A1"' },
        values: {
          type: 'array',
          description: 'Rows of cell values, e.g. [["Name","Total"],["A",1],["Sum","=SUM(B2:B2)"]]',
          items: { type: 'array', items: { type: ['string', 'number', 'boolean', 'null'] } },
        },
        parse: { type: 'boolean', description: 'Parse strings like typed input (default true). Set false to preserve non-formula strings literally, especially identifiers with leading zeroes and Business Central Code fields. Strings starting with "=" are still formulas.', default: true },
      },
      required: ['start', 'values'],
    },
    run: (ctx, input) => {
      const start = str(input, 'start');
      const sheet = sheetOf(ctx, input, start);
      const { row, col } = cellOf(start);
      const values = input.values;
      if (!Array.isArray(values) || !values.every(Array.isArray)) throw new ToolInputError('"values" must be a 2D array (array of rows).');
      const opts: { parse?: boolean } = {};
      if (input.parse === false) opts.parse = false;
      setValues(ctx.wb, sheet.name, row, col, values as InputValue[][], opts);
      const width = Math.max(1, ...values.map((r) => (r as unknown[]).length));
      const range = { start: { row, col }, end: { row: row + values.length - 1, col: col + width - 1 } };
      const computed: Record<string, unknown> = {};
      (values as InputValue[][]).forEach((line, i) =>
        line.forEach((v, j) => {
          if (typeof v === 'string' && v.startsWith('=')) computed[formatAddress(row + i, col + j)] = jsonValue(sheet.getValue(row + i, col + j));
        }),
      );
      return { written: formatRange(range, sheet.name), ...(Object.keys(computed).length ? { formulaResults: computed } : {}) };
    },
  },
  {
    name: 'fill_formula',
    description:
      'Writes one formula into every cell of a range, adjusting relative references like Excel\'s fill-down/right. The formula is written as it should appear in the top-left cell, e.g. range "D2:D100", formula "=B2*C2".',
    inputSchema: {
      type: 'object',
      properties: {
        sheet: SHEET_PARAM,
        range: { type: 'string', description: 'Target range, e.g. "D2:D100"' },
        formula: { type: 'string', description: 'Formula for the top-left cell, starting with "="' },
      },
      required: ['range', 'formula'],
    },
    run: (ctx, input) => {
      const text = str(input, 'range');
      const sheet = sheetOf(ctx, input, text);
      const range = clip(sheet, rangeOf(text));
      const formula = str(input, 'formula').replace(/^=/, '');
      parseFormula(formula);
      const cells: CellPatch[] = [];
      for (let r = range.start.row; r <= range.end.row; r++) {
        for (let c = range.start.col; c <= range.end.col; c++) {
          const style = sheet.getCell(r, c)?.style;
          const f = shiftFormula(formula, r - range.start.row, c - range.start.col);
          cells.push({ row: r, col: c, data: style ? { f, s: style } : { f } });
        }
      }
      ctx.wb.apply({ type: 'setCells', sheet: sheet.name, cells });
      const sample = cells.slice(0, 5).map((p) => ({ cell: formatAddress(p.row, p.col), formula: `=${p.data!.f}`, value: jsonValue(sheet.getValue(p.row, p.col)) }));
      return { filled: formatRange(range, sheet.name), cells: cells.length, sample };
    },
  },
  {
    name: 'clear_range',
    description: 'Clears a range. mode "contents" keeps formatting, "formats" keeps values, "all" removes everything.',
    inputSchema: {
      type: 'object',
      properties: {
        sheet: SHEET_PARAM,
        range: { type: 'string' },
        mode: { type: 'string', enum: ['contents', 'formats', 'all'], default: 'contents' },
      },
      required: ['range'],
    },
    run: (ctx, input) => {
      const text = str(input, 'range');
      const sheet = sheetOf(ctx, input, text);
      clearRange(ctx.wb, sheet.name, rangeOf(text), (optStr(input, 'mode') ?? 'contents') as ClearMode);
      return { cleared: formatRange(rangeOf(text), sheet.name) };
    },
  },
  {
    name: 'format_range',
    description: 'Applies formatting (font, colors, alignment, number format, borders) to a range.',
    inputSchema: {
      type: 'object',
      properties: {
        sheet: SHEET_PARAM,
        range: { type: 'string' },
        style: STYLE_SCHEMA,
        borders: { type: 'string', enum: ['all', 'outline', 'none'], description: 'Optional thin borders around/inside the range' },
      },
      required: ['range'],
    },
    run: (ctx, input) => {
      const text = str(input, 'range');
      const sheet = sheetOf(ctx, input, text);
      const range = rangeOf(text);
      ctx.wb.transaction(() => {
        if (input.style && typeof input.style === 'object') applyStyle(ctx.wb, sheet.name, range, input.style as StylePatch);
        const b = optStr(input, 'borders');
        if (b) applyBorders(ctx.wb, sheet.name, clip(sheet, range), b as 'all' | 'outline' | 'none');
      });
      return { formatted: formatRange(range, sheet.name) };
    },
  },
  {
    name: 'format_table',
    description: 'Formats a data range as a table with a header and alternating row colors. Table styles stay attached to row positions when sorting. A single undo step; saved as a native XLSX table. Existing direct cell formats take priority.',
    inputSchema: { type: 'object', properties: {
      sheet: SHEET_PARAM, range: { type: 'string' }, name: { type: 'string' },
      theme: { type: 'string', enum: Object.keys(TABLE_THEMES), default: 'blue' },
      header_row: { type: 'boolean', default: true }, banded_rows: { type: 'boolean', default: true },
    }, required: ['range'] },
    run: (ctx, input) => {
      const text = str(input, 'range'); const sheet = sheetOf(ctx, input, text);
      const theme = optStr(input, 'theme');
      if (theme && !TABLE_THEMES[theme]) throw new ToolInputError('Unknown table theme');
      return formatTable(ctx.wb, sheet.name, clip(sheet, rangeOf(text)), { name: optStr(input, 'name'), headerRow: input.header_row !== false, bandedRows: input.banded_rows !== false, theme: theme ? TABLE_THEMES[theme] : undefined });
    },
  },
  {
    name: 'list_tables', description: 'Lists the named tables and their ranges, header flags, stripe settings and styles.', readOnly: true,
    inputSchema: { type: 'object', properties: { sheet: SHEET_PARAM } },
    run: (ctx, input) => ({ sheet: sheetOf(ctx, input).name, tables: structuredClone(sheetOf(ctx, input).tables) }),
  },
  {
    name: 'remove_table', description: 'Removes a table and its highlighting, keeping its data and direct cell formats. Supports undo.',
    inputSchema: { type: 'object', properties: { sheet: SHEET_PARAM, name: { type: 'string' } }, required: ['name'] },
    run: (ctx, input) => { removeTable(ctx.wb, sheetOf(ctx, input).name, str(input, 'name')); return { removed: true }; },
  },
  {
    name: 'set_conditional_formats',
    description: 'Replaces all conditional formatting rules on a sheet. Earlier rules take priority. Supports cellIs comparisons, expression formulas relative to the range top-left, two/three-color scales, data bars (color, optional min/max thresholds) and icon sets (iconSet, optional thresholds lowest first, reverse). Pass [] to remove all rules.',
    inputSchema: {
      type: 'object', properties: {
        sheet: SHEET_PARAM,
        rules: { type: 'array', items: { type: 'object', properties: {
          id: { type: 'string' }, range: { type: 'string' },
          type: { type: 'string', enum: ['cellIs', 'expression', 'colorScale', 'dataBar', 'iconSet'] },
          operator: { type: 'string', enum: CONDITIONAL_OPERATORS },
          formulas: { type: 'array', items: { type: 'string' } },
          style: STYLE_SCHEMA, colors: { type: 'array', items: { type: 'string' } },
          color: { type: 'string', description: 'Data bar color (#RRGGBB)' },
          min: THRESHOLD_SCHEMA, max: THRESHOLD_SCHEMA,
          iconSet: { type: 'string', enum: ICON_SETS },
          thresholds: { type: 'array', items: THRESHOLD_SCHEMA, description: 'One per icon, lowest first; default equal percent steps' },
          reverse: { type: 'boolean' }, showValue: { type: 'boolean' },
          stopIfTrue: { type: 'boolean' },
        }, required: ['id', 'range', 'type'] } },
      }, required: ['rules'],
    },
    run: (ctx, input) => {
      if (!Array.isArray(input.rules)) throw new ToolInputError('rules must be an array');
      const sheet = sheetOf(ctx, input);
      setConditionalFormats(ctx.wb, sheet.name, input.rules as ConditionalFormatRule[]);
      return { sheet: sheet.name, rules: structuredClone(sheet.conditionalFormats) };
    },
  },
  {
    name: 'list_conditional_formats',
    description: 'Lists conditional formatting rules in priority order.',
    readOnly: true,
    inputSchema: { type: 'object', properties: { sheet: SHEET_PARAM } },
    run: (ctx, input) => {
      const sheet = sheetOf(ctx, input);
      return { sheet: sheet.name, rules: structuredClone(sheet.conditionalFormats) };
    },
  },
  {
    name: 'create_pivot_table',
    description: 'Creates or refreshes a pivot report with header-based fields, row/column groups, filters and aggregates. The source must include a header row. Unrelated target cells are protected. Refresh explicitly after source edits. XLSX stores results as ordinary cells and keeps the Sheets definition, without a native Excel pivot cache.',
    inputSchema: {
      type: 'object', properties: {
        id: { type: 'string', description: 'Unique pivot ID; reuse to refresh the same report.' },
        source_sheet: SHEET_PARAM, source_range: { type: 'string' },
        target_sheet: SHEET_PARAM, target_cell: { type: 'string' },
        rows: { type: 'array', items: { type: ['string', 'integer'] } },
        columns: { type: 'array', items: { type: ['string', 'integer'] } },
        values: { type: 'array', items: { type: 'object', properties: {
          field: { type: ['string', 'integer'] }, aggregate: { type: 'string', enum: PIVOT_AGGREGATES }, label: { type: 'string' },
        }, required: ['field', 'aggregate'] } },
        filters: { type: 'array', items: { type: 'object', properties: { field: { type: ['string', 'integer'] }, criteria: { type: ['string', 'number', 'boolean', 'null'] } }, required: ['field', 'criteria'] } },
        grand_total_row: { type: 'boolean' }, grand_total_column: { type: 'boolean' },
      }, required: ['id', 'source_range', 'target_cell', 'rows', 'values'],
    },
    run: (ctx, input) => {
      const sourceText = str(input, 'source_range');
      const source = sheetOf(ctx, { sheet: input.source_sheet }, sourceText);
      const targetText = str(input, 'target_cell');
      const target = sheetOf(ctx, { sheet: input.target_sheet }, targetText);
      const spec: PivotTableSpec = {
        id: str(input, 'id'), source: { sheet: source.name, range: formatRange(rangeOf(sourceText)) },
        target: { sheet: target.name, cell: formatAddress(cellOf(targetText).row, cellOf(targetText).col) },
        rows: input.rows as PivotTableSpec['rows'], columns: input.columns as PivotTableSpec['columns'],
        values: input.values as PivotTableSpec['values'], filters: input.filters as PivotTableSpec['filters'],
        grandTotalRow: input.grand_total_row !== false, grandTotalColumn: input.grand_total_column !== false,
      };
      return createPivotTable(ctx.wb, spec);
    },
  },
  {
    name: 'refresh_pivot_table', description: 'Rebuilds a pivot report from the current source values as one undoable change.',
    inputSchema: { type: 'object', properties: { id: { type: 'string' } }, required: ['id'] },
    run: (ctx, input) => refreshPivotTable(ctx.wb, str(input, 'id')),
  },
  {
    name: 'list_pivot_tables', description: 'Lists pivot definitions, source/target ranges and last output rectangles.', readOnly: true,
    inputSchema: { type: 'object', properties: {} },
    run: (ctx) => ({ tables: structuredClone(ctx.wb.pivotTables) }),
  },
  {
    name: 'remove_pivot_table', description: 'Removes a pivot definition and its output, preserving cell formatting. Supports undo.',
    inputSchema: { type: 'object', properties: { id: { type: 'string' } }, required: ['id'] },
    run: (ctx, input) => { removePivotTable(ctx.wb, str(input, 'id')); return { removed: true }; },
  },
  {
    name: 'insert_rows',
    description: 'Inserts empty rows before the given 1-based row number. Formulas are adjusted.',
    inputSchema: {
      type: 'object',
      properties: { sheet: SHEET_PARAM, before_row: { type: 'integer', minimum: 1 }, count: { type: 'integer', minimum: 1, default: 1 } },
      required: ['before_row'],
    },
    run: (ctx, input) => {
      const sheet = sheetOf(ctx, input);
      ctx.wb.apply({ type: 'insertRows', sheet: sheet.name, index: rowIndex(input.before_row), count: optNum(input, 'count') ?? 1 });
      return { ok: true };
    },
  },
  {
    name: 'delete_rows',
    description: 'Deletes rows starting at the given 1-based row number. References to deleted cells become #REF!.',
    inputSchema: {
      type: 'object',
      properties: { sheet: SHEET_PARAM, row: { type: 'integer', minimum: 1 }, count: { type: 'integer', minimum: 1, default: 1 } },
      required: ['row'],
    },
    run: (ctx, input) => {
      const sheet = sheetOf(ctx, input);
      ctx.wb.apply({ type: 'deleteRows', sheet: sheet.name, index: rowIndex(input.row), count: optNum(input, 'count') ?? 1 });
      return { ok: true };
    },
  },
  {
    name: 'insert_columns',
    description: 'Inserts empty columns before the given column letter. Formulas are adjusted.',
    inputSchema: {
      type: 'object',
      properties: { sheet: SHEET_PARAM, before_column: { type: 'string', description: 'Column letter, e.g. "C"' }, count: { type: 'integer', minimum: 1, default: 1 } },
      required: ['before_column'],
    },
    run: (ctx, input) => {
      const sheet = sheetOf(ctx, input);
      ctx.wb.apply({ type: 'insertCols', sheet: sheet.name, index: columnIndex(str(input, 'before_column')), count: optNum(input, 'count') ?? 1 });
      return { ok: true };
    },
  },
  {
    name: 'delete_columns',
    description: 'Deletes columns starting at the given column letter.',
    inputSchema: {
      type: 'object',
      properties: { sheet: SHEET_PARAM, column: { type: 'string' }, count: { type: 'integer', minimum: 1, default: 1 } },
      required: ['column'],
    },
    run: (ctx, input) => {
      const sheet = sheetOf(ctx, input);
      ctx.wb.apply({ type: 'deleteCols', sheet: sheet.name, index: columnIndex(str(input, 'column')), count: optNum(input, 'count') ?? 1 });
      return { ok: true };
    },
  },
  {
    name: 'add_sheet',
    description: 'Adds a new worksheet.',
    inputSchema: {
      type: 'object',
      properties: { name: { type: 'string' }, position: { type: 'integer', description: '0-based position; defaults to the end' }, activate: { type: 'boolean', default: true } },
      required: ['name'],
    },
    run: (ctx, input) => {
      const name = str(input, 'name');
      const pos = optNum(input, 'position');
      ctx.wb.apply(pos === undefined ? { type: 'addSheet', name } : { type: 'addSheet', name, index: pos });
      if (input.activate !== false) activate(ctx, name);
      return { sheets: ctx.wb.sheets.map((s) => s.name) };
    },
  },
  {
    name: 'rename_sheet',
    description: 'Renames a sheet. Formulas referencing it are updated.',
    inputSchema: { type: 'object', properties: { sheet: { type: 'string' }, new_name: { type: 'string' } }, required: ['sheet', 'new_name'] },
    run: (ctx, input) => {
      ctx.wb.apply({ type: 'renameSheet', from: str(input, 'sheet'), to: str(input, 'new_name') });
      return { sheets: ctx.wb.sheets.map((s) => s.name) };
    },
  },
  {
    name: 'delete_sheet',
    description: 'Deletes a sheet. References to it become #REF!.',
    inputSchema: { type: 'object', properties: { sheet: { type: 'string' } }, required: ['sheet'] },
    run: (ctx, input) => {
      ctx.wb.apply({ type: 'removeSheet', name: str(input, 'sheet') });
      return { sheets: ctx.wb.sheets.map((s) => s.name) };
    },
  },
  {
    name: 'activate_sheet',
    description: 'Switches the visible sheet in the user interface.',
    inputSchema: { type: 'object', properties: { sheet: { type: 'string' } }, required: ['sheet'] },
    run: (ctx, input) => {
      const sheet = sheetOf(ctx, input);
      activate(ctx, sheet.name);
      return { activeSheet: sheet.name };
    },
  },
  {
    name: 'select_range',
    description: 'Selects and scrolls to a range in the user interface, e.g. to show the user a result.',
    readOnly: true,
    inputSchema: { type: 'object', properties: { sheet: SHEET_PARAM, range: { type: 'string' } }, required: ['range'] },
    run: (ctx, input) => {
      const text = str(input, 'range');
      const sheet = sheetOf(ctx, input, text);
      if (!ctx.host?.select) return { ok: false, reason: 'No user interface attached.' };
      ctx.host.select(sheet.name, rangeOf(text));
      return { selected: formatRange(rangeOf(text), sheet.name) };
    },
  },
  {
    name: 'sort_range',
    description: 'Sorts the rows of a range by one or more columns.',
    inputSchema: {
      type: 'object',
      properties: {
        sheet: SHEET_PARAM,
        range: { type: 'string', description: 'Range including all columns that belong to a row, e.g. "A1:F200"' },
        by: {
          type: 'array',
          items: { type: 'object', properties: { column: { type: 'string', description: 'Column letter' }, descending: { type: 'boolean' } }, required: ['column'] },
        },
        has_header: { type: 'boolean', description: 'Keep the first row in place', default: true },
      },
      required: ['range', 'by'],
    },
    run: (ctx, input) => {
      const text = str(input, 'range');
      const sheet = sheetOf(ctx, input, text);
      const by = input.by as { column: string; descending?: boolean }[];
      if (!Array.isArray(by) || !by.length) throw new ToolInputError('"by" must list at least one column.');
      sortRange(ctx.wb, sheet.name, rangeOf(text), by.map((k) => ({ col: columnIndex(k.column), descending: !!k.descending })), { hasHeader: input.has_header !== false });
      return { sorted: formatRange(rangeOf(text), sheet.name) };
    },
  },
  {
    name: 'copy_range',
    description: 'Copies a range (values, formulas and formats) to another location. Relative formula references are adjusted.',
    inputSchema: {
      type: 'object',
      properties: {
        sheet: SHEET_PARAM,
        source: { type: 'string' },
        destination: { type: 'string', description: 'Top-left target cell, optionally sheet-qualified ("Sheet2!A1")' },
        mode: { type: 'string', enum: ['all', 'values', 'formats', 'formulas'], default: 'all' },
        move: { type: 'boolean', description: 'Cut instead of copy', default: false },
        transpose: { type: 'boolean', description: 'Swap rows and columns', default: false },
        skip_blanks: { type: 'boolean', description: 'Keep destination cells where the source is empty', default: false },
      },
      required: ['source', 'destination'],
    },
    run: (ctx, input) => {
      const srcText = str(input, 'source');
      const src = sheetOf(ctx, input, srcText);
      const dstText = str(input, 'destination');
      const dst = parseRange(dstText)?.sheet ? sheetOf(ctx, {}, dstText) : src;
      const clipData = copyRange(ctx.wb, src.name, rangeOf(srcText));
      if (input.move) clipData.cut = true;
      const start = cellOf(dstText);
      const mode = (optStr(input, 'mode') ?? 'all') as 'all';
      const pasted = pasteRange(ctx.wb, clipData, dst.name, { start, end: start }, { mode, transpose: !!input.transpose, skipBlanks: !!input.skip_blanks });
      return { pasted: formatRange(pasted, dst.name) };
    },
  },
  {
    name: 'find',
    description: 'Searches cell contents (displayed values or formulas) and returns matching cells.',
    readOnly: true,
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string' },
        sheet: { type: 'string', description: 'Restrict to one sheet' },
        regex: { type: 'boolean', default: false },
        in_formulas: { type: 'boolean', default: false },
        whole_cell: { type: 'boolean', default: false },
        case_sensitive: { type: 'boolean', default: false },
        limit: { type: 'integer', default: 100 },
      },
      required: ['query'],
    },
    run: (ctx, input) => {
      const opts: Parameters<typeof findCells>[2] = { limit: optNum(input, 'limit') ?? 100 };
      const sheet = optStr(input, 'sheet');
      if (sheet) opts.sheet = sheet;
      if (input.regex) opts.regex = true;
      if (input.in_formulas) opts.inFormulas = true;
      if (input.whole_cell) opts.wholeCell = true;
      if (input.case_sensitive) opts.caseSensitive = true;
      const results = findCells(ctx.wb, str(input, 'query'), opts);
      return { count: results.length, matches: results.map((r) => ({ sheet: r.sheet, cell: r.address, value: jsonValue(r.value), display: r.display, ...(r.formula ? { formula: r.formula } : {}) })) };
    },
  },
  {
    name: 'replace',
    description: 'Finds and replaces text like Excel\'s Replace dialog. Constants are changed in their content; formulas only with in_formulas. All changes are one undo step.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string' },
        replacement: { type: 'string', description: 'Replacement text; with regex, $1 etc. refer to groups' },
        sheet: { type: 'string', description: 'Restrict to one sheet' },
        range: { type: 'string', description: 'Restrict to a range of the sheet(s), e.g. "A1:D20"' },
        regex: { type: 'boolean', default: false },
        in_formulas: { type: 'boolean', default: false },
        whole_cell: { type: 'boolean', default: false },
        case_sensitive: { type: 'boolean', default: false },
      },
      required: ['query', 'replacement'],
    },
    run: (ctx, input) => {
      const sheet = optStr(input, 'sheet');
      const range = optStr(input, 'range');
      const count = replaceCells(ctx.wb, str(input, 'query'), str(input, 'replacement'), {
        ...(sheet ? { sheet } : {}),
        ...(range ? { range: rangeOf(range) } : {}),
        ...(input.regex ? { regex: true } : {}),
        ...(input.in_formulas ? { inFormulas: true } : {}),
        ...(input.whole_cell ? { wholeCell: true } : {}),
        ...(input.case_sensitive ? { caseSensitive: true } : {}),
      });
      return { replaced: count };
    },
  },
  {
    name: 'evaluate_formula',
    description: 'Evaluates a formula without writing it to the sheet, e.g. "=AVERAGE(B2:B50)". Useful for checking results or exploring data.',
    readOnly: true,
    inputSchema: { type: 'object', properties: { formula: { type: 'string' }, sheet: SHEET_PARAM }, required: ['formula'] },
    run: (ctx, input) => {
      const sheet = sheetOf(ctx, input);
      const v = ctx.wb.evaluate(str(input, 'formula'), sheet.name);
      return { value: jsonValue(v), display: formatValue(v) };
    },
  },
  {
    name: 'inspect_cell',
    description: 'Detailed information about one cell: value, displayed text, formula, style, cells it references (precedents) and cells that depend on it.',
    readOnly: true,
    inputSchema: { type: 'object', properties: { sheet: SHEET_PARAM, cell: { type: 'string' } }, required: ['cell'] },
    run: (ctx, input) => {
      const text = str(input, 'cell');
      const sheet = sheetOf(ctx, input, text);
      const { row, col } = cellOf(text);
      const cell = sheet.getCell(row, col);
      const effectiveStyle = ctx.wb.getCellStyle(sheet.name, row, col);
      const precedents: string[] = [];
      if (cell?.formula) {
        try {
          walk(parseFormula(cell.formula), (n) => {
            if (n.type === 'cell') precedents.push(formatRange({ start: n.ref, end: n.ref }, n.sheet));
            if (n.type === 'range') precedents.push(formatRange({ start: n.start, end: n.end }, n.sheet));
          });
        } catch {
          /* invalid formula */
        }
      }
      const dependents = ctx.wb.calc.dependentsOf(sheet, row, col).map((d) => {
        const s = ctx.wb.getSheetById(d.sheetId);
        return formatRange({ start: d, end: d }, s && s !== sheet ? s.name : undefined);
      });
      return {
        cell: formatAddress(row, col),
        sheet: sheet.name,
        value: jsonValue(cell?.value ?? null),
        display: cell ? formatValue(cell.value, effectiveStyle?.numFmt) : '',
        formula: cell?.formula !== undefined ? `=${cell.formula}` : null,
        style: cell?.style ?? null,
        effectiveStyle: effectiveStyle ?? null,
        spillRange: cell?.spillRange ? formatRange(cell.spillRange) : null,
        spillParent: cell?.spillParent !== undefined ? (() => { const p = keyToAddress(cell.spillParent); return formatAddress(p.row, p.col); })() : null,
        precedents,
        dependents,
        ...(sheet.notes.has(cellKey(row, col)) ? { note: sheet.notes.get(cellKey(row, col)) } : {}),
        ...(sheet.links.has(cellKey(row, col)) ? { link: sheet.links.get(cellKey(row, col)) } : {}),
        ...(validationAt(sheet, row, col) ? { validation: validationAt(sheet, row, col) } : {}),
      };
    },
  },
  {
    name: 'create_chart',
    description:
      'Creates a chart from a data range. A text header row becomes series names and a leading text column becomes category labels (like Excel). For scatter charts the first column holds the x values.',
    inputSchema: {
      type: 'object',
      properties: {
        sheet: SHEET_PARAM,
        type: { type: 'string', enum: CHART_TYPES },
        data_range: { type: 'string', description: 'e.g. "A1:C13" including headers' },
        title: { type: 'string' },
        position: { type: 'string', description: 'Top-left cell of the chart, e.g. "E2". Defaults to right of the data.' },
        width_columns: { type: 'integer', default: 8 },
        height_rows: { type: 'integer', default: 16 },
        series_in_rows: { type: 'boolean', default: false },
        stacked: { type: 'boolean', default: false },
        legend: { type: 'string', enum: ['top', 'bottom', 'left', 'right', 'none'] },
        x_axis_title: { type: 'string' },
        y_axis_title: { type: 'string' },
      },
      required: ['type', 'data_range'],
    },
    run: (ctx, input) => {
      const text = str(input, 'data_range');
      const sheet = sheetOf(ctx, input, text);
      const type = str(input, 'type') as ChartType;
      if (!CHART_TYPES.includes(type)) throw new ToolInputError(`Unknown chart type "${type}". Use one of ${CHART_TYPES.join(', ')}.`);
      const range = clip(sheet, rangeOf(text));
      const opts: Parameters<typeof insertChartFromRange>[4] = {};
      const title = optStr(input, 'title');
      if (title) opts.title = title;
      if (input.series_in_rows) opts.seriesInRows = true;
      const pos = optStr(input, 'position');
      if (pos) {
        const from = cellOf(pos);
        opts.anchor = { from, to: { row: from.row + (optNum(input, 'height_rows') ?? 16), col: from.col + (optNum(input, 'width_columns') ?? 8) } };
      }
      return ctx.wb.transaction(() => {
        const spec = insertChartFromRange(ctx.wb, sheet.name, range, type, opts);
        const patch = chartPatch(input);
        const final = Object.keys(patch).length ? updateChart(ctx.wb, spec.id, patch) : spec;
        return chartSummary(ctx.wb, sheet.name, final);
      });
    },
  },
  {
    name: 'update_chart',
    description: 'Changes an existing chart (type, title, data, position, legend, stacking).',
    inputSchema: {
      type: 'object',
      properties: {
        chart_id: { type: 'string' },
        type: { type: 'string', enum: CHART_TYPES },
        title: { type: 'string' },
        data_range: { type: 'string', description: 'New data range; series are rebuilt from it' },
        position: { type: 'string', description: 'New top-left cell' },
        width_columns: { type: 'integer' },
        height_rows: { type: 'integer' },
        stacked: { type: 'boolean' },
        legend: { type: 'string', enum: ['top', 'bottom', 'left', 'right', 'none'] },
        x_axis_title: { type: 'string' },
        y_axis_title: { type: 'string' },
      },
      required: ['chart_id'],
    },
    run: (ctx, input) => {
      const id = str(input, 'chart_id');
      const found = ctx.wb.getChart(id);
      if (!found) throw new ToolInputError(`Chart "${id}" not found. Use list_charts to get ids.`);
      const patch = chartPatch(input);
      const dataRange = optStr(input, 'data_range');
      if (dataRange) {
        const s = sheetOf(ctx, { sheet: found.sheet.name }, dataRange);
        patch.series = chartFromRange(ctx.wb, s.name, clip(s, rangeOf(dataRange)), (patch.type ?? found.chart.type) as ChartType).series;
      }
      const pos = optStr(input, 'position');
      const w = optNum(input, 'width_columns') ?? found.chart.anchor.to.col - found.chart.anchor.from.col;
      const h = optNum(input, 'height_rows') ?? found.chart.anchor.to.row - found.chart.anchor.from.row;
      const from = pos ? cellOf(pos) : found.chart.anchor.from;
      patch.anchor = { from, to: { row: from.row + h, col: from.col + w } };
      return chartSummary(ctx.wb, found.sheet.name, updateChart(ctx.wb, id, patch));
    },
  },
  {
    name: 'delete_chart',
    description: 'Deletes a chart.',
    inputSchema: { type: 'object', properties: { chart_id: { type: 'string' } }, required: ['chart_id'] },
    run: (ctx, input) => {
      removeChart(ctx.wb, str(input, 'chart_id'));
      return { ok: true };
    },
  },
  {
    name: 'list_charts',
    description: 'Lists charts with their ids, data references and a preview of the plotted data.',
    readOnly: true,
    inputSchema: { type: 'object', properties: { sheet: { type: 'string', description: 'Restrict to one sheet' } } },
    run: (ctx, input) => {
      const only = optStr(input, 'sheet')?.toUpperCase();
      return {
        charts: ctx.wb.sheets
          .filter((s) => !only || s.name.toUpperCase() === only)
          .flatMap((s) => s.charts.map((c) => chartSummary(ctx.wb, s.name, c))),
      };
    },
  },
  {
    name: 'set_column_width',
    description: 'Sets the width of one or more columns in pixels (default width is 64). Use null to reset.',
    inputSchema: {
      type: 'object',
      properties: { sheet: SHEET_PARAM, columns: { type: 'string', description: 'e.g. "B" or "A:D"' }, width: { type: ['number', 'null'] } },
      required: ['columns', 'width'],
    },
    run: (ctx, input) => {
      const sheet = sheetOf(ctx, input);
      const spec = str(input, 'columns');
      const [a, b = a] = spec.split(':') as [string, string?];
      const from = columnIndex(a);
      const to = columnIndex(b!);
      const width = input.width === null ? null : optNum(input, 'width') ?? null;
      const sizes = [];
      for (let c = Math.min(from, to); c <= Math.max(from, to); c++) sizes.push({ index: c, size: width });
      ctx.wb.apply({ type: 'setColWidths', sheet: sheet.name, sizes });
      return { ok: true };
    },
  },
  {
    name: 'set_row_height',
    description: 'Sets the height of one or more rows in pixels (default 20). Use null to reset.',
    inputSchema: {
      type: 'object',
      properties: { sheet: SHEET_PARAM, rows: { type: 'string', description: 'e.g. "1" or "2:10"' }, height: { type: ['number', 'null'] } },
      required: ['rows', 'height'],
    },
    run: (ctx, input) => {
      const sheet = sheetOf(ctx, input);
      const [a, b = a] = str(input, 'rows').split(':') as [string, string?];
      const from = rowIndex(a);
      const to = rowIndex(b);
      const height = input.height === null ? null : optNum(input, 'height') ?? null;
      const sizes = [];
      for (let r = Math.min(from, to); r <= Math.max(from, to); r++) sizes.push({ index: r, size: height });
      ctx.wb.apply({ type: 'setRowHeights', sheet: sheet.name, sizes });
      return { ok: true };
    },
  },
  {
    name: 'set_hidden',
    description: 'Hides or shows rows or columns, e.g. rows "2:5" or columns "C:D". Hidden rows are skipped by SUBTOTAL(101-111).',
    inputSchema: {
      type: 'object',
      properties: {
        sheet: SHEET_PARAM,
        rows: { type: 'string', description: 'e.g. "3" or "2:10"' },
        columns: { type: 'string', description: 'e.g. "B" or "C:F"' },
        hidden: { type: 'boolean', default: true },
      },
      required: ['hidden'],
    },
    run: (ctx, input) => {
      const sheet = sheetOf(ctx, input);
      const hidden = input.hidden !== false;
      const rows = optStr(input, 'rows');
      const cols = optStr(input, 'columns');
      if (!rows && !cols) throw new ToolInputError('Pass "rows" and/or "columns".');
      ctx.wb.transaction(() => {
        if (rows) {
          const [a, b = a] = rows.split(':') as [string, string?];
          const [from, to] = [rowIndex(a), rowIndex(b)].sort((x, y) => x - y) as [number, number];
          (hidden ? hideRows : unhideRows)(ctx.wb, sheet.name, from, to);
        }
        if (cols) {
          const [a, b = a] = cols.split(':') as [string, string?];
          const [from, to] = [columnIndex(a), columnIndex(b!)].sort((x, y) => x - y) as [number, number];
          (hidden ? hideCols : unhideCols)(ctx.wb, sheet.name, from, to);
        }
      });
      return {
        hidden_rows: [...sheet.rowHeights].filter(([, h]) => h === 0).map(([r]) => r + 1).sort((x, y) => x - y),
        hidden_columns: [...sheet.colWidths].filter(([, w]) => w === 0).map(([c]) => colToName(c)).sort(),
      };
    },
  },
  {
    name: 'set_data_validation',
    description: 'Sets or removes data validation for a range, e.g. a dropdown list (type "list" with values or a source range) or number limits. The range is cut out of existing rules first.',
    inputSchema: {
      type: 'object',
      properties: {
        sheet: SHEET_PARAM,
        range: { type: 'string' },
        type: { type: 'string', enum: [...VALIDATION_TYPES, 'none'], description: '"none" removes validation from the range' },
        operator: { type: 'string', enum: [...VALIDATION_OPERATORS], default: 'between' },
        values: { type: 'array', items: { type: ['string', 'number'] }, description: 'Allowed values for a list' },
        formula1: { type: 'string', description: 'Value, source range ("=$D$1:$D$9") or formula; relative references are relative to the top-left cell' },
        formula2: { type: 'string', description: 'Maximum for between/notBetween' },
        allow_blank: { type: 'boolean', default: true },
        dropdown: { type: 'boolean', default: true },
        error_style: { type: 'string', enum: ['stop', 'warning', 'information'], default: 'stop' },
        error_message: { type: 'string' },
        input_message: { type: 'string' },
      },
      required: ['range', 'type'],
    },
    run: (ctx, input) => {
      const text = str(input, 'range');
      const sheet = sheetOf(ctx, input, text);
      const type = str(input, 'type');
      if (type === 'none') {
        applyDataValidation(ctx.wb, sheet.name, rangeOf(text), null);
        return { rules: sheet.validations.length };
      }
      const strip = (f: string) => f.trim().replace(/^=/, '');
      const values = Array.isArray(input.values) ? (input.values as unknown[]).map(String) : null;
      const formula1 = values ? `"${values.join(',').replace(/"/g, '""')}"` : optStr(input, 'formula1');
      if (!formula1) throw new ToolInputError('Pass "values" (lists) or "formula1".');
      const rule: Omit<DataValidation, 'range'> = { type: type as ValidationType, formula1: values ? formula1 : strip(formula1) };
      if (type !== 'list' && type !== 'custom') rule.operator = (optStr(input, 'operator') ?? 'between') as ValidationOperator;
      const formula2 = optStr(input, 'formula2');
      if (formula2) rule.formula2 = strip(formula2);
      if (input.allow_blank === false) rule.allowBlank = false;
      if (input.dropdown === false) rule.dropdown = false;
      const style = optStr(input, 'error_style');
      if (style === 'warning' || style === 'information') rule.errorStyle = style;
      const error = optStr(input, 'error_message');
      if (error) rule.error = error;
      const prompt = optStr(input, 'input_message');
      if (prompt) rule.prompt = prompt;
      applyDataValidation(ctx.wb, sheet.name, rangeOf(text), rule);
      return { rules: sheet.validations.length, invalid_cells: invalidCells(ctx.wb, sheet.name, 20).map((c) => formatAddress(c.row, c.col)) };
    },
  },
  {
    name: 'list_data_validations',
    description: 'Lists the data validation rules of a sheet and cells whose current values violate them.',
    readOnly: true,
    inputSchema: { type: 'object', properties: { sheet: SHEET_PARAM } },
    run: (ctx, input) => {
      const sheet = sheetOf(ctx, input);
      return { rules: sheet.validations, invalid_cells: invalidCells(ctx.wb, sheet.name, 100).map((c) => formatAddress(c.row, c.col)) };
    },
  },
  {
    name: 'merge_cells',
    description: 'Merges a range into one cell (only the top-left value is kept), or unmerges it.',
    inputSchema: { type: 'object', properties: { sheet: SHEET_PARAM, range: { type: 'string' }, unmerge: { type: 'boolean', default: false } }, required: ['range'] },
    run: (ctx, input) => {
      const text = str(input, 'range');
      const sheet = sheetOf(ctx, input, text);
      if (input.unmerge) unmergeCells(ctx.wb, sheet.name, rangeOf(text));
      else mergeCells(ctx.wb, sheet.name, rangeOf(text));
      return { merges: sheet.merges.map((m) => formatRange(m)) };
    },
  },
  {
    name: 'freeze_panes',
    description: 'Freezes the top rows and/or left columns so they stay visible while scrolling. Use 0 to unfreeze.',
    inputSchema: { type: 'object', properties: { sheet: SHEET_PARAM, rows: { type: 'integer', minimum: 0 }, columns: { type: 'integer', minimum: 0 } } },
    run: (ctx, input) => {
      const sheet = sheetOf(ctx, input);
      ctx.wb.apply({ type: 'setFrozen', sheet: sheet.name, rows: optNum(input, 'rows') ?? 0, cols: optNum(input, 'columns') ?? 0 });
      return { frozen: sheet.frozen };
    },
  },
  {
    name: 'define_name',
    description: 'Creates or deletes a named range usable in formulas, e.g. name "TaxRate" ref "Sheet1!$B$1".',
    inputSchema: { type: 'object', properties: { name: { type: 'string' }, ref: { type: ['string', 'null'] } }, required: ['name', 'ref'] },
    run: (ctx, input) => {
      const name = str(input, 'name');
      if (!/^[A-Za-z_][\w.]*$/.test(name) || parseRange(name)) throw new ToolInputError(`Invalid name "${name}".`);
      ctx.wb.apply({ type: 'setName', name, ref: optStr(input, 'ref') ?? null });
      return { names: Object.fromEntries([...ctx.wb.names.values()].map((n) => [n.name, n.ref])) };
    },
  },
  {
    name: 'list_functions',
    description: 'Lists supported spreadsheet functions with syntax, optionally filtered by category or name.',
    readOnly: true,
    inputSchema: {
      type: 'object',
      properties: {
        category: { type: 'string', enum: ['math', 'statistical', 'financial', 'logical', 'lookup', 'text', 'date', 'info'] },
        search: { type: 'string' },
      },
    },
    run: (_ctx, input) => {
      const cat = optStr(input, 'category');
      const q = optStr(input, 'search')?.toUpperCase();
      return {
        functions: listFunctions()
          .filter((f) => (!cat || f.category === cat) && (!q || f.name.includes(q)))
          .map((f) => ({ name: f.name, syntax: f.syntax, description: f.description })),
      };
    },
  },
  {
    name: 'undo',
    description: 'Undoes the last change (including changes made by the user).',
    inputSchema: { type: 'object', properties: {} },
    run: (ctx) => ({ undone: ctx.wb.undo() }),
  },
  {
    name: 'redo',
    description: 'Redoes the last undone change.',
    inputSchema: { type: 'object', properties: {} },
    run: (ctx) => ({ redone: ctx.wb.redo() }),
  },
];

function activate(ctx: ToolContext, name: string): void {
  if (ctx.host?.activateSheet) ctx.host.activateSheet(name);
  else ctx.wb.activeSheetIndex = ctx.wb.sheets.findIndex((s) => s.name === name);
}

function chartPatch(input: Input): Partial<Omit<ChartSpec, 'id'>> {
  const patch: Partial<Omit<ChartSpec, 'id'>> = {};
  const type = optStr(input, 'type');
  if (type) patch.type = type as ChartType;
  const title = optStr(input, 'title');
  if (title !== undefined) patch.title = title;
  if (typeof input.stacked === 'boolean') patch.stacked = input.stacked;
  const legend = optStr(input, 'legend');
  if (legend) patch.legend = legend as LegendPosition;
  const xt = optStr(input, 'x_axis_title');
  if (xt) patch.xAxisTitle = xt;
  const yt = optStr(input, 'y_axis_title');
  if (yt) patch.yAxisTitle = yt;
  return patch;
}

/** Plain tool definitions (without implementations), e.g. for manifests. */
export const TOOL_DEFINITIONS: ToolDefinition[] = TOOLS.map(({ name, description, inputSchema, readOnly }) =>
  readOnly ? { name, description, inputSchema, readOnly } : { name, description, inputSchema },
);
