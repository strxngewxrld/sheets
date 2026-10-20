import { numberFormatCategory } from './format-names.js';
import {
  applyBorders,
  applyStyle,
  cellKey,
  clearRange,
  copyRange,
  ConditionalFormatEvaluator,
  evaluate,
  toText,
  deleteCols,
  deleteRows,
  Emitter,
  fillRange,
  formatAddress,
  formatRange,
  formatValue,
  getDefaultLocale,
  hasHidden,
  hideCols,
  hideRows,
  insertChartFromRange,
  insertCols,
  insertRows,
  inputToCellData,
  isDateFormat,
  isValidValue,
  scalarFromJson,
  validationAt,
  validationListValues,
  isError,
  keyToAddress,
  parseFormula,
  walk,
  MAX_COLS,
  MAX_ROWS,
  mergeCells,
  normalizeRange,
  parseCsv,
  parseInput,
  parseRange,
  pasteRange,
  rangeToTSV,
  refreshPivotTable,
  removePivotTable,
  setCellInput,
  setValues,
  sortRange,
  tableAt,
  setTableFilter,
  unhideCols,
  unhideRows,
  unmergeCells,
  CHART_TYPES,
  LOCALE_DE,
  type CellAddress,
  type ChangeEvent,
  type ClipboardData,
  type FormatLocale,
  type PasteSpecial,
  type RangeAddress,
  type Scalar,
  type Sheet,
  type StylePatch,
  type Workbook,
} from '@sheets/core';
import { ChartLayer } from './chart-layer.js';
import { closeAnalysisDialog, openCellFormatDialog, openConditionalDialog, openPasteSpecialDialog, openPivotDialog, openSortDialog, openTableFormatDialog, openTableFilterDialog, openValidationDialog } from './analysis-dialogs.js';
import { currentRegion, resolveSortContext } from './sheet-actions.js';
export { currentRegion } from './sheet-actions.js';
import type { ChartTheme } from './charts.js';
import { CellEditor, formulaHighlights } from './editor.js';
import { openFindDialog, type FindDialog } from './find-dialog.js';
import { Geometry } from './geometry.js';
import { DEFAULT_CELL_FONT, GridRenderer, Viewport, type GridTheme } from './grid.js';
import { getStrings, type Strings } from './i18n.js';
import { ICONS } from './icons.js';
import { closeMenu, showMenu, type MenuEntry } from './menu.js';
import { injectStyles } from './styles.js';

export interface ToolbarAction {
  id: string;
  label: string;
  title?: string;
  /** Inline SVG markup. */
  icon?: string;
  onClick: () => void;
}

export interface SpreadsheetViewOptions {
  /** Number/date display locale. Defaults to the core default locale. */
  locale?: FormatLocale;
  /** UI language ("de", "en"). Defaults to the browser language. */
  language?: string;
  readOnly?: boolean;
  /** Handle Ctrl+Z / Ctrl+Y inside the view (disable when the host handles undo, e.g. VS Code). */
  keyboardUndo?: boolean;
  toolbar?: boolean;
  formulaBar?: boolean;
  sheetTabs?: boolean;
  statusBar?: boolean;
  /** Extra buttons at the start of the toolbar (e.g. Open / Save). */
  toolbarActions?: ToolbarAction[];
  theme?: 'auto' | 'light' | 'dark';
  /** Opens an external hyperlink (default: a new browser window). */
  onOpenLink?: (url: string) => void;
}

export interface SelectionInfo {
  sheet: string;
  range: RangeAddress;
  active: CellAddress;
}

/** Commands a host (e.g. VS Code keybindings) can trigger in the view. */
export type ViewCommand = 'find' | 'replace' | 'hideRows' | 'unhideRows' | 'hideCols' | 'unhideCols' | 'goTo' | 'zoomIn' | 'zoomOut' | 'zoomReset';

export type ViewEvents = {
  selectionChange: SelectionInfo;
  sheetChange: { sheet: string };
};

const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36];
const ZOOM_STEPS = [0.25, 0.5, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4];
export const MIN_ZOOM = ZOOM_STEPS[0]!;
export const MAX_ZOOM = ZOOM_STEPS.at(-1)!;

type DragMode =
  | { kind: 'select' }
  | { kind: 'cols'; anchor: number }
  | { kind: 'rows'; anchor: number }
  | { kind: 'resize-col'; index: number; start: number; size: number }
  | { kind: 'resize-row'; index: number; start: number; size: number }
  | { kind: 'fill'; source: RangeAddress; target: RangeAddress | null }
  | { kind: 'ref'; anchor: CellAddress };

export class SpreadsheetView {
  readonly root: HTMLDivElement;
  readonly events = new Emitter<ViewEvents>();
  readonly strings: Strings;
  private wb: Workbook;
  private options: SpreadsheetViewOptions;
  private locale: FormatLocale;

  private gridEl!: HTMLDivElement;
  private canvas!: HTMLCanvasElement;
  private scroller!: HTMLDivElement;
  private spacer!: HTMLDivElement;
  private renderer!: GridRenderer;
  private editor!: CellEditor;
  private charts!: ChartLayer;
  private nameBox?: HTMLInputElement;
  private formulaInput?: HTMLInputElement;
  private tabsEl?: HTMLDivElement;
  private statusEl?: HTMLDivElement;
  private toolbarEl?: HTMLDivElement;
  private tableControls!: HTMLDivElement;
  private controls: Record<string, HTMLElement> = {};

  private vp = new Viewport();
  private geoSheet: Sheet | null = null;
  private geoDirty = true;
  private extentRows = 200;
  private extentCols = 40;

  private anchor: CellAddress = { row: 0, col: 0 };
  private focusCell: CellAddress = { row: 0, col: 0 };
  private active: CellAddress = { row: 0, col: 0 };
  private clip: ClipboardData | null = null;
  private clipText = '';
  private drag: DragMode | null = null;
  private frame = 0;
  private chartsDirty = true;
  private theme!: GridTheme;
  private disposers: (() => void)[] = [];
  private wbOff: (() => void) | null = null;
  private findDialog: FindDialog | null = null;
  private zoom = 1;
  private zoomLabel?: HTMLButtonElement;
  private tip!: HTMLDivElement;
  /** An alert from data validation is open (it stays until answered). */
  private alertOpen = false;
  private validationButton!: HTMLButtonElement;
  private noteEl!: HTMLDivElement;
  private noticeEl!: HTMLDivElement;
  private skipValidation = false;

  constructor(container: HTMLElement, workbook: Workbook, options: SpreadsheetViewOptions = {}) {
    injectStyles(container.ownerDocument);
    this.wb = workbook;
    this.options = { keyboardUndo: true, toolbar: true, formulaBar: true, sheetTabs: true, statusBar: true, ...options };
    this.locale = options.locale ?? getDefaultLocale();
    this.strings = getStrings(options.language);
    this.root = document.createElement('div');
    this.root.className = 'sx-root';
    if (options.theme && options.theme !== 'auto') this.root.dataset.theme = options.theme;
    container.appendChild(this.root);
    this.build();
    this.bindWorkbook();
    this.readTheme();

    const ro = new ResizeObserver(() => this.scheduleRender(true));
    ro.observe(this.gridEl);
    this.disposers.push(() => ro.disconnect());
    const mq = window.matchMedia?.('(prefers-color-scheme: dark)');
    const onScheme = () => {
      this.readTheme();
      this.scheduleRender(true);
    };
    mq?.addEventListener('change', onScheme);
    this.disposers.push(() => mq?.removeEventListener('change', onScheme));
    // VS Code switches themes by changing body classes / CSS variables.
    const mo = new MutationObserver(onScheme);
    mo.observe(document.body, { attributes: true, attributeFilter: ['class', 'style', 'data-vscode-theme-kind'] });
    this.disposers.push(() => mo.disconnect());
    this.scheduleRender(true);
  }

  // =========================================================================
  // Public API
  // =========================================================================

  get workbook(): Workbook {
    return this.wb;
  }

  get sheet(): Sheet {
    return this.wb.activeSheet;
  }

  /** Opens the find panel (or the find and replace panel) for the selection's text. */
  openFind(replace = false): void {
    if (this.editor.editing) this.commitEdit();
    closeMenu();
    closeAnalysisDialog(this.root);
    this.charts.closeEditor();
    if (this.findDialog?.element.isConnected) {
      this.findDialog.setReplace(replace);
      return;
    }
    this.findDialog = openFindDialog(this.root, {
      workbook: () => this.wb,
      position: () => ({ sheet: this.sheet.name, ...this.active }),
      reveal: (m) => {
        if (m.sheet.toUpperCase() !== this.sheet.name.toUpperCase()) this.activateSheet(m.sheet);
        const cell = { row: m.row, col: m.col };
        this.setSelection(cell, cell);
        this.scrollIntoView(m.row, m.col);
      },
      readOnly: !!this.options.readOnly,
      onClose: () => {
        this.findDialog = null;
        this.focus();
      },
    }, this.strings, replace);
  }

  /** Current zoom factor (1 = 100 %). */
  get zoomLevel(): number {
    return this.zoom;
  }

  /** Sets the zoom factor (0.25-4), keeping the top-left cell in view. */
  setZoom(zoom: number): void {
    const next = Math.round(Math.max(MIN_ZOOM, Math.min(MAX_ZOOM, zoom)) * 100) / 100;
    if (next === this.zoom) return;
    const ratio = next / this.zoom;
    const { scrollLeft, scrollTop } = this.scroller;
    this.zoom = next;
    this.geoDirty = true;
    this.chartsDirty = true;
    this.ensureGeometry();
    this.render();
    this.scroller.scrollLeft = scrollLeft * ratio;
    this.scroller.scrollTop = scrollTop * ratio;
    this.scheduleRender();
  }

  private stepZoom(direction: 1 | -1): void {
    const next = direction > 0 ? ZOOM_STEPS.find((z) => z > this.zoom + 1e-9) : [...ZOOM_STEPS].reverse().find((z) => z < this.zoom - 1e-9);
    if (next !== undefined) this.setZoom(next);
  }

  /** Focuses the name box to jump to a cell, range or defined name (Ctrl+G). */
  openGoTo(): void {
    if (this.editor.editing) this.commitEdit();
    if (this.nameBox) {
      this.nameBox.focus();
      this.nameBox.select();
      return;
    }
    const target = prompt(this.strings.goTo, formatRange(this.selectionRange()));
    if (target) this.goTo(target);
  }

  private goTo(text: string): void {
    const named = this.wb.names.get(text.trim().toUpperCase());
    this.select(named ? named.ref : text.trim());
  }

  /**
   * Shows a warning above the grid, e.g. that the current changes cannot be
   * saved without loss (`detail` is the technical reason); null hides it.
   */
  setNotice(detail: string | null): void {
    if (!detail) {
      this.noticeEl.hidden = true;
      return;
    }
    const strong = document.createElement('strong');
    strong.textContent = `⚠ ${this.strings.saveBlocked}`;
    const reason = document.createElement('span');
    reason.textContent = detail;
    this.noticeEl.replaceChildren(strong, reason);
    this.noticeEl.hidden = false;
    this.scheduleRender(true);
  }

  /** Runs a host command such as find or replace. */
  runCommand(command: ViewCommand): void {
    switch (command) {
      case 'goTo': this.openGoTo(); break;
      case 'zoomIn': this.stepZoom(1); break;
      case 'zoomOut': this.stepZoom(-1); break;
      case 'zoomReset': this.setZoom(1); break;
      case 'find': this.openFind(false); break;
      case 'replace': this.openFind(true); break;
      case 'hideRows': case 'unhideRows': case 'hideCols': case 'unhideCols':
        if (this.options.readOnly || this.editor.editing) return;
        this.setHidden(command.endsWith('Rows') ? 'row' : 'col', command.startsWith('hide'));
        break;
    }
  }

  /** Replaces the displayed workbook (e.g. after opening a file). */
  setWorkbook(wb: Workbook): void {
    closeAnalysisDialog(this.root);
    this.findDialog?.close();
    this.wbOff?.();
    this.wb = wb;
    this.bindWorkbook();
    this.clip = null;
    this.setSelection({ row: 0, col: 0 }, { row: 0, col: 0 });
    this.geoDirty = true;
    this.renderTabs();
    this.scheduleRender(true);
  }

  activateSheet(nameOrIndex: string | number): void {
    closeAnalysisDialog(this.root);
    const sheet = this.wb.getSheet(nameOrIndex);
    if (!sheet) return;
    if (this.editor.editing && !this.commitEdit()) return;
    this.wb.activeSheetIndex = this.wb.sheets.indexOf(sheet);
    this.geoDirty = true;
    this.extentRows = 200;
    this.extentCols = 40;
    this.charts.select(null);
    this.charts.closeEditor();
    this.setSelection({ row: 0, col: 0 }, { row: 0, col: 0 });
    this.scroller.scrollTo(0, 0);
    this.renderTabs();
    this.events.emit('sheetChange', { sheet: sheet.name });
    this.scheduleRender(true);
  }

  getSelection(): SelectionInfo {
    return { sheet: this.sheet.name, range: normalizeRange(this.anchor, this.focusCell), active: { ...this.active } };
  }

  /** Selects a range (A1 string or address), optionally switching sheets, and scrolls to it. */
  select(range: RangeAddress | string, sheet?: string): void {
    let r: RangeAddress;
    let target = sheet;
    if (typeof range === 'string') {
      const p = parseRange(range);
      if (!p) return;
      r = p.range;
      target = p.sheet ?? sheet;
    } else r = range;
    if (target && target.toUpperCase() !== this.sheet.name.toUpperCase()) this.activateSheet(target);
    const used = this.sheet.usedRange();
    const end = {
      row: r.end.row === MAX_ROWS - 1 ? Math.max(r.start.row, used?.end.row ?? 0) : r.end.row,
      col: r.end.col === MAX_COLS - 1 ? Math.max(r.start.col, used?.end.col ?? 0) : r.end.col,
    };
    this.setSelection(r.start, end, r.start);
    this.scrollIntoView(r.start.row, r.start.col);
  }

  focus(): void {
    this.editor.textarea.focus({ preventScroll: true });
  }

  /** Adapter for @sheets/agent: lets agent tools read and drive the UI selection. */
  get agentHost() {
    return {
      getSelection: () => {
        const s = this.getSelection();
        return { sheet: s.sheet, range: formatRange(s.range) };
      },
      select: (sheet: string, range: RangeAddress) => this.select(range, sheet),
      activateSheet: (sheet: string) => this.activateSheet(sheet),
    };
  }

  destroy(): void {
    cancelAnimationFrame(this.frame);
    this.findDialog?.close();
    this.disposers.forEach((d) => d());
    this.wbOff?.();
    this.root.remove();
    closeMenu();
  }

  // =========================================================================
  // Construction
  // =========================================================================

  private bindWorkbook(): void {
    this.wbOff = this.wb.events.on('change', (e) => this.onWorkbookChange(e));
  }

  private onWorkbookChange(e: ChangeEvent): void {
    this.geoDirty = true;
    this.chartsDirty = true;
    if (this.wb.activeSheetIndex >= this.wb.sheets.length) this.wb.activeSheetIndex = this.wb.sheets.length - 1;
    if (e.structural) this.renderTabs();
    this.findDialog?.refresh();
    this.scheduleRender();
  }

  private build(): void {
    const s = this.strings;
    if (this.options.toolbar) this.root.appendChild(this.buildToolbar());
    if (this.options.formulaBar) {
      const bar = document.createElement('div');
      bar.className = 'sx-formulabar';
      this.nameBox = document.createElement('input');
      this.nameBox.className = 'sx-namebox';
      this.nameBox.setAttribute('aria-label', 'Name box');
      this.nameBox.spellcheck = false;
      const fx = document.createElement('div');
      fx.className = 'sx-fx';
      fx.textContent = 'fx';
      this.formulaInput = document.createElement('input');
      this.formulaInput.className = 'sx-formula';
      this.formulaInput.setAttribute('aria-label', 'Formula');
      this.formulaInput.spellcheck = false;
      bar.append(this.nameBox, fx, this.formulaInput);
      this.root.appendChild(bar);
      this.bindFormulaBar();
    }

    this.noticeEl = document.createElement('div');
    this.noticeEl.className = 'sx-notice';
    this.noticeEl.setAttribute('role', 'status');
    this.noticeEl.hidden = true;
    this.root.appendChild(this.noticeEl);

    this.gridEl = document.createElement('div');
    this.gridEl.className = 'sx-grid';
    this.canvas = document.createElement('canvas');
    this.scroller = document.createElement('div');
    this.scroller.className = 'sx-scroller';
    this.spacer = document.createElement('div');
    this.spacer.className = 'sx-spacer';
    this.scroller.appendChild(this.spacer);
    this.gridEl.append(this.canvas, this.scroller);
    this.root.appendChild(this.gridEl);
    this.renderer = new GridRenderer(this.canvas, this.readTheme());
    this.vp.sheet = this.sheet;
    this.vp.geo = new Geometry(this.sheet);
    this.charts = new ChartLayer(
      this.gridEl,
      () => this.wb,
      this.vp,
      () => this.chartTheme(),
      s,
      (e) => {
        e.preventDefault();
        this.scroller.scrollBy(e.deltaX, e.deltaY);
      },
      () => this.scheduleRender(),
    );
    this.editor = new CellEditor(this.gridEl);
    this.tip = document.createElement('div');
    this.tip.className = 'sx-validation-tip';
    this.tip.hidden = true;
    this.tip.addEventListener('mousedown', (e) => e.stopPropagation());
    this.tip.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Escape') {
        e.preventDefault();
        this.tip.querySelector('button')?.click();
      }
    });
    this.validationButton = document.createElement('button');
    this.validationButton.type = 'button';
    this.validationButton.className = 'sx-validation-btn';
    this.validationButton.textContent = '▾';
    this.validationButton.title = this.strings.chooseValue;
    this.validationButton.setAttribute('aria-label', this.strings.chooseValue);
    this.validationButton.setAttribute('aria-haspopup', 'menu');
    this.validationButton.hidden = true;
    this.validationButton.addEventListener('mousedown', (e) => { e.preventDefault(); e.stopPropagation(); });
    this.validationButton.addEventListener('click', () => this.openValidationList());
    this.noteEl = document.createElement('div');
    this.noteEl.className = 'sx-note';
    this.noteEl.setAttribute('role', 'tooltip');
    this.noteEl.hidden = true;
    this.gridEl.append(this.tip, this.validationButton, this.noteEl);
    this.bindGrid();
    this.bindKeyboard();

    if (this.options.sheetTabs || this.options.statusBar) {
      const bottom = document.createElement('div');
      bottom.className = 'sx-bottom';
      if (this.options.sheetTabs) {
        const add = document.createElement('button');
        add.className = 'sx-btn';
        add.title = s.addSheet;
        add.setAttribute('aria-label', s.addSheet);
        add.textContent = '+';
        add.addEventListener('click', () => this.mutate(() => this.activateSheet(this.wb.addSheet().name)));
        this.tabsEl = document.createElement('div');
        this.tabsEl.className = 'sx-tabs';
        this.tabsEl.setAttribute('role', 'tablist');
        bottom.append(add, this.tabsEl);
      }
      if (this.options.statusBar) {
        this.statusEl = document.createElement('div');
        this.statusEl.className = 'sx-status';
        this.statusEl.setAttribute('aria-live', 'polite');
        if (!this.options.sheetTabs) this.statusEl.style.marginLeft = 'auto';
        bottom.appendChild(this.statusEl);
        bottom.appendChild(this.buildZoomControl());
      }
      this.root.appendChild(bottom);
      this.renderTabs();
    }
  }

  private buildZoomControl(): HTMLDivElement {
    const s = this.strings;
    const wrap = document.createElement('div');
    wrap.className = 'sx-zoom';
    const button = (label: string, title: string, action: () => void) => {
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'sx-btn';
      b.textContent = label;
      b.title = title;
      b.setAttribute('aria-label', title);
      b.addEventListener('mousedown', (e) => e.preventDefault());
      b.addEventListener('click', action);
      wrap.append(b);
      return b;
    };
    button('−', s.zoomOut, () => this.stepZoom(-1));
    this.zoomLabel = button('100 %', s.zoomReset, () => this.setZoom(1));
    this.zoomLabel.classList.add('sx-zoom-level');
    button('+', s.zoomIn, () => this.stepZoom(1));
    return wrap;
  }

  private buildToolbar(): HTMLDivElement {
    const s = this.strings;
    const tb = document.createElement('div');
    tb.className = 'sx-toolbar';
    tb.setAttribute('role', 'toolbar');
    this.toolbarEl = tb;
    const sep = () => {
      const d = document.createElement('div');
      d.className = 'sx-sep';
      tb.appendChild(d);
    };
    const button = (id: string, title: string, content: string, onClick: (b: HTMLButtonElement) => void, html = false) => {
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'sx-btn';
      b.title = title;
      b.setAttribute('aria-label', title);
      if (html) b.innerHTML = content;
      else b.textContent = content;
      b.addEventListener('mousedown', (e) => e.preventDefault());
      b.addEventListener('click', () => onClick(b));
      tb.appendChild(b);
      this.controls[id] = b;
      return b;
    };

    for (const a of this.options.toolbarActions ?? []) {
      button(`action:${a.id}`, a.title ?? a.label, a.icon ? `${a.icon}<span>${escapeHtml(a.label)}</span>` : a.label, () => a.onClick(), !!a.icon);
    }
    if (this.options.toolbarActions?.length) sep();

    button('undo', s.undo, ICONS.undo, () => this.wb.undo(), true);
    button('redo', s.redo, ICONS.redo, () => this.wb.redo(), true);
    sep();
    const fmt = document.createElement('select');
    fmt.className = 'sx-select';
    fmt.title = s.numberFormat;
    fmt.setAttribute('aria-label', s.numberFormat);
    const de = this.locale === LOCALE_DE || this.strings === getStrings('de');
    const formats: [string, string][] = [
      ['', s.general],
      ['0.00', s.number],
      ['#,##0.00', s.thousands],
      [de ? '#,##0.00 "€"' : '"$"#,##0.00', s.currency],
      ['0.00%', s.percent],
      [de ? 'dd.mm.yyyy' : 'yyyy-mm-dd', s.date],
      ['hh:mm', s.time],
      ['0.00E+00', s.scientific],
      ['@', s.text],
    ];
    for (const [code, label] of formats) fmt.add(new Option(label, code));
    // Names a custom format of the selection (e.g. a currency code); shown only while it applies.
    const current = new Option('', '__current');
    current.hidden = true;
    fmt.add(current);
    fmt.add(new Option('…', '__custom', false, false));
    fmt.addEventListener('change', () => {
      if (fmt.value === '__custom') { this.openCellFormats(); return; }
      if (fmt.value === '__current') return;
      this.styleSelection({ numFmt: fmt.value || null });
      this.focus();
    });
    tb.appendChild(fmt);
    this.controls.numFmt = fmt;
    button('formatCells', s.formatCells, s.formatCells, () => this.openCellFormats());
    button('formatTable', s.formatTable, s.formatTable, () => this.openTableFormats());

    const size = document.createElement('select');
    size.className = 'sx-select';
    size.title = s.fontSize;
    size.setAttribute('aria-label', s.fontSize);
    for (const n of FONT_SIZES) size.add(new Option(String(n), String(n)));
    size.addEventListener('change', () => {
      this.styleSelection({ fontSize: Number(size.value) === 11 ? null : Number(size.value) });
      this.focus();
    });
    tb.appendChild(size);
    this.controls.fontSize = size;
    sep();
    const toggle = (key: 'bold' | 'italic' | 'underline' | 'strike', title: string, label: string, css: string) => {
      const b = button(key, title, label, () => this.toggleStyle(key));
      b.style.cssText = css;
      b.setAttribute('aria-pressed', 'false');
    };
    toggle('bold', s.bold, 'B', 'font-weight:700');
    toggle('italic', s.italic, 'I', 'font-style:italic;font-family:Georgia,serif');
    toggle('underline', s.underline, 'U', 'text-decoration:underline');
    toggle('strike', s.strike, 'S', 'text-decoration:line-through');
    const color = (id: 'color' | 'fill', title: string, label: string, initial: string) => {
      const wrap = document.createElement('label');
      wrap.className = 'sx-btn sx-color';
      wrap.title = title;
      const inner = document.createElement('span');
      inner.className = 'sx-color-inner';
      const glyph = document.createElement('span');
      if (label.startsWith('<svg')) glyph.innerHTML = label;
      else {
        glyph.textContent = label;
        glyph.style.fontWeight = '600';
      }
      const swatch = document.createElement('span');
      swatch.className = 'sx-swatch';
      swatch.style.background = initial;
      inner.append(glyph, swatch);
      const input = document.createElement('input');
      input.type = 'color';
      input.value = initial;
      input.setAttribute('aria-label', title);
      input.addEventListener('input', () => (swatch.style.background = input.value));
      input.addEventListener('change', () => {
        this.styleSelection({ [id]: input.value.toUpperCase() } as StylePatch);
        this.focus();
      });
      wrap.append(inner, input);
      tb.appendChild(wrap);
    };
    color('color', s.textColor, 'A', '#C00000');
    color('fill', s.fillColor, ICONS.fill, '#FFD966');
    sep();
    for (const [id, align, title] of [['alignLeft', 'left', s.alignLeft], ['alignCenter', 'center', s.alignCenter], ['alignRight', 'right', s.alignRight]] as const) {
      const b = button(id, title, ICONS[id], () => {
        const current = this.sheet.getCell(this.active.row, this.active.col)?.style?.hAlign;
        this.styleSelection({ hAlign: current === align ? null : align });
      }, true);
      b.setAttribute('aria-pressed', 'false');
    }
    const wrapBtn = button('wrap', s.wrap, ICONS.wrap, () => this.toggleStyle('wrap'), true);
    wrapBtn.setAttribute('aria-pressed', 'false');
    button('merge', s.merge, ICONS.merge, () => this.toggleMerge(), true);
    button('borders', s.borders, ICONS.borders, (b) => {
      const r = b.getBoundingClientRect();
      const range = this.selectionRange();
      showMenu(this.root, r.left, r.bottom + 2, [
        { label: s.bordersAll, action: () => this.mutate(() => applyBorders(this.wb, this.sheet.name, range, 'all')) },
        { label: s.bordersOutline, action: () => this.mutate(() => applyBorders(this.wb, this.sheet.name, range, 'outline')) },
        { label: s.bordersNone, action: () => this.mutate(() => applyBorders(this.wb, this.sheet.name, range, 'none')) },
      ]);
    }, true);
    sep();
    button('sortAsc', s.sortAsc, ICONS.sortAsc, () => this.sortSelection(false), true);
    button('sortDesc', s.sortDesc, ICONS.sortDesc, () => this.sortSelection(true), true);
    button('sort', s.sort, s.sort, () => this.openSort());
    button('freeze', s.freeze, ICONS.freeze, () => this.toggleFreeze(), true);
    button('find', s.findReplace, ICONS.search, () => this.openFind(!this.options.readOnly), true);
    sep();
    const chartBtn = button('chart', s.chart, `${ICONS.chart}<span>${escapeHtml(s.chartEditor)}</span>`, (b) => {
      const r = b.getBoundingClientRect();
      showMenu(this.root, r.left, r.bottom + 2, CHART_TYPES.map((t) => ({ label: s[t], action: () => this.insertChart(t) })));
    }, true);
    chartBtn.setAttribute('aria-haspopup', 'menu');
    button('conditionalFormatting', s.conditionalFormatting, s.conditionalFormatting, () => {
      if (this.options.readOnly) return;
      this.charts.closeEditor();
      openConditionalDialog(this.root, this.wb, this.sheet.name, this.selectionRange(), s, () => this.focus());
    });
    const pivotButton = button('pivot', s.pivot, s.pivot, (b) => {
      if (this.options.readOnly) return;
      const rect = b.getBoundingClientRect();
      const entries: MenuEntry[] = [{ label: s.createPivot, action: () => {
        let range = this.selectionRange();
        if (range.start.row === range.end.row && range.start.col === range.end.col) range = currentRegion(this.sheet, this.active);
        this.charts.closeEditor();
        openPivotDialog(this.root, this.wb, this.sheet.name, range, s, () => this.focus());
      } }];
      if (this.wb.pivotTables.length) {
        entries.push({ label: s.refreshPivots, action: () => this.mutate(() => this.wb.transaction(() => {
          for (const pivot of [...this.wb.pivotTables]) refreshPivotTable(this.wb, pivot.id);
        })) });
        for (const pivot of this.wb.pivotTables) entries.push({ label: `${s.delete}: ${pivot.source.sheet}!${pivot.source.range} → ${pivot.target.cell}`, action: () => this.mutate(() => removePivotTable(this.wb, pivot.id)) });
      }
      showMenu(this.root, rect.left, rect.bottom + 2, entries);
    });
    pivotButton.setAttribute('aria-haspopup', 'menu');
    return tb;
  }

  // =========================================================================
  // Rendering
  // =========================================================================

  private readTheme(): GridTheme {
    const cs = getComputedStyle(this.root);
    const v = (name: string, fallback: string) => cs.getPropertyValue(name).trim() || fallback;
    const bg = v('--sx-bg', '#ffffff');
    const theme: GridTheme = {
      bg,
      fg: v('--sx-fg', '#1f1f1f'),
      muted: v('--sx-muted', '#616161'),
      grid: v('--sx-grid', '#e1e1e1'),
      headerBg: v('--sx-header-bg', '#f5f5f5'),
      headerFg: v('--sx-header-fg', '#444'),
      headerActiveBg: v('--sx-header-active-bg', '#dfe7f3'),
      headerActiveFg: v('--sx-header-active-fg', '#0f4c81'),
      accent: v('--sx-accent', '#1a73e8'),
      selection: v('--sx-selection', 'rgba(26,115,232,0.12)'),
      fontFamily: DEFAULT_CELL_FONT,
      dark: isDark(bg),
    };
    this.theme = theme;
    if (this.renderer) this.renderer.theme = theme;
    this.chartsDirty = true;
    return theme;
  }

  private chartTheme(): ChartTheme {
    const t = this.theme;
    return { dark: t.dark, background: t.bg, text: t.fg, muted: t.muted, grid: t.grid, fontFamily: getComputedStyle(this.root).fontFamily };
  }

  private scheduleRender(resized = false): void {
    if (resized) this.chartsDirty = true;
    if (this.frame) return;
    this.frame = requestAnimationFrame(() => {
      this.frame = 0;
      this.render();
    });
  }

  private renderNow(): void {
    if (this.frame) cancelAnimationFrame(this.frame);
    this.frame = 0;
    this.render();
  }

  private ensureGeometry(): void {
    const sheet = this.sheet;
    if (this.geoDirty || this.geoSheet !== sheet) {
      this.vp.sheet = sheet;
      this.vp.geo = new Geometry(sheet, this.zoom);
      this.geoSheet = sheet;
      this.geoDirty = false;
    }
  }

  private render(): void {
    this.ensureGeometry();
    const vp = this.vp;
    const sheet = this.sheet;
    const width = this.gridEl.clientWidth;
    const height = this.gridEl.clientHeight;
    if (width !== vp.width || height !== vp.height) {
      vp.width = width;
      vp.height = height;
      this.renderer.resize(width, height);
    }
    // Header width adapts to the number of digits.
    const lastRow = vp.geo.rows.indexAt(vp.frozenH + vp.scrollY + height);
    vp.headerW = Math.max(40, String(lastRow + 1).length * 8 + 18);

    // Scroll extents (grow as the user scrolls).
    const used = sheet.usedRange();
    const chartRows = Math.max(0, ...sheet.charts.map((c) => c.anchor.to.row));
    const chartCols = Math.max(0, ...sheet.charts.map((c) => c.anchor.to.col));
    this.extentRows = Math.min(MAX_ROWS, Math.max(this.extentRows, (used?.end.row ?? 0) + 100, chartRows + 20, this.active.row + 50));
    this.extentCols = Math.min(MAX_COLS, Math.max(this.extentCols, (used?.end.col ?? 0) + 20, chartCols + 10, this.active.col + 10));
    const contentW = vp.geo.cols.offset(this.extentCols) - vp.frozenW;
    const contentH = vp.geo.rows.offset(this.extentRows) - vp.frozenH;
    this.spacer.style.width = `${contentW}px`;
    this.spacer.style.height = `${contentH}px`;
    const left = vp.headerW + vp.frozenW;
    const top = vp.headerH + vp.frozenH;
    Object.assign(this.scroller.style, { left: `${left}px`, top: `${top}px`, width: `${Math.max(0, width - left)}px`, height: `${Math.max(0, height - top)}px` });
    vp.scrollX = this.scroller.scrollLeft;
    vp.scrollY = this.scroller.scrollTop;

    const highlights = this.editor.editing ? formulaHighlights(this.editor.value, sheet.name) : [];
    if (!this.editor.editing) {
      const selected = sheet.getCell(this.active.row, this.active.col);
      const parent = selected?.spillParent !== undefined ? keyToAddress(selected.spillParent) : this.active;
      const spill = sheet.getCell(parent.row, parent.col)?.spillRange;
      if (spill) highlights.push({ range: spill, color: this.theme.accent });
    }
    const conditional = new ConditionalFormatEvaluator(this.wb);
    const resolvedStyles = new Map<number, ReturnType<ConditionalFormatEvaluator['style']>>();
    const cellStyle = (row: number, col: number) => {
      const key = row * MAX_COLS + col;
      if (!resolvedStyles.has(key)) resolvedStyles.set(key, conditional.style(sheet, row, col));
      return resolvedStyles.get(key);
    };
    this.renderer.draw(vp, {
      cellStyle,
      selection: normalizeRange(this.anchor, this.focusCell),
      active: this.active,
      highlights,
      copyRange: this.clip && this.clip.sheet === sheet.name ? this.clip.range : null,
      fillPreview: this.drag?.kind === 'fill' ? this.drag.target : null,
      locale: this.locale,
      showFillHandle: !this.editor.editing && !this.options.readOnly,
      isLink: (r, c) => sheet.links.has(cellKey(r, c)) || /^\s*HYPERLINK\s*\(/i.test(sheet.getCell(r, c)?.formula ?? ''),
      visual: (r, c) => conditional.visual(sheet, r, c),
    });
    this.renderTableControls();

    // Editor follows its cell (or the active cell as an invisible focus sink).
    const ec = this.editor.editing ? this.editor.cell : this.active;
    const merge = sheet.mergeAt(ec.row, ec.col);
    const rect = vp.rect(merge ?? { start: ec, end: ec });
    this.editor.place(rect, sheet.getCell(ec.row, ec.col)?.style, this.theme, width, this.zoom);
    this.renderValidationUi();

    if (this.chartsDirty) {
      this.charts.update(true);
      this.chartsDirty = false;
    } else this.charts.update(false);

    this.updateChrome();
  }

  /** Name box, formula bar, toolbar states and status bar. */
  private updateChrome(): void {
    const sheet = this.sheet;
    const cell = sheet.getCell(this.active.row, this.active.col);
    const range = normalizeRange(this.anchor, this.focusCell);
    if (this.nameBox && document.activeElement !== this.nameBox) {
      const single = range.start.row === range.end.row && range.start.col === range.end.col;
      this.nameBox.value = single ? formatAddress(this.active.row, this.active.col) : formatRange(range);
    }
    if (this.formulaInput && !(this.editor.editing && this.editor.activeInput === this.formulaInput)) {
      this.formulaInput.value = this.editor.editing ? this.editor.value : this.editText(this.active.row, this.active.col);
    }
    const style = this.wb.getCellStyle(sheet.name, this.active.row, this.active.col);
    const pressed = (id: string, on: boolean) => this.controls[id]?.setAttribute('aria-pressed', String(on));
    pressed('bold', !!style?.bold);
    pressed('italic', !!style?.italic);
    pressed('underline', !!style?.underline);
    pressed('strike', !!style?.strike);
    pressed('wrap', !!style?.wrap);
    pressed('alignLeft', style?.hAlign === 'left');
    pressed('alignCenter', style?.hAlign === 'center');
    pressed('alignRight', style?.hAlign === 'right');
    pressed('freeze', !!(sheet.frozen.rows || sheet.frozen.cols));
    this.controls.freeze?.setAttribute('title', sheet.frozen.rows || sheet.frozen.cols ? this.strings.unfreeze : this.strings.freeze);
    (this.controls.undo as HTMLButtonElement | undefined)?.toggleAttribute('disabled', !this.wb.canUndo);
    (this.controls.redo as HTMLButtonElement | undefined)?.toggleAttribute('disabled', !this.wb.canRedo);
    const fmt = this.controls.numFmt as HTMLSelectElement | undefined;
    if (fmt) {
      const code = style?.numFmt ?? '';
      const known = code !== '__current' && code !== '__custom' && [...fmt.options].some((o) => o.value === code);
      const current = [...fmt.options].find((o) => o.value === '__current');
      if (current) {
        const category = numberFormatCategory(code);
        const name = category ? this.strings[category] : '';
        current.text = name ? `${name} (${code.length > 24 ? `${code.slice(0, 23)}…` : code})` : code.length > 32 ? `${code.slice(0, 31)}…` : code;
        current.title = code;
        current.hidden = known;
      }
      fmt.value = known ? code : current ? '__current' : '__custom';
    }
    const size = this.controls.fontSize as HTMLSelectElement | undefined;
    if (size) size.value = String(style?.fontSize ?? 11);
    if (this.zoomLabel) this.zoomLabel.textContent = `${Math.round(this.zoom * 100)} %`;
    this.updateStatus(range);
  }

  private updateStatus(range: RangeAddress): void {
    if (!this.statusEl) return;
    const s = this.strings;
    const sheet = this.sheet;
    const cells = (range.end.row - range.start.row + 1) * (range.end.col - range.start.col + 1);
    if (cells < 2) {
      const link = this.linkAt(this.active.row, this.active.col);
      const hint = /Mac|iPhone|iPad/.test(navigator.platform) ? s.openLinkHint.replace(/^(Ctrl|Strg)/, '⌘') : s.openLinkHint;
      this.statusEl.textContent = link ? `${link.url ?? link.location} · ${hint}` : '';
      return;
    }
    let sum = 0;
    let nums = 0;
    let count = 0;
    const visit = (r: number, c: number) => {
      const v = sheet.getValue(r, c);
      if (v === null) return;
      count++;
      if (typeof v === 'number') {
        sum += v;
        nums++;
      }
    };
    if (cells <= sheet.cells.size * 2 && cells < 200_000) {
      for (let r = range.start.row; r <= range.end.row; r++) for (let c = range.start.col; c <= range.end.col; c++) visit(r, c);
    } else {
      for (const key of sheet.cells.keys()) {
        const { row, col } = keyToAddress(key);
        if (row >= range.start.row && row <= range.end.row && col >= range.start.col && col <= range.end.col) visit(row, col);
      }
    }
    const fmt = (n: number) => formatValue(Number(n.toPrecision(12)), '#,##0.##########', this.locale).replace(/[.,]$/, '');
    const parts: string[] = [];
    if (nums) {
      parts.push(`<span class="sx-optional">${s.average}: ${fmt(sum / nums)}</span>`);
      parts.push(`<span>${s.sum}: ${fmt(sum)}</span>`);
    }
    parts.push(`<span>${s.count}: ${count}</span>`);
    this.statusEl.innerHTML = parts.join('');
  }

  private renderTabs(): void {
    if (!this.tabsEl) return;
    const s = this.strings;
    this.tabsEl.replaceChildren(
      ...this.wb.sheets.map((sheet, i) => {
        const tab = document.createElement('button');
        tab.type = 'button';
        tab.className = 'sx-tab';
        tab.setAttribute('role', 'tab');
        tab.setAttribute('aria-selected', String(i === this.wb.activeSheetIndex));
        tab.textContent = sheet.name;
        tab.addEventListener('click', () => this.activateSheet(i));
        tab.addEventListener('dblclick', () => this.renameSheetInline(tab, sheet.name));
        tab.addEventListener('contextmenu', (e) => {
          e.preventDefault();
          showMenu(this.root, e.clientX, e.clientY, [
            { label: s.renameSheet, action: () => this.renameSheetInline(tab, sheet.name), disabled: this.options.readOnly },
            {
              label: s.deleteSheet,
              disabled: this.wb.sheets.length < 2 || this.options.readOnly,
              action: () => {
                if (confirm(s.confirmDeleteSheet.replace('{0}', sheet.name))) this.mutate(() => this.wb.apply({ type: 'removeSheet', name: sheet.name }));
              },
            },
            'separator',
            { label: s.moveLeft, disabled: i === 0 || this.options.readOnly, action: () => this.mutate(() => this.wb.apply({ type: 'moveSheet', name: sheet.name, index: i - 1 })) },
            { label: s.moveRight, disabled: i === this.wb.sheets.length - 1 || this.options.readOnly, action: () => this.mutate(() => this.wb.apply({ type: 'moveSheet', name: sheet.name, index: i + 1 })) },
          ]);
        });
        return tab;
      }),
    );
    this.tabsEl.children[this.wb.activeSheetIndex]?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }

  private renameSheetInline(tab: HTMLElement, name: string): void {
    if (this.options.readOnly) return;
    const input = document.createElement('input');
    input.value = name;
    tab.replaceChildren(input);
    input.focus();
    input.select();
    let done = false;
    const finish = (save: boolean) => {
      if (done) return;
      done = true;
      const next = input.value.trim();
      if (save && next && next !== name) {
        try {
          this.wb.apply({ type: 'renameSheet', from: name, to: next });
        } catch (e) {
          alert((e as Error).message);
        }
      }
      this.renderTabs();
      this.focus();
    };
    input.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Enter') finish(true);
      if (e.key === 'Escape') finish(false);
    });
    input.addEventListener('blur', () => finish(true));
    input.addEventListener('click', (e) => e.stopPropagation());
  }

  // =========================================================================
  // Selection & navigation
  // =========================================================================

  private selectionRange(): RangeAddress {
    return normalizeRange(this.anchor, this.focusCell);
  }

  private setSelection(anchor: CellAddress, focus: CellAddress, active: CellAddress = anchor): void {
    this.anchor = clampCell(anchor);
    this.focusCell = clampCell(focus);
    this.active = clampCell(active);
    this.charts?.select(null);
    this.events.emit('selectionChange', this.getSelection());
    this.scheduleRender();
  }

  scrollIntoView(row: number, col: number): void {
    this.ensureGeometry();
    const vp = this.vp;
    this.extentRows = Math.min(MAX_ROWS, Math.max(this.extentRows, row + 50));
    this.extentCols = Math.min(MAX_COLS, Math.max(this.extentCols, col + 10));
    this.spacer.style.width = `${vp.geo.cols.offset(this.extentCols) - vp.frozenW}px`;
    this.spacer.style.height = `${vp.geo.rows.offset(this.extentRows) - vp.frozenH}px`;
    const viewW = this.scroller.clientWidth;
    const viewH = this.scroller.clientHeight;
    if (row >= vp.frozenRows) {
      const top = vp.geo.rows.offset(row) - vp.frozenH;
      const bottom = top + vp.geo.rows.size(row);
      if (top < this.scroller.scrollTop) this.scroller.scrollTop = top;
      else if (bottom > this.scroller.scrollTop + viewH) this.scroller.scrollTop = bottom - viewH;
    }
    if (col >= vp.frozenCols) {
      const left = vp.geo.cols.offset(col) - vp.frozenW;
      const right = left + vp.geo.cols.size(col);
      if (left < this.scroller.scrollLeft) this.scroller.scrollLeft = left;
      else if (right > this.scroller.scrollLeft + viewW) this.scroller.scrollLeft = Math.min(left, right - viewW);
    }
    this.scheduleRender();
  }

  private move(dr: number, dc: number, extend: boolean): void {
    this.ensureGeometry();
    const visibleTarget = (cell: CellAddress): CellAddress => {
      const next = clampCell(cell);
      if (dr) {
        const step = Math.sign(dr);
        while (this.vp.geo.rows.size(next.row) === 0 && next.row + step >= 0 && next.row + step < MAX_ROWS) next.row += step;
      }
      if (dc) {
        const step = Math.sign(dc);
        while (this.vp.geo.cols.size(next.col) === 0 && next.col + step >= 0 && next.col + step < MAX_COLS) next.col += step;
      }
      return next;
    };
    if (extend) {
      const f = visibleTarget({ row: this.focusCell.row + dr, col: this.focusCell.col + dc });
      this.setSelection(this.anchor, f, this.active);
      this.scrollIntoView(f.row, f.col);
      return;
    }
    let { row, col } = this.active;
    const merge = this.sheet.mergeAt(row, col);
    if (merge) {
      if (dr > 0) row = merge.end.row;
      if (dc > 0) col = merge.end.col;
      if (dr < 0) row = merge.start.row;
      if (dc < 0) col = merge.start.col;
    }
    let next = visibleTarget({ row: row + dr, col: col + dc });
    const target = this.sheet.mergeAt(next.row, next.col);
    if (target) next = { ...target.start };
    this.setSelection(next, next);
    this.scrollIntoView(next.row, next.col);
  }

  /** Ctrl+Arrow: jump to the edge of the current data block. */
  private jump(dr: number, dc: number, extend: boolean): void {
    const sheet = this.sheet;
    const from = extend ? this.focusCell : this.active;
    const used = sheet.usedRange();
    const maxR = Math.max(used?.end.row ?? 0, from.row);
    const maxC = Math.max(used?.end.col ?? 0, from.col);
    const filled = (r: number, c: number) => sheet.getValue(r, c) !== null;
    const inside = (r: number, c: number) => r >= 0 && c >= 0 && r <= maxR && c <= maxC;
    let r = from.row;
    let c = from.col;
    if (filled(r, c) && inside(r + dr, c + dc) && filled(r + dr, c + dc)) {
      while (inside(r + dr, c + dc) && filled(r + dr, c + dc)) {
        r += dr;
        c += dc;
      }
    } else {
      r += dr;
      c += dc;
      while (inside(r, c) && !filled(r, c)) {
        r += dr;
        c += dc;
      }
      if (!inside(r, c)) {
        r = Math.max(0, Math.min(r, maxR));
        c = Math.max(0, Math.min(c, maxC));
      }
    }
    const target = clampCell({ row: r, col: c });
    if (extend) this.setSelection(this.anchor, target, this.active);
    else this.setSelection(target, target);
    this.scrollIntoView(target.row, target.col);
  }

  // =========================================================================
  // Editing
  // =========================================================================

  private mutate(fn: () => void): void {
    if (this.options.readOnly) return;
    try {
      fn();
    } catch (e) {
      console.error(e);
      alert((e as Error).message);
    }
  }

  /** Text shown when editing a cell. */
  private editText(row: number, col: number): string {
    const cell = this.sheet.getCell(row, col);
    if (!cell) return '';
    if (cell.formula !== undefined) return `=${cell.formula}`;
    const v = cell.value;
    if (v === null) return '';
    if (isError(v)) return v.code;
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') {
      const fmt = cell.style?.numFmt;
      if (isDateFormat(fmt)) return formatValue(v, v % 1 ? (v < 1 ? 'hh:mm:ss' : 'yyyy-mm-dd hh:mm:ss') : 'yyyy-mm-dd', this.locale);
      if (fmt?.includes('%')) return `${String(Number((v * 100).toPrecision(15))).replace('.', this.locale.decimal)}%`;
      return String(v).replace('.', this.locale.decimal);
    }
    const parsed = parseInput(v, this.locale);
    return parsed.value === v && parsed.formula === undefined ? v : `'${v}`;
  }

  beginEdit(initial?: string, mode: 'enter' | 'edit' = 'edit', input?: HTMLInputElement | HTMLTextAreaElement): void {
    if (this.options.readOnly) return;
    const { row, col } = this.active;
    const original = this.editText(row, col);
    this.editor.begin(row, col, original, initial ?? original, mode, input ?? this.editor.textarea);
    if (!input) this.editor.textarea.focus({ preventScroll: true });
    this.editor.updateAssist();
    // Position the editor immediately so a quick click does not hit a stale editor.
    this.renderNow();
  }

  /** Ends editing and writes the input; returns false if data validation keeps the editor open. */
  commitEdit(move?: [number, number]): boolean {
    if (!this.editor.editing) return true;
    const text = this.editor.value;
    const { row, col } = this.editor.cell;
    const changed = text !== this.editor.original;
    if (changed && !this.skipValidation && !this.acceptInput(row, col, text, move)) return false;
    this.skipValidation = false;
    this.editor.end();
    if (changed) this.mutate(() => setCellInput(this.wb, this.sheet.name, row, col, text));
    if (move) this.move(move[0], move[1], false);
    this.focus();
    this.scheduleRender();
    return true;
  }

  cancelEdit(): void {
    if (!this.editor.editing) return;
    this.hideTip(true);
    this.editor.end();
    this.focus();
    this.scheduleRender();
  }

  // =========================================================================
  // Data validation
  // =========================================================================

  /** Checks typed input against the cell's rule; shows the rule's alert if it is invalid. */
  private acceptInput(row: number, col: number, text: string, move?: [number, number]): boolean {
    const rule = validationAt(this.sheet, row, col);
    if (!rule || rule.showError === false) return true;
    const data = inputToCellData(text, this.sheet.getCell(row, col)?.style, this.locale);
    let value: Scalar = data?.v === undefined ? null : scalarFromJson(data.v);
    if (data?.f !== undefined) value = this.wb.calc.evaluateFormula(`=${data.f}`, this.sheet, row, col);
    if (isValidValue(this.wb, this.sheet, rule, row, col, value)) return true;
    const s = this.strings;
    const style = rule.errorStyle ?? 'stop';
    const title = rule.errorTitle || s.invalidTitle;
    const message = rule.error || s.invalidValue;
    if (style === 'information') {
      this.showTip('info', title, message, [{ label: s.ok, primary: true, action: () => this.hideTip(true) }]);
      return true;
    }
    const keepEditing = () => { this.hideTip(true); this.editor.activeInput.focus(); };
    if (style === 'warning') {
      this.showTip('warning', title, message, [
        { label: s.editValue, action: keepEditing },
        { label: s.continueAnyway, primary: true, action: () => { this.hideTip(true); this.skipValidation = true; this.commitEdit(move); } },
      ]);
    } else {
      this.showTip('error', title, message, [
        { label: s.cancel, action: () => this.cancelEdit() },
        { label: s.retry, primary: true, action: () => { keepEditing(); this.editor.activeInput.select(); } },
      ]);
    }
    return false;
  }

  /** Shows a message below the active cell (input message or validation alert). */
  private showTip(kind: 'prompt' | 'info' | 'warning' | 'error', title: string, message: string, actions: { label: string; primary?: boolean; action: () => void }[] = []): void {
    const tip = this.tip;
    tip.className = `sx-validation-tip sx-tip-${kind}`;
    tip.setAttribute('role', kind === 'prompt' ? 'tooltip' : 'alertdialog');
    const heading = document.createElement('strong');
    heading.textContent = title;
    const body = document.createElement('div');
    body.textContent = message;
    tip.replaceChildren(...(title ? [heading] : []), body);
    if (actions.length) {
      const bar = document.createElement('div');
      bar.className = 'sx-actions';
      for (const a of actions) {
        const b = document.createElement('button');
        b.type = 'button';
        b.textContent = a.label;
        if (a.primary) b.className = 'sx-primary';
        b.addEventListener('mousedown', (e) => e.stopPropagation());
        b.addEventListener('click', a.action);
        bar.append(b);
      }
      tip.append(bar);
      this.alertOpen = true;
    }
    tip.hidden = false;
    this.placeTip();
    if (actions.length) (tip.querySelector('.sx-primary') as HTMLButtonElement | null)?.focus();
  }

  private hideTip(alerts = false): void {
    if (this.alertOpen && !alerts) return;
    this.alertOpen = false;
    if (this.tip) this.tip.hidden = true;
  }

  private placeTip(): void {
    if (this.tip.hidden) return;
    const cell = this.editor.editing ? this.editor.cell : this.active;
    const rect = this.vp.rect(this.sheet.mergeAt(cell.row, cell.col) ?? { start: cell, end: cell });
    const width = this.tip.offsetWidth || 240;
    this.tip.style.left = `${Math.max(4, Math.min(rect.x, this.vp.width - width - 8))}px`;
    const below = rect.y + rect.h + 4;
    this.tip.style.top = `${below + this.tip.offsetHeight > this.vp.height ? Math.max(4, rect.y - this.tip.offsetHeight - 4) : below}px`;
  }

  /** Dropdown arrow for list rules and the input message of the active cell. */
  private renderValidationUi(): void {
    const button = this.validationButton;
    const { row, col } = this.active;
    const rule = validationAt(this.sheet, row, col);
    const list = rule?.type === 'list' && rule.dropdown !== false && !this.options.readOnly && !this.editor.editing;
    button.hidden = !list;
    if (list) {
      const merge = this.sheet.mergeAt(row, col);
      const rect = this.vp.rect(merge ?? { start: this.active, end: this.active });
      const size = Math.max(14, Math.min(20, rect.h));
      button.style.left = `${rect.x + rect.w + 1}px`;
      button.style.top = `${rect.y + rect.h - size}px`;
      button.style.height = `${size}px`;
      button.hidden = rect.w <= 0 || rect.x + rect.w > this.vp.width - 4 || rect.y < this.vp.headerH;
    }
    if (this.alertOpen) {
      this.placeTip();
      return;
    }
    if (!this.editor.editing && rule && rule.showPrompt !== false && (rule.prompt || rule.promptTitle)) {
      const key = `${this.sheet.id}:${row}:${col}`;
      if (this.tip.hidden || this.tip.dataset.cell !== key) {
        this.showTip('prompt', rule.promptTitle ?? '', rule.prompt ?? '');
        this.tip.dataset.cell = key;
      } else this.placeTip();
    } else this.hideTip();
  }

  /** Opens the value list of the active cell's list rule. */
  private openValidationList(): void {
    if (this.options.readOnly || this.editor.editing) return;
    const { row, col } = this.active;
    const rule = validationAt(this.sheet, row, col);
    if (rule?.type !== 'list') return;
    const values = validationListValues(this.wb, this.sheet, rule, row, col);
    const rect = this.validationButton.getBoundingClientRect();
    const grid = this.gridEl.getBoundingClientRect();
    const cell = this.vp.rect({ start: this.active, end: this.active });
    const name = this.sheet.name;
    showMenu(this.root, grid.left + cell.x, (rect.height ? rect.bottom : grid.top + cell.y + cell.h) + 1, values.length
      ? values.map((v) => ({ label: formatValue(v, undefined, this.locale), action: () => this.mutate(() => setValues(this.wb, name, row, col, [[v as string | number | boolean]], { parse: false })) }))
      : [{ label: this.strings.chooseValue, disabled: true, action: () => undefined }]);
  }

  private openValidation(): void {
    if (this.options.readOnly) return;
    if (!this.commitEdit()) return;
    this.charts.closeEditor();
    openValidationDialog(this.root, this.wb, this.sheet.name, this.selectionRange(), this.active, this.locale, this.strings, () => this.focus());
  }

  // =========================================================================
  // Notes and hyperlinks
  // =========================================================================

  /** Hyperlink of a cell: a stored link or the target of a HYPERLINK formula. */
  private linkAt(row: number, col: number): { url?: string; location?: string } | null {
    const stored = this.sheet.links.get(cellKey(row, col));
    if (stored) return { ...(stored.target ? { url: stored.target } : {}), ...(stored.location ? { location: stored.location } : {}) };
    const formula = this.sheet.getCell(row, col)?.formula;
    if (!formula || !/HYPERLINK\s*\(/i.test(formula)) return null;
    try {
      let target: string | null = null;
      walk(parseFormula(formula), (n) => {
        if (target === null && n.type === 'func' && n.name === 'HYPERLINK' && n.args[0]) {
          const ctx = this.wb.calc.contextFor(this.sheet, row, col);
          const v = evaluate(n.args[0], ctx);
          const text = toText(Array.isArray(v) ? v[0]?.[0] ?? null : typeof v === 'object' && v && 'kind' in v ? ctx.getCell(v.sheet, v.range.start.row, v.range.start.col) : v);
          if (typeof text === 'string') target = text;
        }
      });
      if (!target) return null;
      const t = target as string;
      return t.startsWith('#') ? { location: t.slice(1) } : { url: t };
    } catch {
      return null;
    }
  }

  private openLink(link: { url?: string; location?: string }): void {
    if (link.location) {
      this.goTo(link.location);
      this.focus();
      return;
    }
    if (!link.url) return;
    if (this.options.onOpenLink) this.options.onOpenLink(link.url);
    else window.open(link.url, '_blank', 'noopener');
  }

  /** Shows the note of the cell under the mouse. */
  private hoverNote(x: number, y: number): void {
    const vp = this.vp;
    const note = x > vp.headerW && y > vp.headerH ? (() => {
      const cell = this.cellAt(x, y);
      const merge = this.sheet.mergeAt(cell.row, cell.col);
      const at = merge?.start ?? cell;
      const n = this.sheet.notes.get(cellKey(at.row, at.col));
      return n ? { n, rect: vp.rect(merge ?? { start: at, end: at }), key: `${at.row}:${at.col}` } : null;
    })() : null;
    if (!note) {
      this.noteEl.hidden = true;
      return;
    }
    if (this.noteEl.hidden || this.noteEl.dataset.cell !== note.key) {
      const author = document.createElement('strong');
      author.textContent = note.n.author ? `${note.n.author}:` : '';
      const body = document.createElement('div');
      body.textContent = note.n.text;
      this.noteEl.replaceChildren(...(note.n.author ? [author] : []), body);
      this.noteEl.dataset.cell = note.key;
      this.noteEl.hidden = false;
    }
    const width = this.noteEl.offsetWidth || 200;
    const left = note.rect.x + note.rect.w + 6;
    this.noteEl.style.left = `${left + width > vp.width ? Math.max(4, note.rect.x - width - 6) : left}px`;
    this.noteEl.style.top = `${Math.max(vp.headerH, Math.min(note.rect.y, vp.height - this.noteEl.offsetHeight - 4))}px`;
  }

  private styleSelection(patch: StylePatch): void {
    this.mutate(() => applyStyle(this.wb, this.sheet.name, this.selectionRange(), patch));
  }

  private toggleStyle(key: 'bold' | 'italic' | 'underline' | 'strike' | 'wrap'): void {
    const current = this.wb.getCellStyle(this.sheet.name, this.active.row, this.active.col)?.[key];
    this.styleSelection({ [key]: !current } as StylePatch);
  }

  private toggleMerge(): void {
    const range = this.selectionRange();
    const existing = this.sheet.mergeAt(range.start.row, range.start.col);
    this.mutate(() => {
      if (existing) unmergeCells(this.wb, this.sheet.name, range);
      else if (range.start.row !== range.end.row || range.start.col !== range.end.col) mergeCells(this.wb, this.sheet.name, range);
    });
  }

  /** Hides or shows the selected rows/columns; hiding moves the cursor to a visible cell. */
  private setHidden(axis: 'row' | 'col', hidden: boolean): void {
    const sel = this.selectionRange();
    const [first, last] = axis === 'row' ? [sel.start.row, sel.end.row] : [sel.start.col, sel.end.col];
    this.mutate(() => {
      if (axis === 'row') (hidden ? hideRows : unhideRows)(this.wb, this.sheet.name, first, last);
      else (hidden ? hideCols : unhideCols)(this.wb, this.sheet.name, first, last);
    });
    if (!hidden) return;
    this.geoDirty = true;
    this.ensureGeometry();
    const axisGeo = axis === 'row' ? this.vp.geo.rows : this.vp.geo.cols;
    const max = axis === 'row' ? MAX_ROWS : MAX_COLS;
    let target = last + 1;
    while (target < max && axisGeo.size(target) === 0) target++;
    if (target >= max) { target = first - 1; while (target > 0 && axisGeo.size(target) === 0) target--; }
    const cell = axis === 'row' ? { row: Math.max(0, target), col: this.active.col } : { row: this.active.row, col: Math.max(0, target) };
    this.setSelection(cell, cell);
  }

  private toggleFreeze(): void {
    const sheet = this.sheet;
    this.mutate(() => {
      if (sheet.frozen.rows || sheet.frozen.cols) this.wb.apply({ type: 'setFrozen', sheet: sheet.name, rows: 0, cols: 0 });
      else {
        const rows = this.active.row || 1;
        const cols = this.active.col;
        this.wb.apply({ type: 'setFrozen', sheet: sheet.name, rows, cols });
      }
    });
    this.scroller.scrollTo(0, 0);
  }

  private sortSelection(descending: boolean): void {
    this.commitEdit();
    const { range, hasHeader } = resolveSortContext(this.sheet, this.selectionRange(), this.active);
    this.mutate(() => sortRange(this.wb, this.sheet.name, range, [{ col: this.active.col, descending }], { hasHeader }));
  }

  private openCellFormats(): void {
    if (this.options.readOnly) return;
    this.commitEdit(); this.charts.closeEditor();
    openCellFormatDialog(this.root, this.wb, this.sheet.name, this.selectionRange(), this.locale, this.strings, () => this.focus(), (range) => this.autoFit(range.start.col, range.end.col));
  }

  private openTableFormats(): void {
    if (this.options.readOnly) return;
    this.commitEdit(); this.charts.closeEditor();
    let range = this.selectionRange();
    const table = tableAt(this.sheet, this.active.row, this.active.col);
    if (table) range = parseRange(table.range)!.range;
    else if (range.start.row === range.end.row && range.start.col === range.end.col) range = currentRegion(this.sheet, this.active);
    openTableFormatDialog(this.root, this.wb, this.sheet.name, range, this.strings, () => this.focus(), (range) => this.autoFit(range.start.col, range.end.col));
  }

  private openSort(): void {
    if (this.options.readOnly) return;
    this.commitEdit(); this.charts.closeEditor();
    const { range, hasHeader } = resolveSortContext(this.sheet, this.selectionRange(), this.active);
    openSortDialog(this.root, this.wb, this.sheet.name, range, hasHeader, this.active.col, this.strings, () => this.focus());
  }

  private renderTableControls(): void {
    if (!this.tableControls) {
      this.tableControls = document.createElement('div'); this.tableControls.className = 'sx-table-controls';
      this.gridEl.append(this.tableControls);
    }
    this.tableControls.replaceChildren();
    if (this.editor.editing || this.options.readOnly) return;
    const vp = this.vp;
    for (const table of this.sheet.tables) {
      if (!table.headerRow) continue;
      const range = parseRange(table.range)!.range;
      const row = range.start.row;
      const y = vp.y(row), height = vp.geo.rows.size(row);
      const minimumY = vp.headerH + (row >= vp.frozenRows ? vp.frozenH : 0);
      if (height < 14 || y < minimumY || y >= vp.height - 16) continue;
      for (let col = range.start.col; col <= range.end.col; col++) {
        const x = vp.x(col), width = vp.geo.cols.size(col);
        const minimumX = vp.headerW + (col >= vp.frozenCols ? vp.frozenW : 0);
        if (width < 22 || x + width - 18 < minimumX || x + width > vp.width - 14) continue;
        const filtered = table.filters?.some((f) => f.column === col - range.start.col);
        const b = document.createElement('button'); b.type = 'button'; b.className = 'sx-table-sort'; b.textContent = filtered ? '⏷' : '▾';
        b.classList.toggle('sx-filter-active', !!filtered);
        b.style.background = filtered ? this.theme.selection : '';
        const title = this.strings.filterColumn.replace('{0}', String(this.sheet.getValue(row, col) ?? formatAddress(row, col)));
        b.title = title; b.setAttribute('aria-label', title); b.setAttribute('aria-haspopup', 'menu');
        b.style.left = `${x + width - 18}px`; b.style.top = `${y + Math.max(0, (height - 18) / 2)}px`;
        b.style.color = this.wb.getCellStyle(this.sheet.name, row, col)?.color ?? this.theme.fg;
        b.addEventListener('mousedown', (event) => { event.preventDefault(); event.stopPropagation(); });
        b.addEventListener('click', (event) => {
          event.stopPropagation();
          this.setSelection(range.start, range.end, { row, col });
          const rect = b.getBoundingClientRect();
          showMenu(this.root, rect.left, rect.bottom + 2, [
            { label: this.strings.ascending, action: () => this.mutate(() => sortRange(this.wb, this.sheet.name, range, [{ col }], { hasHeader: true })) },
            { label: this.strings.descending, action: () => this.mutate(() => sortRange(this.wb, this.sheet.name, range, [{ col, descending: true }], { hasHeader: true })) },
            { label: this.strings.sort, action: () => this.openSort() },
            'separator',
            { label: this.strings.filterValues, action: () => openTableFilterDialog(this.root, this.wb, this.sheet.name, table, col, this.locale, this.strings, () => this.focus()) },
            { label: this.strings.clearFilter, disabled: !filtered, action: () => setTableFilter(this.wb, this.sheet.name, table.name, col - range.start.col, null) },
            { label: this.strings.formatCells, action: () => { this.select({ start: { row: Math.min(row + 1, range.end.row), col }, end: { row: range.end.row, col } }); this.openCellFormats(); } },
          ]);
        });
        this.tableControls.append(b);
      }
    }
  }

  private insertChart(type: (typeof CHART_TYPES)[number]): void {
    let range = this.selectionRange();
    if (range.start.row === range.end.row && range.start.col === range.end.col) range = currentRegion(this.sheet, this.active);
    this.mutate(() => {
      const spec = insertChartFromRange(this.wb, this.sheet.name, range, type);
      requestAnimationFrame(() => this.charts.select(spec.id));
    });
  }

  private autoFit(firstCol: number, lastCol = firstCol): void {
    const sheet = this.sheet;
    const widths = new Map<number, number>();
    const conditional = new ConditionalFormatEvaluator(this.wb);
    for (const [key, cell] of sheet.cells) {
      const a = keyToAddress(key);
      if (a.col < firstCol || a.col > lastCol || cell.value === null) continue;
      const style = conditional.style(sheet, a.row, a.col);
      const table = tableAt(sheet, a.row, a.col);
      const header = table?.headerRow && parseRange(table.range)!.range.start.row === a.row;
      const width = this.renderer.measure(formatValue(cell.value, style?.numFmt, this.locale), style) + 12 + (header ? 18 : 0);
      widths.set(a.col, Math.max(widths.get(a.col) ?? 0, width));
    }
    if (widths.size) this.mutate(() => this.wb.apply({ type: 'setColWidths', sheet: sheet.name, sizes: [...widths].map(([index, width]) => ({ index, size: Math.ceil(width) })) }));
  }

  // =========================================================================
  // Clipboard
  // =========================================================================

  private onCopy(e: ClipboardEvent, cut: boolean): void {
    if (this.editor.editing) return;
    e.preventDefault();
    const range = this.selectionRange();
    this.clip = copyRange(this.wb, this.sheet.name, range);
    if (cut && !this.options.readOnly) this.clip.cut = true;
    this.clipText = rangeToTSV(this.wb, this.sheet.name, this.clip.range);
    e.clipboardData?.setData('text/plain', this.clipText);
    e.clipboardData?.setData('text/html', this.toHtml(this.clip.range));
    this.scheduleRender();
  }

  private toHtml(range: RangeAddress): string {
    const rows: string[] = [];
    for (let r = range.start.row; r <= range.end.row; r++) {
      const cells: string[] = [];
      for (let c = range.start.col; c <= range.end.col; c++) {
        const cell = this.sheet.getCell(r, c);
        const s = cell?.style;
        const css = [s?.bold ? 'font-weight:bold' : '', s?.italic ? 'font-style:italic' : '', s?.color ? `color:${s.color}` : '', s?.fill ? `background:${s.fill}` : ''].filter(Boolean).join(';');
        cells.push(`<td${css ? ` style="${css}"` : ''}>${escapeHtml(cell ? formatValue(cell.value, s?.numFmt, this.locale) : '')}</td>`);
      }
      rows.push(`<tr>${cells.join('')}</tr>`);
    }
    return `<table>${rows.join('')}</table>`;
  }

  private onPaste(text: string | null, special: PasteSpecial = {}): void {
    if (this.editor.editing || this.options.readOnly) return;
    const target = this.selectionRange();
    this.mutate(() => {
      if (this.clip && (text === null || normalizeNewlines(text) === normalizeNewlines(this.clipText))) {
        const pasted = pasteRange(this.wb, this.clip, this.sheet.name, target, special);
        if (this.clip.cut) this.clip = null;
        this.setSelection(pasted.start, pasted.end);
        return;
      }
      if (!text) return;
      this.clip = null;
      // External text becomes a clipboard block so paste options apply to it too;
      // like typing, it keeps the target cells' formatting.
      const rows = parseCsv(text.replace(/\r?\n$/, ''), '\t');
      const width = Math.max(1, ...rows.map((r) => r.length));
      const cells = rows.map((r) => Array.from({ length: width }, (_, i) => (r[i] ? inputToCellData(r[i]!, undefined, this.locale) : null)));
      const external = { sheet: this.sheet.name, range: { start: target.start, end: { row: target.start.row + rows.length - 1, col: target.start.col + width - 1 } }, cells };
      const pasted = pasteRange(this.wb, external, this.sheet.name, { start: target.start, end: target.start }, { ...special, mode: 'formulas' });
      this.setSelection(pasted.start, pasted.end);
    });
  }

  /** Pastes from the system clipboard (or the internal clipboard) with paste options. */
  async pasteSpecial(special?: PasteSpecial): Promise<void> {
    if (this.editor.editing || this.options.readOnly) return;
    let text: string | null = null;
    try {
      text = await navigator.clipboard.readText();
    } catch {
      text = null;
    }
    if (special) {
      this.onPaste(text, special);
      return;
    }
    this.charts.closeEditor();
    openPasteSpecialDialog(this.root, this.strings, (options) => this.onPaste(text, options), () => this.focus());
  }

  // =========================================================================
  // Event binding
  // =========================================================================

  private local(e: MouseEvent): { x: number; y: number } {
    const r = this.gridEl.getBoundingClientRect();
    return { x: e.clientX - r.left, y: e.clientY - r.top };
  }

  private cellAt(x: number, y: number): CellAddress {
    return clampCell({ row: this.vp.rowAt(Math.max(this.vp.headerH, y)), col: this.vp.colAt(Math.max(this.vp.headerW, x)) });
  }

  private bindGrid(): void {
    const grid = this.gridEl;
    this.scroller.addEventListener('scroll', () => {
      const sc = this.scroller;
      if (sc.scrollTop + sc.clientHeight > sc.scrollHeight - 400 && this.extentRows < MAX_ROWS) this.extentRows = Math.min(MAX_ROWS, this.extentRows + 500);
      if (sc.scrollLeft + sc.clientWidth > sc.scrollWidth - 400 && this.extentCols < MAX_COLS) this.extentCols = Math.min(MAX_COLS, this.extentCols + 20);
      this.scheduleRender();
    });
    // Ctrl/Cmd+wheel and trackpad pinch zoom the grid instead of the page.
    grid.addEventListener('wheel', (e) => {
      if (!e.ctrlKey && !e.metaKey) return;
      e.preventDefault();
      e.stopPropagation();
      this.setZoom(this.zoom * Math.exp(-Math.max(-50, Math.min(50, e.deltaY)) / 400));
    }, { passive: false, capture: true });
    grid.addEventListener('wheel', (e) => {
      if (this.scroller.contains(e.target as Node)) return;
      e.preventDefault();
      this.scroller.scrollBy(e.deltaX, e.deltaY);
    }, { passive: false });

    grid.addEventListener('mousedown', (e) => this.onMouseDown(e));
    grid.addEventListener('mousemove', (e) => {
      if (this.drag) return;
      const { x, y } = this.local(e);
      const onLink = (e.ctrlKey || e.metaKey) && x > this.vp.headerW && y > this.vp.headerH && (() => { const c = this.cellAt(x, y); return !!this.linkAt(c.row, c.col); })();
      grid.style.cursor = onLink ? 'pointer' : this.cursorAt(x, y);
      if (!(e.target as HTMLElement).closest('.sx-dialog, .sx-validation-tip')) this.hoverNote(x, y);
    });
    grid.addEventListener('mouseleave', () => { this.noteEl.hidden = true; });
    grid.addEventListener('dblclick', (e) => {
      const { x, y } = this.local(e);
      const vp = this.vp;
      if (y < vp.headerH && x > vp.headerW) {
        const border = this.colBorderAt(x);
        if (border !== null) this.autoFit(border);
        return;
      }
      if (x < vp.headerW || y < vp.headerH) return;
      this.beginEdit(undefined, 'edit');
    });
    grid.addEventListener('contextmenu', (e) => this.onContextMenu(e));
  }

  private colBorderAt(x: number): number | null {
    const vp = this.vp;
    const c = vp.colAt(x);
    const right = vp.x(c) + vp.geo.cols.size(c);
    if (Math.abs(x - right) <= 4) return c;
    if (c > 0 && Math.abs(x - vp.x(c)) <= 4) return c - 1;
    return null;
  }

  private rowBorderAt(y: number): number | null {
    const vp = this.vp;
    const r = vp.rowAt(y);
    const bottom = vp.y(r) + vp.geo.rows.size(r);
    if (Math.abs(y - bottom) <= 3) return r;
    if (r > 0 && Math.abs(y - vp.y(r)) <= 3) return r - 1;
    return null;
  }

  private onFillHandle(x: number, y: number): boolean {
    if (this.options.readOnly || this.editor.editing) return false;
    const rc = this.vp.rect(this.selectionRange());
    return Math.abs(x - (rc.x + rc.w)) <= 5 && Math.abs(y - (rc.y + rc.h)) <= 5;
  }

  private cursorAt(x: number, y: number): string {
    const vp = this.vp;
    if (y < vp.headerH && x > vp.headerW) return this.colBorderAt(x) !== null ? 'col-resize' : 's-resize';
    if (x < vp.headerW && y > vp.headerH) return this.rowBorderAt(y) !== null ? 'row-resize' : 'e-resize';
    if (this.onFillHandle(x, y)) return 'crosshair';
    return 'cell';
  }

  private onMouseDown(e: MouseEvent): void {
    closeMenu();
    if (e.target === this.editor.textarea && this.editor.editing) return;
    if ((e.target as HTMLElement).closest('.sx-suggest, .sx-dialog, .sx-chart')) return;
    if (e.target === this.scroller) {
      // Ignore clicks on the native scrollbars.
      const r = this.scroller.getBoundingClientRect();
      if (e.clientX - r.left >= this.scroller.clientWidth || e.clientY - r.top >= this.scroller.clientHeight) return;
    }
    const { x, y } = this.local(e);
    const vp = this.vp;
    this.charts.select(null);
    if (e.button === 2) {
      // Right click keeps the selection if the cell is inside it.
      if (x > vp.headerW && y > vp.headerH) {
        const cell = this.cellAt(x, y);
        const sel = this.selectionRange();
        if (!(cell.row >= sel.start.row && cell.row <= sel.end.row && cell.col >= sel.start.col && cell.col <= sel.end.col)) this.setSelection(cell, cell);
      }
      return;
    }
    e.preventDefault();

    if (y < vp.headerH && x < vp.headerW) {
      if (!this.commitEdit()) return;
      const used = this.sheet.usedRange();
      this.setSelection({ row: 0, col: 0 }, { row: Math.max(used?.end.row ?? 0, 0), col: Math.max(used?.end.col ?? 0, 0) });
      this.focus();
      return;
    }
    if (y < vp.headerH) {
      const border = this.colBorderAt(x);
      if (border !== null && !this.options.readOnly) {
        this.drag = { kind: 'resize-col', index: border, start: x, size: vp.geo.cols.size(border) / this.zoom };
      } else {
        if (!this.commitEdit()) return;
        const c = vp.colAt(x);
        const anchorCol = e.shiftKey ? this.anchor.col : c;
        this.setSelection({ row: 0, col: anchorCol }, { row: this.lastRowForSelection(), col: c }, { row: vp.frozenRows ? 0 : vp.rowAt(vp.headerH + vp.frozenH), col: anchorCol });
        this.drag = { kind: 'cols', anchor: anchorCol };
        this.focus();
      }
    } else if (x < vp.headerW) {
      const border = this.rowBorderAt(y);
      if (border !== null && !this.options.readOnly) {
        this.drag = { kind: 'resize-row', index: border, start: y, size: vp.geo.rows.size(border) / this.zoom };
      } else {
        if (!this.commitEdit()) return;
        const r = vp.rowAt(y);
        const anchorRow = e.shiftKey ? this.anchor.row : r;
        this.setSelection({ row: anchorRow, col: 0 }, { row: r, col: this.lastColForSelection() }, { row: anchorRow, col: 0 });
        this.drag = { kind: 'rows', anchor: anchorRow };
        this.focus();
      }
    } else {
      const cell = this.cellAt(x, y);
      if (this.editor.editing && this.editor.canInsertRef()) {
        this.editor.insertRef(cell);
        this.drag = { kind: 'ref', anchor: cell };
        this.afterEditorInput();
      } else if (this.onFillHandle(x, y)) {
        this.drag = { kind: 'fill', source: this.selectionRange(), target: null };
      } else {
        if (!this.commitEdit()) return;
        const link = (e.ctrlKey || e.metaKey) && !e.shiftKey ? this.linkAt(cell.row, cell.col) : null;
        if (link) {
          this.setSelection(cell, cell);
          this.openLink(link);
          return;
        }
        const merge = this.sheet.mergeAt(cell.row, cell.col);
        if (e.shiftKey) this.setSelection(this.anchor, cell, this.active);
        else if (merge) this.setSelection(merge.start, merge.end, merge.start);
        else this.setSelection(cell, cell);
        this.drag = { kind: 'select' };
        this.focus();
      }
    }

    const onMove = (ev: MouseEvent) => this.onDragMove(ev);
    const onUp = () => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
      this.onDragEnd();
    };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
  }

  private lastRowForSelection(): number {
    return Math.max(this.sheet.usedRange()?.end.row ?? 0, this.vp.visibleRows()[1]);
  }

  private lastColForSelection(): number {
    return Math.max(this.sheet.usedRange()?.end.col ?? 0, this.vp.visibleCols()[1]);
  }

  private onDragMove(e: MouseEvent): void {
    const drag = this.drag;
    if (!drag) return;
    const { x, y } = this.local(e);
    const vp = this.vp;
    // Auto-scroll when dragging beyond the edges.
    if (drag.kind !== 'resize-col' && drag.kind !== 'resize-row') {
      if (x > vp.width - 10) this.scroller.scrollLeft += 20;
      else if (x < vp.headerW + vp.frozenW && drag.kind !== 'rows' && this.scroller.scrollLeft > 0) this.scroller.scrollLeft -= 20;
      if (y > vp.height - 10) this.scroller.scrollTop += 20;
      else if (y < vp.headerH + vp.frozenH && drag.kind !== 'cols' && this.scroller.scrollTop > 0) this.scroller.scrollTop -= 20;
    }
    const cell = this.cellAt(Math.min(x, vp.width - 1), Math.min(y, vp.height - 1));
    switch (drag.kind) {
      case 'select':
        this.focusCell = cell;
        this.scheduleRender();
        break;
      case 'cols':
        this.focusCell = { row: this.focusCell.row, col: cell.col };
        this.anchor = { row: 0, col: drag.anchor };
        this.scheduleRender();
        break;
      case 'rows':
        this.focusCell = { row: cell.row, col: this.focusCell.col };
        this.anchor = { row: drag.anchor, col: 0 };
        this.scheduleRender();
        break;
      case 'resize-col': {
        const size = Math.max(0, Math.round(drag.size + (x - drag.start) / this.zoom));
        this.sheet.colWidths.set(drag.index, size); // live preview, committed on mouseup
        this.geoDirty = true;
        this.chartsDirty = true;
        this.scheduleRender();
        break;
      }
      case 'resize-row': {
        const size = Math.max(0, Math.round(drag.size + (y - drag.start) / this.zoom));
        this.sheet.rowHeights.set(drag.index, size);
        this.geoDirty = true;
        this.chartsDirty = true;
        this.scheduleRender();
        break;
      }
      case 'fill': {
        const s = drag.source;
        const dRow = cell.row > s.end.row ? cell.row - s.end.row : cell.row < s.start.row ? cell.row - s.start.row : 0;
        const dCol = cell.col > s.end.col ? cell.col - s.end.col : cell.col < s.start.col ? cell.col - s.start.col : 0;
        if (!dRow && !dCol) drag.target = null;
        else if (Math.abs(dRow) >= Math.abs(dCol)) drag.target = dRow > 0 ? { start: s.start, end: { row: cell.row, col: s.end.col } } : { start: { row: cell.row, col: s.start.col }, end: s.end };
        else drag.target = dCol > 0 ? { start: s.start, end: { row: s.end.row, col: cell.col } } : { start: { row: s.start.row, col: cell.col }, end: s.end };
        this.scheduleRender();
        break;
      }
      case 'ref': {
        this.editor.insertRef(drag.anchor, cell);
        this.afterEditorInput();
        break;
      }
    }
  }

  private onDragEnd(): void {
    const drag = this.drag;
    this.drag = null;
    if (!drag) return;
    if (drag.kind === 'resize-col' || drag.kind === 'resize-row') {
      const map = drag.kind === 'resize-col' ? this.sheet.colWidths : this.sheet.rowHeights;
      const size = map.get(drag.index)!;
      // Restore and apply through an operation so it is undoable.
      if (drag.size === (drag.kind === 'resize-col' ? this.sheet.defaultColWidth : this.sheet.defaultRowHeight)) map.delete(drag.index);
      else map.set(drag.index, drag.size);
      // Resize all selected columns/rows when the resized one is part of a full selection.
      const sel = this.selectionRange();
      const isCol = drag.kind === 'resize-col';
      const inSel = isCol ? drag.index >= sel.start.col && drag.index <= sel.end.col : drag.index >= sel.start.row && drag.index <= sel.end.row;
      const indices: number[] = [];
      if (inSel) for (let i = isCol ? sel.start.col : sel.start.row; i <= (isCol ? sel.end.col : sel.end.row); i++) indices.push(i);
      else indices.push(drag.index);
      this.mutate(() => this.wb.apply({ type: isCol ? 'setColWidths' : 'setRowHeights', sheet: this.sheet.name, sizes: indices.map((index) => ({ index, size })) }));
    } else if (drag.kind === 'fill' && drag.target) {
      const target = drag.target;
      this.mutate(() => fillRange(this.wb, this.sheet.name, drag.source, target));
      this.setSelection(target.start, target.end, this.active);
    } else if (drag.kind === 'ref') {
      this.editor.activeInput.focus();
    }
    this.scheduleRender();
  }

  private onContextMenu(e: MouseEvent): void {
    e.preventDefault();
    if ((e.target as HTMLElement).closest('.sx-dialog')) return;
    const s = this.strings;
    const sel = this.selectionRange();
    const ro = !!this.options.readOnly;
    const rows = sel.end.row - sel.start.row + 1;
    const cols = sel.end.col - sel.start.col + 1;
    const name = this.sheet.name;
    const entries: MenuEntry[] = [
      { label: s.cut, disabled: ro, action: () => document.execCommand('cut') },
      { label: s.copy, action: () => document.execCommand('copy') },
      { label: s.paste, disabled: ro, action: () => void this.pasteSpecial({}) },
      { label: s.pasteValues, disabled: ro, action: () => void this.pasteSpecial({ mode: 'values' }) },
      { label: s.pasteSpecial, disabled: ro, action: () => void this.pasteSpecial() },
      'separator',
      { label: s.insertRowAbove, disabled: ro, action: () => this.mutate(() => insertRows(this.wb, name, sel.start.row, rows)) },
      { label: s.insertRowBelow, disabled: ro, action: () => this.mutate(() => insertRows(this.wb, name, sel.end.row + 1, rows)) },
      { label: s.deleteRows, disabled: ro, action: () => this.mutate(() => deleteRows(this.wb, name, sel.start.row, rows)) },
      'separator',
      { label: s.insertColLeft, disabled: ro, action: () => this.mutate(() => insertCols(this.wb, name, sel.start.col, cols)) },
      { label: s.insertColRight, disabled: ro, action: () => this.mutate(() => insertCols(this.wb, name, sel.end.col + 1, cols)) },
      { label: s.deleteCols, disabled: ro, action: () => this.mutate(() => deleteCols(this.wb, name, sel.start.col, cols)) },
      'separator',
      { label: s.hideRows, disabled: ro, action: () => this.setHidden('row', true) },
      { label: s.unhideRows, disabled: ro || !hasHidden(this.sheet, 'row', sel.start.row, sel.end.row), action: () => this.setHidden('row', false) },
      { label: s.hideCols, disabled: ro, action: () => this.setHidden('col', true) },
      { label: s.unhideCols, disabled: ro || !hasHidden(this.sheet, 'col', sel.start.col, sel.end.col), action: () => this.setHidden('col', false) },
      'separator',
      { label: s.sortAsc, disabled: ro, action: () => this.sortSelection(false) },
      { label: s.sortDesc, disabled: ro, action: () => this.sortSelection(true) },
      { label: s.sort, disabled: ro, action: () => this.openSort() },
      { label: s.formatCells, disabled: ro, action: () => this.openCellFormats() },
      { label: s.formatTable, disabled: ro, action: () => this.openTableFormats() },
      { label: s.chart, disabled: ro, action: () => this.insertChart('column') },
      { label: s.dataValidation, disabled: ro, action: () => this.openValidation() },
      { label: `${s.findReplace}…`, action: () => this.openFind(!ro) },
      'separator',
      { label: s.clearContents, disabled: ro, action: () => this.mutate(() => clearRange(this.wb, name, sel, 'contents')) },
      { label: s.clearFormats, disabled: ro, action: () => this.mutate(() => clearRange(this.wb, name, sel, 'formats')) },
    ];
    showMenu(this.root, e.clientX, e.clientY, entries);
  }

  private afterEditorInput(): void {
    if (this.formulaInput && this.editor.activeInput !== this.formulaInput) this.formulaInput.value = this.editor.value;
    this.editor.updateAssist();
    this.scheduleRender();
  }

  private bindKeyboard(): void {
    const ta = this.editor.textarea;
    ta.addEventListener('input', () => {
      if (!this.editor.editing) {
        if (this.options.readOnly) {
          ta.value = '';
          return;
        }
        this.beginEdit(ta.value, 'enter');
      }
      this.editor.clearRefSpan();
      this.afterEditorInput();
    });
    ta.addEventListener('sx-change', () => this.afterEditorInput());
    ta.addEventListener('compositionstart', () => {
      if (!this.editor.editing && !this.options.readOnly) this.beginEdit('', 'enter');
    });
    ta.addEventListener('keydown', (e) => this.onKeyDown(e, ta));
    ta.addEventListener('keyup', () => this.editor.editing && this.editor.updateAssist());
    ta.addEventListener('copy', (e) => this.onCopy(e, false));
    ta.addEventListener('cut', (e) => this.onCopy(e, true));
    ta.addEventListener('paste', (e) => {
      if (this.editor.editing) return;
      e.preventDefault();
      this.onPaste(e.clipboardData?.getData('text/plain') ?? '');
    });
    ta.addEventListener('blur', () => {
      // Commit when focus leaves the grid entirely (not to the formula bar or popups).
      setTimeout(() => {
        const active = document.activeElement;
        if (this.editor.editing && this.editor.activeInput === ta && active !== ta && !this.root.contains(active)) this.commitEdit();
      });
    });
    // Charts are focusable for keyboard deletion.
    this.gridEl.addEventListener('keydown', (e) => {
      if (e.target !== ta && this.charts.selectedId && (e.key === 'Delete' || e.key === 'Backspace') && !this.options.readOnly) {
        e.preventDefault();
        this.charts.deleteSelected();
        this.focus();
      }
    });
  }

  private onKeyDown(e: KeyboardEvent, input: HTMLInputElement | HTMLTextAreaElement): void {
    // AltGr can report Ctrl+Alt on Windows/Linux. Keep character entry (e.g.
    // @, brackets and backslashes) out of shortcut handling, including AltGr+V.
    if (e.isComposing || e.getModifierState('AltGraph')) return;
    const mod = e.ctrlKey || e.metaKey;
    const ed = this.editor;
    if (ed.editing) {
      if (ed.suggestOpen) {
        if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
          e.preventDefault();
          ed.moveSuggest(e.key === 'ArrowDown' ? 1 : -1);
          return;
        }
        if (e.key === 'Tab' || (e.key === 'Enter' && !e.altKey)) {
          e.preventDefault();
          ed.acceptSuggest();
          return;
        }
        if (e.key === 'Escape') {
          e.preventDefault();
          ed.closeSuggest();
          return;
        }
      }
      switch (e.key) {
        case 'Enter':
          e.preventDefault();
          if (e.altKey) {
            const pos = input.selectionStart ?? input.value.length;
            input.setRangeText('\n', pos, input.selectionEnd ?? pos, 'end');
            ed.syncFrom(input);
            this.afterEditorInput();
          } else this.commitEdit(e.shiftKey ? [-1, 0] : [1, 0]);
          return;
        case 'Tab':
          e.preventDefault();
          this.commitEdit(e.shiftKey ? [0, -1] : [0, 1]);
          return;
        case 'Escape':
          e.preventDefault();
          this.cancelEdit();
          return;
        case 'F2':
          e.preventDefault();
          ed.mode = ed.mode === 'enter' ? 'edit' : 'enter';
          return;
        case 'ArrowUp':
        case 'ArrowDown':
        case 'ArrowLeft':
        case 'ArrowRight': {
          if (ed.mode !== 'enter' || input !== ed.textarea) return;
          const [dr, dc] = arrowDelta(e.key);
          if (ed.canInsertRef()) {
            // "Point" mode: arrow keys move a reference in the formula, Shift extends it to a range.
            e.preventDefault();
            const current = ed.currentRef() ?? { anchor: this.active, active: this.active };
            const next = clampCell({ row: current.active.row + dr, col: current.active.col + dc });
            ed.insertRef(e.shiftKey ? current.anchor : next, next);
            this.afterEditorInput();
            return;
          }
          e.preventDefault();
          this.commitEdit([dr, dc]);
          return;
        }
      }
      return;
    }

    // ---- Not editing: navigation and shortcuts ----
    if (mod && !e.shiftKey && !e.altKey && e.code === 'KeyG') {
      e.preventDefault();
      this.openGoTo();
      return;
    }
    if (mod && !e.shiftKey && (e.code === 'KeyF' || e.code === 'KeyH')) {
      // Ctrl+F finds; Ctrl+H (and Cmd+Option+F on macOS, as in VS Code) replaces.
      e.preventDefault();
      this.openFind(e.code === 'KeyH' || e.altKey);
      return;
    }
    if (mod && e.code === 'KeyV' && (e.shiftKey || e.altKey)) {
      // Ctrl+Shift+V pastes values, Ctrl+Alt+V opens the paste options.
      e.preventDefault();
      void this.pasteSpecial(e.altKey ? undefined : { mode: 'values' });
      return;
    }
    if (mod && !e.altKey && (e.code === 'Digit9' || e.code === 'Digit0')) {
      // Ctrl+9 / Ctrl+0 hide rows / columns; with Shift they unhide.
      e.preventDefault();
      if (!this.options.readOnly) this.setHidden(e.code === 'Digit9' ? 'row' : 'col', !e.shiftKey);
      return;
    }
    if (mod && !e.altKey) {
      const k = e.key.toLowerCase();
      if (k === '1') { e.preventDefault(); this.openCellFormats(); return; }
      if ((k === 'z' && !e.shiftKey) || k === 'y' || (k === 'z' && e.shiftKey)) {
        if (!this.options.keyboardUndo) return;
        e.preventDefault();
        if (k === 'y' || e.shiftKey) this.wb.redo();
        else this.wb.undo();
        return;
      }
      const styleKeys: Record<string, 'bold' | 'italic' | 'underline'> = { b: 'bold', i: 'italic', u: 'underline' };
      if (styleKeys[k]) {
        e.preventDefault();
        this.toggleStyle(styleKeys[k]!);
        return;
      }
      if (k === 'a') {
        e.preventDefault();
        const used = this.sheet.usedRange();
        this.setSelection({ row: 0, col: 0 }, used?.end ?? { row: 0, col: 0 }, this.active);
        return;
      }
      if (k === 'd' || k === 'r') {
        e.preventDefault();
        const sel = this.selectionRange();
        const down = k === 'd';
        if ((down && sel.end.row > sel.start.row) || (!down && sel.end.col > sel.start.col)) {
          const source = down ? { start: sel.start, end: { row: sel.start.row, col: sel.end.col } } : { start: sel.start, end: { row: sel.end.row, col: sel.start.col } };
          this.mutate(() => fillRange(this.wb, this.sheet.name, source, sel));
        }
        return;
      }
      if (k === 'home') {
        e.preventDefault();
        this.setSelection({ row: 0, col: 0 }, { row: 0, col: 0 });
        this.scrollIntoView(0, 0);
        return;
      }
      if (e.key.startsWith('Arrow')) {
        e.preventDefault();
        const [dr, dc] = arrowDelta(e.key);
        this.jump(dr, dc, e.shiftKey);
        return;
      }
      return; // let copy/paste/cut reach the clipboard handlers
    }
    if (e.altKey && e.key === 'ArrowDown' && validationAt(this.sheet, this.active.row, this.active.col)?.type === 'list') {
      e.preventDefault();
      this.openValidationList();
      return;
    }
    switch (e.key) {
      case 'ArrowUp':
      case 'ArrowDown':
      case 'ArrowLeft':
      case 'ArrowRight': {
        e.preventDefault();
        const [dr, dc] = arrowDelta(e.key);
        this.move(dr, dc, e.shiftKey);
        return;
      }
      case 'Enter':
        e.preventDefault();
        this.move(e.shiftKey ? -1 : 1, 0, false);
        return;
      case 'Tab':
        e.preventDefault();
        this.move(0, e.shiftKey ? -1 : 1, false);
        return;
      case 'Home':
        e.preventDefault();
        this.setSelection({ row: this.active.row, col: 0 }, { row: this.active.row, col: 0 });
        this.scrollIntoView(this.active.row, 0);
        return;
      case 'PageDown':
      case 'PageUp': {
        e.preventDefault();
        const [r0, r1] = this.vp.visibleRows();
        this.move((e.key === 'PageDown' ? 1 : -1) * Math.max(1, r1 - r0 - 1), 0, e.shiftKey);
        return;
      }
      case 'Delete':
        e.preventDefault();
        this.mutate(() => clearRange(this.wb, this.sheet.name, this.selectionRange(), 'contents'));
        return;
      case 'Backspace':
        e.preventDefault();
        this.beginEdit('', 'enter');
        return;
      case 'F2':
        e.preventDefault();
        this.beginEdit(undefined, 'edit');
        return;
      case 'Escape':
        if (this.clip) {
          this.clip = null;
          this.scheduleRender();
        }
        return;
    }
  }

  private bindFormulaBar(): void {
    const bar = this.formulaInput!;
    bar.addEventListener('input', () => {
      if (!this.editor.editing) {
        if (this.options.readOnly) {
          this.scheduleRender();
          return;
        }
        this.beginEdit(bar.value, 'edit', bar);
      } else if (this.editor.activeInput !== bar) {
        this.editor.activeInput = bar;
      }
      this.editor.clearRefSpan();
      this.editor.syncFrom(bar);
      this.afterEditorInput();
    });
    bar.addEventListener('focus', () => {
      if (this.editor.editing) this.editor.activeInput = bar;
    });
    bar.addEventListener('keydown', (e) => {
      if (!this.editor.editing) {
        if (e.key === 'Enter' || e.key === 'Escape') {
          e.preventDefault();
          this.focus();
        }
        return;
      }
      this.onKeyDown(e, bar);
    });
    bar.addEventListener('keyup', () => this.editor.editing && this.editor.updateAssist());
    bar.addEventListener('click', () => this.editor.editing && this.editor.updateAssist());

    const nb = this.nameBox!;
    nb.addEventListener('focus', () => nb.select());
    nb.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        this.goTo(nb.value);
        this.focus();
      } else if (e.key === 'Escape') {
        e.preventDefault();
        this.focus();
      }
    });
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function clampCell(c: CellAddress): CellAddress {
  return { row: Math.max(0, Math.min(MAX_ROWS - 1, c.row)), col: Math.max(0, Math.min(MAX_COLS - 1, c.col)) };
}

function arrowDelta(key: string): [number, number] {
  switch (key) {
    case 'ArrowUp': return [-1, 0];
    case 'ArrowDown': return [1, 0];
    case 'ArrowLeft': return [0, -1];
    default: return [0, 1];
  }
}

function isDark(color: string): boolean {
  const m = /^#?([0-9a-f]{6})$/i.exec(color.trim()) ?? /^#?([0-9a-f]{3})$/i.exec(color.trim());
  let r = 255, g = 255, b = 255;
  if (m) {
    const h = m[1]!.length === 3 ? m[1]!.split('').map((c) => c + c).join('') : m[1]!;
    [r, g, b] = [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16)) as [number, number, number];
  } else {
    const rgb = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(color);
    if (rgb) [r, g, b] = [Number(rgb[1]), Number(rgb[2]), Number(rgb[3])];
  }
  return 0.2126 * r + 0.7152 * g + 0.0722 * b < 128;
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function normalizeNewlines(s: string): string {
  return s.replace(/\r\n/g, '\n').replace(/\n$/, '');
}
