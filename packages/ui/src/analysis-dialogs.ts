import { applyStyle, clearRange, colToName, createPivotTable, formatAddress, formatRange, formatTable, formatValue, LOCALE_DE, MAX_COLS, parseRange, removeTable, setConditionalFormats, sortRange, tableAt, TABLE_THEMES, type ConditionalFormatRule, type ConditionalOperator, type FormatLocale, type PivotAggregate, type RangeAddress, type SortKey, type StylePatch, type Workbook } from '@sheets/core';
import type { Strings } from './i18n.js';
import { setTableFilter, tableFilterValue, type SheetTable } from '@sheets/core';

export function openTableFilterDialog(root: HTMLElement, wb: Workbook, sheetName: string, table: SheetTable, col: number, locale: FormatLocale, s: Strings, onClose: () => void): void {
  const sheet = wb.requireSheet(sheetName);
  const range = parseRange(table.range)!.range;
  const column = col - range.start.col;
  const title = s.filterColumn.replace('{0}', String(sheet.getValue(range.start.row, col) ?? ''));
  const p = panel(root, title, s, onClose);
  const search = p.field(s.searchValues, input('')); search.type = 'search';
  const values = new Map<string, string>();
  for (let row = range.start.row + 1; row <= range.end.row; row++) {
    const value = sheet.getValue(row, col), key = tableFilterValue(value);
    if (!values.has(key)) values.set(key, key === '' ? s.blankValue : formatValue(value, wb.getCellStyle(sheetName, row, col)?.numFmt, locale));
  }
  const filter = table.filters?.find((f) => f.column === column);
  const selected = new Set(filter?.values ?? values.keys());
  const all = p.field(s.selectAllValues, input('', 'checkbox')); all.parentElement!.className = 'sx-check';
  const list = document.createElement('div'); list.className = 'sx-filter-values';
  list.style.cssText = 'max-height:260px;overflow:auto;min-height:80px'; p.dialog.append(list);
  const matching = () => [...values].filter(([, label]) => label.toLocaleLowerCase().includes(search.value.toLocaleLowerCase()));
  const updateAll = () => { const keys = matching().map(([key]) => key); all.checked = keys.length > 0 && keys.every((key) => selected.has(key)); all.indeterminate = !all.checked && keys.some((key) => selected.has(key)); };
  const render = () => {
    list.replaceChildren();
    for (const [key, value] of matching().sort((a, b) => a[1].localeCompare(b[1], undefined, { numeric: true }))) {
      const label = document.createElement('label'); label.className = 'sx-check';
      const checkbox = input('', 'checkbox'); checkbox.checked = selected.has(key);
      checkbox.addEventListener('change', () => { checkbox.checked ? selected.add(key) : selected.delete(key); updateAll(); });
      label.append(checkbox, document.createTextNode(value)); list.append(label);
    }
    updateAll();
  };
  search.addEventListener('input', render);
  all.addEventListener('change', () => { for (const [key] of matching()) all.checked ? selected.add(key) : selected.delete(key); render(); });
  p.button(s.clearFilter, () => { setTableFilter(wb, sheetName, table.name, column, null); p.close(); });
  p.finish(() => setTableFilter(wb, sheetName, table.name, column, [...values.keys()].every((v) => selected.has(v)) ? null : [...selected]));
  render();
}

export function closeAnalysisDialog(root: HTMLElement): void {
  root.querySelector('.sx-analysis-dialog')?.remove();
}

function panel(root: HTMLElement, title: string, s: Strings, onClose: () => void) {
  closeAnalysisDialog(root);
  const dialog = document.createElement('form');
  dialog.className = 'sx-dialog sx-analysis-dialog';
  dialog.setAttribute('role', 'dialog');
  dialog.setAttribute('aria-label', title);
  const heading = document.createElement('h3');
  heading.textContent = title;
  dialog.append(heading);
  const close = () => { dialog.remove(); onClose(); };
  dialog.addEventListener('keydown', (event) => {
    event.stopPropagation();
    if (event.key === 'Escape') { event.preventDefault(); close(); }
  });
  dialog.addEventListener('mousedown', (event) => event.stopPropagation());
  const field = <T extends HTMLElement>(label: string, input: T): T => {
    const el = document.createElement('label');
    el.append(label, input); dialog.append(el); return input;
  };
  const error = document.createElement('div');
  error.setAttribute('role', 'alert');
  const actions = document.createElement('div');
  actions.className = 'sx-actions';
  const button = (label: string, action: () => void) => {
    const b = document.createElement('button'); b.type = 'button'; b.textContent = label;
    b.addEventListener('click', () => { try { action(); } catch (e) { error.textContent = String(e instanceof Error ? e.message : e); } });
    actions.append(b); return b;
  };
  const finish = (apply: () => void) => {
    button(s.close, close);
    const submit = button(s.apply, () => { if (!dialog.reportValidity()) return; apply(); close(); }); submit.className = 'sx-primary';
    dialog.addEventListener('submit', (event) => { event.preventDefault(); submit.click(); });
    dialog.append(error, actions); root.append(dialog);
    dialog.querySelector<HTMLElement>('input,select')?.focus();
  };
  return { dialog, field, button, finish, close };
}

function input(value: string, type = 'text'): HTMLInputElement {
  const el = document.createElement('input'); el.type = type; el.value = value; return el;
}

function select(options: [string, string][], value: string): HTMLSelectElement {
  const el = document.createElement('select');
  for (const [key, label] of options) el.add(new Option(label, key, false, key === value));
  return el;
}

export function openCellFormatDialog(root: HTMLElement, wb: Workbook, sheetName: string, range: RangeAddress, locale: FormatLocale, s: Strings, onClose: () => void, autoFit?: (range: RangeAddress) => void): void {
  const p = panel(root, s.formatCells, s, onClose);
  const sheet = wb.requireSheet(sheetName);
  const style = wb.getCellStyle(sheetName, range.start.row, range.start.col) ?? {};
  const patch: StylePatch = {};
  const formats: [string, string][] = [['general', s.general], ['number', s.number], ['currency', s.currency], ['percent', s.percent], ['date', s.date], ['time', s.time], ['text', s.text], ['scientific', s.scientific], ['custom', s.customFormat]];
  const category = p.field(s.numberFormat, select(formats, 'custom'));
  const code = p.field(s.customFormat, input(style.numFmt ?? ''));
  if (!code.value || code.value === 'General') category.value = 'general';
  else if (/^[#0,]+(?:\.[0#]+)?$/.test(code.value)) category.value = 'number';
  else if (code.value.endsWith('%')) category.value = 'percent';
  else if (/[€$]/.test(code.value)) category.value = 'currency';
  code.placeholder = '#,##0.00 "€"';
  const decimals = p.field(s.decimalPlaces, input(String(/\.([0#]+)/.exec(code.value)?.[1]?.length ?? 2), 'number')); decimals.min = '0'; decimals.max = '15'; decimals.step = '1';
  const thousands = p.field(s.useThousands, input('', 'checkbox')); thousands.parentElement!.className = 'sx-check'; thousands.checked = /#,##/.test(code.value);
  const preview = document.createElement('output'); preview.className = 'sx-format-preview'; preview.setAttribute('aria-label', s.preview);
  preview.setAttribute('aria-live', 'polite'); p.dialog.append(preview);
  const fit = autoFit ? p.field(s.autoFitColumns, input('', 'checkbox')) : undefined;
  if (fit) { fit.checked = true; fit.parentElement!.className = 'sx-check'; }
  let sample: number | string | boolean = 1234.567;
  const selectedValue = sheet.getValue(range.start.row, range.start.col);
  if (typeof selectedValue === 'number' || typeof selectedValue === 'boolean') sample = selectedValue;
  const updatePreview = () => { preview.textContent = formatValue(sample, code.value || undefined, locale); };
  const numericFields = () => {
    const numeric = ['number', 'currency', 'percent', 'scientific'].includes(category.value);
    decimals.parentElement!.hidden = !numeric; thousands.parentElement!.hidden = !['number', 'currency'].includes(category.value);
  };
  const preset = () => {
    const count = Number(decimals.value);
    if (!Number.isInteger(count) || count < 0 || count > 15) return;
    const digits = count ? '.' + '0'.repeat(count) : '';
    const number = (thousands.checked ? '#,##0' : '0') + digits;
    const de = locale === LOCALE_DE;
    const codes: Record<string, string> = { general: '', number, currency: de ? number + ' "€"' : '"$"' + number, percent: '0' + digits + '%', date: de ? 'dd.mm.yyyy' : 'yyyy-mm-dd', time: 'hh:mm:ss', text: '@', scientific: '0' + digits + 'E+00' };
    if (category.value !== 'custom') { code.value = codes[category.value] ?? ''; patch.numFmt = code.value || null; }
    numericFields(); updatePreview();
  };
  category.addEventListener('change', preset);
  decimals.addEventListener('input', preset); thousands.addEventListener('change', preset);
  code.addEventListener('input', () => { category.value = 'custom'; patch.numFmt = code.value || null; numericFields(); updatePreview(); });
  const font = p.field(s.fontName, input(style.fontName ?? 'Calibri'));
  font.addEventListener('input', () => { patch.fontName = font.value.trim() || null; });
  const size = p.field(s.fontSize, input(String(style.fontSize ?? 11), 'number')); size.min = '1'; size.max = '409';
  size.addEventListener('input', () => { patch.fontSize = Number(size.value); });
  for (const [key, title] of [['bold', s.bold], ['italic', s.italic], ['underline', s.underline], ['strike', s.strike], ['wrap', s.wrap]] as const) {
    const checkbox = p.field(title, input('', 'checkbox')); checkbox.checked = !!style[key]; checkbox.parentElement!.className = 'sx-check';
    checkbox.addEventListener('change', () => { patch[key] = checkbox.checked; });
  }
  for (const [key, title, fallback] of [['color', s.textColor, '#1F1F1F'], ['fill', s.fillColor, '#FFFFFF']] as const) {
    const color = p.field(title, input(style[key] ?? fallback, 'color'));
    color.addEventListener('change', () => { patch[key] = color.value.toUpperCase(); });
  }
  const horizontal = p.field(s.horizontalAlign, select([['', s.automatic], ['left', s.left], ['center', s.alignCenter], ['right', s.right], ['justify', s.justify]], style.hAlign ?? ''));
  horizontal.addEventListener('change', () => { patch.hAlign = horizontal.value ? horizontal.value as 'left' : null; });
  const vertical = p.field(s.verticalAlign, select([['', s.automatic], ['top', s.top], ['middle', s.alignCenter], ['bottom', s.bottom]], style.vAlign ?? ''));
  vertical.addEventListener('change', () => { patch.vAlign = vertical.value ? vertical.value as 'top' : null; });
  p.button(s.clearFormats, () => { clearRange(wb, sheetName, range, 'formats'); p.close(); });
  p.finish(() => {
    if (patch.fontSize !== undefined && patch.fontSize !== null && (!Number.isFinite(patch.fontSize) || patch.fontSize < 1 || patch.fontSize > 409)) throw new Error(s.fontSize);
    wb.transaction(() => {
      if (Object.keys(patch).length) applyStyle(wb, sheetName, range, patch);
      if (fit?.checked) autoFit?.(range);
    });
  });
  numericFields(); updatePreview();
}

export function openTableFormatDialog(root: HTMLElement, wb: Workbook, sheetName: string, selected: RangeAddress, s: Strings, onClose: () => void, autoFit?: (range: RangeAddress) => void): void {
  const p = panel(root, s.formatTable, s, onClose);
  const sheet = wb.requireSheet(sheetName);
  const existing = tableAt(sheet, selected.start.row, selected.start.col);
  const range = p.field(s.dataRange, input(existing?.range ?? formatRange(selected)));
  const name = p.field(s.tableName, input(existing?.name ?? '')); name.placeholder = s.automatic;
  const header = p.field(s.headerRow, input('', 'checkbox')); header.checked = existing?.headerRow ?? true; header.parentElement!.className = 'sx-check';
  const stripes = p.field(s.bandedRows, input('', 'checkbox')); stripes.checked = existing?.bandedRows ?? true; stripes.parentElement!.className = 'sx-check';
  const fit = autoFit ? p.field(s.autoFitColumns, input('', 'checkbox')) : undefined;
  if (fit) { fit.checked = true; fit.parentElement!.className = 'sx-check'; }
  const names: Record<string, string> = { blue: s.blue, green: s.green, orange: s.orange, purple: s.purple, gray: s.gray };
  const selectedTheme = Object.entries(TABLE_THEMES).find(([, t]) => JSON.stringify(t) === JSON.stringify(existing?.theme))?.[0];
  const styles: [string, string][] = Object.keys(TABLE_THEMES).map((key) => [key, names[key]!]);
  if (existing && !selectedTheme) styles.push(['custom', s.tableStyle]);
  const palette = p.field(s.tableStyle, select(styles, selectedTheme ?? (existing ? 'custom' : 'blue')));
  const preview = document.createElement('table'); preview.className = 'sx-table-preview'; preview.setAttribute('aria-label', s.preview); p.dialog.append(preview);
  const chosenTheme = () => palette.value === 'custom' ? existing!.theme : TABLE_THEMES[palette.value]!;
  const show = () => {
    const theme = chosenTheme(); preview.replaceChildren();
    for (let row = 0; row < 4; row++) {
      const tr = preview.insertRow();
      const isHeader = header.checked && row === 0;
      const striped = stripes.checked && (row - (header.checked ? 1 : 0)) % 2 === 0;
      const style = { ...theme.body, ...(isHeader ? theme.header : striped ? theme.stripe : {}) };
      for (let col = 0; col < 3; col++) {
        const td = tr.insertCell(); td.textContent = isHeader ? colToName(col) : String((row + 1) * (col + 1) * 10);
        td.style.background = style.fill ?? ''; td.style.color = style.color ?? ''; td.style.fontWeight = style.bold ? '700' : '400';
      }
    }
  };
  for (const el of [palette, header, stripes]) el.addEventListener('change', show);
  if (existing) p.button(s.removeTableStyle, () => { removeTable(wb, sheetName, existing.name); p.close(); });
  p.finish(() => {
    const ref = parseRange(range.value);
    if (!ref || (ref.sheet && ref.sheet.toLowerCase() !== sheetName.toLowerCase())) throw new Error(s.invalidRange);
    wb.transaction(() => {
      formatTable(wb, sheetName, ref.range, { name: name.value.trim() || undefined, headerRow: header.checked, bandedRows: stripes.checked, theme: chosenTheme() });
      if (fit?.checked) autoFit?.(ref.range);
    });
  });
  show();
}

export function openSortDialog(root: HTMLElement, wb: Workbook, sheetName: string, selected: RangeAddress, hasHeader: boolean, activeCol: number, s: Strings, onClose: () => void): void {
  const p = panel(root, s.sort, s, onClose);
  const sheet = wb.requireSheet(sheetName);
  const range = p.field(s.dataRange, input(formatRange(selected)));
  const header = p.field(s.headerRow, input('', 'checkbox')); header.checked = hasHeader; header.parentElement!.className = 'sx-check';
  const levels = document.createElement('div'); levels.className = 'sx-sort-levels'; p.dialog.append(levels);
  const keys: SortKey[] = [{ col: Math.max(selected.start.col, Math.min(selected.end.col, activeCol)) }];
  const currentRange = () => {
    const ref = parseRange(range.value);
    if (!ref || (ref.sheet && ref.sheet.toLowerCase() !== sheetName.toLowerCase())) throw new Error(s.invalidRange);
    return ref.range;
  };
  const drawLevels = () => {
    let ref: RangeAddress;
    try { ref = currentRange(); } catch { return; }
    levels.replaceChildren();
    keys.forEach((key, index) => {
      const group = document.createElement('fieldset');
      const legend = document.createElement('legend'); legend.textContent = index === 0 ? s.sortBy : s.thenBy; group.append(legend);
      const options: [string, string][] = [];
      for (let col = ref.start.col; col <= ref.end.col; col++) options.push([String(col), header.checked ? String(sheet.getValue(ref.start.row, col) ?? colToName(col)) : colToName(col)]);
      if (key.col < ref.start.col || key.col > ref.end.col) key.col = ref.start.col;
      const column = select(options, String(key.col)); column.setAttribute('aria-label', `${legend.textContent} ${index + 1}`);
      column.addEventListener('change', () => { key.col = Number(column.value); });
      const order = select([['asc', s.ascending], ['desc', s.descending]], key.descending ? 'desc' : 'asc'); order.setAttribute('aria-label', `${s.sortOrder} ${index + 1}`);
      order.addEventListener('change', () => { key.descending = order.value === 'desc'; });
      const remove = document.createElement('button'); remove.type = 'button'; remove.textContent = '×'; remove.title = s.removeSortLevel; remove.setAttribute('aria-label', `${s.removeSortLevel} ${index + 1}`); remove.disabled = keys.length === 1;
      remove.addEventListener('click', () => { keys.splice(index, 1); drawLevels(); });
      group.append(column, order, remove); levels.append(group);
    });
  };
  range.addEventListener('change', drawLevels); header.addEventListener('change', drawLevels);
  p.button(s.addSortLevel, () => {
    const ref = currentRange();
    let col = ref.start.col;
    while (keys.some((k) => k.col === col) && col <= ref.end.col) col++;
    if (col <= ref.end.col) { keys.push({ col }); drawLevels(); }
  });
  p.finish(() => {
    if (new Set(keys.map((k) => k.col)).size !== keys.length) throw new Error(s.duplicateSortColumn);
    sortRange(wb, sheetName, currentRange(), keys, { hasHeader: header.checked });
  });
  drawLevels();
}

export function openConditionalDialog(root: HTMLElement, wb: Workbook, sheetName: string, selection: RangeAddress, s: Strings, onClose: () => void): void {
  const p = panel(root, s.conditionalFormatting, s, onClose);
  const sheet = wb.requireSheet(sheetName);
  const choose = p.field(s.rule, select([['', s.newRule], ...sheet.conditionalFormats.map((r): [string, string] => [r.id, `${r.range} · ${r.type}`])], ''));
  const range = p.field(s.dataRange, input(formatRange(selection)));
  const mode = p.field(s.ruleType, select([['greaterThan', s.greaterThan], ['lessThan', s.lessThan], ['between', s.between], ['expression', s.expression], ['colorScale', s.colorScale]], 'greaterThan'));
  const first = p.field(s.firstFormula, input('0'));
  const second = p.field(s.secondFormula, input('100'));
  const fill = p.field(s.fillColor, input('#FFF2CC', 'color'));
  const low = p.field(s.lowColor, input('#F8696B', 'color'));
  const middle = p.field(s.middleColor, input('#FFEB84', 'color'));
  const high = p.field(s.highColor, input('#63BE7B', 'color'));
  const stop = p.field(s.stopIfTrue, input('', 'checkbox'));
  stop.parentElement!.className = 'sx-check';
  const showFields = () => {
    const scale = mode.value === 'colorScale';
    for (const el of [first, fill, stop]) el.parentElement!.hidden = scale;
    second.parentElement!.hidden = mode.value !== 'between' && mode.value !== 'notBetween';
    for (const el of [low, middle, high]) el.parentElement!.hidden = !scale;
  };
  mode.addEventListener('change', showFields);
  choose.addEventListener('change', () => {
    const rule = sheet.conditionalFormats.find((r) => r.id === choose.value);
    if (!rule) return;
    range.value = rule.range; stop.checked = !!rule.stopIfTrue;
    const nextMode = rule.type === 'cellIs' ? rule.operator : rule.type;
    if (![...mode.options].some((o) => o.value === nextMode)) mode.add(new Option(nextMode, nextMode));
    mode.value = nextMode;
    if (rule.type === 'colorScale') {
      low.value = rule.colors[0]!; high.value = rule.colors.at(-1)!; middle.value = rule.colors[1]!;
    } else {
      first.value = rule.formulas[0]!; second.value = rule.formulas[1] ?? '100'; fill.value = rule.style.fill ?? '#FFF2CC';
    }
    showFields();
  });
  const move = (delta: number) => {
    const rules = [...sheet.conditionalFormats];
    const index = rules.findIndex((r) => r.id === choose.value);
    if (index < 0 || index + delta < 0 || index + delta >= rules.length) return;
    const [rule] = rules.splice(index, 1); rules.splice(index + delta, 0, rule!);
    setConditionalFormats(wb, sheetName, rules);
  };
  const priorities = document.createElement('div'); priorities.className = 'sx-actions';
  for (const [title, delta] of [[s.priorityUp, -1], [s.priorityDown, 1]] as const) {
    const b = document.createElement('button'); b.type = 'button'; b.textContent = delta < 0 ? '↑' : '↓'; b.title = title; b.setAttribute('aria-label', title);
    b.addEventListener('click', () => move(delta)); priorities.append(b);
  }
  p.dialog.append(priorities);
  p.button(s.delete, () => { if (choose.value) { setConditionalFormats(wb, sheetName, sheet.conditionalFormats.filter((r) => r.id !== choose.value)); p.close(); } });
  p.finish(() => {
    const previous = sheet.conditionalFormats.find((r) => r.id === choose.value);
    const base = { id: choose.value || `cf-${crypto.randomUUID()}`, range: range.value.trim(), stopIfTrue: stop.checked };
    const rule: ConditionalFormatRule = mode.value === 'colorScale'
      ? { ...base, type: 'colorScale', colors: previous?.type === 'colorScale' && previous.colors.length === 2 ? [low.value, high.value] : [low.value, middle.value, high.value] }
      : mode.value === 'expression'
        ? { ...base, type: 'expression', formulas: [first.value], style: { ...(previous && previous.type !== 'colorScale' ? previous.style : {}), fill: fill.value.toUpperCase() } }
        : { ...base, type: 'cellIs', operator: mode.value as ConditionalOperator, formulas: ['between', 'notBetween'].includes(mode.value) ? [first.value, second.value] : [first.value], style: { ...(previous && previous.type !== 'colorScale' ? previous.style : {}), fill: fill.value.toUpperCase() } };
    setConditionalFormats(wb, sheetName, previous ? sheet.conditionalFormats.map((r) => r.id === previous.id ? rule : r) : [...sheet.conditionalFormats, rule]);
  });
  showFields();
}

export function openPivotDialog(root: HTMLElement, wb: Workbook, sheetName: string, selection: RangeAddress, s: Strings, onClose: () => void): void {
  const p = panel(root, s.createPivot, s, onClose);
  const sheet = wb.requireSheet(sheetName);
  const source = p.field(s.dataRange, input(formatRange(selection, sheetName)));
  const target = p.field(s.targetCell, input(formatRange({ start: { row: 0, col: Math.min(MAX_COLS - 1, (sheet.usedRange()?.end.col ?? 0) + 2) }, end: { row: 0, col: Math.min(MAX_COLS - 1, (sheet.usedRange()?.end.col ?? 0) + 2) } }, sheetName)));
  const rows = p.field(s.rowFields, select([], '')); rows.multiple = true; rows.size = 3;
  const columns = p.field(s.columnFields, select([], '')); columns.multiple = true; columns.size = 3;
  const value = p.field(s.valueField, select([], ''));
  const aggregation = p.field(s.aggregation, select([['sum', s.sum], ['count', s.count], ['average', s.average], ['min', s.minimum], ['max', s.maximum]], 'sum'));
  const updateFields = () => {
    const ref = parseRange(source.value);
    const data = ref ? wb.getSheet(ref.sheet ?? sheetName) : undefined;
    if (!ref || !data) return;
    for (const el of [rows, columns, value]) el.replaceChildren();
    for (let c = ref.range.start.col; c <= ref.range.end.col; c++) {
      const name = String(data.getValue(ref.range.start.row, c) ?? formatAddress(ref.range.start.row, c));
      const index = c - ref.range.start.col;
      rows.add(new Option(name, String(index), false, index === 0));
      columns.add(new Option(name, String(index)));
      value.add(new Option(name, String(index), false, c === ref.range.end.col));
    }
  };
  source.addEventListener('change', updateFields);
  updateFields();
  p.finish(() => {
    const from = parseRange(source.value), to = parseRange(target.value);
    if (!from || !to || to.range.start.row !== to.range.end.row || to.range.start.col !== to.range.end.col) throw new Error(s.dataRange);
    createPivotTable(wb, {
      id: `pivot-${crypto.randomUUID()}`, source: { sheet: from.sheet ?? sheetName, range: formatRange(from.range) },
      target: { sheet: to.sheet ?? sheetName, cell: formatAddress(to.range.start.row, to.range.start.col) },
      rows: [...rows.selectedOptions].map((o) => Number(o.value)), columns: [...columns.selectedOptions].map((o) => Number(o.value)),
      values: [{ field: Number(value.value), aggregate: aggregation.value as PivotAggregate }],
    });
  });
}
