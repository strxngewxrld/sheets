import { findCells, replaceCells, type FindOptions, type FindResult, type Workbook } from '@sheets/core';
import type { Strings } from './i18n.js';

export interface FindHost {
  workbook(): Workbook;
  /** Sheet and cell the search continues from. */
  position(): { sheet: string; row: number; col: number };
  /** Selects and reveals a match. */
  reveal(match: FindResult): void;
  readOnly: boolean;
  onClose(): void;
}

export interface FindDialog {
  readonly element: HTMLFormElement;
  /** Shows or hides the replace row and focuses the search field. */
  setReplace(replace: boolean): void;
  /** Re-runs the search after the workbook changed. */
  refresh(): void;
  close(): void;
}

const ORDER = (wb: Workbook) => new Map(wb.sheets.map((s, i) => [s.name.toUpperCase(), i]));

/** Index of the first match after (or, backwards, before) the position. */
export function nextMatchIndex(wb: Workbook, matches: FindResult[], from: { sheet: string; row: number; col: number }, backwards: boolean): number {
  if (!matches.length) return -1;
  const order = ORDER(wb);
  const key = (sheet: string, row: number, col: number) => [order.get(sheet.toUpperCase()) ?? 0, row, col] as const;
  const here = key(from.sheet, from.row, from.col);
  const cmp = (m: FindResult) => {
    const k = key(m.sheet, m.row, m.col);
    return k[0] - here[0] || k[1] - here[1] || k[2] - here[2];
  };
  if (backwards) {
    for (let i = matches.length - 1; i >= 0; i--) if (cmp(matches[i]!) < 0) return i;
    return matches.length - 1;
  }
  const i = matches.findIndex((m) => cmp(m) > 0);
  return i < 0 ? 0 : i;
}

/** Floating find/replace panel, kept open while the user works in the grid. */
export function openFindDialog(root: HTMLElement, host: FindHost, s: Strings, replace: boolean): FindDialog {
  root.querySelector('.sx-find-dialog')?.remove();
  const form = document.createElement('form');
  form.className = 'sx-dialog sx-find-dialog';
  form.setAttribute('role', 'dialog');
  form.setAttribute('aria-label', s.findReplace);

  const row = (label: string, el: HTMLElement) => {
    const l = document.createElement('label');
    l.append(label, el);
    form.append(l);
    return l;
  };
  const text = (label: string) => {
    const input = document.createElement('input');
    input.type = 'text';
    input.spellcheck = false;
    input.setAttribute('aria-label', label);
    return input;
  };
  const check = (label: string) => {
    const l = document.createElement('label');
    l.className = 'sx-check';
    const input = document.createElement('input');
    input.type = 'checkbox';
    l.append(input, label);
    return { label: l, input };
  };
  const button = (label: string, action: () => void, primary = false) => {
    const b = document.createElement('button');
    b.type = 'button';
    b.textContent = label;
    if (primary) b.className = 'sx-primary';
    b.addEventListener('click', action);
    return b;
  };

  const heading = document.createElement('h3');
  heading.textContent = s.findReplace;
  form.append(heading);
  const query = text(s.findWhat);
  row(s.findWhat, query);
  const replacement = text(s.replaceWith);
  const replaceRow = row(s.replaceWith, replacement);
  const scope = document.createElement('select');
  scope.add(new Option(s.scopeSheet, 'sheet'));
  scope.add(new Option(s.scopeWorkbook, 'workbook'));
  row(s.searchIn, scope);
  const matchCase = check(s.matchCase);
  const wholeCell = check(s.wholeCell);
  const inFormulas = check(s.inFormulas);
  const regex = check(s.useRegex);
  form.append(matchCase.label, wholeCell.label, inFormulas.label, regex.label);
  const status = document.createElement('div');
  status.className = 'sx-find-status';
  status.setAttribute('aria-live', 'polite');

  let matches: FindResult[] = [];
  let index = -1;
  let message = '';

  const options = (): FindOptions => {
    const o: FindOptions = { limit: 100_000 };
    if (scope.value === 'sheet') o.sheet = host.position().sheet;
    if (matchCase.input.checked) o.caseSensitive = true;
    if (wholeCell.input.checked) o.wholeCell = true;
    if (inFormulas.input.checked) o.inFormulas = true;
    if (regex.input.checked) o.regex = true;
    return o;
  };
  const search = () => {
    message = '';
    try {
      matches = query.value ? findCells(host.workbook(), query.value, options()) : [];
      query.removeAttribute('aria-invalid');
    } catch {
      matches = [];
      query.setAttribute('aria-invalid', 'true');
    }
    if (index >= matches.length) index = -1;
  };
  const update = () => {
    const text = !query.value ? '' : !matches.length ? s.noMatches : index >= 0 ? s.matchCount.replace('{0}', String(index + 1)).replace('{1}', String(matches.length)) : String(matches.length);
    status.textContent = [message, text].filter(Boolean).join(' · ');
    for (const b of [replaceOne, replaceEvery]) b.disabled = host.readOnly || !matches.length;
  };
  const go = (backwards: boolean) => {
    search();
    index = nextMatchIndex(host.workbook(), matches, host.position(), backwards);
    if (index >= 0) host.reveal(matches[index]!);
    update();
  };
  const isCurrent = (m: FindResult | undefined) => {
    const p = host.position();
    return !!m && m.sheet.toUpperCase() === p.sheet.toUpperCase() && m.row === p.row && m.col === p.col;
  };
  const replaceCurrent = () => {
    if (host.readOnly) return;
    search();
    const current = matches.find(isCurrent);
    if (!current) { go(false); return; }
    const count = runReplace([{ sheet: current.sheet, row: current.row, col: current.col }]);
    if (count) go(false);
    else update();
  };
  const runReplace = (cells?: { sheet: string; row: number; col: number }[]) => {
    try {
      const count = replaceCells(host.workbook(), query.value, replacement.value, { ...options(), ...(cells ? { cells } : {}) });
      message = s.replacedCount.replace('{0}', String(count));
      return count;
    } catch (e) {
      message = e instanceof Error ? e.message : String(e);
      return 0;
    }
  };
  const replaceAllMatches = () => {
    if (host.readOnly) return;
    runReplace();
    const keep = message;
    search();
    message = keep;
    index = -1;
    update();
  };

  const nav = document.createElement('div');
  nav.className = 'sx-actions';
  const previous = button(s.findPrevious, () => go(true));
  const next = button(s.findNext, () => go(false), true);
  const replaceOne = button(s.replace, replaceCurrent);
  const replaceEvery = button(s.replaceAll, replaceAllMatches);
  const replaceActions = document.createElement('div');
  replaceActions.className = 'sx-actions';
  replaceActions.append(replaceOne, replaceEvery);
  nav.append(button(s.close, () => dialog.close()), previous, next);
  form.append(status, replaceActions, nav);

  for (const el of [scope, matchCase.input, wholeCell.input, inFormulas.input, regex.input]) el.addEventListener('change', () => { index = -1; search(); update(); });
  query.addEventListener('input', () => { index = -1; search(); update(); });
  form.addEventListener('submit', (e) => { e.preventDefault(); go(false); });
  form.addEventListener('keydown', (e) => {
    e.stopPropagation();
    if (e.key === 'Escape') { e.preventDefault(); dialog.close(); }
    else if (e.key === 'Enter' && (e.target as HTMLElement).tagName === 'INPUT') {
      e.preventDefault();
      if (e.target === replacement && !e.shiftKey) replaceCurrent();
      else go(e.shiftKey);
    }
  });
  form.addEventListener('mousedown', (e) => e.stopPropagation());

  const dialog: FindDialog = {
    element: form,
    setReplace(on: boolean) {
      replaceRow.hidden = !on;
      replaceActions.hidden = !on || host.readOnly;
      heading.textContent = on ? s.findReplace : s.find;
      query.focus();
      query.select();
    },
    refresh() {
      if (!form.isConnected) return;
      search();
      update();
    },
    close() {
      if (!form.isConnected) return;
      form.remove();
      host.onClose();
    },
  };
  root.append(form);
  dialog.setReplace(replace);
  update();
  return dialog;
}
