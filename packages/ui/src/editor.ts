import { getFunction, listFunctions, parseRange, tokenize, type CellStyle, type RangeAddress } from '@sheets/core';
import { cssFont, type GridTheme, type Highlight } from './grid.js';

export const REF_COLORS = ['#1a73e8', '#e8710a', '#188038', '#a142f4', '#d01884', '#007b83', '#b06000', '#c5221f'];

/** Colored reference ranges in a formula being edited (for the current sheet). */
export function formulaHighlights(text: string, sheetName: string): Highlight[] {
  if (!text.startsWith('=')) return [];
  const out: Highlight[] = [];
  let i = 0;
  for (const tok of tokenize(text.slice(1))) {
    if (tok.type !== 'ref') continue;
    const parsed = parseRange(tok.text);
    const color = REF_COLORS[i++ % REF_COLORS.length]!;
    if (!parsed || (parsed.sheet && parsed.sheet.toUpperCase() !== sheetName.toUpperCase())) continue;
    out.push({ range: parsed.range, color });
  }
  return out;
}

type InputEl = HTMLTextAreaElement | HTMLInputElement;

/**
 * In-cell editor. The textarea doubles as the keyboard focus sink of the
 * grid while not editing (it is invisible then), so typing starts editing
 * and IME composition works.
 */
export class CellEditor {
  readonly textarea: HTMLTextAreaElement;
  private suggest: HTMLDivElement;
  private hint: HTMLDivElement;
  editing = false;
  /** "enter": arrow keys commit (typing over a cell); "edit": arrows move the caret (F2 / double click). */
  mode: 'enter' | 'edit' = 'enter';
  cell = { row: 0, col: 0 };
  original = '';
  /** Range of the last inserted reference in the active input (for replacing while dragging). */
  private refSpan: { start: number; end: number } | null = null;
  private suggestions: string[] = [];
  private suggestIndex = 0;
  /** The input currently being typed into (cell editor or formula bar). */
  activeInput: InputEl;

  constructor(parent: HTMLElement) {
    this.textarea = document.createElement('textarea');
    this.textarea.className = 'sx-editor sx-hidden';
    this.textarea.spellcheck = false;
    this.textarea.setAttribute('autocomplete', 'off');
    this.textarea.setAttribute('aria-label', 'Cell editor');
    this.textarea.wrap = 'off';
    parent.appendChild(this.textarea);
    this.suggest = document.createElement('div');
    this.suggest.className = 'sx-suggest';
    this.suggest.hidden = true;
    this.suggest.setAttribute('role', 'listbox');
    parent.appendChild(this.suggest);
    this.hint = document.createElement('div');
    this.hint.className = 'sx-hint';
    this.hint.hidden = true;
    parent.appendChild(this.hint);
    this.activeInput = this.textarea;
  }

  get value(): string {
    return this.activeInput.value;
  }

  begin(row: number, col: number, original: string, initial: string, mode: 'enter' | 'edit', input: InputEl = this.textarea): void {
    this.editing = true;
    this.mode = mode;
    this.cell = { row, col };
    this.original = original;
    this.activeInput = input;
    this.textarea.value = initial;
    if (input !== this.textarea) input.value = initial;
    this.textarea.classList.remove('sx-hidden');
    this.refSpan = null;
    const end = initial.length;
    input.setSelectionRange(end, end);
  }

  end(): void {
    this.editing = false;
    this.textarea.value = '';
    this.textarea.classList.add('sx-hidden');
    this.activeInput = this.textarea;
    this.refSpan = null;
    this.closeSuggest();
    this.hint.hidden = true;
  }

  /** Positions the editor over a cell rect (grid container coordinates). */
  place(rect: { x: number; y: number; w: number; h: number }, style: CellStyle | undefined, theme: GridTheme, maxWidth: number, zoom = 1): void {
    const ta = this.textarea;
    if (!this.editing) {
      ta.style.left = `${rect.x}px`;
      ta.style.top = `${rect.y}px`;
      return;
    }
    ta.style.font = cssFont(style, theme, zoom);
    ta.style.left = `${rect.x - 1}px`;
    ta.style.top = `${rect.y - 1}px`;
    ta.style.minWidth = `${rect.w + 2}px`;
    ta.style.minHeight = `${rect.h + 2}px`;
    ta.style.textAlign = 'left';
    ta.style.background = style?.fill ?? '';
    ta.style.color = style?.color ?? '';
    // Grow to fit the content.
    ta.style.width = '0px';
    ta.style.height = '0px';
    ta.style.width = `${Math.min(maxWidth - rect.x, Math.max(rect.w + 2, ta.scrollWidth + 6))}px`;
    ta.style.height = `${Math.max(rect.h + 2, ta.scrollHeight + 2)}px`;
    this.positionPopups();
  }

  syncFrom(input: InputEl): void {
    if (input !== this.textarea) this.textarea.value = input.value;
  }

  // -------------------------------------------------------------------------
  // Reference picking
  // -------------------------------------------------------------------------

  /** True if clicking a cell should insert a reference at the caret. */
  canInsertRef(): boolean {
    if (!this.editing) return false;
    const input = this.activeInput;
    const text = input.value;
    if (!text.startsWith('=')) return false;
    const pos = input.selectionStart ?? text.length;
    if (this.refSpan && this.refSpan.end === pos) return true;
    const before = text.slice(0, pos).trimEnd();
    return /[=(,;+\-*/^&<>:]$/.test(before);
  }

  insertRef(ref: string): void {
    const input = this.activeInput;
    const text = input.value;
    let start: number;
    let end: number;
    if (this.refSpan && this.refSpan.end === (input.selectionStart ?? -1)) {
      ({ start, end } = this.refSpan);
    } else {
      start = end = input.selectionStart ?? text.length;
    }
    input.value = text.slice(0, start) + ref + text.slice(end);
    const caret = start + ref.length;
    input.setSelectionRange(caret, caret);
    this.refSpan = { start, end: caret };
    this.syncFrom(input);
  }

  clearRefSpan(): void {
    this.refSpan = null;
  }

  // -------------------------------------------------------------------------
  // Function autocomplete & argument hints
  // -------------------------------------------------------------------------

  updateAssist(): void {
    const input = this.activeInput;
    const text = input.value;
    const pos = input.selectionStart ?? text.length;
    if (!this.editing || !text.startsWith('=')) {
      this.closeSuggest();
      this.hint.hidden = true;
      return;
    }
    const m = /(?:^=|[=(,;+\-*/^&<>:\s])([A-Za-z][A-Za-z0-9.]*)$/.exec(text.slice(0, pos));
    if (m) {
      const prefix = m[1]!.toUpperCase();
      this.suggestions = listFunctions()
        .filter((f) => f.name.startsWith(prefix))
        .slice(0, 10)
        .map((f) => f.name);
      if (this.suggestions.length && !(this.suggestions.length === 1 && this.suggestions[0] === prefix)) {
        this.suggestIndex = 0;
        this.renderSuggest();
      } else this.closeSuggest();
    } else this.closeSuggest();

    // Innermost open function call before the caret.
    const stack: string[] = [];
    let pending = '';
    for (const tok of tokenize(text.slice(1, pos))) {
      if (tok.type === 'func') pending = tok.text;
      else if (tok.type === 'lparen') {
        stack.push(pending);
        pending = '';
      } else if (tok.type === 'rparen') stack.pop();
      else if (tok.type !== 'ws') pending = '';
    }
    const fn = stack.length ? getFunction(stack[stack.length - 1]!) : undefined;
    if (fn && this.suggest.hidden) {
      this.hint.textContent = fn.syntax;
      this.hint.hidden = false;
    } else this.hint.hidden = true;
    this.positionPopups();
  }

  get suggestOpen(): boolean {
    return !this.suggest.hidden;
  }

  moveSuggest(delta: number): void {
    if (!this.suggestions.length) return;
    this.suggestIndex = (this.suggestIndex + delta + this.suggestions.length) % this.suggestions.length;
    this.renderSuggest();
  }

  acceptSuggest(name = this.suggestions[this.suggestIndex]): void {
    if (!name) return;
    const input = this.activeInput;
    const pos = input.selectionStart ?? input.value.length;
    const before = input.value.slice(0, pos).replace(/[A-Za-z][A-Za-z0-9.]*$/, '');
    const after = input.value.slice(pos);
    input.value = `${before}${name}(${after}`;
    const caret = before.length + name.length + 1;
    input.setSelectionRange(caret, caret);
    this.syncFrom(input);
    this.closeSuggest();
    this.updateAssist();
    input.dispatchEvent(new Event('sx-change'));
  }

  closeSuggest(): void {
    this.suggest.hidden = true;
    this.suggestions = [];
  }

  private renderSuggest(): void {
    this.suggest.replaceChildren(
      ...this.suggestions.map((name, i) => {
        const item = document.createElement('div');
        item.setAttribute('role', 'option');
        item.setAttribute('aria-selected', String(i === this.suggestIndex));
        item.textContent = name;
        const small = document.createElement('small');
        small.textContent = getFunction(name)?.description ?? '';
        item.appendChild(small);
        item.addEventListener('mousedown', (e) => {
          e.preventDefault();
          this.acceptSuggest(name);
        });
        return item;
      }),
    );
    this.suggest.hidden = false;
    this.hint.hidden = true;
  }

  private positionPopups(): void {
    const ta = this.textarea;
    const left = parseFloat(ta.style.left) || 0;
    const top = (parseFloat(ta.style.top) || 0) + ta.offsetHeight + 2;
    this.suggest.style.left = `${left}px`;
    this.suggest.style.top = `${top}px`;
    this.hint.style.left = `${left}px`;
    this.hint.style.top = `${top}px`;
  }

  /** Visible reference range of the current text, for ref-aware arrow navigation. */
  currentRefRange(): RangeAddress | null {
    if (!this.refSpan) return null;
    return parseRange(this.activeInput.value.slice(this.refSpan.start, this.refSpan.end))?.range ?? null;
  }
}
