import { describe, expect, it, vi } from 'vitest';
import { SpreadsheetView } from '../src/view.js';

// Exercise shortcut dispatch without constructing a canvas or DOM.
function key(event: Partial<KeyboardEvent>) {
  const view = Object.assign(Object.create(SpreadsheetView.prototype), {
    editor: { editing: false }, openFind: vi.fn(), pasteSpecial: vi.fn(),
    options: { keyboardUndo: false },
  });
  const preventDefault = vi.fn();
  view.onKeyDown({ ctrlKey: false, metaKey: false, shiftKey: false, altKey: false,
    key: '', code: '', getModifierState: () => false, preventDefault, ...event }, {});
  return { view, preventDefault };
}

describe('cross-platform keyboard input', () => {
  it.each(['KeyF', 'KeyH', 'KeyV'])('does not consume AltGr text input on %s', (code) => {
    const { view, preventDefault } = key({ code, ctrlKey: true, altKey: true, getModifierState: (name) => name === 'AltGraph' });
    expect(preventDefault).not.toHaveBeenCalled();
    expect(view.openFind).not.toHaveBeenCalled();
    expect(view.pasteSpecial).not.toHaveBeenCalled();
  });

  it('leaves composition input to the input method', () => {
    expect(key({ key: 'Enter', isComposing: true }).preventDefault).not.toHaveBeenCalled();
  });

  it.each([{ ctrlKey: true }, { metaKey: true }])('retains Ctrl/Cmd find shortcuts: %o', (mod) => {
    const { view, preventDefault } = key({ ...mod, code: 'KeyF' });
    expect(preventDefault).toHaveBeenCalled();
    expect(view.openFind).toHaveBeenCalledWith(false);
  });

  it('retains Cmd+Option+F and Ctrl+Alt+V', () => {
    expect(key({ metaKey: true, altKey: true, code: 'KeyF' }).view.openFind).toHaveBeenCalledWith(true);
    expect(key({ ctrlKey: true, altKey: true, code: 'KeyV' }).view.pasteSpecial).toHaveBeenCalled();
  });

  it.each([
    { ctrlKey: true, key: 'z' },
    { ctrlKey: true, key: 'y' },
    { ctrlKey: true, shiftKey: true, key: 'z' },
    { metaKey: true, key: 'z' },
    { metaKey: true, shiftKey: true, key: 'z' },
    { metaKey: true, key: 'y' },
  ])('leaves host undo/redo to VS Code for %o', event => {
    expect(key(event).preventDefault).not.toHaveBeenCalled();
  });
});
