/**
 * Component styles. Colors are CSS variables so hosts can theme the view;
 * inside VS Code they fall back to the editor theme variables.
 */
export const CSS = `
.sx-root {
  --sx-bg: var(--vscode-editor-background, #ffffff);
  --sx-fg: var(--vscode-editor-foreground, #1f1f1f);
  --sx-muted: var(--vscode-descriptionForeground, #616161);
  --sx-grid: var(--vscode-editorGroup-border, #e1e1e1);
  --sx-header-bg: var(--vscode-sideBar-background, #f5f5f5);
  --sx-header-fg: var(--vscode-sideBar-foreground, #444444);
  --sx-header-active-bg: var(--vscode-list-inactiveSelectionBackground, #dfe7f3);
  --sx-header-active-fg: var(--vscode-list-activeSelectionForeground, #0f4c81);
  --sx-accent: var(--vscode-focusBorder, #1a73e8);
  --sx-selection: rgba(26, 115, 232, 0.12);
  --sx-toolbar-bg: var(--vscode-editorWidget-background, #fafafa);
  --sx-input-bg: var(--vscode-input-background, #ffffff);
  --sx-input-fg: var(--vscode-input-foreground, #1f1f1f);
  --sx-input-border: var(--vscode-input-border, #d0d0d0);
  --sx-button-hover: var(--vscode-toolbar-hoverBackground, rgba(0,0,0,0.07));
  --sx-button-active: var(--vscode-inputOption-activeBackground, rgba(26,115,232,0.18));
  --sx-menu-bg: var(--vscode-menu-background, #ffffff);
  --sx-menu-fg: var(--vscode-menu-foreground, #1f1f1f);
  --sx-menu-hover: var(--vscode-menu-selectionBackground, #e8f0fe);
  --sx-shadow: 0 4px 16px rgba(0,0,0,0.18);
  --sx-font: var(--vscode-font-family, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif);
  display: flex;
  flex-direction: column;
  height: 100%;
  width: 100%;
  min-height: 0;
  background: var(--sx-bg);
  color: var(--sx-fg);
  font-family: var(--sx-font);
  font-size: 13px;
  overflow: hidden;
  position: relative;
  user-select: none;
  -webkit-user-select: none;
}
@media (prefers-color-scheme: dark) {
  .sx-root:not([data-theme="light"]) {
    --sx-bg: var(--vscode-editor-background, #1e1e1e);
    --sx-fg: var(--vscode-editor-foreground, #e3e3e3);
    --sx-muted: var(--vscode-descriptionForeground, #a0a0a0);
    --sx-grid: var(--vscode-editorGroup-border, #3a3a3a);
    --sx-header-bg: var(--vscode-sideBar-background, #252526);
    --sx-header-fg: var(--vscode-sideBar-foreground, #c8c8c8);
    --sx-header-active-bg: var(--vscode-list-inactiveSelectionBackground, #37404d);
    --sx-header-active-fg: var(--vscode-list-activeSelectionForeground, #9cc3ff);
    --sx-accent: var(--vscode-focusBorder, #4f9cff);
    --sx-selection: rgba(79, 156, 255, 0.18);
    --sx-toolbar-bg: var(--vscode-editorWidget-background, #252526);
    --sx-input-bg: var(--vscode-input-background, #2d2d2d);
    --sx-input-fg: var(--vscode-input-foreground, #e3e3e3);
    --sx-input-border: var(--vscode-input-border, #474747);
    --sx-button-hover: var(--vscode-toolbar-hoverBackground, rgba(255,255,255,0.08));
    --sx-button-active: var(--vscode-inputOption-activeBackground, rgba(79,156,255,0.25));
    --sx-menu-bg: var(--vscode-menu-background, #2b2b2b);
    --sx-menu-fg: var(--vscode-menu-foreground, #e3e3e3);
    --sx-menu-hover: var(--vscode-menu-selectionBackground, #094771);
  }
}
.sx-root[data-theme="dark"] {
  --sx-bg: #1e1e1e; --sx-fg: #e3e3e3; --sx-muted: #a0a0a0; --sx-grid: #3a3a3a; --sx-header-bg: #252526; --sx-header-fg: #c8c8c8;
  --sx-header-active-bg: #37404d; --sx-header-active-fg: #9cc3ff; --sx-accent: #4f9cff; --sx-selection: rgba(79,156,255,0.18);
  --sx-toolbar-bg: #252526; --sx-input-bg: #2d2d2d; --sx-input-fg: #e3e3e3; --sx-input-border: #474747;
  --sx-button-hover: rgba(255,255,255,0.08); --sx-button-active: rgba(79,156,255,0.25); --sx-menu-bg: #2b2b2b; --sx-menu-fg: #e3e3e3; --sx-menu-hover: #094771;
}
.sx-toolbar {
  display: flex; flex-wrap: wrap; align-items: center; gap: 2px; padding: 4px 6px;
  background: var(--sx-toolbar-bg); border-bottom: 1px solid var(--sx-grid);
}
.sx-toolbar .sx-sep { width: 1px; height: 20px; background: var(--sx-grid); margin: 0 4px; }
.sx-btn {
  min-width: 28px; height: 28px; padding: 0 6px; border: 0; border-radius: 4px; background: transparent; color: inherit;
  font: inherit; cursor: pointer; display: inline-flex; align-items: center; justify-content: center; gap: 4px;
}
.sx-btn:hover { background: var(--sx-button-hover); }
.sx-btn[aria-pressed="true"] { background: var(--sx-button-active); }
.sx-btn:disabled { opacity: 0.4; cursor: default; }
.sx-btn:focus-visible, .sx-select:focus-visible { outline: 2px solid var(--sx-accent); outline-offset: -2px; }
.sx-btn svg { width: 16px; height: 16px; }
.sx-select {
  height: 26px; border: 1px solid var(--sx-input-border); border-radius: 4px; background: var(--sx-input-bg); color: var(--sx-input-fg);
  font: inherit; padding: 0 4px;
}
.sx-color { position: relative; }
.sx-color input[type=color] { position: absolute; inset: 0; opacity: 0; cursor: pointer; width: 100%; height: 100%; border: 0; padding: 0; }
.sx-color .sx-swatch { display: block; width: 16px; height: 3px; margin-top: 1px; border-radius: 1px; }
.sx-color-inner { display: flex; flex-direction: column; align-items: center; line-height: 1; }
.sx-formulabar { display: flex; align-items: stretch; border-bottom: 1px solid var(--sx-grid); height: 28px; }
.sx-namebox {
  width: 96px; border: 0; border-right: 1px solid var(--sx-grid); background: var(--sx-bg); color: var(--sx-fg);
  font: inherit; padding: 0 8px; outline: none;
}
.sx-fx { display: flex; align-items: center; padding: 0 8px; color: var(--sx-muted); font-style: italic; border-right: 1px solid var(--sx-grid); }
.sx-formula {
  flex: 1; border: 0; background: var(--sx-bg); color: var(--sx-fg); font: 13px ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
  padding: 0 8px; outline: none; min-width: 0; user-select: text; -webkit-user-select: text;
}
.sx-namebox:focus, .sx-formula:focus { box-shadow: inset 0 -2px 0 var(--sx-accent); }
.sx-grid { position: relative; flex: 1; min-height: 0; overflow: hidden; outline: none; }
.sx-grid canvas { position: absolute; left: 0; top: 0; }
.sx-scroller { position: absolute; overflow: auto; }
.sx-spacer { pointer-events: none; }
.sx-charts { position: absolute; overflow: hidden; pointer-events: none; }
.sx-charts-inner { position: absolute; left: 0; top: 0; }
.sx-chart {
  position: absolute; pointer-events: auto; background: var(--sx-bg); border: 1px solid var(--sx-grid);
  box-shadow: 0 1px 3px rgba(0,0,0,0.08); cursor: move; border-radius: 2px;
}
.sx-chart.sx-selected { outline: 2px solid var(--sx-accent); }
.sx-chart svg { display: block; width: 100%; height: 100%; }
.sx-chart .sx-resize {
  position: absolute; right: -5px; bottom: -5px; width: 10px; height: 10px; background: var(--sx-accent);
  border: 2px solid var(--sx-bg); border-radius: 50%; cursor: nwse-resize; display: none;
}
.sx-chart.sx-selected .sx-resize { display: block; }
.sx-editor {
  position: absolute; z-index: 5; margin: 0; padding: 2px 4px; border: 2px solid var(--sx-accent); resize: none; outline: none;
  background: var(--sx-input-bg); color: var(--sx-input-fg); overflow: hidden; white-space: pre; line-height: 1.25;
  box-sizing: border-box; user-select: text; -webkit-user-select: text;
}
.sx-editor.sx-hidden { opacity: 0; pointer-events: none; width: 1px !important; height: 1px !important; padding: 0; border: 0; }
.sx-suggest {
  position: absolute; z-index: 20; min-width: 220px; max-height: 220px; overflow: auto; background: var(--sx-menu-bg); color: var(--sx-menu-fg);
  border: 1px solid var(--sx-grid); border-radius: 4px; box-shadow: var(--sx-shadow); font-size: 12px; padding: 2px 0;
}
.sx-suggest div { padding: 3px 8px; cursor: pointer; white-space: nowrap; }
.sx-suggest div small { color: var(--sx-muted); margin-left: 6px; }
.sx-suggest div[aria-selected="true"], .sx-suggest div:hover { background: var(--sx-menu-hover); }
.sx-hint {
  position: absolute; z-index: 19; background: var(--sx-menu-bg); color: var(--sx-menu-fg); border: 1px solid var(--sx-grid);
  border-radius: 4px; padding: 3px 8px; font: 12px ui-monospace, Menlo, Consolas, monospace; box-shadow: var(--sx-shadow); pointer-events: none;
}
.sx-bottom { display: flex; align-items: stretch; border-top: 1px solid var(--sx-grid); height: 30px; background: var(--sx-toolbar-bg); }
.sx-tabs { display: flex; align-items: stretch; overflow-x: auto; flex: 1; min-width: 0; scrollbar-width: thin; }
.sx-tab {
  padding: 0 14px; display: flex; align-items: center; cursor: pointer; border-right: 1px solid var(--sx-grid);
  white-space: nowrap; color: var(--sx-muted); background: transparent; border-top: 0; border-left: 0; border-bottom: 0; font: inherit;
}
.sx-tab:hover { background: var(--sx-button-hover); }
.sx-tab[aria-selected="true"] { color: var(--sx-fg); background: var(--sx-bg); box-shadow: inset 0 -2px 0 var(--sx-accent); font-weight: 600; }
.sx-tab input { font: inherit; width: 100px; border: 1px solid var(--sx-accent); background: var(--sx-input-bg); color: var(--sx-input-fg); }
.sx-status { display: flex; align-items: center; gap: 14px; padding: 0 12px; color: var(--sx-muted); font-size: 12px; white-space: nowrap; }
.sx-zoom { display: flex; align-items: center; gap: 2px; padding: 0 8px; margin-left: auto; border-left: 1px solid var(--sx-grid); }
.sx-zoom .sx-btn { min-width: 22px; height: 22px; padding: 0 4px; font-size: 12px; }
.sx-zoom .sx-zoom-level { min-width: 48px; color: var(--sx-muted); }
.sx-menu {
  position: fixed; z-index: 1000; min-width: 200px; background: var(--sx-menu-bg); color: var(--sx-menu-fg); border: 1px solid var(--sx-grid);
  border-radius: 6px; box-shadow: var(--sx-shadow); padding: 4px 0; font-size: 13px;
}
.sx-menu button {
  display: block; width: 100%; text-align: left; padding: 5px 14px; border: 0; background: transparent; color: inherit; font: inherit; cursor: pointer;
}
.sx-menu button:hover, .sx-menu button:focus { background: var(--sx-menu-hover); outline: none; }
.sx-menu hr { border: 0; border-top: 1px solid var(--sx-grid); margin: 4px 0; }
.sx-dialog {
  position: absolute; z-index: 50; right: 16px; top: 16px; width: 280px; background: var(--sx-menu-bg); color: var(--sx-menu-fg);
  border: 1px solid var(--sx-grid); border-radius: 8px; box-shadow: var(--sx-shadow); padding: 12px 12px 0; display: grid; gap: 8px; font-size: 13px;
  /* Never taller than the editor: the content scrolls and the buttons stay reachable at the bottom. */
  box-sizing: border-box; max-height: calc(100% - 32px); overflow-y: auto; overscroll-behavior: contain;
}
.sx-dialog h3 { margin: 0 0 4px; font-size: 14px; }
.sx-dialog label { display: grid; gap: 3px; font-size: 12px; color: var(--sx-muted); }
.sx-dialog label.sx-check { display: flex; align-items: center; gap: 6px; }
.sx-dialog input[type=text], .sx-dialog select {
  height: 26px; border: 1px solid var(--sx-input-border); border-radius: 4px; background: var(--sx-input-bg); color: var(--sx-input-fg); font: inherit; padding: 0 6px;
  user-select: text; -webkit-user-select: text;
}
.sx-dialog input[type=number] { height: 26px; border: 1px solid var(--sx-input-border); border-radius: 4px; background: var(--sx-input-bg); color: var(--sx-input-fg); font: inherit; padding: 0 6px; user-select: text; }
.sx-dialog .sx-actions { flex-wrap: wrap; }
.sx-format-preview { display: block; padding: 12px; border: 1px solid var(--sx-grid); border-radius: 4px; background: var(--sx-input-bg); color: var(--sx-input-fg); font-size: 16px; text-align: right; user-select: text; }
.sx-table-preview { width: 100%; border-collapse: collapse; font-size: 12px; }
.sx-table-preview td { padding: 5px 8px; border: 1px solid var(--sx-grid); }
.sx-sort-levels { display: grid; gap: 8px; }
.sx-sort-levels fieldset { display: grid; grid-template-columns: 1fr 28px; gap: 5px; padding: 8px; border: 1px solid var(--sx-grid); border-radius: 4px; min-width: 0; }
.sx-sort-levels legend { color: var(--sx-muted); font-size: 12px; }
.sx-sort-levels select { min-width: 0; grid-column: 1; }
.sx-sort-levels button { grid-column: 2; grid-row: 1 / span 2; border: 0; border-radius: 3px; background: var(--sx-button-hover); color: inherit; cursor: pointer; }
.sx-table-controls { position: absolute; inset: 0; pointer-events: none; z-index: 4; }
.sx-table-sort { position: absolute; pointer-events: auto; width: 16px; height: 18px; padding: 0; border: 1px solid currentColor; border-radius: 2px; background: transparent; font: 11px sans-serif; cursor: pointer; opacity: 0.85; }
.sx-table-sort:hover, .sx-table-sort:focus-visible { opacity: 1; background: rgba(128,128,128,0.3); outline: 1px solid var(--sx-accent); }
.sx-dialog .sx-actions { display: flex; gap: 6px; justify-content: flex-end; margin-top: 4px; position: sticky; bottom: 0; z-index: 1; padding: 8px 0 12px; background: var(--sx-menu-bg); }
.sx-dialog > :last-child:not(.sx-actions) { margin-bottom: 12px; }
.sx-dialog .sx-actions button { height: 28px; padding: 0 12px; border-radius: 4px; border: 1px solid var(--sx-input-border); background: var(--sx-input-bg); color: inherit; cursor: pointer; font: inherit; }
.sx-dialog .sx-actions button.sx-primary { background: var(--sx-accent); border-color: var(--sx-accent); color: #fff; }
.sx-validation-btn { position: absolute; z-index: 6; width: 18px; padding: 0; border: 1px solid var(--sx-input-border); border-radius: 2px; background: var(--sx-menu-bg); color: var(--sx-fg); font: 11px sans-serif; cursor: pointer; }
.sx-validation-btn:hover { background: var(--sx-menu-hover); }
.sx-validation-tip { position: absolute; z-index: 30; max-width: 280px; padding: 8px 10px; border: 1px solid var(--sx-grid); border-radius: 6px; background: var(--sx-menu-bg); color: var(--sx-menu-fg); box-shadow: var(--sx-shadow); font-size: 12px; display: grid; gap: 4px; white-space: pre-wrap; }
.sx-validation-tip[hidden] { display: none; }
.sx-validation-tip.sx-tip-prompt { pointer-events: none; background: #FFFFE1; color: #1f1f1f; border-color: #c8c8a0; }
.sx-validation-tip.sx-tip-error { border-left: 4px solid #C00000; }
.sx-validation-tip.sx-tip-warning { border-left: 4px solid #D9822B; }
.sx-validation-tip.sx-tip-info { border-left: 4px solid var(--sx-accent); }
.sx-validation-tip .sx-actions { display: flex; gap: 6px; justify-content: flex-end; margin-top: 4px; }
.sx-validation-tip .sx-actions button { height: 26px; padding: 0 10px; border-radius: 4px; border: 1px solid var(--sx-input-border); background: var(--sx-input-bg); color: inherit; cursor: pointer; font: inherit; }
.sx-validation-tip .sx-actions button.sx-primary { background: var(--sx-accent); border-color: var(--sx-accent); color: #fff; }
.sx-notice { display: flex; flex-wrap: wrap; gap: 4px 8px; padding: 5px 10px; font-size: 12px; background: #FFF4CE; color: #4A3B00; border-bottom: 1px solid #E8D48A; }
.sx-notice[hidden] { display: none; }
.sx-notice span { color: #6B5800; }
.sx-note { position: absolute; z-index: 29; max-width: 260px; padding: 6px 8px; background: #FFFFE1; color: #1f1f1f; border: 1px solid #c8c8a0; box-shadow: var(--sx-shadow); font-size: 12px; white-space: pre-wrap; pointer-events: none; display: grid; gap: 2px; }
.sx-note[hidden] { display: none; }
.sx-find-dialog { width: 300px; z-index: 51; }
.sx-find-dialog [hidden] { display: none; }
.sx-find-dialog input[aria-invalid=true] { border-color: #C00000; }
.sx-find-status { min-height: 16px; color: var(--sx-muted); font-size: 12px; }
.sx-analysis-dialog { max-height: calc(100% - 32px); overflow: auto; }
.sx-dialog [hidden] { display: none !important; }
.sx-analysis-dialog select[multiple] { height: auto; min-height: 72px; padding: 4px; }
.sx-analysis-dialog [role=alert] { color: #C00000; white-space: normal; }
@media (max-width: 600px) {
  .sx-namebox { width: 64px; }
  .sx-status .sx-optional { display: none; }
}
`;

let injected = false;

export function injectStyles(doc: Document = document): void {
  if (injected || doc.getElementById('sx-styles')) return;
  const style = doc.createElement('style');
  style.id = 'sx-styles';
  style.textContent = CSS;
  doc.head.appendChild(style);
  injected = true;
}
