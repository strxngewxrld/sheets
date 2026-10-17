import {
  detectDelimiter,
  formatFromFileName,
  LOCALE_DE,
  LOCALE_EN,
  openWorkbook,
  saveWorkbook,
  setDefaultLocale,
  type Workbook,
} from '@sheets/core';
import { SpreadsheetAgent } from '@sheets/agent';
import { SpreadsheetView } from '@sheets/ui';
import type { HostToWebview, InitMessage, WebviewToHost } from '../protocol.js';

declare function acquireVsCodeApi(): { postMessage(msg: WebviewToHost): void };

const vscode = acquireVsCodeApi();
const app = document.getElementById('app')!;
let view: SpreadsheetView | null = null;
let agent: SpreadsheetAgent | null = null;
let offChange: (() => void) | null = null;
let fileName = 'Book1.xlsx';
let csvDelimiter = ',';

function post(msg: WebviewToHost): void {
  vscode.postMessage(msg);
}

function attach(wb: Workbook): void {
  offChange?.();
  // Report edits to VS Code so it tracks dirty state and routes undo/redo to us.
  offChange = wb.events.on('change', (e) => {
    if (e.origin === 'undo' || e.origin === 'redo') return;
    post({ type: 'edit', label: e.origin === 'agent' ? 'Agent edit' : 'Edit' });
  });
}

function init(msg: InitMessage): void {
  fileName = msg.fileName;
  const de = msg.language.toLowerCase().startsWith('de');
  const isCsv = formatFromFileName(fileName) !== 'xlsx';
  if (isCsv) csvDelimiter = /\.tsv$/i.test(fileName) ? '\t' : detectDelimiter(new TextDecoder().decode(msg.bytes.slice(0, 20_000)));
  const locale = de || csvDelimiter === ';' ? LOCALE_DE : LOCALE_EN;
  setDefaultLocale(locale);
  let wb: Workbook;
  try {
    wb = openWorkbook(msg.bytes, fileName, { csv: { delimiter: csvDelimiter, locale: csvDelimiter === ';' ? LOCALE_DE : LOCALE_EN } });
  } catch (e) {
    post({ type: 'error', message: `Could not open ${fileName}: ${(e as Error).message}` });
    return;
  }
  if (!view) {
    view = new SpreadsheetView(app, wb, { locale, language: msg.language, keyboardUndo: false, readOnly: msg.readOnly });
    agent = new SpreadsheetAgent(() => view!.workbook, { host: view.agentHost });
  } else {
    view.setWorkbook(wb);
  }
  attach(wb);
  view.focus();
}

window.addEventListener('message', (event: MessageEvent<HostToWebview>) => {
  const msg = event.data;
  switch (msg.type) {
    case 'init':
      init(msg);
      break;
    case 'undo':
      view?.workbook.undo();
      break;
    case 'redo':
      view?.workbook.redo();
      break;
    case 'request':
      try {
        if (!view || !agent) throw new Error('Workbook not loaded');
        if (msg.method === 'getFileData') {
          const format = formatFromFileName(fileName);
          const bytes = saveWorkbook(view.workbook, format, {
            csv: { delimiter: csvDelimiter, locale: csvDelimiter === ';' ? LOCALE_DE : LOCALE_EN, lineEnding: '\n' },
          });
          post({ type: 'response', id: msg.id, result: bytes });
        } else {
          post({ type: 'response', id: msg.id, result: agent.execute(msg.params.name, msg.params.input) });
        }
      } catch (e) {
        post({ type: 'response', id: msg.id, error: (e as Error).message });
      }
      break;
  }
});

post({ type: 'ready' });
