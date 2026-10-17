import {
  detectDelimiter,
  formatFromFileName,
  LOCALE_DE,
  LOCALE_EN,
  openWorkbook,
  saveWorkbook,
  setDefaultLocale,
  Workbook,
  writeCsv,
} from '@sheets/core';
import { SpreadsheetAgent } from '@sheets/agent';
import { SpreadsheetView } from '@sheets/ui';
import type { HostToWebview, InitMessage, WebviewToHost } from '../protocol.js';
import { decodeBytes, encodeBytes } from '../binary.js';

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

let saveCheck: ReturnType<typeof setTimeout> | undefined;

/** Warns in the editor when the current changes would be refused on XLSX save. */
function scheduleSaveCheck(wb: Workbook): void {
  clearTimeout(saveCheck);
  if (formatFromFileName(fileName) !== 'xlsx' || !wb.xlsxSource) {
    view?.setNotice(null);
    return;
  }
  saveCheck = setTimeout(() => {
    if (view?.workbook !== wb) return;
    try {
      saveWorkbook(wb, 'xlsx');
      view.setNotice(null);
    } catch (e) {
      const message = (e as Error).message;
      view.setNotice(message.startsWith('XLSX-Schutz:') ? message.replace(/^XLSX-Schutz:\s*/, '').replace(/\s*Diese Änderung kann noch nicht.*$/s, '') : message);
    }
  }, 600);
}

function attach(wb: Workbook): void {
  offChange?.();
  scheduleSaveCheck(wb);
  // Report edits to VS Code so it tracks dirty state and routes undo/redo to us.
  offChange = wb.events.on('change', (e) => {
    scheduleSaveCheck(wb);
    if (e.origin === 'undo' || e.origin === 'redo') return;
    post({ type: 'edit', label: e.origin === 'agent' ? 'Agent edit' : 'Edit' });
  });
}

function init(msg: InitMessage): void {
  fileName = msg.fileName;
  try {
    const bytes = decodeBytes(msg.bytes);
    const backupText = msg.backup ? new TextDecoder().decode(bytes) : '';
    // Older extension versions backed up ordinary XLSX/CSV bytes.
    const snapshot = backupText.startsWith('{"type":"sheets-backup-v1",') ? JSON.parse(backupText) : undefined;
    const de = msg.language.toLowerCase().startsWith('de');
    const format = formatFromFileName(fileName);
    const isCsv = format === 'csv' || format === 'tsv';
    csvDelimiter = isCsv ? (format === 'tsv' ? '\t' : detectDelimiter(new TextDecoder().decode(bytes.subarray(0, 20_000)))) : ',';
    if (snapshot && typeof snapshot.csvDelimiter === 'string' && snapshot.csvDelimiter.length === 1) csvDelimiter = snapshot.csvDelimiter;
    const locale = de || csvDelimiter === ';' ? LOCALE_DE : LOCALE_EN;
    setDefaultLocale(locale);
    // A file newly created through the VS Code explorer has no ZIP container yet.
    // Initialize its model in memory; only an explicit save writes the XLSX file.
    if (snapshot && snapshot.type !== 'sheets-backup-v1') throw new Error('Invalid Sheets backup');
    const wb = snapshot ? Workbook.fromJSON(snapshot.workbook) : bytes.length === 0 && format === 'xlsx'
      ? new Workbook()
      : openWorkbook(bytes, fileName, { csv: { delimiter: csvDelimiter, locale: csvDelimiter === ';' ? LOCALE_DE : LOCALE_EN } });
    if (!view) {
      view = new SpreadsheetView(app, wb, { locale, language: msg.language, keyboardUndo: false, readOnly: msg.readOnly, onOpenLink: (url) => post({ type: 'openLink', url }) });
      agent = new SpreadsheetAgent(() => view!.workbook, { host: view.agentHost });
    } else {
      view.setWorkbook(wb);
    }
    attach(wb);
    view.focus();
    post({ type: 'initialized' });
  } catch (e) {
    post({ type: 'error', message: `Could not open ${fileName}: ${(e as Error).message}` });
  }
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
    case 'command':
      view?.runCommand(msg.command);
      break;
    case 'request':
      try {
        if (!view || !agent) throw new Error('Workbook not loaded');
        if (msg.method === 'getBackupData') {
          const bytes = new TextEncoder().encode(JSON.stringify({ type: 'sheets-backup-v1', csvDelimiter, workbook: view.workbook.toJSON() }));
          post({ type: 'response', id: msg.id, result: encodeBytes(bytes) });
        } else if (msg.method === 'getFileData') {
          const format = formatFromFileName(fileName);
          const bytes = saveWorkbook(view.workbook, format, {
            csv: { delimiter: csvDelimiter, locale: csvDelimiter === ';' ? LOCALE_DE : LOCALE_EN, lineEnding: '\n' },
          });
          post({ type: 'response', id: msg.id, result: encodeBytes(bytes) });
        } else if (msg.method === 'exportCsv') {
          const wb = view.workbook;
          const bytes = new TextEncoder().encode(writeCsv(wb, wb.activeSheetIndex, { delimiter: msg.params.delimiter, locale: msg.params.delimiter === ';' ? LOCALE_DE : LOCALE_EN, lineEnding: '\r\n' }));
          post({ type: 'response', id: msg.id, result: { sheet: wb.activeSheet.name, bytes: encodeBytes(bytes) } });
        } else if (msg.method === 'executeTool') {
          post({ type: 'response', id: msg.id, result: agent.execute(msg.params.name, msg.params.input) });
        }
      } catch (e) {
        post({ type: 'response', id: msg.id, error: (e as Error).message });
      }
      break;
  }
});

post({ type: 'ready' });
