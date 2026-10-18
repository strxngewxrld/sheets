import {
  csvUnsupportedContent,
  detectCsvLayout,
  detectDelimiter,
  formatFromFileName,
  isXlsxFormat,
  LOCALE_DE,
  LOCALE_EN,
  openWorkbook,
  saveWorkbook,
  setDefaultLocale,
  setXlsxMessageLanguage,
  Workbook,
  writeCsv,
  XlsxProtectionError,
  type CsvLayout,
} from '@sheets/core';
import { SpreadsheetAgent } from '@sheets/agent';
import { SpreadsheetView } from '@sheets/ui';
import type { HostToWebview, InitMessage, WebviewToHost } from '../protocol.js';
import { decodeBytes, encodeBytes } from '../binary.js';
import { fill, shellStrings } from './strings.js';

declare function acquireVsCodeApi(): { postMessage(msg: WebviewToHost): void };

const vscode = acquireVsCodeApi();
const app = document.getElementById('app')!;
let view: SpreadsheetView | null = null;
let agent: SpreadsheetAgent | null = null;
let offChange: (() => void) | null = null;
let fileName = 'Book1.xlsx';
let language = 'en';
let csvDelimiter = ',';
/** Line endings, final line break and BOM of the opened CSV, reproduced on save. */
let csvLayout: CsvLayout | undefined;
/** Content a CSV cannot hold that already existed when the file was opened. */
let csvBaseline: Record<string, number> | undefined;
let csvLossWarned = false;

function post(msg: WebviewToHost): void {
  vscode.postMessage(msg);
}

/** Replaces the editor area with a status panel (loading or error) while no workbook is shown. */
function showStatus(kind: 'loading' | 'error', title: string, detail: string): void {
  const panel = document.createElement('div');
  panel.className = `sheets-status sheets-status-${kind}`;
  panel.setAttribute('role', kind === 'error' ? 'alert' : 'status');
  const heading = document.createElement('strong');
  heading.textContent = title;
  const text = document.createElement('div');
  text.textContent = detail;
  panel.append(heading, text);
  app.replaceChildren(panel);
}

let saveCheck: ReturnType<typeof setTimeout> | undefined;

/** Warns in the editor when the current changes would be refused on XLSX save. */
function scheduleSaveCheck(wb: Workbook): void {
  clearTimeout(saveCheck);
  const format = formatFromFileName(fileName);
  if (!isXlsxFormat(format) || !wb.xlsxSource) {
    view?.setNotice(null);
    return;
  }
  saveCheck = setTimeout(() => {
    if (view?.workbook !== wb) return;
    try {
      saveWorkbook(wb, format);
      view.setNotice(null);
    } catch (e) {
      view.setNotice(e instanceof XlsxProtectionError ? e.reason : (e as Error).message);
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

/** Content that would be dropped by saving as CSV/TSV and did not exist when the file was opened. */
function csvLosses(wb: Workbook): string[] {
  const now = csvUnsupportedContent(wb);
  return Object.keys(now).filter((key) => now[key]! > (csvBaseline?.[key] ?? 0));
}

function init(msg: InitMessage): void {
  fileName = msg.fileName;
  language = msg.language;
  const strings = shellStrings(language);
  setXlsxMessageLanguage(language);
  try {
    const bytes = decodeBytes(msg.bytes);
    const backupText = msg.backup ? new TextDecoder().decode(bytes) : '';
    // Older extension versions backed up ordinary XLSX/CSV bytes.
    const snapshot = backupText.startsWith('{"type":"sheets-backup-v1",') ? JSON.parse(backupText) : undefined;
    const de = msg.language.toLowerCase().startsWith('de');
    const format = formatFromFileName(fileName);
    const isCsv = format === 'csv' || format === 'tsv';
    const head = isCsv && !snapshot ? new TextDecoder().decode(bytes.subarray(0, 20_000)) : '';
    csvDelimiter = isCsv ? (format === 'tsv' ? '\t' : detectDelimiter(head)) : ',';
    csvLayout = undefined;
    if (isCsv && !snapshot) {
      // The layout is taken from the start (BOM, first line break) and the last byte.
      const last = bytes[bytes.length - 1];
      const bom = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf;
      csvLayout = { ...detectCsvLayout(head), finalNewline: bytes.length === 0 || last === 10 || last === 13, bom };
    }
    if (snapshot && typeof snapshot.csvDelimiter === 'string' && snapshot.csvDelimiter.length === 1) csvDelimiter = snapshot.csvDelimiter;
    if (snapshot?.csvLayout && typeof snapshot.csvLayout === 'object') csvLayout = snapshot.csvLayout;
    const locale = de || csvDelimiter === ';' ? LOCALE_DE : LOCALE_EN;
    setDefaultLocale(locale);
    // A file newly created through the VS Code explorer has no ZIP container yet.
    // Initialize its model in memory; only an explicit save writes the XLSX file.
    if (snapshot && snapshot.type !== 'sheets-backup-v1') throw new Error('Invalid Sheets backup');
    const wb = snapshot ? Workbook.fromJSON(snapshot.workbook) : bytes.length === 0 && isXlsxFormat(format)
      ? new Workbook()
      : openWorkbook(bytes, fileName, { csv: { delimiter: csvDelimiter, locale: csvDelimiter === ';' ? LOCALE_DE : LOCALE_EN } });
    csvBaseline = isCsv ? (snapshot?.csvBaseline ?? csvUnsupportedContent(wb)) : undefined;
    csvLossWarned = !!snapshot?.csvLossWarned;
    if (!view) {
      app.replaceChildren();
      view = new SpreadsheetView(app, wb, { locale, language: msg.language, keyboardUndo: false, readOnly: msg.readOnly, onOpenLink: (url) => post({ type: 'openLink', url }) });
      agent = new SpreadsheetAgent(() => view!.workbook, { host: view.agentHost });
    } else {
      view.setWorkbook(wb);
    }
    attach(wb);
    view.focus();
    post({ type: 'initialized' });
  } catch (e) {
    const message = (e as Error).message;
    // Without a loaded workbook nothing can be saved; the host refuses saves for this document.
    offChange?.();
    offChange = null;
    view?.destroy();
    view = null;
    agent = null;
    showStatus('error', fill(strings.openFailed, fileName), `${message}\n${strings.openFailedDetail}`);
    post({ type: 'error', message: `Could not open ${fileName}: ${message}` });
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
          const bytes = new TextEncoder().encode(JSON.stringify({ type: 'sheets-backup-v1', csvDelimiter, csvLayout, csvBaseline, csvLossWarned, workbook: view.workbook.toJSON() }));
          post({ type: 'response', id: msg.id, result: encodeBytes(bytes) });
        } else if (msg.method === 'getFileData') {
          const target = msg.params?.fileName ?? fileName;
          const format = formatFromFileName(target);
          const sameFormat = format === formatFromFileName(fileName);
          const wb = view.workbook;
          // Another target format gets its own delimiter (TSV: tab, CSV: by UI language).
          const delimiter = sameFormat ? csvDelimiter : format === 'tsv' ? '\t' : language.toLowerCase().startsWith('de') ? ';' : ',';
          const bytes = saveWorkbook(wb, format, {
            csv: { delimiter, locale: delimiter === ';' ? LOCALE_DE : LOCALE_EN, ...(sameFormat && csvLayout ? csvLayout : { lineEnding: '\n' }) },
          });
          if ((format === 'csv' || format === 'tsv') && !csvLossWarned) {
            const lost = sameFormat ? csvLosses(wb) : Object.entries(csvUnsupportedContent(wb)).filter(([, n]) => n > 0).map(([key]) => key);
            if (lost.length) {
              // Once per document: saving as CSV is allowed, but the user learns what is not kept.
              csvLossWarned = true;
              const strings = shellStrings(language);
              post({ type: 'warning', message: fill(strings.csvLoss, target.split(/[\\/]/).pop() ?? target, format.toUpperCase(), wb.activeSheet.name, lost.map((key) => strings.csvLossParts[key] ?? key).join(', ')) });
            }
          }
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
