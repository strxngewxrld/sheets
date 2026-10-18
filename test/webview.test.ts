import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { formatTable, openWorkbook, saveWorkbook, setTableFilter, setValues, tableFilteredRows, toRange, Workbook } from '@sheets/core';
import { decodeBytes, encodeBytes } from '../src/binary.js';
import type { HostToWebview, WebviewToHost } from '../src/protocol.js';
import { strToU8, unzipSync, zipSync } from 'fflate';

const bridge = vi.hoisted(() => ({
  post: vi.fn(),
  commands: [] as string[],
  options: undefined as undefined | { onOpenLink?: (url: string) => void },
  notices: [] as (string | null)[],
  receive: undefined as ((event: { data: HostToWebview }) => void) | undefined,
  app: undefined as FakeElement | undefined,
}));

/** Just enough DOM for the status panel of the webview shell. */
class FakeElement {
  children: FakeElement[] = [];
  className = '';
  textContent = '';
  attributes: Record<string, string> = {};
  setAttribute(name: string, value: string) { this.attributes[name] = value; }
  append(...nodes: FakeElement[]) { this.children.push(...nodes); }
  replaceChildren(...nodes: FakeElement[]) { this.children = nodes; }
  text(): string { return [this.textContent, ...this.children.map((c) => c.text())].filter(Boolean).join(' '); }
}

// Keep the real webview message handler, workbook, I/O and agent. Only the
// canvas/DOM view is replaced so this regression runs without VS Code itself.
vi.mock('@sheets/ui', () => ({
  SpreadsheetView: class {
    agentHost = {};
    constructor(_app: unknown, public workbook: Workbook, options: { onOpenLink?: (url: string) => void }) { bridge.options = options; }
    setWorkbook(wb: Workbook) { this.workbook = wb; }
    focus() {}
    runCommand(command: string) { bridge.commands.push(command); }
    setNotice(text: string | null) { bridge.notices.push(text); }
    destroy() {}
  },
}));

function send(message: HostToWebview): void {
  bridge.receive!({ data: JSON.parse(JSON.stringify(message)) });
}

function response(id: number): Extract<WebviewToHost, { type: 'response' }> {
  const messages = bridge.post.mock.calls.map(([m]) => m as WebviewToHost);
  const result = messages.find((m) => m.type === 'response' && m.id === id);
  if (!result || result.type !== 'response') throw new Error('Missing webview response');
  return JSON.parse(JSON.stringify(result));
}

beforeEach(async () => {
  vi.resetModules();
  bridge.post.mockClear();
  bridge.commands.length = 0;
  bridge.notices.length = 0;
  bridge.receive = undefined;
  vi.stubGlobal('acquireVsCodeApi', () => ({ postMessage: bridge.post }));
  bridge.app = new FakeElement();
  vi.stubGlobal('document', { getElementById: () => bridge.app, createElement: () => new FakeElement() });
  vi.stubGlobal('window', { addEventListener: (_name: string, listener: typeof bridge.receive) => { bridge.receive = listener; } });
  await import('../src/webview/main.js');
});

afterEach(() => vi.unstubAllGlobals());

describe('VS Code webview commands', () => {
  it('forwards keybinding commands to the view', () => {
    send({ type: 'init', bytes: encodeBytes(saveWorkbook(new Workbook(), 'xlsx')), fileName: 'book.xlsx', language: 'en', readOnly: false });
    send({ type: 'command', command: 'find' });
    send({ type: 'command', command: 'replace' });
    expect(bridge.commands).toEqual(['find', 'replace']);
  });

  it('asks the host to open external links', () => {
    send({ type: 'init', bytes: encodeBytes(saveWorkbook(new Workbook(), 'xlsx')), fileName: 'book.xlsx', language: 'en', readOnly: false });
    bridge.options!.onOpenLink!('https://example.com');
    expect(bridge.post).toHaveBeenCalledWith({ type: 'openLink', url: 'https://example.com' });
  });
});

describe('VS Code webview save warnings and CSV export', () => {
  it('warns while changes would be refused on save and clears the warning after undo', () => {
    vi.useFakeTimers();
    try {
      const wb = new Workbook();
      setValues(wb, 0, 0, 0, [['Name', 'Value'], ['A', 1]]);
      const parts = unzipSync(saveWorkbook(wb, 'xlsx'));
      parts['customXml/unimplemented.xml'] = strToU8('<preserve/>');
      send({ type: 'init', bytes: encodeBytes(zipSync(parts)), fileName: 'complex.xlsx', language: 'de', readOnly: false });
      vi.advanceTimersByTime(1000);
      expect(bridge.notices.at(-1)).toBe(null);
      send({ type: 'request', id: 1, method: 'executeTool', params: { name: 'format_table', input: { range: 'A1:B2' } } });
      vi.advanceTimersByTime(1000);
      expect(bridge.notices.at(-1)).toMatch(/^Struktur-, Tabellen-/);
      send({ type: 'undo' });
      vi.advanceTimersByTime(1000);
      expect(bridge.notices.at(-1)).toBe(null);
    } finally {
      vi.useRealTimers();
    }
  });

  it('exports the active sheet as CSV', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['Name', 'Value'], ['A', 1.5]]);
    send({ type: 'init', bytes: encodeBytes(saveWorkbook(wb, 'xlsx')), fileName: 'data.xlsx', language: 'de', readOnly: false });
    send({ type: 'request', id: 7, method: 'exportCsv', params: { delimiter: ';' } });
    const result = response(7).result as { sheet: string; bytes: string };
    expect(result.sheet).toBe('Sheet1');
    expect(new TextDecoder().decode(decodeBytes(result.bytes))).toBe('Name;Value\r\nA;1,5\r\n');
  });
});

describe('VS Code webview file lifecycle', () => {
  it('refuses a lossy save but retains unsaved changes and original parts in hot-exit backup', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['Name', 'Value'], ['A', 1]]);
    const parts = unzipSync(saveWorkbook(wb, 'xlsx'));
    parts['customXml/unimplemented.xml'] = strToU8('<preserve>Original metadata</preserve>');
    const original = zipSync(parts);
    send({ type: 'init', bytes: encodeBytes(original), fileName: 'complex.xlsx', language: 'de', readOnly: false });
    send({ type: 'request', id: 1, method: 'executeTool', params: { name: 'format_table', input: { range: 'A1:B2', name: 'UnsavedTable' } } });
    expect(response(1).result).toMatchObject({ ok: true });
    send({ type: 'request', id: 2, method: 'getFileData' });
    expect(response(2).error).toContain('XLSX-Schutz');
    send({ type: 'request', id: 3, method: 'getBackupData' });
    const backup = response(3).result as string;
    const snapshot = JSON.parse(new TextDecoder().decode(decodeBytes(backup)));
    expect(decodeBytes(snapshot.workbook.xlsxSource.archive)).toEqual(original);
    send({ type: 'init', bytes: backup, backup: true, fileName: 'complex.xlsx', language: 'de', readOnly: false });
    send({ type: 'request', id: 4, method: 'executeTool', params: { name: 'list_tables', input: {} } });
    expect(response(4).result).toMatchObject({ ok: true, result: { tables: [expect.objectContaining({ name: 'UnsavedTable' })] } });
    send({ type: 'request', id: 5, method: 'getFileData' });
    expect(response(5).error).toContain('XLSX-Schutz');
  });

  it('preserves imported tables, highlighting and filters through opening, editing and saving', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['Segment', 'Country'], ['Government', 'Canada'], ['Midmarket', 'Germany']]);
    formatTable(wb, 0, toRange('A1:B3'), { name: 'financials' });
    setTableFilter(wb, 0, 'financials', 1, ['Canada']);
    send({ type: 'init', bytes: encodeBytes(saveWorkbook(wb, 'xlsx')), fileName: 'Financial Sample.xlsx', language: 'en', readOnly: false });
    send({ type: 'request', id: 1, method: 'executeTool', params: { name: 'list_tables', input: {} } });
    expect(response(1).result).toMatchObject({ ok: true, result: { tables: [expect.objectContaining({ name: 'financials' })] } });
    send({ type: 'request', id: 2, method: 'executeTool', params: { name: 'write_range', input: { start: 'A2', values: [['Enterprise']] } } });
    expect(response(2).result).toMatchObject({ ok: true });
    send({ type: 'request', id: 3, method: 'getFileData' });
    const reopened = openWorkbook(decodeBytes(response(3).result as string), 'Financial Sample.xlsx');
    expect(reopened.activeSheet.tables).toEqual(wb.activeSheet.tables);
    expect(reopened.getCellStyle(0, 0, 0)?.fill).toBe(wb.getCellStyle(0, 0, 0)?.fill);
    expect(reopened.getCellStyle(0, 1, 0)?.fill).not.toBe(reopened.getCellStyle(0, 2, 0)?.fill);
    expect([...tableFilteredRows(reopened.activeSheet)]).toEqual([2]);
    expect(reopened.getValue(0, 1, 0)).toBe('Enterprise');
  });

  it('opens a transferred workbook, edits it through tools and returns a valid XLSX', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [[21, '=A1*2']]);
    send({ type: 'init', bytes: encodeBytes(Buffer.from(saveWorkbook(wb, 'xlsx'))), fileName: 'test.xlsx', language: 'de', readOnly: false });
    expect(bridge.post.mock.calls.some(([m]) => m.type === 'error')).toBe(false);
    expect(bridge.post.mock.calls.some(([m]) => m.type === 'initialized')).toBe(true);
    send({ type: 'request', id: 1, method: 'executeTool', params: { name: 'write_range', input: { start: 'A1', values: [[50]] } } });
    expect(response(1).result).toMatchObject({ ok: true });
    send({ type: 'request', id: 2, method: 'getFileData' });
    const saved = response(2);
    expect(saved.error).toBeUndefined();
    const reopened = openWorkbook(decodeBytes(saved.result as string), 'test.xlsx');
    expect(reopened.getValue(0, 0, 1)).toBe(100);
  });

  it('initializes a zero-byte XLSX and creates its ZIP on save', () => {
    send({ type: 'init', bytes: '', fileName: 'test.xlsx', language: 'de', readOnly: false });
    expect(bridge.post.mock.calls.some(([m]) => m.type === 'error')).toBe(false);
    // Opening the empty file alone must not mark it dirty or write to disk.
    expect(bridge.post.mock.calls.some(([m]) => m.type === 'edit')).toBe(false);
    send({ type: 'request', id: 1, method: 'getFileData' });
    const bytes = decodeBytes(response(1).result as string);
    expect([...bytes.subarray(0, 2)]).toEqual([80, 75]);
    expect(openWorkbook(bytes, 'test.xlsx').sheets).toHaveLength(1);
  });

  it('still rejects a nonempty invalid archive without replacing its contents', () => {
    send({ type: 'init', bytes: encodeBytes(new TextEncoder().encode('not an xlsx')), fileName: 'test.xlsx', language: 'de', readOnly: false });
    expect(bridge.post.mock.calls).toContainEqual([expect.objectContaining({ type: 'error', message: expect.stringContaining('Could not open test.xlsx') })]);
    expect(bridge.post.mock.calls.some(([m]) => m.type === 'initialized')).toBe(false);
    send({ type: 'request', id: 1, method: 'getFileData' });
    expect(response(1).error).toBe('Workbook not loaded');
  });
});

describe('VS Code webview loading errors and CSV saves', () => {
  const csv = (text: string, fileName = 'data.csv', language = 'de') => send({ type: 'init', bytes: encodeBytes(new TextEncoder().encode(text)), fileName, language, readOnly: false });
  const fileData = (id: number, fileName?: string) => {
    send({ type: 'request', id, method: 'getFileData', ...(fileName ? { params: { fileName } } : {}) });
    return response(id);
  };

  it('shows a localized error in the editor area when a file cannot be opened', () => {
    send({ type: 'init', bytes: encodeBytes(new TextEncoder().encode('not an xlsx')), fileName: 'kaputt.xlsx', language: 'de', readOnly: false });
    const panel = bridge.app!.children[0]!;
    expect(panel.attributes.role).toBe('alert');
    expect(panel.text()).toContain('kaputt.xlsx konnte nicht geöffnet werden.');
    expect(panel.text()).toContain('Speichern ist in diesem Editor deaktiviert');
    expect(fileData(1).error).toBe('Workbook not loaded');
    send({ type: 'init', bytes: encodeBytes(new TextEncoder().encode('not an xlsx')), fileName: 'broken.xlsx', language: 'en', readOnly: false });
    expect(bridge.app!.children[0]!.text()).toContain('broken.xlsx could not be opened.');
  });

  it('reproduces CRLF, a missing final line break and the BOM of the opened CSV', () => {
    for (const text of ['\uFEFFName;Wert\r\nMüller;1,5', 'a,b\n1,2\n', 'a,b\r\n0004711,2\r\n']) {
      csv(text);
      const bytes = decodeBytes(fileData(1).result as string);
      expect(new TextDecoder('utf-8', { ignoreBOM: true }).decode(bytes)).toBe(text);
      bridge.post.mockClear();
    }
  });

  it('warns once, without blocking, when a CSV save drops content', () => {
    csv('Name;Wert\r\nA;1\r\n');
    expect(fileData(1).error).toBeUndefined();
    expect(bridge.post.mock.calls.some(([m]) => m.type === 'warning')).toBe(false);
    send({ type: 'request', id: 2, method: 'executeTool', params: { name: 'add_sheet', input: { name: 'Zweites' } } });
    expect(response(2).result).toMatchObject({ ok: true });
    expect(fileData(3).error).toBeUndefined();
    const warnings = bridge.post.mock.calls.map(([m]) => m as WebviewToHost).filter((m) => m.type === 'warning');
    expect(warnings).toHaveLength(1);
    expect(warnings[0]).toMatchObject({ message: expect.stringMatching(/data\.csv wird als CSV gespeichert\. .*nur die Werte des Blatts „data“/) });
    expect(warnings[0]).toMatchObject({ message: expect.stringContaining('weitere Blätter') });
    fileData(4);
    expect(bridge.post.mock.calls.filter(([m]) => m.type === 'warning')).toHaveLength(1);
  });

  it('saves the data sheet of a CSV even when another sheet is active', () => {
    csv('Name;Wert\r\nA;1\r\n');
    send({ type: 'request', id: 1, method: 'executeTool', params: { name: 'add_sheet', input: { name: 'Zweites' } } });
    send({ type: 'request', id: 2, method: 'executeTool', params: { name: 'write_range', input: { sheet: 'Zweites', start: 'A1', values: [['nur Notizen']] } } });
    expect(new TextDecoder().decode(decodeBytes(fileData(3).result as string))).toBe('Name;Wert\r\nA;1\r\n');
  });

  it('keeps a Windows-1252 CSV byte for byte and warns once when it has to switch to UTF-8', () => {
    // "Name;Wert\r\nMüller;1,5\r\n" in Windows-1252 (ü = 0xFC).
    const cp1252 = Uint8Array.from([...'Name;Wert\r\nM'].map((c) => c.charCodeAt(0)).concat([0xfc], [...'ller;1,5\r\n'].map((c) => c.charCodeAt(0))));
    send({ type: 'init', bytes: encodeBytes(cp1252), fileName: 'alt.csv', language: 'de', readOnly: false });
    expect([...decodeBytes(fileData(1).result as string)]).toEqual([...cp1252]);
    send({ type: 'request', id: 2, method: 'executeTool', params: { name: 'write_range', input: { start: 'A3', values: [['東京']] } } });
    const utf8 = decodeBytes(fileData(3).result as string);
    expect([...utf8.subarray(0, 3)]).toEqual([0xef, 0xbb, 0xbf]);
    expect(new TextDecoder().decode(utf8)).toContain('Müller;1,5\r\n東京');
    fileData(4);
    const warnings = bridge.post.mock.calls.map(([m]) => m as WebviewToHost).filter((m) => m.type === 'warning');
    expect(warnings).toHaveLength(1);
    expect(warnings[0]).toMatchObject({ message: expect.stringMatching(/alt\.csv wurde als UTF-8 gespeichert, .*"東", "京"/) });
  });

  it('writes the format of the save target for "Save As"', () => {
    csv('Name;Wert\r\nA;1,5\r\n');
    const xlsx = decodeBytes(fileData(1, 'data.xlsx').result as string);
    expect(openWorkbook(xlsx, 'data.xlsx').getValue(0, 1, 1)).toBe(1.5);
    const tsv = new TextDecoder().decode(decodeBytes(fileData(2, 'data.tsv').result as string));
    expect(tsv).toBe('Name\tWert\nA\t1.5\n');
  });

  it('shows protection reasons in the UI language', () => {
    vi.useFakeTimers();
    try {
      const wb = new Workbook();
      setValues(wb, 0, 0, 0, [['Name', 'Value'], ['A', 1]]);
      const parts = unzipSync(saveWorkbook(wb, 'xlsx'));
      parts['customXml/unimplemented.xml'] = strToU8('<preserve/>');
      send({ type: 'init', bytes: encodeBytes(zipSync(parts)), fileName: 'complex.xlsx', language: 'en', readOnly: false });
      send({ type: 'request', id: 1, method: 'executeTool', params: { name: 'format_table', input: { range: 'A1:B2' } } });
      vi.advanceTimersByTime(1000);
      expect(bridge.notices.at(-1)).toMatch(/^Structure, table, chart or rule change/);
      send({ type: 'request', id: 2, method: 'getFileData' });
      expect(response(2).error).toMatch(/^XLSX protection: /);
    } finally {
      vi.useRealTimers();
    }
  });
});
