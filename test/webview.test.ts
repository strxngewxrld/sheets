import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { formatTable, openWorkbook, saveWorkbook, setTableFilter, setValues, tableFilteredRows, toRange, Workbook } from '@sheets/core';
import { decodeBytes, encodeBytes } from '../src/binary.js';
import type { HostToWebview, WebviewToHost } from '../src/protocol.js';
import { strToU8, unzipSync, zipSync } from 'fflate';

const bridge = vi.hoisted(() => ({
  post: vi.fn(),
  commands: [] as string[],
  receive: undefined as ((event: { data: HostToWebview }) => void) | undefined,
}));

// Keep the real webview message handler, workbook, I/O and agent. Only the
// canvas/DOM view is replaced so this regression runs without VS Code itself.
vi.mock('@sheets/ui', () => ({
  SpreadsheetView: class {
    agentHost = {};
    constructor(_app: unknown, public workbook: Workbook) {}
    setWorkbook(wb: Workbook) { this.workbook = wb; }
    focus() {}
    runCommand(command: string) { bridge.commands.push(command); }
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
  bridge.receive = undefined;
  vi.stubGlobal('acquireVsCodeApi', () => ({ postMessage: bridge.post }));
  vi.stubGlobal('document', { getElementById: () => ({}) });
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
