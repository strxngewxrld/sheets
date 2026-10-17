import { describe, expect, it } from 'vitest';
import { openWorkbook, saveWorkbook, setValues, Workbook } from '@sheets/core';
import { decodeBytes, encodeBytes } from '../src/binary.js';
import type { InitMessage, WebviewToHost } from '../src/protocol.js';

const transfer = <T>(message: T): T => JSON.parse(JSON.stringify(message));

describe('VS Code file byte transport', () => {
  it('opens and saves an XLSX after JSON serialization in both directions', () => {
    const original = new Workbook();
    setValues(original, 0, 0, 0, [['Müller 😀', 42, '=B1*2']]);
    // workspace.fs may return a Node Buffer; its toJSON differs from Uint8Array.
    const diskBytes = Buffer.from(saveWorkbook(original, 'xlsx'));
    const init = transfer<InitMessage>({ type: 'init', bytes: encodeBytes(diskBytes), fileName: 'test.xlsx', language: 'de', readOnly: false });
    const loaded = openWorkbook(decodeBytes(init.bytes), init.fileName);
    expect(loaded.getValue(0, 0, 0)).toBe('Müller 😀');
    expect(loaded.getValue(0, 0, 2)).toBe(84);
    setValues(loaded, 0, 0, 1, [[100]]);
    const saved = transfer<WebviewToHost>({ type: 'response', id: 1, result: encodeBytes(saveWorkbook(loaded, 'xlsx')) });
    if (saved.type !== 'response') throw new Error('Expected a file response');
    const reopened = openWorkbook(decodeBytes(saved.result as string), 'test.xlsx');
    expect(reopened.getValue(0, 0, 2)).toBe(200);
  });

  it('preserves all byte values and sliced buffers without leaking surrounding data', () => {
    const backing = Buffer.from([99, ...Array.from({ length: 256 }, (_, i) => i), 77]);
    const bytes = backing.subarray(1, 257);
    expect(decodeBytes(transfer(encodeBytes(bytes)))).toEqual(Uint8Array.from({ length: 256 }, (_, i) => i));
  });

  it('handles empty payloads and large workbooks without argument overflows', () => {
    expect(decodeBytes(transfer(encodeBytes(new Uint8Array())))).toHaveLength(0);
    const bytes = Uint8Array.from({ length: 200_000 }, (_, i) => i % 256);
    expect(Buffer.from(decodeBytes(transfer(encodeBytes(bytes)))).equals(Buffer.from(bytes))).toBe(true);
  });

  it('reports stale non-base64 messages instead of treating an object as an empty ZIP', () => {
    expect(() => decodeBytes({ 0: 80, 1: 75 } as unknown as string)).toThrow(/Reload the Sheets editor/);
  });
});
