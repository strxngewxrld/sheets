import { describe, expect, it } from 'vitest';
import { detectCsvLayout, openWorkbook, saveWorkbook } from '../src/index.js';

// Opening and saving a CSV without editing must not change what the file says.
const RAW = ['0004711', '007', '1e5', '+5', '-0', '-0,0', '-0.0', '1 000', 'true', 'False', '2024-01-05', '05.01.2024', '1/2/2024', '00:30', '12:00:00',
  '12345678901234567890', '0.1234567890123456789', '$5', '5%', '5 %', '€5', "'quoted", '0e0', '00.5', '-007', '1.5E+3', 'TRUE ', '12.', '5.',
  'TRUE', 'FALSE', '12', '-3.25', '0.5', '1.234,50', '1.234', '-1.234', '1,234.5', '1.23', '1,23', 'Müller', '=SUM(1)', '+SUM(1)', '-SUM(1)', '@x',
  'N/A', '#N/A', '#DIV/0!', 'null', ' 5', '5 ', '"x"'];

const roundTrip = (raw: string, sep: string) => {
  const field = raw.includes(sep) || /["\n]/.test(raw) || /^\s|\s$/.test(raw) ? `"${raw.replace(/"/g, '""')}"` : raw;
  const wb = openWorkbook(new TextEncoder().encode(`h\n${field}\n`), 'a.csv', { csv: { delimiter: sep } });
  const out = new TextDecoder().decode(saveWorkbook(wb, 'csv', { csv: { delimiter: sep, lineEnding: '\n' } }));
  return { wb, line: out.split('\n')[1]!, field };
};

describe('csv open + save round trip', () => {
  for (const sep of [',', ';']) {
    const decimal = sep === ';' ? ',' : '.';
    it(`keeps every value ("${sep}")`, () => {
      for (const raw of RAW) {
        const { wb, line, field } = roundTrip(raw, sep);
        const value = wb.getValue(0, 1, 0);
        if (typeof value === 'number') {
          // Only plain or correctly grouped decimals in the file's own notation become numbers;
          // they may lose grouping and insignificant zeros (1.234,50 -> 1234,5) but keep their value.
          const group = sep === ';' ? '.' : ',';
          expect(raw, `${raw} became a number`).toMatch(sep === ';' ? /^-?(\d{1,3}(\.\d{3})+|\d*)(,\d+)?$/ : /^-?(\d{1,3}(,\d{3})+|\d*)(\.\d+)?$/);
          expect(Number(line.replace(decimal, '.'))).toBe(Number(raw.split(group).join('').replace(decimal, '.')));
        } else expect(line, raw).toBe(field);
      }
    });
  }

  it('keeps plain numbers numeric and normalises only cosmetic decimals', () => {
    const de = roundTrip('12,50', ';');
    expect(de.wb.getValue(0, 1, 0)).toBe(12.5);
    expect(de.line).toBe('12,5');
    expect(roundTrip('-3.25', ',').wb.getValue(0, 1, 0)).toBe(-3.25);
    expect(roundTrip('TRUE', ',').wb.getValue(0, 1, 0)).toBe(true);
  });

  it('keeps foreign decimal separators and malformed grouping as text', () => {
    // A German file never writes "1.23" for 1,23: rewriting it would change the file.
    for (const raw of ['1.23', '1.2345', '12.34,5']) expect(roundTrip(raw, ';').wb.getValue(0, 1, 0)).toBe(raw);
    for (const raw of ['1,23', '1,2345']) expect(roundTrip(raw, ',').wb.getValue(0, 1, 0)).toBe(raw);
    // Correct grouping stays numeric, as when typed.
    expect(roundTrip('1.234,50', ';').wb.getValue(0, 1, 0)).toBe(1234.5);
    expect(roundTrip('1,234.5', ',').wb.getValue(0, 1, 0)).toBe(1234.5);
  });

  it('reproduces line endings, a missing final line break and the byte order mark', () => {
    for (const text of ['a;b\r\n1;2\r\n', 'a;b\n1;2', '﻿Name;Wert\r\nÄ;1', 'a\rb\r', '']) {
      const layout = detectCsvLayout(text);
      const wb = openWorkbook(new TextEncoder().encode(text), 'a.csv', { csv: { delimiter: ';' } });
      expect(new TextDecoder('utf-8', { ignoreBOM: true }).decode(saveWorkbook(wb, 'csv', { csv: { delimiter: ';', ...layout } })), JSON.stringify(text))
        .toBe(text.replace(/\r(?!\n)/g, '\n'));
    }
    expect(detectCsvLayout('a\r\nb')).toEqual({ lineEnding: '\r\n', finalNewline: false, bom: false });
    expect(detectCsvLayout('﻿a\n')).toEqual({ lineEnding: '\n', finalNewline: true, bom: true });
  });
});
