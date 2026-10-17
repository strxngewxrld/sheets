import { describe, expect, it } from 'vitest';
import { setValues, Workbook } from '../src/index.js';

function result(formula: string) {
  const wb = new Workbook();
  setValues(wb, 0, 0, 0, [[formula]]);
  const range = wb.activeSheet.getCell(0, 0)?.spillRange;
  if (!range) return wb.getValue(0, 0, 0);
  return Array.from({ length: range.end.row + 1 }, (_, r) => Array.from({ length: range.end.col + 1 }, (_, c) => wb.getValue(0, r, c)));
}

describe('TEXTSPLIT', () => {
  it('splits into columns and rows', () => {
    expect(result('=TEXTSPLIT("a,b,c", ",")')).toEqual([['a', 'b', 'c']]);
    expect(result('=TEXTSPLIT("a=1;b=2", "=", ";")')).toEqual([['a', '1'], ['b', '2']]);
    expect(result('=TEXTSPLIT("a1b", , "1")')).toEqual([['a'], ['b']]);
  });

  it('handles multiple delimiters, empty parts, case and padding', () => {
    expect(result('=TEXTSPLIT("a, b;c", {", ",";"})')).toEqual([['a', 'b', 'c']]);
    expect(result('=TEXTSPLIT("a,,b", ",")')).toEqual([['a', '', 'b']]);
    expect(result('=TEXTSPLIT("a,,b", ",", , TRUE)')).toEqual([['a', 'b']]);
    expect(result('=TEXTSPLIT("aXbxc", "x", , , 1)')).toEqual([['a', 'b', 'c']]);
    expect(result('=TEXTSPLIT("a,b;c", ",", ";", , , "-")')).toEqual([['a', 'b'], ['c', '-']]);
    expect(result('=TEXTSPLIT("a,b;c", ",", ";")')).toEqual([['a', 'b'], ['c', expect.objectContaining({ code: '#N/A' })]]);
    expect(result('=TEXTSPLIT("abc", "")')).toMatchObject({ code: '#VALUE!' });
  });
});

describe('TEXTBEFORE and TEXTAFTER', () => {
  it('returns text around the n-th delimiter', () => {
    expect(result('=TEXTBEFORE("john.doe@example.com", "@")')).toBe('john.doe');
    expect(result('=TEXTAFTER("john.doe@example.com", "@")')).toBe('example.com');
    expect(result('=TEXTAFTER("a-b-c", "-", 2)')).toBe('c');
    expect(result('=TEXTBEFORE("a-b-c", "-", -1)')).toBe('a-b');
    expect(result('=TEXTAFTER("a-b-c", "-", -1)')).toBe('c');
    expect(result('=TEXTBEFORE("a, b; c", {";", ","}, 2)')).toBe('a, b');
  });

  it('supports case-insensitive matching, match_end and fallbacks', () => {
    expect(result('=TEXTAFTER("RedXblue", "x")')).toMatchObject({ code: '#N/A' });
    expect(result('=TEXTAFTER("RedXblue", "x", 1, 1)')).toBe('blue');
    expect(result('=TEXTBEFORE("abc", "-", 1, 0, TRUE)')).toBe('abc');
    expect(result('=TEXTBEFORE("abc", "-", 1, 0, FALSE, "none")')).toBe('none');
    expect(result('=TEXTBEFORE("abc", "-", 0)')).toMatchObject({ code: '#VALUE!' });
  });

  it('broadcasts over arrays of text', () => {
    expect(result('=TEXTBEFORE({"a.b";"c.d"}, ".")')).toEqual([['a'], ['c']]);
  });
});
