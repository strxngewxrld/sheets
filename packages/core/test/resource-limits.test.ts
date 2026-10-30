import { describe, expect, it, vi } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { CellError, MAX_CHART_POINTS, openWorkbook, resolveChart, setValues, Workbook, writeXlsx } from '../src/index.js';
import { calcLimits } from '../src/formula/limits.js';
import { serialToParts } from '../src/dates.js';

const raw = (wb: Workbook, formula: string) => wb.calc.evaluateRaw(formula, wb.activeSheet);

describe('formula resource limits', () => {
  it('rejects oversized references before reading cells and remains usable', () => {
    const wb = new Workbook();
    const read = vi.spyOn(wb.activeSheet, 'getValue');
    expect(raw(wb, '=SUM(B1:XFC1048575)')).toMatchObject({ code: '#CALC!' });
    expect(read).not.toHaveBeenCalled();
    setValues(wb, 0, 0, 0, [[7]]);
    expect(wb.evaluate('=SUM(A:A)')).toBe(7);
    expect(wb.evaluate('=SUM(A1:B2)')).toBe(7);
    expect(wb.evaluate('=INDEX(A1:XFD1048576,1,1)')).toBe(7);
    expect(wb.evaluate('=ROWS(A1:XFD1048576)')).toBe(1_048_576);
    expect(wb.evaluate('=COLUMNS(A1:XFD1048576)')).toBe(16_384);
  });

  it.each([
    '=WRAPROWS({1},1000001)',
    '=WRAPCOLS({1},1000001)',
    '=WRAPROWS({1},1e20)',
    '=WRAPCOLS({1},1e309)',
    '=SEQUENCE(1100)+TRANSPOSE(SEQUENCE(1100))',
    '=TEXTSPLIT(REPT(",",1100)&REPT(";",1100),",",";")',
    '=CHOOSEROWS(SEQUENCE(1,1100),SEQUENCE(1100,1,1,0))',
    '=CHOOSECOLS(SEQUENCE(1100),SEQUENCE(1100,1,1,0))',
    '=COUNTIFS(B:XFC,"")',
    '=MATCH(0,B:B,0)',
    '=MAKEARRAY(1e309,1,LAMBDA(r,c,1))',
  ])('refuses an oversized intermediate allocation: %s', (formula) => {
    const wb = new Workbook();
    expect(raw(wb, formula)).toMatchObject({ code: '#CALC!' });
    expect(wb.evaluate('=1+1')).toBe(2);
  });

  it('rejects invalid OFFSET bounds instead of looping at an unrepresentable row', () => {
    const wb = new Workbook();
    for (const formula of ['=SUM(OFFSET(B1,1e20,0,2,2))', '=OFFSET(B1,0,0,1e309,1)', '=OFFSET(B1,0,0,1.5,1)']) {
      expect(raw(wb, formula)).toMatchObject({ code: '#REF!' });
    }
  });

  it('keeps valid wrapping, broadcasting and ragged text splitting', () => {
    const wb = new Workbook();
    expect(raw(wb, '=WRAPROWS({1,2,3},2,0)')).toEqual([[1, 2], [3, 0]]);
    expect(raw(wb, '=WRAPCOLS({1,2,3},2,0)')).toEqual([[1, 3], [2, 0]]);
    expect(raw(wb, '=SEQUENCE(2)+TRANSPOSE(SEQUENCE(2))')).toEqual([[2, 3], [3, 4]]);
    expect(raw(wb, '=TEXTSPLIT("a,b;c",",",";",FALSE,0,"pad")')).toEqual([['a', 'b'], ['c', 'pad']]);
    const edge = raw(wb, '=WRAPROWS({1},1000000)');
    expect(Array.isArray(edge) && edge[0]!.length).toBe(1_000_000);
  });

  it.each([
    ['WORKDAY(45000,1e20)', '#NUM!'],
    ['NETWORKDAYS(1e20,1e20)', '#NUM!'],
    ['SUM(B1:XFC1048575)', '#CALC!'],
    ['WRAPROWS({1},1000001)', '#CALC!'],
    ['SEQUENCE(1100)+TRANSPOSE(SEQUENCE(1100))', '#CALC!'],
    ['TEXTSPLIT(REPT(",",1100)&REPT(";",1100),",",";")', '#CALC!'],
    ['FACT(1e12)', '#NUM!'],
    ['CUMIPMT(0.01,1e12,1,1,1e12,0)', '#CALC!'],
    ['UPPER(REPT("a",4e8))', '#VALUE!'],
    ['FIXED(1,3e8)', '#VALUE!'],
  ])('opens a small XLSX containing %s safely', (formula, code) => {
    const parts = unzipSync(writeXlsx(new Workbook()));
    const sheet = strFromU8(parts['xl/worksheets/sheet1.xml']!);
    const escaped = formula.replace(/&/g, '&amp;').replace(/</g, '&lt;');
    parts['xl/worksheets/sheet1.xml'] = strToU8(sheet.replace('<sheetData>', `<sheetData><row r="1"><c r="A1"><f>${escaped}</f><v>0</v></c></row>`));
    const wb = openWorkbook(zipSync(parts), 'untrusted.xlsx');
    expect(wb.getValue(0, 0, 0)).toBeInstanceOf(CellError);
    expect(wb.getValue(0, 0, 0)).toMatchObject({ code });
    setValues(wb, 0, 1, 0, [[42]]);
    expect(wb.getValue(0, 1, 0)).toBe(42);
  });
});

describe('bounded workday calculations', () => {
  it.each([
    '=WORKDAY(45000,1e20)', '=WORKDAY(45000,-1e20)', '=WORKDAY(45000,1e309)',
    '=NETWORKDAYS(1e20,1e20)', '=NETWORKDAYS(0,1e309)', '=NETWORKDAYS(-1,10)',
    '=WORKDAY(2958465,1)', '=WORKDAY(0,-1)', '=WORKDAY(45000,2958465)',
    '=WORKDAY.INTL(45000,1e20,11)', '=NETWORKDAYS.INTL(0,1e20,11)',
    '=WORKDAY(45000,1,{1e309})',
  ])('returns a numeric error for invalid dates or unreachable results: %s', (formula) => {
    expect(new Workbook().evaluate(formula)).toMatchObject({ code: '#NUM!' });
  });

  it('handles the complete date domain without iterating millions of days', () => {
    const wb = new Workbook();
    expect(wb.evaluate('=NETWORKDAYS.INTL(0,2958465,"0000000")')).toBe(2_958_466);
    expect(wb.evaluate('=WORKDAY.INTL(0,2958465,"0000000")')).toBe(2_958_465);
    expect(wb.evaluate('=WORKDAY.INTL(2958465,-2958465,"0000000")')).toBe(0);
  });

  it('matches a day-by-day reference across weekends, holidays and the fictitious leap day', () => {
    const wb = new Workbook();
    const holidays = new Set([1, 59, 60, 61, 64, 45001, 45003]);
    for (const mask of ['0000011', '0010001', '0000001', '0000000', '1111110']) {
      const work = (day: number) => mask[(serialToParts(day).weekday + 6) % 7] === '0' && !holidays.has(day);
      for (const start of [0, 1, 55, 59, 60, 61, 45000, 2958400]) {
        const end = start + 40;
        let expected = 0;
        for (let day = start; day <= end; day++) if (work(day)) expected++;
        const args = `"${mask}",{${[...holidays].join(',')}}`;
        expect(wb.evaluate(`=NETWORKDAYS.INTL(${start},${end},${args})`)).toBe(expected);
        expect(wb.evaluate(`=NETWORKDAYS.INTL(${end},${start},${args})`)).toBe(-expected);
        for (const n of [-5, 0, 5]) {
          if (n < 0 && start < 55) continue;
          let day = start, left = Math.abs(n);
          while (left > 0) { day += Math.sign(n); if (work(day)) left--; }
          expect(wb.evaluate(`=WORKDAY.INTL(${start},${n},${args})`)).toBe(day);
        }
      }
    }
  });
});

describe('bounded loops', () => {
  it.each([
    ['=FACT(1e12)', '#NUM!'],
    ['=FACT(171)', '#NUM!'],
    ['=PERMUT(1e12,1e12)', '#NUM!'],
    ['=COMBIN(2e12,1e12)', '#NUM!'],
    ['=CUMIPMT(0.01,1e12,1,1,1e12,0)', '#CALC!'],
    ['=CUMPRINC(0.01,1e12,1,1,1e12,0)', '#CALC!'],
    ['=DB(100,1,1e12,1e12,12)', '#CALC!'],
  ])('stops %s instead of iterating', (formula, code) => {
    expect(raw(new Workbook(), formula)).toMatchObject({ code });
  });

  it('keeps results at the edge of the number range', () => {
    const wb = new Workbook();
    expect(wb.evaluate('=FACT(170)')).toBe(7.257415615307994e306);
    expect(wb.evaluate('=COMBIN(1000,500)')).toBe(2.7028824094543666e299);
    expect(wb.evaluate('=PERMUT(10,3)')).toBe(720);
  });
});

describe('text length limit', () => {
  it.each([
    '=REPT("a",32768)',
    '=REPT("ab",4e8)',
    '=REPT("a",20000)&REPT("a",20000)',
    '=CONCAT(REPT("a",20000),REPT("a",20000))',
    '=CONCATENATE(REPT("a",20000),REPT("a",20000))',
    '=TEXTJOIN(",",1,REPT("a",20000),REPT("a",20000))',
    '=SUBSTITUTE(REPT("a",30000),"a",REPT("b",30000))',
    '=ARRAYTOTEXT(REPT("a",20000)&{1,2})',
    '=FIXED(1,128)',
    '=TEXT(1,REPT("0",20000)&REPT("0",20000))',
  ])('refuses texts over 32,767 characters: %s', (formula) => {
    expect(raw(new Workbook(), formula)).toMatchObject({ code: '#VALUE!' });
  });

  it('keeps texts up to the limit and long number formats', () => {
    const wb = new Workbook();
    expect(wb.evaluate('=LEN(REPT("a",32767))')).toBe(32_767);
    expect(wb.evaluate('=LEN(REPT("a",32766)&"b")')).toBe(32_767);
    expect(wb.evaluate('=FIXED(1.5,3)')).toBe('1.500');
    expect(wb.evaluate('=TEXT(1,"0."&REPT("0",120))')).toBe(`1.${'0'.repeat(120)}`);
  });
});

describe('calculation budget', () => {
  const exponential = 'LET(f,LAMBDA(f,n,IF(n=0,0,f(f,n-1)+f(f,n-1))),f(f,40))';
  const withLimits = (limits: Partial<typeof calcLimits>, fn: () => void) => {
    const saved = { ...calcLimits };
    Object.assign(calcLimits, limits);
    try { fn(); } finally { Object.assign(calcLimits, saved); }
  };

  it('stops exponential LAMBDA recursion by work and by time', () => {
    withLimits({ cellWork: 100_000 }, () => expect(raw(new Workbook(), `=${exponential}`)).toMatchObject({ code: '#CALC!' }));
    withLimits({ cellMs: 0 }, () => expect(raw(new Workbook(), `=${exponential}`)).toMatchObject({ code: '#CALC!' }));
    expect(raw(new Workbook(), '=LET(f,LAMBDA(f,n,IF(n=0,1,f(f,n-1)+f(f,n-1))),f(f,10))')).toBe(1024);
  });

  it('does not let functions turn an exhausted budget into a value', () => {
    withLimits({ cellWork: 100_000 }, () => {
      expect(raw(new Workbook(), `=IFERROR(${exponential},1)`)).toMatchObject({ code: '#CALC!' });
      expect(raw(new Workbook(), '=SUM(MAKEARRAY(1000,1000,LAMBDA(r,c,SUM(SEQUENCE(1000)))))')).toMatchObject({ code: '#CALC!' });
    });
  });

  it('limits a whole recalculation and recovers on the next one', () => {
    const wb = new Workbook();
    withLimits({ passWork: 50_000 }, () => {
      setValues(wb, 0, 0, 0, Array.from({ length: 6 }, () => ['=SUM(SEQUENCE(100,100))']));
      const values = Array.from({ length: 6 }, (_, r) => wb.getValue(0, r, 0));
      expect(values[0]).toBe(50_005_000);
      expect(values[5]).toMatchObject({ code: '#CALC!' });
      setValues(wb, 0, 5, 0, [['=SUM(SEQUENCE(100,100))']]);
      expect(wb.getValue(0, 5, 0)).toBe(50_005_000);
    });
  });

  it('gives each formula its own budget while it computes a dependency', () => {
    const wb = new Workbook();
    withLimits({ cellWork: 30_000 }, () => {
      setValues(wb, 0, 0, 0, [['=SUM(SEQUENCE(100,100))', '=SUM(SEQUENCE(100,100))', '=A1+B1']]);
      expect(wb.getValue(0, 0, 2)).toBe(100_010_000);
    });
  });

  it('bounds the evaluation of a formula part, as for links in the editor', () => {
    const wb = new Workbook();
    const node = { type: 'func', name: 'SUM', args: [{ type: 'func', name: 'SEQUENCE', args: [{ type: 'number', value: 1000 }, { type: 'number', value: 1000 }] }] };
    withLimits({ cellWork: 100_000 }, () => {
      expect(wb.calc.evaluateNode(node as never, wb.activeSheet, 0, 0)).toMatchObject({ code: '#CALC!' });
    });
    expect(wb.calc.evaluateNode(node as never, wb.activeSheet, 0, 0)).toBe(500_000_500_000);
  });
});

describe('chart data limits', () => {
  const chart = (values: string, categories: string) =>
    ({ id: 'c', type: 'column' as const, series: [{ values, categories }], anchor: { from: { row: 1, col: 1 }, to: { row: 10, col: 8 } } });

  it('reads category references only within the used area', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [[1, 'a'], [2, 'b']]);
    const read = vi.spyOn(wb.activeSheet, 'getCell');
    const data = resolveChart(wb, chart('Sheet1!A1:A1048576', 'Sheet1!B1:B1048576'));
    expect(read.mock.calls.length).toBeLessThanOrEqual(2);
    expect(data.categories).toEqual(['a', 'b']);
    expect(data.series[0]!.values).toEqual([1, 2]);
  });

  it('caps references into a sheet whose used area is huge', () => {
    const wb = new Workbook();
    setValues(wb, 0, 1_048_575, 16_383, [['x']]);
    const data = resolveChart(wb, chart('Sheet1!A1:XFD1048576', 'Sheet1!A1:XFD1048576'));
    expect(data.categories).toHaveLength(MAX_CHART_POINTS);
    expect(data.series[0]!.values).toHaveLength(MAX_CHART_POINTS);
  });

  it('opens an XLSX whose chart categories reference the whole sheet', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [[1]]);
    wb.activeSheet.charts.push(chart('Sheet1!A1:A1', 'Sheet1!B1:XFD1048576'));
    const back = openWorkbook(writeXlsx(wb), 'chart.xlsx');
    expect(back.activeSheet.charts[0]!.series[0]!.categories).toBe('Sheet1!B1:XFD1048576');
    // Column B lies outside the used area, so the category is numbered.
    expect(resolveChart(back, back.activeSheet.charts[0]!).categories).toEqual(['1']);
  });
});
