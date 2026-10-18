import { readdirSync, readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { listFunctions, setValues, Workbook } from '../src/index.js';

/**
 * Reference results as Excel computes them, for functions without a dedicated test.
 * Data (rows 1-5):  A: 3 1 4 1 5   B: a b a c a   C: 10 20 30 40 50   D: TRUE "x" (empty) 0 =1/0
 */
function book(): Workbook {
  const wb = new Workbook();
  setValues(wb, 0, 0, 0, [
    [3, 'a', 10, true],
    [1, 'b', 20, 'x'],
    [4, 'a', 30, null],
    [1, 'c', 40, 0],
    [5, 'a', 50, '=1/0'],
  ]);
  return wb;
}

const PI = Math.PI;
// [formula, expected]; numbers are compared to 9 decimals, errors by their code.
const CASES: [string, number | string | boolean][] = [
  // Math and trigonometry
  ['=ABS(-3)', 3], ['=ACOS(1)', 0], ['=ASIN(1)', PI / 2], ['=ATAN(1)', PI / 4],
  ['=ATAN2(1,0)', 0], ['=ATAN2(0,1)', PI / 2], ['=ATAN2(-1,0)', PI],
  ['=COS(0)', 1], ['=SIN(PI()/2)', 1], ['=TAN(PI()/4)', 1], ['=DEGREES(PI())', 180], ['=RADIANS(180)', PI], ['=PI()', PI],
  ['=CEILING(2.5,1)', 3], ['=CEILING(4.42,0.05)', 4.45], ['=CEILING(-2.5,-2)', -4],
  ['=CEILING.MATH(6.7)', 7], ['=CEILING.MATH(-5.5)', -5], ['=CEILING.MATH(-5.5,2,-1)', -6],
  ['=FLOOR.MATH(5.5)', 5], ['=FLOOR.MATH(-5.5)', -6], ['=FLOOR.MATH(-5.5,2,-1)', -4],
  ['=CEILING(2.5,-1)', '#NUM!'], ['=FLOOR(2.5,-1)', '#NUM!'], ['=CEILING(-2.5,2)', -2], ['=FLOOR(-2.5,2)', -4], ['=FLOOR(1,0)', '#DIV/0!'],
  // Binary floating point must not leak into results (0.3 / 0.1 = 2.9999999999999996 in IEEE 754).
  ['=FLOOR(0.3,0.1)=0.3', true], ['=CEILING(0.3,0.1)=0.3', true], ['=MROUND(0.3,0.1)=0.3', true],
  ['=FLOOR.MATH(0.3,0.1)=0.3', true], ['=CEILING.MATH(0.7,0.1)=0.7', true], ['=CEILING.MATH(6.7,,1)', 7],
  ['=LOG(100)', 2], ['=LOG(8,2)', 3], ['=LOG(0)', '#NUM!'], ['=LOG(8,1)', '#NUM!'],
  ['=EVEN(1.5)', 2], ['=EVEN(3)', 4], ['=EVEN(-1)', -2], ['=ODD(1.5)', 3], ['=ODD(2)', 3], ['=ODD(-2)', -3],
  ['=EXP(1)', Math.E], ['=LN(EXP(2))', 2], ['=LOG10(1000)', 3], ['=FACT(5)', 120], ['=FACT(0)', 1],
  ['=GCD(24,36)', 12], ['=LCM(4,6)', 12], ['=INT(-2.5)', -3], ['=INT(2.5)', 2],
  ['=MOD(-3,2)', 1], ['=MOD(3,-2)', -1], ['=MOD(1,0)', '#DIV/0!'], ['=QUOTIENT(-7,2)', -3],
  ['=MROUND(10,3)', 9], ['=MROUND(-10,-3)', -9], ['=MROUND(10,-3)', '#NUM!'],
  ['=POWER(2,10)', 1024], ['=PRODUCT(A1:A5)', 60], ['=SIGN(-2)', -1], ['=SIGN(0)', 0],
  ['=SQRT(16)', 4], ['=SQRT(-1)', '#NUM!'], ['=SUMSQ(A1:A5)', 52],
  ['=ROUNDDOWN(-2.567,2)', -2.56], ['=ROUNDUP(2.121,2)', 2.13], ['=ROUNDUP(-2.121,2)', -2.13], ['=ROUNDUP(1234,-2)', 1300],
  ['=TRUNC(-2.7)', -2], ['=TRUNC(3.14159,2)', 3.14],
  // Statistics
  ['=AVERAGEA(D1:D4)', 1 / 3], ['=AVERAGEIF(B1:B5,"a",C1:C5)', 30], ['=AVERAGEIFS(C1:C5,B1:B5,"a",A1:A5,">3")', 40],
  ['=COUNTBLANK(D1:D5)', 1], ['=COUNTIFS(B1:B5,"a",A1:A5,">=4")', 2],
  ['=MAXIFS(C1:C5,B1:B5,"a")', 50], ['=MINIFS(C1:C5,B1:B5,"a")', 10],
  ['=LARGE(A1:A5,2)', 4], ['=SMALL(A1:A5,2)', 1], ['=MEDIAN(A1:A5)', 3], ['=MEDIAN(1,2,3,4)', 2.5], ['=MODE(A1:A5)', 1],
  ['=CORREL(A1:A5,C1:C5)', 40 / Math.sqrt(12800)],
  ['=STDEV(A1:A5)', Math.sqrt(3.2)], ['=STDEV.S(A1:A5)', Math.sqrt(3.2)], ['=STDEV.P(A1:A5)', 1.6], ['=STDEVP(A1:A5)', 1.6],
  ['=VAR(A1:A5)', 3.2], ['=VAR.S(A1:A5)', 3.2], ['=VAR.P(A1:A5)', 2.56], ['=VARP(A1:A5)', 2.56],
  ['=PERCENTILE(A1:A5,0.5)', 3], ['=PERCENTILE.INC({1,2,3,4},0.25)', 1.75],
  ['=QUARTILE({1,2,3,4,5,6,7,8},1)', 2.75], ['=QUARTILE.INC({1,2,3,4,5,6,7,8},3)', 6.25],
  ['=RANK(4,A1:A5)', 2], ['=RANK(1,A1:A5)', 4], ['=RANK.EQ(1,A1:A5,1)', 1], ['=RANK.EQ(5,A1:A5,1)', 5],
  // Logic and information
  ['=AND(TRUE,1,0)', false], ['=AND(D1:D2)', true], ['=OR(FALSE,0,1)', true], ['=NOT(TRUE)', false],
  ['=XOR(TRUE,TRUE)', false], ['=XOR(TRUE,TRUE,TRUE)', true], ['=TRUE()', true], ['=FALSE()', false],
  ['=IFNA(NA(),"x")', 'x'], ['=IFNA(1,"x")', 1], ['=NA()', '#N/A'],
  ['=IFS(1>2,"a",TRUE,"b")', 'b'], ['=IFS(FALSE,1)', '#N/A'],
  ['=SWITCH(2,1,"a",2,"b","c")', 'b'], ['=SWITCH(9,1,"a","c")', 'c'], ['=SWITCH(9,1,"a")', '#N/A'],
  ['=CHOOSE(2,"x","y","z")', 'y'], ['=CHOOSE(4,"x","y","z")', '#VALUE!'],
  ['=ISBLANK(D3)', true], ['=ISBLANK(D2)', false], ['=ISERR(NA())', false], ['=ISERR(1/0)', true], ['=ISERROR(NA())', true],
  ['=ISEVEN(4)', true], ['=ISEVEN(3)', false], ['=ISODD(3)', true], ['=ISFORMULA(D5)', true], ['=ISFORMULA(D4)', false],
  ['=ISLOGICAL(TRUE)', true], ['=ISLOGICAL(1)', false], ['=ISNA(NA())', true], ['=ISNA(1/0)', false],
  ['=ISNONTEXT(1)', true], ['=ISNONTEXT("a")', false], ['=ISNUMBER("1")', false], ['=ISNUMBER(1)', true], ['=ISTEXT("a")', true],
  ['=ISREF(A1)', true], ['=ISREF(1)', false], ['=LAMBDA(a,b,ISOMITTED(b))(1)', true], ['=LAMBDA(a,b,ISOMITTED(b))(1,2)', false],
  ['=LAMBDA(x,x*2)(21)', 42], ['=LET(add,LAMBDA(a,LAMBDA(b,a+b)),add(1)(41))', 42], ['=LAMBDA(x,x)(1/0)', '#DIV/0!'], ['=SUM(1)(2)', '#VALUE!'],
  ['=ERROR.TYPE(1/0)', 2], ['=ERROR.TYPE(NA())', 7], ['=ERROR.TYPE(1)', '#N/A'],
  ['=N(TRUE)', 1], ['=N("a")', 0], ['=TYPE(1)', 1], ['=TYPE("a")', 2], ['=TYPE(TRUE)', 4], ['=TYPE(NA())', 16], ['=TYPE({1,2})', 64],
  // Lookup and reference
  ['=HLOOKUP("b",{"a","b";1,2},2,FALSE)', 2], ['=HLOOKUP("z",{"a","b";1,2},2,FALSE)', '#N/A'],
  ['=LOOKUP(4.5,{1,3,4,6},{"a","b","c","d"})', 'c'], ['=LOOKUP(0,{1,3})', '#N/A'],
  ['=XMATCH(4,A1:A5)', 3], ['=XMATCH(2,{1,3,5},-1)', 1], ['=XMATCH(2,{1,3,5},1)', 2], ['=XMATCH(9,{1,3,5})', '#N/A'],
  ['=INDIRECT("C2")', 20], ['=INDIRECT("R2C3",FALSE)', 20], ['=SUM(INDIRECT("C1:C3"))', 60], ['=INDIRECT("C2",TRUE)', 20],
  // Relative R1C1 counts from the formula's cell (A1 for an ad-hoc evaluation).
  ['=INDIRECT("R[1]C[2]",FALSE)', 20], ['=SUM(INDIRECT("R1C3:R3C3",FALSE))', 60], ['=SUM(INDIRECT("Sheet1!R1C3:R5C3",FALSE))', 150],
  ['=SUM(INDIRECT("C3",FALSE))', 150], ['=SUM(INDIRECT("R2",FALSE))', 21 + 0], ['=INDIRECT("R0C1",FALSE)', '#REF!'], ['=INDIRECT("XYZ",FALSE)', '#REF!'],
  ['=COLUMN(C5)', 3], ['=ROW(C5)', 5],
  // Text
  ['=CHAR(65)', 'A'], ['=CODE("A")', 65], ['=CLEAN("a"&CHAR(7)&"b")', 'ab'], ['=CONCATENATE("a",1,TRUE)', 'a1TRUE'],
  ['=EXACT("a","A")', false], ['=EXACT("a","a")', true], ['=FIXED(1234.567,1,TRUE)', '1234.6'], ['=FIXED(1234.567,1)', '1,234.6'],
  ['=LOWER("AbC")', 'abc'], ['=PROPER("hello wORLD")', 'Hello World'], ['=MID("Hello",2,3)', 'ell'], ['=MID("Hello",9,3)', ''],
  ['=RIGHT("Hello",2)', 'lo'], ['=RIGHT("Hello")', 'o'], ['=SEARCH("l","Hello")', 3], ['=SEARCH("L*O","Hello")', 3], ['=SEARCH("z","Hello")', '#VALUE!'],
  ['=SUBSTITUTE("a-b-c","-","+",2)', 'a-b+c'], ['=SUBSTITUTE("aaa","a","b")', 'bbb'],
  ['=TRIM("  a   b  ")', 'a b'], ['=VALUE("1.5")', 1.5], ['=VALUE("abc")', '#VALUE!'],
  // Date and time (serials: 2024-01-01 = 45292)
  ['=DATEVALUE("2024-01-15")', 45306], ['=DAY(45306)', 15], ['=MONTH(45306)', 1], ['=YEAR(45306)', 2024],
  ['=DAYS(DATE(2024,3,1),DATE(2024,1,15))', 46], ['=EDATE(DATE(2024,1,31),1)', 45351], ['=EDATE(DATE(2024,3,31),-1)', 45351],
  ['=TIME(10,30,15)', 37815 / 86400], ['=HOUR(0.75)', 18], ['=MINUTE(TIME(10,30,15))', 30], ['=SECOND(TIME(10,30,15))', 15],
  ['=TIMEVALUE("12:00")', 0.5],
  ['=WEEKDAY(DATE(2024,1,1))', 2], ['=WEEKDAY(DATE(2024,1,1),2)', 1], ['=WEEKDAY(DATE(2024,1,7),3)', 6],
  ['=WEEKNUM(DATE(2024,1,1))', 1], ['=WEEKNUM(DATE(2024,1,7))', 2], ['=WEEKNUM(DATE(2024,12,31))', 53],
  ['=ISOWEEKNUM(DATE(2024,1,1))', 1], ['=ISOWEEKNUM(DATE(2021,1,1))', 53],
  ['=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31))', 23], ['=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31),DATE(2024,1,1))', 22],
  ['=WORKDAY(DATE(2024,1,5),1)', 45299], ['=WORKDAY(DATE(2024,1,8),-1)', 45296],
];

describe('Excel functions match Excel results', () => {
  const wb = book();
  it.each(CASES)('%s', (formula, expected) => {
    const value = wb.evaluate(formula);
    if (typeof expected === 'number') {
      expect(typeof value, `${formula} -> ${String(value)}`).toBe('number');
      expect(value as number).toBeCloseTo(expected, 9);
    } else if (typeof expected === 'string' && expected.startsWith('#')) expect(String(value)).toBe(expected);
    else expect(value).toBe(expected);
  });

  it('RAND and RANDBETWEEN stay within their bounds', () => {
    for (let i = 0; i < 200; i++) {
      const r = wb.evaluate('=RAND()') as number;
      expect(r >= 0 && r < 1).toBe(true);
      const n = wb.evaluate('=RANDBETWEEN(1,3)') as number;
      expect([1, 2, 3]).toContain(n);
    }
  });
});

describe('test coverage of the function library', () => {
  it('every registered function is called in at least one test', () => {
    new Workbook();
    const dirs = ['packages/core/test', 'packages/agent/test', 'packages/ui/test', 'test'].map((d) => new URL(`../../../${d}/`, import.meta.url));
    let text = '';
    for (const dir of dirs) for (const file of readdirSync(dir)) if (file.endsWith('.ts')) text += readFileSync(new URL(file, dir), 'utf8');
    const untested = listFunctions().map((f) => f.name).filter((name) => !new RegExp(`\\b${name.replace(/\./g, '\\.')}\\(`, 'i').test(text));
    expect(untested).toEqual([]);
  });
});
