import { readdirSync, readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import { listFunctions, readXlsx, setValues, Workbook, writeXlsx } from '../src/index.js';

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

  // Added functions; the values are the examples of Microsoft's function reference, rounded as printed there.
  // Date
  ['=DAYS360(DATE(2011,1,30),DATE(2011,12,31))', 330], ['=DAYS360(DATE(2011,1,1),DATE(2011,12,31))', 360], ['=DAYS360(DATE(2011,1,1),DATE(2011,12,31),TRUE)', 359],
  ['=DAYS360(DATE(2011,1,30),DATE(2011,2,28))', 28], ['=DAYS360(DATE(2011,1,15),DATE(2011,3,31))', 76],
  ['=ROUND(YEARFRAC(DATE(2012,1,1),DATE(2012,7,30)),8)', 0.58055556], ['=ROUND(YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),1),8)', 0.57650273],
  ['=ROUND(YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),3),8)', 0.57808219], ['=YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),2)', 211 / 360],
  ['=YEARFRAC(DATE(2012,7,30),DATE(2012,1,1),4)', 209 / 360], ['=YEARFRAC(DATE(2011,2,28),DATE(2012,2,29))', 1], ['=YEARFRAC(1,2,5)', '#NUM!'],
  ['=YEARFRAC(DATE(2010,3,15),DATE(2014,9,1),1)', 1631 / (1826 / 5)],
  ['=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31))', 22], ['=NETWORKDAYS.INTL(DATE(2006,2,28),DATE(2006,1,31))', -21],
  ['=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),7,{"2006/1/2","2006/1/16"})', 22],
  ['=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),"0010001",{"2006/1/2","2006/1/16"})', 20],
  ['=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31),11)', 26], ['=NETWORKDAYS.INTL(1,2,8)', '#NUM!'], ['=NETWORKDAYS.INTL(1,2,"101")', '#VALUE!'],
  ['=NETWORKDAYS.INTL(1,31,"1111111")', 0], ['=NETWORKDAYS(1,31,"x")', '#VALUE!'],
  ['=WORKDAY.INTL(DATE(2012,1,1),30,0)', '#NUM!'], ['=WORKDAY.INTL(DATE(2012,1,1),90,11)', 41013], ['=WORKDAY.INTL(DATE(2012,1,1),30,17)', 40944],
  ['=WORKDAY.INTL(DATE(2012,1,1),1,"0000011")', 40910], ['=WORKDAY.INTL(DATE(2012,1,1),1,"1111111")', '#VALUE!'],
  ['=DATEVALUE("2006/1/2")', 38719],
  // Financial
  ['=ROUND(XNPV(0.09,{-10000,2750,4250,3250,2750},{39448,39508,39751,39859,39904}),2)', 2086.65],
  ['=XNPV(0.09,{-10000,2750},{39448,39000})', '#NUM!'], ['=XNPV(0.09,{-10000,2750},{39448})', '#NUM!'],
  ['=SLN(30000,7500,10)', 2250], ['=SLN(1,0,0)', '#DIV/0!'],
  ['=ROUND(SYD(30000,7500,10,1),2)', 4090.91], ['=ROUND(SYD(30000,7500,10,10),2)', 409.09], ['=SYD(1,0,10,11)', '#NUM!'],
  ['=ROUND(DB(1000000,100000,6,1,7),2)', 186083.33], ['=ROUND(DB(1000000,100000,6,2,7),2)', 259639.42], ['=ROUND(DB(1000000,100000,6,7,7),2)', 15845.1],
  ['=DB(1000000,100000,6,7)', '#NUM!'], ['=DB(1000000,100000,6,1,13)', '#NUM!'],
  ['=ROUND(DDB(2400,300,10*365,1),2)', 1.32], ['=DDB(2400,300,10*12,1,2)', 40], ['=DDB(2400,300,10,1,2)', 480], ['=DDB(2400,300,10,2,1.5)', 306],
  ['=ROUND(DDB(2400,300,10,10),2)', 22.12], ['=DDB(2400,300,10,11)', '#NUM!'],
  ['=ROUND(EFFECT(0.0525,4),6)', 0.053543], ['=ROUND(NOMINAL(0.053543,4),6)', 0.0525], ['=EFFECT(0,4)', '#NUM!'], ['=NOMINAL(0.05,0)', '#NUM!'],
  ['=ROUND(CUMIPMT(0.09/12,30*12,125000,13,24,0),5)', -11135.23213], ['=ROUND(CUMIPMT(0.09/12,30*12,125000,1,1,0),2)', -937.5],
  ['=ROUND(CUMPRINC(0.09/12,30*12,125000,13,24,0),5)', -934.10712], ['=ROUND(CUMPRINC(0.09/12,30*12,125000,1,1,0),5)', -68.27827],
  ['=CUMIPMT(0.09/12,30*12,125000,1,1,1)', 0], ['=CUMIPMT(0.09/12,360,125000,24,13,0)', '#NUM!'], ['=CUMPRINC(0,360,125000,1,1,0)', '#NUM!'],
  // Statistics
  ['=MAXA({1,2},TRUE)', 2], ['=MINA(C1:D4)', 0], ['=MAXA(-1,"a")', '#VALUE!'], ['=MAXA(A1:D5)', '#DIV/0!'], ['=MINA(-1,TRUE)', -1],
  ['=ROUND(GEOMEAN(4,5,8,7,11,4,3),6)', 5.476987], ['=ROUND(HARMEAN(4,5,8,7,11,4,3),6)', 5.028376], ['=GEOMEAN(1,0)', '#NUM!'], ['=HARMEAN(-1)', '#NUM!'],
  ['=ROUND(AVEDEV(4,5,6,7,5,4,3),6)', 1.020408], ['=DEVSQ(4,5,8,7,11,4,3)', 48],
  ['=ROUND(SKEW(3,4,5,2,3,4,5,6,4,7),6)', 0.359543], ['=ROUND(KURT(3,4,5,2,3,4,5,6,4,7),6)', -0.151800], ['=SKEW(1,2)', '#DIV/0!'], ['=KURT(1,2,3)', '#DIV/0!'],
  ['=MODE.SNGL(A1:A5)', 1], ['=MODE.SNGL(1,2)', '#N/A'],
  ['=TEXTJOIN(",",,MODE.MULT({1,2,3,4,3,2,1,2,3}))', '2,3'], ['=MODE.MULT(1,2)', '#N/A'],
  ['=RANK.AVG(94,{89,88,92,101,94,97,95})', 4], ['=RANK.AVG(1,A1:A5)', 4.5], ['=RANK.AVG(1,A1:A5,1)', 1.5], ['=RANK.AVG(2,A1:A5)', '#N/A'],
  ['=PERCENTILE.EXC({1,2,3,6,6,6,7,8,9},0.25)', 2.5], ['=PERCENTILE.EXC({1,2,3,6,6,6,7,8,9},0.01)', '#NUM!'], ['=PERCENTILE.EXC({1,2,3},1)', '#NUM!'],
  ['=QUARTILE.EXC({6,7,15,36,39,40,41,42,43,47,49},1)', 15], ['=QUARTILE.EXC({6,7,15,36,39,40,41,42,43,47,49},3)', 43], ['=QUARTILE.EXC({1,2,3},4)', '#NUM!'],
  ['=PERCENTRANK.INC({13,12,11,8,4,3,2,1,1,1},2)', 0.333], ['=PERCENTRANK.INC({13,12,11,8,4,3,2,1,1,1},4)', 0.555],
  ['=PERCENTRANK.INC({13,12,11,8,4,3,2,1,1,1},8)', 0.666], ['=PERCENTRANK.INC({13,12,11,8,4,3,2,1,1,1},5)', 0.583],
  ['=PERCENTRANK({13,12,11,8,4,3,2,1,1,1},2)', 0.333], ['=PERCENTRANK({1,2},3)', '#N/A'], ['=PERCENTRANK({1,2},1,0)', '#NUM!'],
  ['=PERCENTRANK.EXC({1,2,3,6,6,6,7,8,9},7)', 0.7], ['=PERCENTRANK.EXC({1,2,3,6,6,6,7,8,9},5.43)', 0.381], ['=PERCENTRANK.EXC({1,2,3,6,6,6,7,8,9},5.43,1)', 0.3],
  ['=ROUND(SLOPE({2,3,9,1,8,7,5},{6,5,11,7,5,4,4}),6)', 0.305556], ['=ROUND(INTERCEPT({2,3,9,1,8},{6,5,11,7,5}),7)', 0.0483871],
  ['=SLOPE({1,2},{3,3})', '#DIV/0!'], ['=SLOPE({1,2},{1,2,3})', '#N/A'], ['=SLOPE(C1:C5,A1:A5)', 3.125],
  ['=ROUND(FORECAST(30,{6,7,9,15,21},{20,28,31,38,40}),6)', 10.607253], ['=ROUND(FORECAST.LINEAR(30,{6,7,9,15,21},{20,28,31,38,40}),6)', 10.607253],
  ['=ROUND(RSQ({2,3,9,1,8,7,5},{6,5,11,7,5,4,4}),5)', 0.05795], ['=ROUND(PEARSON({9,7,5,3,1},{10,6,1,5,3}),6)', 0.699379],
  ['=ROUND(STEYX({2,3,9,1,8,7,5},{6,5,11,7,5,4,4}),6)', 3.305719], ['=STEYX({1,2},{1,2})', '#DIV/0!'],
  ['=COVARIANCE.S({2,4,8},{5,11,12})', 29 / 3], ['=COVARIANCE.P({3,2,4,5,6},{9,7,12,15,17})', 5.2], ['=COVAR({3,2,4,5,6},{9,7,12,15,17})', 5.2], ['=COVARIANCE.S(1,2)', '#DIV/0!'],
  ['=ROUND(TREND({133890,135000,135790,137300,138130,139100,139900,141120,141890,143230,144000,145290},{1,2,3,4,5,6,7,8,9,10,11,12},13),0)', 146172],
  ['=TEXTJOIN(",",,TREND({1,2,3,4},{1,2,3,4},{5,6}))', '5,6'], ['=TEXTJOIN(",",,TREND({2;4;6}))', '2,4,6'], ['=TREND({2;4;6},,4)', 8],
  ['=TREND({2;4;6},{1;2;3},4,FALSE)', 8], ['=TREND({1;2;3},{1,2;2,1;3,3},{4,4})', 4], ['=TREND({1;2;3},{1,2;2,1},{4,4})', '#REF!'],
  ['=TEXTJOIN(",",,ROUND(GROWTH({33100,47300,69000,102000,150000,220000},{11,12,13,14,15,16},{17,18}),0))', '320197,468536'], ['=GROWTH({1,-1})', '#NUM!'],
  ['=TEXTJOIN(",",,FREQUENCY({79,85,78,85,50,81,95,88,97},{70,79,89}))', '1,2,4,2'], ['=TEXTJOIN(",",,FREQUENCY({1,2,3},{2,1}))', '1,1,1'],
  ['=ROWS(FREQUENCY(A1:A5,{2,4}))', 3],
  ['=ROUND(NORM.DIST(42,40,1.5,TRUE),7)', 0.9087888], ['=ROUND(NORM.DIST(42,40,1.5,FALSE),7)', 0.1093400], ['=NORM.DIST(1,0,0,TRUE)', '#NUM!'],
  ['=ROUND(NORMDIST(42,40,1.5,TRUE),7)', 0.9087888], ['=ROUND(NORM.S.DIST(1.333333,TRUE),6)', 0.908789], ['=ROUND(NORM.S.DIST(1.333333,FALSE),6)', 0.164010],
  ['=NORM.S.DIST(0,TRUE)', 0.5], ['=ROUND(NORMSDIST(1.333333),6)', 0.908789], ['=NORMSDIST(-40)', 0],
  ['=ROUND(NORM.INV(0.908789,40,1.5),6)', 42.000002], ['=ROUND(NORMINV(0.908789,40,1.5),6)', 42.000002], ['=NORM.INV(0,40,1.5)', '#NUM!'],
  ['=ROUND(NORM.S.INV(0.908789),6)', 1.333335], ['=ROUND(NORMSINV(0.908789),6)', 1.333335], ['=NORM.S.INV(0.5)', 0], ['=NORM.S.INV(1)', '#NUM!'],
  ['=NORM.S.DIST(NORM.S.INV(1E-10),TRUE)', 1e-10], ['=NORM.S.INV(NORM.S.DIST(-7,TRUE))', -7],
  ['=ROUND(STANDARDIZE(42,40,1.5),6)', 1.333333], ['=STANDARDIZE(1,1,0)', '#NUM!'],
  ['=ROWS(RANDARRAY(3,2))', 3], ['=COLUMNS(RANDARRAY(3,2))', 2], ['=RANDARRAY(0)', '#VALUE!'], ['=RANDARRAY(1,1,5,1)', '#VALUE!'], ['=RANDARRAY(1,1,0,1.5,TRUE)', '#VALUE!'],
  // Math and engineering
  ['=COMBIN(8,2)', 28], ['=COMBIN(30,15)', 155117520], ['=COMBIN(2,3)', '#NUM!'], ['=PERMUT(100,3)', 970200], ['=PERMUT(3,2)', 6], ['=PERMUT(0,0)', '#NUM!'],
  ['=SINH(1)', Math.sinh(1)], ['=COSH(4)', Math.cosh(4)], ['=TANH(0.5)', Math.tanh(0.5)], ['=ASINH(-2.5)', Math.asinh(-2.5)], ['=ACOSH(10)', Math.acosh(10)],
  ['=ACOSH(0.5)', '#NUM!'], ['=ATANH(0.76159416)', Math.atanh(0.76159416)], ['=ATANH(1)', '#NUM!'],
  ['=BASE(7,2)', '111'], ['=BASE(100,16,4)', '0064'], ['=BASE(15,2,10)', '0000001111'], ['=BASE(-1,2)', '#NUM!'], ['=BASE(1,37)', '#NUM!'],
  ['=DECIMAL("FF",16)', 255], ['=DECIMAL("111",2)', 7], ['=DECIMAL("zap",36)', 45745], ['=DECIMAL("2",2)', '#NUM!'],
  ['=ROMAN(499,0)', 'CDXCIX'], ['=ROMAN(499,1)', 'LDVLIV'], ['=ROMAN(499,2)', 'XDIX'], ['=ROMAN(499,3)', 'VDIV'], ['=ROMAN(499,4)', 'ID'],
  ['=ROMAN(499,FALSE)', 'ID'], ['=ROMAN(2024)', 'MMXXIV'], ['=ROMAN(3999)', 'MMMCMXCIX'], ['=ROMAN(0)', ''], ['=ROMAN(4000)', '#VALUE!'], ['=ROMAN(1,5)', '#VALUE!'],
  ['=ARABIC("LVII")', 57], ['=ARABIC("mcmxii")', 1912], ['=ARABIC("-XIV")', -14], ['=ARABIC("")', 0], ['=ARABIC(ROMAN(1999,4))', 1999], ['=ARABIC("ABC")', '#VALUE!'],
  ['=BIN2DEC(1100100)', 100], ['=BIN2DEC("1111111111")', -1], ['=BIN2DEC("12")', '#NUM!'], ['=BIN2DEC("11111111111")', '#NUM!'],
  ['=BIN2HEX(11111011,4)', '00FB'], ['=BIN2HEX("1110")', 'E'], ['=BIN2HEX("1111111111")', 'FFFFFFFFFF'],
  ['=BIN2OCT(1001,3)', '011'], ['=BIN2OCT("1100100")', '144'], ['=BIN2OCT("1111111111")', '7777777777'],
  ['=DEC2BIN(9,4)', '1001'], ['=DEC2BIN(-100)', '1110011100'], ['=DEC2BIN(512)', '#NUM!'], ['=DEC2BIN(9,3)', '#NUM!'], ['=DEC2BIN(9,-1)', '#NUM!'],
  ['=DEC2HEX(100,4)', '0064'], ['=DEC2HEX(-54)', 'FFFFFFFFCA'], ['=DEC2HEX(28)', '1C'], ['=DEC2HEX(64,1)', '#NUM!'],
  ['=DEC2OCT(58,3)', '072'], ['=DEC2OCT(-100)', '7777777634'],
  ['=HEX2BIN("F",8)', '00001111'], ['=HEX2BIN("B7")', '10110111'], ['=HEX2BIN("FFFFFFFE00")', '1000000000'], ['=HEX2BIN("200")', '#NUM!'],
  ['=HEX2DEC("A5")', 165], ['=HEX2DEC("FFFFFFFF5B")', -165], ['=HEX2DEC("3DA408B9")', 1034160313], ['=HEX2DEC("G")', '#NUM!'],
  ['=HEX2OCT("F",3)', '017'], ['=HEX2OCT("3B4E")', '35516'], ['=HEX2OCT("FFFFFFFF00")', '7777777400'],
  ['=OCT2BIN(3,3)', '011'], ['=OCT2BIN("7777777000")', '1000000000'], ['=OCT2DEC(54)', 44], ['=OCT2DEC("7777777533")', -165],
  ['=OCT2HEX(100,4)', '0040'], ['=OCT2HEX("7777777533")', 'FFFFFFFF5B'], ['=OCT2HEX(8)', '#NUM!'],
  // Text and lookup
  ['=LENB("Müller")', 6], ['=UNICHAR(66)', 'B'], ['=UNICHAR(32)', ' '], ['=UNICHAR(128512)', '😀'], ['=UNICHAR(0)', '#VALUE!'], ['=UNICHAR(55296)', '#N/A'],
  ['=UNICODE("B")', 66], ['=UNICODE("😀")', 128512], ['=UNICODE("")', '#VALUE!'],
  ['=NUMBERVALUE("2.500,27",",",".")', 2500.27], ['=NUMBERVALUE("3.5%")', 0.035], ['=NUMBERVALUE(" 1 234.5 ")', 1234.5], ['=NUMBERVALUE("")', 0],
  ['=NUMBERVALUE("1.2.3")', '#VALUE!'], ['=NUMBERVALUE("1,000.5,0")', '#VALUE!'], ['=NUMBERVALUE("1,5",",",",")', '#VALUE!'], ['=NUMBERVALUE("5%%")', 0.0005],
  ['=VALUETOTEXT("a")', 'a'], ['=VALUETOTEXT("a",1)', '"a"'], ['=VALUETOTEXT(1.5,1)', '1.5'], ['=VALUETOTEXT(TRUE)', 'TRUE'], ['=VALUETOTEXT(1/0)', '#DIV/0!'],
  ['=VALUETOTEXT(1,2)', '#VALUE!'], ['=ARRAYTOTEXT(A1:B2)', '3, a, 1, b'], ['=ARRAYTOTEXT(A1:B2,1)', '{3,"a";1,"b"}'], ['=ARRAYTOTEXT({1,"a";TRUE,""},1)', '{1,"a";TRUE,""}'],
  ['=ADDRESS(2,3)', '$C$2'], ['=ADDRESS(2,3,2)', 'C$2'], ['=ADDRESS(2,3,3)', '$C2'], ['=ADDRESS(2,3,2,FALSE)', 'R2C[3]'],
  ['=ADDRESS(2,3,1,FALSE,"[Book1]Sheet1")', "'[Book1]Sheet1'!R2C3"], ['=ADDRESS(2,3,1,FALSE,"EXCEL SHEET")', "'EXCEL SHEET'!R2C3"],
  ['=ADDRESS(2,3,4,TRUE,"Sheet1")', 'Sheet1!C2'], ['=ADDRESS(0,1)', '#VALUE!'], ['=ADDRESS(1,16385)', '#VALUE!'], ['=ADDRESS(1,1,5)', '#VALUE!'],
  ['=INDIRECT(ADDRESS(2,3))', 20], ['=AREAS(A1:B2)', 1], ['=AREAS(1)', '#VALUE!'],
  ['=SHEET()', 1], ['=SHEET("sheet1")', 1], ['=SHEET("nope")', '#N/A'], ['=SHEET(C3)', 1], ['=SHEETS()', 1], ['=SHEETS(A1)', 1],
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

  it('SHEET and SHEETS count the sheets in workbook order', () => {
    const multi = book();
    multi.addSheet('Zwei');
    multi.addSheet('Drei Teil');
    expect(multi.evaluate('=SHEETS()')).toBe(3);
    expect(multi.evaluate("=SHEET('Drei Teil'!A1)")).toBe(3);
    expect(multi.evaluate('=SHEET("zwei")')).toBe(2);
    expect(multi.evaluate('=ADDRESS(1,1,1,TRUE,"Drei Teil")')).toBe("'Drei Teil'!$A$1");
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

describe('newer functions in XLSX files', () => {
  it('are written with the _xlfn. prefix Excel needs and read back without it', () => {
    const wb = new Workbook();
    const formulas = ['=NORM.S.INV(0.5)', '=PERCENTILE.EXC({1,2,3},0.5)', '=NETWORKDAYS.INTL(1,7,11)', '=YEARFRAC(1,366)', '=BIN2DEC("101")'];
    setValues(wb, 0, 0, 0, [formulas]);
    const xml = strFromU8(unzipSync(writeXlsx(wb))['xl/worksheets/sheet1.xml']!);
    expect(xml).toContain('_xlfn.NORM.S.INV(');
    expect(xml).toContain('_xlfn.PERCENTILE.EXC(');
    expect(xml).toContain('_xlfn.NETWORKDAYS.INTL(');
    // Functions Excel 2007 already knew are stored without a prefix.
    expect(xml).not.toContain('_xlfn.YEARFRAC');
    expect(xml).not.toContain('_xlfn.BIN2DEC');
    const back = readXlsx(writeXlsx(wb));
    expect(formulas.map((_, c) => back.sheets[0]!.getCell(0, c)?.formula)).toEqual(formulas.map((f) => f.slice(1)));
    expect(formulas.map((_, c) => back.getValue(0, 0, c))).toEqual([0, 2, 6, 1, 5]);
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
