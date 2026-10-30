// Run after `npm run typecheck`. Every probe runs in an isolated child process.
import { spawnSync } from 'node:child_process';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const script = fileURLToPath(import.meta.url);
const root = join(dirname(script), '..');

if (process.argv[2] === '--probe') {
  const { Workbook, openWorkbook, resolveChart, writeXlsx } = await import('../packages/core/dist/index.js');
  const { strFromU8, strToU8, zipSync, unzipSync } = await import('fflate');
  const formula = process.argv[3];
  if (process.argv[4] === 'chart') {
    // The argument is the category reference of a chart in an otherwise small file.
    const wb = new Workbook();
    wb.activeSheet.charts.push({ id: 'c1', type: 'column', series: [{ values: 'Sheet1!A1:A1', categories: formula }], anchor: { from: { row: 1, col: 1 }, to: { row: 10, col: 8 } } });
    const bytes = writeXlsx(wb);
    console.log(JSON.stringify({ phase: 'opening', bytes: bytes.length }));
    const back = openWorkbook(bytes, 'probe.xlsx');
    console.log(JSON.stringify({ result: resolveChart(back, back.activeSheet.charts[0]).categories.length }));
  } else if (process.argv[4] === 'xlsx') {
    const files = unzipSync(writeXlsx(new Workbook()));
    const xml = strFromU8(files['xl/worksheets/sheet1.xml']);
    const escaped = formula.replace(/^=/, '').replace(/&/g, '&amp;').replace(/</g, '&lt;');
    files['xl/worksheets/sheet1.xml'] = strToU8(xml.replace('<sheetData>', `<sheetData><row r="1"><c r="A1"><f>${escaped}</f><v>0</v></c></row>`));
    const bytes = zipSync(files);
    console.log(JSON.stringify({ phase: 'opening', bytes: bytes.length }));
    const wb = openWorkbook(bytes, 'probe.xlsx');
    console.log(JSON.stringify({ result: wb.getValue(0, 0, 0) }));
  } else {
    const wb = new Workbook();
    const result = wb.calc.evaluateRaw(formula, wb.activeSheet);
    console.log(JSON.stringify(Array.isArray(result)
      ? { rows: result.length, cols: result[0]?.length, cells: result.length * (result[0]?.length ?? 0) }
      : { result }));
  }
} else {
  const cases = [
    ['workday-normal', '=WORKDAY(45000,1)', 'xlsx', 45001],
    ['workday-nonprogress', '=WORKDAY(45000,1e20)', 'xlsx', '#NUM!'],
    ['networkdays-nonprogress', '=NETWORKDAYS(1e20,1e20)', 'xlsx', '#NUM!'],
    ['oversized-range', '=SUM(B1:XFC1048575)', 'xlsx', '#CALC!'],
    ['wraprows-budget', '=WRAPROWS({1},1000001)', 'raw', '#CALC!'],
    ['broadcast-budget', '=SEQUENCE(1100)+TRANSPOSE(SEQUENCE(1100))', 'raw', '#CALC!'],
    ['textsplit-budget', '=TEXTSPLIT(REPT(",",1100)&REPT(";",1100),",",";")', 'raw', '#CALC!'],
    ['fact-loop', '=FACT(1e12)', 'xlsx', '#NUM!'],
    ['permut-loop', '=PERMUT(1e12,1e12)', 'xlsx', '#NUM!'],
    ['combin-loop', '=COMBIN(2e12,1e12)', 'xlsx', '#NUM!'],
    ['cumipmt-loop', '=CUMIPMT(0.01,1e12,1,1,1e12,0)', 'xlsx', '#CALC!'],
    ['db-loop', '=DB(100,1,1e12,1e12,12)', 'xlsx', '#CALC!'],
    ['rept-length', '=UPPER(REPT("a",4e8))', 'xlsx', '#VALUE!'],
    ['fixed-length', '=FIXED(1,3e8)', 'xlsx', '#VALUE!'],
    ['text-format-length', '=TEXT(1,REPT("0",3e8))', 'xlsx', '#VALUE!'],
    ['lambda-exponential', '=LET(f,LAMBDA(f,n,IF(n=0,0,f(f,n-1)+f(f,n-1))),f(f,40))', 'xlsx', '#CALC!'],
    ['makearray-nested', '=SUM(MAKEARRAY(1000,1000,LAMBDA(r,c,SUM(SEQUENCE(1000)))))', 'xlsx', '#CALC!'],
    ['chart-categories', 'Sheet1!B1:XFD1048576', 'chart', 1],
  ];
  for (const [name, formula, mode, expected] of cases) {
    const started = performance.now();
    const result = spawnSync(process.execPath,
      ['--max-old-space-size=160', script, '--probe', formula, ...(mode ? [mode] : [])],
      { cwd: root, encoding: 'utf8', timeout: 5000, maxBuffer: 64 * 1024, windowsHide: true });
    let actual;
    try {
      const last = JSON.parse(result.stdout.trim().split(/\r?\n/).at(-1));
      actual = last.result && typeof last.result === 'object' ? last.result.error : last.result;
    } catch { /* A timeout or process failure may leave no complete output. */ }
    const passed = result.status === 0 && actual === expected;
    if (!passed) process.exitCode = 1;
    console.log(JSON.stringify({ name, formula, expected, actual, passed, milliseconds: Math.round(performance.now() - started),
      status: result.status, signal: result.signal, error: result.error?.code,
      heapExhausted: /heap out of memory|Reached heap limit|Ineffective mark-compacts/i.test(result.stderr ?? ''),
      stdout: result.stdout?.trim() }));
  }
}
