import { describe, expect, it } from 'vitest';
import { Workbook } from '@sheets/core';
import { SpreadsheetAgent, toAnthropicTools } from '../src/index.js';

describe('SpreadsheetAgent', () => {
  it('builds a small report via tools', () => {
    const wb = new Workbook();
    const agent = new SpreadsheetAgent(wb);
    const ok = (name: string, input: unknown) => {
      const r = agent.execute(name, input);
      if (!r.ok) throw new Error(r.error);
      return r.result as Record<string, unknown>;
    };

    ok('write_range', {
      start: 'A1',
      values: [
        ['Product', 'Qty', 'Price'],
        ['Apples', 3, 1.5],
        ['Pears', 5, 2],
        ['Plums', 2, 4],
      ],
    });
    ok('write_range', { start: 'D1', values: [['Total']] });
    const fill = ok('fill_formula', { range: 'D2:D4', formula: '=B2*C2' });
    expect(fill.cells).toBe(3);
    expect(wb.getValue(0, 3, 3)).toBe(8);

    const sum = ok('write_range', { start: 'C5', values: [['Sum', '=SUM(D2:D4)']] });
    expect(sum.formulaResults).toEqual({ D5: 22.5 });

    ok('format_range', { range: 'A1:D1', style: { bold: true, fill: '#DDEBF7' }, borders: 'outline' });
    expect(wb.activeSheet.getCell(0, 0)!.style?.bold).toBe(true);

    ok('sort_range', { range: 'A1:D4', by: [{ column: 'D', descending: true }] });
    expect(wb.getValue(0, 1, 0)).toBe('Pears');
    expect(wb.activeSheet.getCell(1, 3)!.formula).toBe('B2*C2');

    const chart = ok('create_chart', { type: 'column', data_range: 'A1:A4', title: 'Qty' });
    expect(chart.id).toBeTruthy();

    const read = ok('read_range', { range: 'A1:D5' });
    expect((read.values as unknown[][])[4]![3]).toBe(22.5);
    expect((read.formulas as Record<string, string>).D5).toBe('=SUM(D2:D4)');

    const info = ok('get_workbook_info', {});
    expect((info.sheets as { usedRange: string }[])[0]!.usedRange).toBe('A1:D5');

    // A whole tool call is a single undo step.
    ok('undo', {});
    expect(wb.activeSheet.charts).toHaveLength(0);
  });

  it('reports helpful errors', () => {
    const agent = new SpreadsheetAgent(new Workbook());
    expect(agent.execute('read_range', {}).error).toMatch(/range/);
    expect(agent.execute('read_range', { range: 'Nope!A1' }).error).toMatch(/not found/);
    expect(agent.execute('nope', {}).ok).toBe(false);
  });

  it('exports provider tool formats', () => {
    const tools = toAnthropicTools(new SpreadsheetAgent(new Workbook()).tools);
    expect(tools.find((t) => t.name === 'write_range')?.input_schema).toBeTruthy();
  });
});
