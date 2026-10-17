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

  it('discovers and uses financial functions through tools', () => {
    const wb = new Workbook();
    const agent = new SpreadsheetAgent(wb);
    const result = agent.execute('list_functions', { category: 'financial' });
    expect(result.ok).toBe(true);
    const functions = (result.result as { functions: { name: string }[] }).functions;
    expect(functions.map((f) => f.name)).toEqual(['FV', 'IPMT', 'IRR', 'NPER', 'NPV', 'PMT', 'PPMT', 'PV', 'RATE', 'XIRR']);
    const schema = agent.tools.find((tool) => tool.name === 'list_functions')!.inputSchema;
    expect(JSON.stringify(schema)).toContain('financial');
    expect(agent.execute('write_range', { start: 'A1', values: [['=PMT(0,12,1200)']] }).ok).toBe(true);
    expect(wb.getValue(0, 0, 0)).toBe(-100);
    expect(agent.execute('undo', {}).ok).toBe(true);
    expect(wb.getValue(0, 0, 0)).toBe(null);
  });

  it('handles spilled arrays and conditional formatting through the same tools as the UI', () => {
    const wb = new Workbook();
    const agent = new SpreadsheetAgent(wb);
    expect(agent.execute('write_range', { start: 'A1', values: [['=SEQUENCE(3)']] }).ok).toBe(true);
    expect(agent.execute('read_range', { range: 'A1:A3' }).result).toMatchObject({ values: [[1], [2], [3]] });
    expect(agent.execute('write_range', { start: 'A2', values: [[42]] }).ok).toBe(false);
    const rules = [{ id: 'high', range: 'A1:A3', type: 'cellIs', operator: 'greaterThan', formulas: ['1'], style: { numFmt: '0.00', fill: '#FF0000' } }];
    expect(agent.execute('set_conditional_formats', { rules }).ok).toBe(true);
    expect(agent.execute('read_range', { range: 'A1:A3', format: 'display' }).result).toMatchObject({ values: [['1'], ['2.00'], ['3.00']] });
    expect(agent.execute('list_conditional_formats', {}).result).toEqual({ sheet: 'Sheet1', rules });
    expect(agent.execute('undo', {}).ok).toBe(true);
    expect(wb.activeSheet.conditionalFormats).toHaveLength(0);
  });

  it('creates, refreshes, lists and removes pivot reports with undo', () => {
    const wb = new Workbook();
    const agent = new SpreadsheetAgent(wb);
    agent.execute('write_range', { start: 'A1', values: [['Team', 'Sales'], ['A', 10], ['A', 20], ['B', 30]] });
    expect(agent.execute('create_pivot_table', { id: 'sales', source_range: 'A1:B4', target_cell: 'D1', rows: ['Team'], values: [{ field: 'Sales', aggregate: 'sum' }] }).ok).toBe(true);
    expect(wb.getValue(0, 3, 4)).toBe(60);
    expect(agent.execute('list_pivot_tables', {}).result).toMatchObject({ tables: [{ id: 'sales' }] });
    agent.execute('write_range', { start: 'B2', values: [[100]] });
    expect(agent.execute('refresh_pivot_table', { id: 'sales' }).ok).toBe(true);
    expect(wb.getValue(0, 3, 4)).toBe(150);
    agent.execute('undo', {});
    expect(wb.getValue(0, 3, 4)).toBe(60);
    expect(agent.execute('remove_pivot_table', { id: 'sales' }).ok).toBe(true);
    expect(wb.getValue(0, 3, 4)).toBe(null);
    agent.execute('undo', {});
    expect(wb.getValue(0, 3, 4)).toBe(60);
  });

  it('formats, discovers and removes tables without changing the data', () => {
    const wb = new Workbook();
    const agent = new SpreadsheetAgent(wb);
    agent.execute('write_range', { start: 'A1', values: [['Name', 'Value'], ['B', 20], ['A', 10]] });
    expect(agent.execute('format_table', { range: 'A1:B3', name: 'Sales', theme: 'green' }).ok).toBe(true);
    expect(agent.execute('list_tables', {}).result).toMatchObject({ tables: [{ name: 'Sales', range: 'A1:B3' }] });
    expect(wb.getCellStyle(0, 0, 0)?.fill).toBe('#548235');
    agent.execute('sort_range', { range: 'A1:B3', by: [{ column: 'A' }], has_header: true });
    expect(wb.getValue(0, 1, 0)).toBe('A');
    expect(agent.execute('remove_table', { name: 'Sales' }).ok).toBe(true);
    expect(wb.activeSheet.tables).toHaveLength(0);
    expect(wb.getValue(0, 1, 1)).toBe(10);
    agent.execute('undo', {});
    expect(wb.activeSheet.tables).toHaveLength(1);
  });
});

describe('replace, visibility and paste options', () => {
  it('replaces text, hides rows and transposes copies through tools', () => {
    const wb = new Workbook();
    const agent = new SpreadsheetAgent(wb);
    expect(agent.execute('write_range', { start: 'A1', values: [['red', 1], ['blue', 2], ['red', 3], ['=SUBTOTAL(109,B1:B3)', null]] }).ok).toBe(true);
    expect(agent.execute('replace', { query: 'red', replacement: 'green' }).result).toEqual({ replaced: 2 });
    expect(wb.getValue(0, 2, 0)).toBe('green');
    const hidden = agent.execute('set_hidden', { rows: '2', columns: 'C:D', hidden: true });
    expect(hidden.result).toEqual({ hidden_rows: [2], hidden_columns: ['C', 'D'] });
    expect(wb.getValue(0, 3, 0)).toBe(4);
    expect(agent.execute('set_hidden', { rows: '1:3', hidden: false }).ok).toBe(true);
    expect(wb.getValue(0, 3, 0)).toBe(6);
    expect(agent.execute('set_hidden', { hidden: true }).ok).toBe(false);
    expect(agent.execute('copy_range', { source: 'A1:B2', destination: 'F1', transpose: true }).result).toEqual({ pasted: 'Sheet1!F1:G2' });
    expect([wb.getValue(0, 0, 5), wb.getValue(0, 0, 6), wb.getValue(0, 1, 5), wb.getValue(0, 1, 6)]).toEqual(['green', 'blue', 1, 2]);
  });
});
