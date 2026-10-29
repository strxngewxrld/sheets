import { describe, expect, it } from 'vitest';
import type { ChartSpec, ResolvedChart } from '@sheets/core';
import { renderChartSvg } from '../src/charts.js';

const theme = { dark: false, background: '#fff', text: '#000', muted: '#666', grid: '#ccc', fontFamily: 'sans-serif' };
const spec: ChartSpec = { id: 'c1', type: 'column', series: [{ values: 'Sheet1!B2:B3' }], anchor: { from: { row: 0, col: 0 }, to: { row: 10, col: 6 } }, legend: 'right' };

describe('chart SVG', () => {
  it('never writes an unsafe series color into the markup', () => {
    const data: ResolvedChart = { categories: ['a', 'b'], series: [{ name: 's', values: [1, 2], color: '#000"/><foreignObject><div>x</div></foreignObject><rect fill="' }] };
    const svg = renderChartSvg(data, spec, 400, 300, theme);
    expect(svg).not.toContain('foreignObject');
    expect(svg).not.toContain('#000"');
  });

  it('keeps valid series colors', () => {
    const data: ResolvedChart = { categories: ['a', 'b'], series: [{ name: 's', values: [1, 2], color: '#1A2B3C' }] };
    expect(renderChartSvg(data, spec, 400, 300, theme)).toContain('fill="#1A2B3C"');
  });

  it('escapes theme colors', () => {
    const data: ResolvedChart = { categories: ['a'], series: [{ name: 's', values: [1] }] };
    expect(renderChartSvg(data, spec, 400, 300, { ...theme, text: 'red"><g' })).not.toContain('red"><g');
  });
});
