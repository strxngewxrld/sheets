import {
  CHART_TYPES,
  chartFromRange,
  formatRange,
  parseRange,
  removeChart,
  resolveChart,
  updateChart,
  type ChartSpec,
  type ChartType,
  type LegendPosition,
  type Workbook,
} from '@sheets/core';
import { renderChartSvg, type ChartTheme } from './charts.js';
import type { Viewport } from './grid.js';
import type { Strings } from './i18n.js';

/**
 * DOM overlay that renders charts above the grid and lets users move,
 * resize, edit and delete them.
 */
export class ChartLayer {
  readonly el: HTMLDivElement;
  private inner: HTMLDivElement;
  private nodes = new Map<string, { el: HTMLDivElement; key: string }>();
  selectedId: string | null = null;
  private dialog: HTMLDivElement | null = null;

  constructor(
    parent: HTMLElement,
    private readonly getWorkbook: () => Workbook,
    private readonly vp: Viewport,
    private readonly getTheme: () => ChartTheme,
    private readonly strings: Strings,
    private readonly onWheel: (e: WheelEvent) => void,
    private readonly onSelect: () => void,
  ) {
    this.el = document.createElement('div');
    this.el.className = 'sx-charts';
    this.inner = document.createElement('div');
    this.inner.className = 'sx-charts-inner';
    this.el.appendChild(this.inner);
    parent.appendChild(this.el);
    this.el.addEventListener('wheel', (e) => this.onWheel(e), { passive: false });
  }

  /** Re-renders all charts of the active sheet. `force` re-renders the SVG even if position is unchanged. */
  update(force = false): void {
    const vp = this.vp;
    const wb = this.getWorkbook();
    const sheet = vp.sheet;
    const mainX = vp.headerW + vp.frozenW;
    const mainY = vp.headerH + vp.frozenH;
    this.el.style.left = `${mainX}px`;
    this.el.style.top = `${mainY}px`;
    this.el.style.width = `${Math.max(0, vp.width - mainX)}px`;
    this.el.style.height = `${Math.max(0, vp.height - mainY)}px`;
    this.inner.style.transform = `translate(${-vp.scrollX - vp.frozenW}px, ${-vp.scrollY - vp.frozenH}px)`;
    const seen = new Set<string>();
    const theme = this.getTheme();
    for (const chart of sheet.charts) {
      seen.add(chart.id);
      const x = vp.geo.cols.offset(chart.anchor.from.col);
      const y = vp.geo.rows.offset(chart.anchor.from.row);
      const w = Math.max(80, vp.geo.cols.offset(chart.anchor.to.col) - x);
      const h = Math.max(60, vp.geo.rows.offset(chart.anchor.to.row) - y);
      let node = this.nodes.get(chart.id);
      if (!node) {
        node = { el: this.createNode(chart.id), key: '' };
        this.nodes.set(chart.id, node);
        this.inner.appendChild(node.el);
      }
      node.el.style.left = `${x}px`;
      node.el.style.top = `${y}px`;
      node.el.style.width = `${w}px`;
      node.el.style.height = `${h}px`;
      node.el.classList.toggle('sx-selected', this.selectedId === chart.id);
      const data = resolveChart(wb, chart);
      const key = JSON.stringify([chart, data, w, h, theme.dark]);
      if (force || key !== node.key) {
        node.key = key;
        const svg = node.el.querySelector('.sx-svg')!;
        svg.innerHTML = renderChartSvg(data, chart, w, h, theme);
        node.el.setAttribute('aria-label', chart.title ?? chart.type);
      }
    }
    for (const [id, node] of this.nodes) {
      if (!seen.has(id)) {
        node.el.remove();
        this.nodes.delete(id);
        if (this.selectedId === id) this.selectedId = null;
      }
    }
  }

  private createNode(id: string): HTMLDivElement {
    const el = document.createElement('div');
    el.className = 'sx-chart';
    el.tabIndex = -1;
    el.setAttribute('role', 'img');
    const svg = document.createElement('div');
    svg.className = 'sx-svg';
    svg.style.width = '100%';
    svg.style.height = '100%';
    const handle = document.createElement('div');
    handle.className = 'sx-resize';
    el.append(svg, handle);

    el.addEventListener('mousedown', (e) => {
      if (e.button !== 0) return;
      e.preventDefault();
      e.stopPropagation();
      this.select(id);
      const resizing = e.target === handle;
      const startX = e.clientX;
      const startY = e.clientY;
      const left = parseFloat(el.style.left);
      const top = parseFloat(el.style.top);
      const width = parseFloat(el.style.width);
      const height = parseFloat(el.style.height);
      const move = (ev: MouseEvent) => {
        const dx = ev.clientX - startX;
        const dy = ev.clientY - startY;
        if (resizing) {
          el.style.width = `${Math.max(80, width + dx)}px`;
          el.style.height = `${Math.max(60, height + dy)}px`;
        } else {
          el.style.left = `${Math.max(0, left + dx)}px`;
          el.style.top = `${Math.max(0, top + dy)}px`;
        }
      };
      const up = () => {
        window.removeEventListener('mousemove', move);
        window.removeEventListener('mouseup', up);
        this.commitGeometry(id, el);
      };
      window.addEventListener('mousemove', move);
      window.addEventListener('mouseup', up);
    });
    el.addEventListener('dblclick', (e) => {
      e.stopPropagation();
      this.openEditor(id);
    });
    return el;
  }

  /** Snaps the dragged element to cells and stores the new anchor. */
  private commitGeometry(id: string, el: HTMLDivElement): void {
    const geo = this.vp.geo;
    const x = parseFloat(el.style.left);
    const y = parseFloat(el.style.top);
    const fromCol = geo.cols.indexAt(x + geo.cols.def / 2);
    const fromRow = geo.rows.indexAt(y + geo.rows.def / 2);
    const toCol = Math.max(fromCol + 1, geo.cols.indexAt(x + parseFloat(el.style.width) + geo.cols.def / 2));
    const toRow = Math.max(fromRow + 1, geo.rows.indexAt(y + parseFloat(el.style.height) + geo.rows.def / 2));
    const chart = this.vp.sheet.charts.find((c) => c.id === id);
    if (!chart) return;
    const a = chart.anchor;
    if (a.from.row === fromRow && a.from.col === fromCol && a.to.row === toRow && a.to.col === toCol) {
      this.update(true);
      return;
    }
    updateChart(this.getWorkbook(), id, { anchor: { from: { row: fromRow, col: fromCol }, to: { row: toRow, col: toCol } } });
  }

  select(id: string | null): void {
    this.selectedId = id;
    for (const [cid, node] of this.nodes) node.el.classList.toggle('sx-selected', cid === id);
    if (id) {
      this.nodes.get(id)?.el.focus({ preventScroll: true });
      this.onSelect();
    }
  }

  deleteSelected(): boolean {
    if (!this.selectedId) return false;
    removeChart(this.getWorkbook(), this.selectedId);
    this.selectedId = null;
    this.closeEditor();
    return true;
  }

  closeEditor(): void {
    this.dialog?.remove();
    this.dialog = null;
  }

  openEditor(id: string): void {
    const wb = this.getWorkbook();
    const found = wb.getChart(id);
    if (!found) return;
    const chart = found.chart;
    const s = this.strings;
    this.closeEditor();
    const d = document.createElement('div');
    d.className = 'sx-dialog';
    d.setAttribute('role', 'dialog');
    d.setAttribute('aria-label', s.chartEditor);
    const title = document.createElement('h3');
    title.textContent = s.chartEditor;
    d.appendChild(title);

    const field = (label: string, input: HTMLElement) => {
      const l = document.createElement('label');
      l.append(label, input);
      d.appendChild(l);
      return input;
    };
    const select = (options: [string, string][], value: string) => {
      const sel = document.createElement('select');
      for (const [v, text] of options) sel.add(new Option(text, v, false, v === value));
      return sel;
    };
    const text = (value: string) => {
      const i = document.createElement('input');
      i.type = 'text';
      i.value = value;
      return i;
    };
    const typeSel = field(s.chartType, select(CHART_TYPES.map((t) => [t, s[t]]), chart.type)) as HTMLSelectElement;
    const titleIn = field(s.chartTitle, text(chart.title ?? '')) as HTMLInputElement;
    const dataRange = guessDataRange(chart);
    const rangeIn = field(s.dataRange, text(dataRange ?? '')) as HTMLInputElement;
    const legendSel = field(
      s.legend,
      select((['right', 'bottom', 'top', 'left', 'none'] as const).map((p) => [p, s[p]]), chart.legend ?? 'right'),
    ) as HTMLSelectElement;
    const xIn = field(s.xAxisTitle, text(chart.xAxisTitle ?? '')) as HTMLInputElement;
    const yIn = field(s.yAxisTitle, text(chart.yAxisTitle ?? '')) as HTMLInputElement;
    const stackedLabel = document.createElement('label');
    stackedLabel.className = 'sx-check';
    const stackedIn = document.createElement('input');
    stackedIn.type = 'checkbox';
    stackedIn.checked = !!chart.stacked;
    stackedLabel.append(stackedIn, s.stacked);
    d.appendChild(stackedLabel);

    const actions = document.createElement('div');
    actions.className = 'sx-actions';
    const del = document.createElement('button');
    del.textContent = s.delete;
    del.addEventListener('click', () => {
      removeChart(wb, id);
      this.closeEditor();
    });
    const close = document.createElement('button');
    close.textContent = s.close;
    close.addEventListener('click', () => this.closeEditor());
    const apply = document.createElement('button');
    apply.textContent = s.apply;
    apply.className = 'sx-primary';
    apply.addEventListener('click', () => {
      const type = typeSel.value as ChartType;
      const patch: Partial<Omit<ChartSpec, 'id'>> = {
        type,
        title: titleIn.value || undefined,
        legend: legendSel.value as LegendPosition,
        stacked: stackedIn.checked || undefined,
        xAxisTitle: xIn.value || undefined,
        yAxisTitle: yIn.value || undefined,
      } as Partial<Omit<ChartSpec, 'id'>>;
      const r = rangeIn.value.trim();
      if (r && r !== dataRange) {
        const parsed = parseRange(r);
        if (parsed) patch.series = chartFromRange(wb, parsed.sheet ?? found.sheet.name, parsed.range, type).series;
        else {
          rangeIn.setCustomValidity('Invalid range');
          rangeIn.reportValidity();
          return;
        }
      }
      updateChart(wb, id, patch);
    });
    actions.append(del, close, apply);
    d.appendChild(actions);
    d.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Escape') this.closeEditor();
    });
    d.addEventListener('mousedown', (e) => e.stopPropagation());
    this.el.parentElement!.appendChild(d);
    this.dialog = d;
    typeSel.focus();
  }
}

/** Reconstructs the bounding data range of a chart's series (for the editor). */
function guessDataRange(chart: ChartSpec): string | undefined {
  const refs = chart.series.flatMap((s) => [s.values, s.categories, s.nameRef].filter((x): x is string => !!x));
  const parsed = refs.map((r) => parseRange(r)).filter((p): p is NonNullable<typeof p> => !!p);
  if (!parsed.length) return undefined;
  const sheet = parsed[0]!.sheet;
  const start = { row: Math.min(...parsed.map((p) => p.range.start.row)), col: Math.min(...parsed.map((p) => p.range.start.col)) };
  const end = { row: Math.max(...parsed.map((p) => p.range.end.row)), col: Math.max(...parsed.map((p) => p.range.end.col)) };
  return formatRange({ start, end }, sheet);
}
