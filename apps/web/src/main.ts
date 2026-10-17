import { formatFromFileName, LOCALE_DE, LOCALE_EN, openWorkbook, saveWorkbook, setDefaultLocale, Workbook, type FileFormat } from '@sheets/core';
import { AGENT_SYSTEM_PROMPT, serveAgent, SpreadsheetAgent } from '@sheets/agent';
import { SpreadsheetView } from '@sheets/ui';
import { createSampleWorkbook } from './sample.js';

const german = navigator.language.toLowerCase().startsWith('de');
const locale = german ? LOCALE_DE : LOCALE_EN;
setDefaultLocale(locale);

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const container = $('sheet');
const fileInput = $<HTMLInputElement>('file');
let fileName = german ? 'Mappe1.xlsx' : 'Book1.xlsx';

const icon = (d: string) =>
  `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${d}</svg>`;

const view = new SpreadsheetView(container, createSampleWorkbook(), {
  locale,
  toolbarActions: [
    { id: 'new', label: german ? 'Neu' : 'New', icon: icon('<path d="M4 1.5h5l3 3v10H4z"/><path d="M9 1.5v3h3"/>'), onClick: () => load(new Workbook({ sheetName: german ? 'Tabelle1' : 'Sheet1' }), german ? 'Mappe1.xlsx' : 'Book1.xlsx') },
    { id: 'open', label: german ? 'Öffnen' : 'Open', icon: icon('<path d="M1.5 4.5v9h11l2-6h-11l-2 6"/><path d="M1.5 4.5v-2h4l1.5 2h5v3"/>'), onClick: () => fileInput.click() },
    { id: 'save', label: german ? 'Speichern' : 'Save', icon: icon('<path d="M8 2v8M4.5 6.5 8 10l3.5-3.5"/><path d="M2 12v2h12v-2"/>'), onClick: () => download('xlsx') },
    { id: 'csv', label: 'CSV', title: german ? 'Aktives Blatt als CSV exportieren' : 'Export active sheet as CSV', onClick: () => download('csv') },
    { id: 'agent', label: 'Agent', title: 'Agent API', icon: icon('<rect x="2.5" y="4.5" width="11" height="8" rx="2"/><path d="M8 2v2.5M6 8.5h.01M10 8.5h.01"/>'), onClick: () => toggleAgent() },
  ],
});

const agent = new SpreadsheetAgent(() => view.workbook, { host: view.agentHost });
let stopBridge: (() => void) | null = null;

function load(wb: Workbook, name: string): void {
  fileName = name;
  view.setWorkbook(wb);
  startBridge();
  document.title = `${name} – Sheets`;
  view.focus();
}

async function openFile(file: File): Promise<void> {
  try {
    const bytes = new Uint8Array(await file.arrayBuffer());
    load(openWorkbook(bytes, file.name), file.name);
  } catch (e) {
    alert(`${german ? 'Datei konnte nicht geöffnet werden' : 'Could not open file'}: ${(e as Error).message}`);
  }
}

function download(format: FileFormat): void {
  const bytes = saveWorkbook(view.workbook, format, { csv: { delimiter: german ? ';' : ',' } });
  const base = fileName.replace(/\.[^.]+$/, '');
  const name = format === 'xlsx' ? `${base}.xlsx` : `${base}${format === 'csv' && view.workbook.sheets.length > 1 ? `-${view.sheet.name}` : ''}.${format}`;
  const type = format === 'xlsx' ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' : 'text/csv';
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type }));
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

fileInput.addEventListener('change', () => {
  const f = fileInput.files?.[0];
  if (f) void openFile(f);
  fileInput.value = '';
});

// Drag & drop
const drop = $('drop');
let dragDepth = 0;
container.addEventListener('dragenter', (e) => {
  e.preventDefault();
  dragDepth++;
  drop.classList.add('on');
});
container.addEventListener('dragover', (e) => e.preventDefault());
container.addEventListener('dragleave', () => {
  if (--dragDepth <= 0) drop.classList.remove('on');
});
container.addEventListener('drop', (e) => {
  e.preventDefault();
  dragDepth = 0;
  drop.classList.remove('on');
  const f = e.dataTransfer?.files[0];
  if (f && /\.(xlsx|xlsm|csv|tsv|txt|json)$/i.test(f.name)) void openFile(f);
});

// Ctrl+S / Ctrl+O
window.addEventListener('keydown', (e) => {
  if (!(e.ctrlKey || e.metaKey)) return;
  if (e.key.toLowerCase() === 's') {
    e.preventDefault();
    download(formatFromFileName(fileName) === 'csv' ? 'csv' : 'xlsx');
  } else if (e.key.toLowerCase() === 'o') {
    e.preventDefault();
    fileInput.click();
  }
});

// ---------------------------------------------------------------------------
// Agent console & programmatic access
// ---------------------------------------------------------------------------

const panel = $('agent');
const toolSel = $<HTMLSelectElement>('agent-tool');
const inputEl = $<HTMLTextAreaElement>('agent-input');
const outputEl = $('agent-output');
const descEl = $('agent-desc');

const EXAMPLES: Record<string, unknown> = {
  get_workbook_info: {},
  read_range: { range: 'A1:E8', format: 'markdown' },
  write_range: { start: 'A13', values: [['Prognose Jul', '=B7*1.05']] },
  fill_formula: { range: 'F2:F7', formula: '=B2*0.19' },
  format_range: { range: 'A13:B13', style: { italic: true, color: '#7F6000', fill: '#FFF2CC' } },
  create_chart: { type: 'line', data_range: 'A1:B7', title: 'Umsatztrend', position: 'G19' },
  evaluate_formula: { formula: '=SUMPRODUCT(B2:B7,E2:E7)/SUM(B2:B7)' },
  sort_range: { range: 'A1:E7', by: [{ column: 'B', descending: true }] },
  find: { query: 'Summe' },
};

for (const t of agent.tools) toolSel.add(new Option(t.name, t.name));
function showTool(): void {
  const t = agent.tools.find((x) => x.name === toolSel.value)!;
  descEl.textContent = t.description;
  const example = EXAMPLES[t.name] ?? Object.fromEntries((t.inputSchema.required ?? []).map((k) => [k, '']));
  inputEl.value = JSON.stringify(example, null, 2);
}
toolSel.addEventListener('change', showTool);
showTool();

$('agent-run').addEventListener('click', () => {
  let input: unknown;
  try {
    input = JSON.parse(inputEl.value || '{}');
  } catch (e) {
    outputEl.textContent = `Invalid JSON: ${(e as Error).message}`;
    return;
  }
  const result = agent.execute(toolSel.value, input);
  outputEl.textContent = JSON.stringify(result, null, 2);
});
$('agent-close').addEventListener('click', () => toggleAgent(false));

function toggleAgent(open = !panel.classList.contains('open')): void {
  panel.classList.toggle('open', open);
  window.dispatchEvent(new Event('resize'));
}

/**
 * Global API for automation (browser devtools, Playwright, browser agents):
 *   sheets.execute('write_range', { start: 'A1', values: [[1, 2]] })
 */
const api = {
  get workbook() {
    return view.workbook;
  },
  view,
  get agent() {
    return agent;
  },
  tools: () => agent.tools,
  execute: (name: string, input?: unknown) => agent.execute(name, input),
  systemPrompt: AGENT_SYSTEM_PROMPT,
  open: (bytes: Uint8Array, name: string) => load(openWorkbook(bytes, name), name),
  save: (format: FileFormat = 'xlsx') => saveWorkbook(view.workbook, format),
};
(window as unknown as { sheets: typeof api }).sheets = api;

// When embedded in an iframe, the parent page can drive the spreadsheet via
// JSON-RPC (see connectAgent in @sheets/agent). Restarted per workbook so
// change notifications follow the loaded file.
function startBridge(): void {
  if (window.parent === window) return;
  stopBridge?.();
  stopBridge = serveAgent(agent, {
    postMessage: (m) => window.parent.postMessage(m, '*'),
    addEventListener: (type, l) => window.addEventListener(type, l),
    removeEventListener: (type, l) => window.removeEventListener(type, l),
  });
}
startBridge();

view.focus();
