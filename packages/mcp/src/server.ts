/**
 * MCP server that lets AI agents read and edit .xlsx / .csv files.
 *
 *   sheets-mcp [--root <dir>] [--readonly] [--no-autosave]
 *
 * Every spreadsheet tool takes a "file" argument. Workbooks are cached in
 * memory, reloaded when the file changes on disk and saved after each
 * modifying call (unless --no-autosave, then call save_workbook).
 */
import { readFile, stat, writeFile, mkdir } from 'node:fs/promises';
import { dirname, isAbsolute, relative, resolve } from 'node:path';
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { CallToolRequestSchema, ListToolsRequestSchema } from '@modelcontextprotocol/sdk/types.js';
import { formatFromFileName, openWorkbook, saveWorkbook, Workbook } from '@sheets/core';
import { AGENT_SYSTEM_PROMPT, resultToText, SpreadsheetAgent, TOOL_DEFINITIONS, type ToolDefinition } from '@sheets/agent';

const args = process.argv.slice(2);
const flag = (name: string) => args.includes(name);
const option = (name: string) => {
  const i = args.indexOf(name);
  return i >= 0 ? args[i + 1] : undefined;
};
const root = resolve(option('--root') ?? process.cwd());
const readOnly = flag('--readonly');
const autosave = !flag('--no-autosave');

interface Entry {
  wb: Workbook;
  mtimeMs: number;
  dirty: boolean;
  original?: Uint8Array;
}
const cache = new Map<string, Entry>();

function resolvePath(file: unknown): string {
  if (typeof file !== 'string' || !file) throw new Error('Parameter "file" is required (path to an .xlsx, .csv or .tsv file).');
  const full = isAbsolute(file) ? resolve(file) : resolve(root, file);
  const rel = relative(root, full);
  if (rel.startsWith('..') || isAbsolute(rel)) throw new Error(`Access outside of ${root} is not allowed.`);
  if (!/\.(xlsx|xlsm|csv|tsv|json)$/i.test(full)) throw new Error('Only .xlsx, .xlsm, .csv, .tsv and .json files are supported.');
  return full;
}

async function load(path: string, create = false): Promise<Entry> {
  let mtimeMs = -1;
  try {
    mtimeMs = (await stat(path)).mtimeMs;
  } catch {
    if (!create) throw new Error(`File not found: ${path}. Use create_workbook to create it.`);
  }
  const cached = cache.get(path);
  if (cached && (cached.dirty || cached.mtimeMs === mtimeMs)) return cached;
  const original = mtimeMs < 0 ? undefined : await readFile(path);
  const wb = original ? openWorkbook(original, path) : new Workbook();
  const entry = { wb, mtimeMs, dirty: false, original };
  cache.set(path, entry);
  return entry;
}

async function save(path: string, entry: Entry, target = path): Promise<void> {
  const bytes = saveWorkbook(entry.wb, formatFromFileName(target));
  if (target === path) {
    const current = await readFile(path).catch((error: NodeJS.ErrnoException) => { if (error.code === 'ENOENT') return undefined; throw error; });
    if (entry.original ? !current || current.length !== entry.original.length || current.some((value, index) => value !== entry.original![index]) : current !== undefined) {
      throw new Error('File changed externally; it was not overwritten. Save to a new file or reload.');
    }
  }
  await mkdir(dirname(target), { recursive: true });
  await writeFile(target, bytes);
  if (target === path) {
    entry.mtimeMs = (await stat(path)).mtimeMs;
    entry.dirty = false;
    entry.original = bytes;
  }
}

const FILE_PARAM = { type: 'string', description: `Path to the spreadsheet (.xlsx, .csv, .tsv), relative to ${root} or absolute.` };

function withFile(def: ToolDefinition): ToolDefinition {
  return {
    ...def,
    inputSchema: {
      ...def.inputSchema,
      properties: { file: FILE_PARAM, ...(def.inputSchema.properties ?? {}) },
      required: ['file', ...(def.inputSchema.required ?? [])],
    },
  };
}

const EXTRA_TOOLS: ToolDefinition[] = [
  {
    name: 'create_workbook',
    description: 'Creates a new empty spreadsheet file (.xlsx or .csv). Fails if the file already exists unless overwrite is true.',
    inputSchema: { type: 'object', properties: { file: FILE_PARAM, sheet_name: { type: 'string' }, overwrite: { type: 'boolean' } }, required: ['file'] },
  },
  {
    name: 'save_workbook',
    description: 'Saves the workbook, optionally to another path/format (e.g. convert CSV to XLSX by passing save_as "data.xlsx").',
    inputSchema: { type: 'object', properties: { file: FILE_PARAM, save_as: { type: 'string', description: 'Optional target path' } }, required: ['file'] },
  },
  {
    name: 'get_agent_guide',
    description: 'Returns usage guidance for operating spreadsheets with these tools.',
    readOnly: true,
    inputSchema: { type: 'object', properties: {} },
  },
];

const spreadsheetTools = TOOL_DEFINITIONS.filter((t) => !['select_range', 'activate_sheet'].includes(t.name) && (!readOnly || t.readOnly)).map(withFile);
const allTools = [...spreadsheetTools, ...EXTRA_TOOLS.filter((t) => !readOnly || t.readOnly)];

const server = new Server({ name: 'sheets', version: '0.1.0' }, { capabilities: { tools: {} }, instructions: AGENT_SYSTEM_PROMPT });

server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: allTools.map((t) => ({ name: t.name, description: t.description, inputSchema: t.inputSchema, annotations: { readOnlyHint: !!t.readOnly } })),
}));

server.setRequestHandler(CallToolRequestSchema, async (req) => {
  const name = req.params.name;
  const input = { ...(req.params.arguments ?? {}) } as Record<string, unknown>;
  const text = (t: string, isError = false) => ({ content: [{ type: 'text' as const, text: t }], isError });
  try {
    if (!allTools.some((t) => t.name === name)) return text(`Unknown tool "${name}"`, true);
    if (name === 'get_agent_guide') return text(AGENT_SYSTEM_PROMPT);
    const path = resolvePath(input.file);
    delete input.file;

    if (name === 'create_workbook') {
      const exists = await stat(path).then(() => true, () => false);
      if (exists && !input.overwrite) return text(`File already exists: ${path}`, true);
      const entry: Entry = { wb: new Workbook({ sheetName: typeof input.sheet_name === 'string' ? input.sheet_name : 'Sheet1' }), mtimeMs: -1, dirty: true, original: exists ? await readFile(path) : undefined };
      cache.set(path, entry);
      await save(path, entry);
      return text(`Created ${path}`);
    }
    if (name === 'save_workbook') {
      const entry = await load(path);
      const target = input.save_as ? resolvePath(input.save_as) : path;
      await save(path, entry, target);
      return text(`Saved ${target}`);
    }

    const def = spreadsheetTools.find((t) => t.name === name)!;
    const entry = await load(path);
    const result = new SpreadsheetAgent(entry.wb).execute(name, input);
    if (result.ok && !def.readOnly) {
      entry.dirty = true;
      if (autosave) await save(path, entry);
    }
    return text(resultToText(result), !result.ok);
  } catch (e) {
    return text(`Error: ${(e as Error).message}`, true);
  }
});

await server.connect(new StdioServerTransport());
console.error(`sheets-mcp ready (root: ${root}${readOnly ? ', read-only' : ''}${autosave ? '' : ', no autosave'})`);
