import type { ChangeEvent, Workbook } from '@sheets/core';
import { TOOLS, ToolInputError, type AgentHost, type ToolContext, type ToolDefinition } from './tools.js';

export interface ToolResult {
  ok: boolean;
  /** Tool-specific JSON result. */
  result?: unknown;
  /** Error message suitable to show to a language model. */
  error?: string;
}

export interface SpreadsheetAgentOptions {
  /** UI integration for selection/navigation tools. */
  host?: AgentHost;
  /** Maximum number of cells returned by read_range (default 2000). */
  maxCells?: number;
  /** Origin tag recorded on workbook change events (default "agent"). */
  origin?: string;
  /** Restrict the tool set, e.g. to read-only tools. */
  filter?: (tool: ToolDefinition) => boolean;
}

/**
 * Programmatic control surface for a workbook. Exposes a set of tools with
 * JSON-Schema inputs that can be handed to an LLM, plus `execute()` to run
 * a tool call. Every mutating call becomes one undo step tagged with the
 * agent origin, so users can see and revert what an agent did.
 */
export class SpreadsheetAgent {
  private readonly getWorkbook: () => Workbook;
  private readonly options: SpreadsheetAgentOptions;

  constructor(workbook: Workbook | (() => Workbook), options: SpreadsheetAgentOptions = {}) {
    this.getWorkbook = typeof workbook === 'function' ? workbook : () => workbook;
    this.options = options;
  }

  get workbook(): Workbook {
    return this.getWorkbook();
  }

  get tools(): ToolDefinition[] {
    return TOOLS.filter((t) => !this.options.filter || this.options.filter(t)).map(({ name, description, inputSchema, readOnly }) =>
      readOnly ? { name, description, inputSchema, readOnly } : { name, description, inputSchema },
    );
  }

  hasTool(name: string): boolean {
    return this.tools.some((t) => t.name === name);
  }

  /** Executes a tool call. Never throws; errors are returned as `{ ok: false, error }`. */
  execute(name: string, input: unknown = {}): ToolResult {
    const tool = TOOLS.find((t) => t.name === name);
    if (!tool || (this.options.filter && !this.options.filter(tool))) {
      return { ok: false, error: `Unknown tool "${name}". Available: ${this.tools.map((t) => t.name).join(', ')}` };
    }
    const args = (input && typeof input === 'object' ? input : {}) as Record<string, unknown>;
    const missing = (tool.inputSchema.required ?? []).filter((k) => args[k] === undefined);
    if (missing.length) return { ok: false, error: `Missing required parameter(s): ${missing.join(', ')}` };
    const ctx: ToolContext = { wb: this.workbook, host: this.options.host, maxCells: this.options.maxCells ?? 2000 };
    try {
      const result = tool.readOnly
        ? tool.run(ctx, args)
        : ctx.wb.transaction(() => tool.run(ctx, args), { origin: this.options.origin ?? 'agent', label: name });
      return { ok: true, result };
    } catch (e) {
      const message = e instanceof Error ? e.message : String(e);
      return { ok: false, error: e instanceof ToolInputError ? message : `${name} failed: ${message}` };
    }
  }

  /** Runs several tool calls in order as a single undo step. Stops at the first error. */
  executeBatch(calls: { name: string; input?: unknown }[]): ToolResult[] {
    const results: ToolResult[] = [];
    this.workbook.transaction(
      () => {
        for (const c of calls) {
          const r = this.execute(c.name, c.input);
          results.push(r);
          if (!r.ok) break;
        }
      },
      { origin: this.options.origin ?? 'agent', label: 'batch' },
    );
    return results;
  }

  /** Subscribes to workbook changes (e.g. to tell an agent what the user edited). */
  onChange(listener: (e: ChangeEvent) => void): () => void {
    return this.workbook.events.on('change', listener);
  }
}
