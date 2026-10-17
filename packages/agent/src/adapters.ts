import type { ToolResult } from './agent.js';
import type { ToolDefinition } from './tools.js';

/** Tool format of the Anthropic Messages API (`tools` parameter). */
export function toAnthropicTools(tools: ToolDefinition[]): { name: string; description: string; input_schema: object }[] {
  return tools.map((t) => ({ name: t.name, description: t.description, input_schema: t.inputSchema }));
}

/** Tool format of the OpenAI Chat Completions / Responses API. */
export function toOpenAITools(tools: ToolDefinition[]): { type: 'function'; function: { name: string; description: string; parameters: object } }[] {
  return tools.map((t) => ({ type: 'function', function: { name: t.name, description: t.description, parameters: t.inputSchema } }));
}

/** Tool format of the Model Context Protocol (`tools/list`). */
export function toMcpTools(tools: ToolDefinition[]): { name: string; description: string; inputSchema: object; annotations?: object }[] {
  return tools.map((t) => ({
    name: t.name,
    description: t.description,
    inputSchema: t.inputSchema,
    annotations: { readOnlyHint: !!t.readOnly },
  }));
}

/** Converts a ToolResult into text content for tool_result / MCP responses. */
export function resultToText(r: ToolResult): string {
  return r.ok ? JSON.stringify(r.result ?? null, null, 2) : `Error: ${r.error}`;
}

/** Recommended system prompt snippet for agents operating the spreadsheet. */
export const AGENT_SYSTEM_PROMPT = `You operate a spreadsheet application through tools.
- Start with get_workbook_info, then read_range to look at data before changing it.
- Cells use A1 notation; rows are 1-based, columns are letters.
- Prefer formulas (e.g. "=SUM(B2:B10)") over computed constants so results stay live.
- Use fill_formula to apply one formula to a whole column.
- Every tool call is one undo step; the user can undo your changes.
- After writing, verify results with read_range or evaluate_formula.`;
