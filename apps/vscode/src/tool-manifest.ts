import { TOOL_DEFINITIONS, type JsonSchema, type ToolDefinition } from '@sheets/agent';

export const TOOL_PREFIX = 'sheets_';

/** Adds the optional "file" parameter that selects which workbook a tool acts on. */
export function withFileParam(schema: ToolDefinition['inputSchema']): ToolDefinition['inputSchema'] {
  const file: JsonSchema = {
    type: 'string',
    description: 'Path (absolute or workspace-relative) or URI of the .xlsx/.csv file. Defaults to the spreadsheet open in the active Sheets editor.',
  };
  return { ...schema, properties: { file, ...(schema.properties ?? {}) } };
}

/** Contribution entries for package.json "contributes.languageModelTools". */
export const languageModelTools = TOOL_DEFINITIONS.map((t) => ({
  name: `${TOOL_PREFIX}${t.name}`,
  displayName: `Sheets: ${t.name.replace(/_/g, ' ')}`,
  toolReferenceName: `sheets_${t.name}`,
  canBeReferencedInPrompt: true,
  userDescription: t.description.split('. ')[0],
  modelDescription: t.description,
  tags: ['spreadsheet', 'excel', 'csv', 'sheets'],
  inputSchema: withFileParam(t.inputSchema),
}));
