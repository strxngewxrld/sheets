import * as vscode from 'vscode';
import { resultToText, TOOL_DEFINITIONS, type ToolResult } from '@sheets/agent';
import { TOOL_PREFIX } from './tool-manifest.js';

type ExecuteTool = (name: string, input: Record<string, unknown>, file?: string) => Promise<ToolResult>;

/** VS Code's tool registry is the integration point used by Copilot agent mode. */
export function registerCopilotTools(context: vscode.ExtensionContext, execute: ExecuteTool, activeFile: () => string | undefined): void {
  if (!vscode.lm?.registerTool) return;
  for (const def of TOOL_DEFINITIONS) {
    context.subscriptions.push(vscode.lm.registerTool<Record<string, unknown>>(`${TOOL_PREFIX}${def.name}`, {
      prepareInvocation: (options) => {
        const file = typeof options.input.file === 'string' ? options.input.file : undefined;
        const target = file ?? activeFile() ?? 'active spreadsheet';
        return {
          invocationMessage: `Sheets: ${def.name.replace(/_/g, ' ')} (${target})`,
          ...(def.readOnly ? {} : {
            confirmationMessages: {
              title: `Sheets: ${def.name.replace(/_/g, ' ')}`,
              message: new vscode.MarkdownString()
                .appendText(`Modify ${target}?`)
                .appendCodeblock(JSON.stringify(options.input, null, 2).slice(0, 2000), 'json'),
            },
          }),
        };
      },
      invoke: async (options, token) => {
        // A cancelled invocation must not begin a mutation or a headless save.
        if (token.isCancellationRequested) throw new vscode.CancellationError();
        const { file, ...input } = options.input;
        const result = await execute(def.name, input, typeof file === 'string' ? file : undefined);
        if (!result.ok) throw new Error(result.error ?? `Sheets tool ${def.name} failed`);
        return new vscode.LanguageModelToolResult([new vscode.LanguageModelTextPart(resultToText(result))]);
      },
    }));
  }
}
