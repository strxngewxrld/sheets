import * as vscode from 'vscode';
import { openWorkbook, saveWorkbook, formatFromFileName, Workbook, writeXlsx } from '@sheets/core';
import { resultToText, SpreadsheetAgent, TOOL_DEFINITIONS, type ToolResult } from '@sheets/agent';
import { SpreadsheetEditorProvider } from './provider.js';
import { TOOL_PREFIX } from './tool-manifest.js';

export interface SheetsApi {
  /** Tool definitions (name, description, JSON schema). */
  tools: typeof TOOL_DEFINITIONS;
  /**
   * Executes a tool on a spreadsheet. If the file is open in a Sheets editor
   * the change is applied live (undoable, marks the editor dirty); otherwise
   * the file is loaded, modified and saved.
   */
  executeTool(name: string, input: Record<string, unknown>, file?: string | vscode.Uri): Promise<ToolResult>;
}

let provider: SpreadsheetEditorProvider;

export function activate(context: vscode.ExtensionContext): SheetsApi {
  provider = SpreadsheetEditorProvider.register(context);

  context.subscriptions.push(
    vscode.commands.registerCommand('sheets.newWorkbook', async () => {
      const target = await vscode.window.showSaveDialog({
        filters: { 'Excel Workbook': ['xlsx'] },
        defaultUri: vscode.workspace.workspaceFolders?.[0] ? vscode.Uri.joinPath(vscode.workspace.workspaceFolders[0].uri, 'Book1.xlsx') : undefined,
      });
      if (!target) return;
      await vscode.workspace.fs.writeFile(target, writeXlsx(new Workbook()));
      await vscode.commands.executeCommand('vscode.openWith', target, 'sheets.spreadsheet');
    }),
    vscode.commands.registerCommand('sheets.openWithSheets', async (uri?: vscode.Uri) => {
      const target = uri ?? vscode.window.activeTextEditor?.document.uri;
      if (!target) return;
      const csv = /\.(csv|tsv)$/i.test(target.path);
      await vscode.commands.executeCommand('vscode.openWith', target, csv ? 'sheets.csv' : 'sheets.spreadsheet');
    }),
    // Command interface for other extensions / automation:
    //   vscode.commands.executeCommand('sheets.executeTool', { name, input, file })
    vscode.commands.registerCommand('sheets.executeTool', (args: { name: string; input?: Record<string, unknown>; file?: string }) =>
      executeTool(args.name, args.input ?? {}, args.file),
    ),
  );

  // Language model tools (usable by Copilot agent mode and other chat participants).
  if (vscode.lm?.registerTool) {
    for (const def of TOOL_DEFINITIONS) {
      context.subscriptions.push(
        vscode.lm.registerTool<Record<string, unknown>>(`${TOOL_PREFIX}${def.name}`, {
          prepareInvocation: (options) => {
            const file = typeof options.input.file === 'string' ? options.input.file : undefined;
            const target = file ?? provider.activeDocument?.uri.path.split('/').pop() ?? 'active spreadsheet';
            return {
              invocationMessage: `Sheets: ${def.name.replace(/_/g, ' ')} (${target})`,
              ...(def.readOnly
                ? {}
                : {
                    confirmationMessages: {
                      title: `Sheets: ${def.name.replace(/_/g, ' ')}`,
                      message: new vscode.MarkdownString(`Modify **${target}**?\n\n\`\`\`json\n${JSON.stringify(options.input, null, 2).slice(0, 2000)}\n\`\`\``),
                    },
                  }),
            };
          },
          invoke: async (options) => {
            const { file, ...input } = options.input;
            const result = await executeTool(def.name, input, typeof file === 'string' ? file : undefined);
            return new vscode.LanguageModelToolResult([new vscode.LanguageModelTextPart(resultToText(result))]);
          },
        }),
      );
    }
  }

  return { tools: TOOL_DEFINITIONS, executeTool };
}

export function deactivate(): void {}

function resolveUri(file: string | vscode.Uri | undefined): vscode.Uri | undefined {
  if (!file) return provider.activeDocument?.uri;
  if (typeof file !== 'string') return file;
  if (/^[a-z][a-z0-9+.-]*:\/\//i.test(file)) return vscode.Uri.parse(file);
  if (file.startsWith('/') || /^[a-zA-Z]:[\\/]/.test(file)) return vscode.Uri.file(file);
  const folder = vscode.workspace.workspaceFolders?.[0];
  return folder ? vscode.Uri.joinPath(folder.uri, file) : vscode.Uri.file(file);
}

async function executeTool(name: string, input: Record<string, unknown>, file?: string | vscode.Uri): Promise<ToolResult> {
  const uri = resolveUri(file);
  if (!uri) return { ok: false, error: 'No spreadsheet is open. Pass "file" with the path of an .xlsx or .csv file.' };
  const def = TOOL_DEFINITIONS.find((t) => t.name === name);
  if (!def) return { ok: false, error: `Unknown tool "${name}"` };

  // Live: the document is open in an editor.
  const doc = provider.documents.get(uri.toString());
  if (doc?.panel) {
    try {
      return await doc.request<ToolResult>('executeTool', { name, input });
    } catch (e) {
      return { ok: false, error: (e as Error).message };
    }
  }

  // Headless: load, run, save.
  if (!def.readOnly && !vscode.workspace.getConfiguration('sheets').get<boolean>('agentTools.headlessWrites', true)) {
    return { ok: false, error: 'Modifying files that are not open is disabled (setting sheets.agentTools.headlessWrites). Open the file in the Sheets editor first.' };
  }
  let wb: Workbook;
  try {
    wb = openWorkbook(await vscode.workspace.fs.readFile(uri), uri.path);
  } catch (e) {
    if (name === 'write_range' && (e as { code?: string }).code === 'FileNotFound') wb = new Workbook();
    else return { ok: false, error: `Cannot open ${uri.fsPath}: ${(e as Error).message}` };
  }
  const result = new SpreadsheetAgent(wb).execute(name, input);
  if (result.ok && !def.readOnly) {
    await vscode.workspace.fs.writeFile(uri, saveWorkbook(wb, formatFromFileName(uri.path)));
  }
  return result;
}
