import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { platform, tmpdir } from 'node:os';
import { basename, join } from 'node:path';
import * as vscode from 'vscode';
import { openWorkbook, saveWorkbook, formatFromFileName, Workbook, writeXlsx } from '@sheets/core';
import { SpreadsheetAgent, TOOL_DEFINITIONS, type ToolResult } from '@sheets/agent';
import { SpreadsheetEditorProvider } from './provider.js';
import { registerCopilotTools } from './copilot.js';
import { decodeBytes } from './binary.js';
import type { CsvExport } from './protocol.js';
import { checkForUpdates, downloadVsix } from './update-check.js';
import { escapesFolders, hasProviderSymlink, headlessAccessError, resolveFileUri } from './file-uri.js';
import { reopenStartupSpreadsheets } from './startup-editors.js';

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
    vscode.commands.registerCommand('sheets.gettingStarted', () =>
      vscode.commands.executeCommand('workbench.action.openWalkthrough', `${context.extension.id}#sheets.gettingStarted`, false),
    ),
    vscode.commands.registerCommand('sheets.openManual', () => {
      const file = vscode.env.language.toLowerCase().startsWith('de') ? 'manual.de.md' : 'manual.en.md';
      return vscode.commands.executeCommand('markdown.showPreview', vscode.Uri.joinPath(context.extensionUri, 'docs', file));
    }),
    vscode.commands.registerCommand('sheets.openFile', async () => {
      const selected = await vscode.window.showOpenDialog({
        canSelectMany: false,
        filters: { Spreadsheets: ['xlsx', 'xlsm', 'csv', 'tsv'] },
      });
      const target = selected?.[0];
      if (!target) return;
      const csv = /\.(csv|tsv)$/i.test(target.path);
      await vscode.commands.executeCommand('vscode.openWith', target, csv ? 'sheets.csv' : 'sheets.spreadsheet');
    }),
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
    ...(['find', 'replace', 'hideRows', 'unhideRows', 'hideCols', 'unhideCols', 'goTo', 'zoomIn', 'zoomOut', 'zoomReset'] as const).map((command) =>
      vscode.commands.registerCommand(`sheets.${command}`, () => provider.activeDocument?.post({ type: 'command', command })),
    ),
    vscode.commands.registerCommand('sheets.exportCsv', () => exportCsv()),
    // Command interface for other extensions / automation:
    //   vscode.commands.executeCommand('sheets.executeTool', { name, input, file })
    vscode.commands.registerCommand('sheets.executeTool', (args: { name: string; input?: Record<string, unknown>; file?: string }) =>
      executeTool(args.name, args.input ?? {}, args.file),
    ),
  );

  registerCopilotTools(context, executeTool, () => provider.activeDocument?.uri.path.split('/').pop());
  void reopenStartupSpreadsheets().catch(error => console.error('Sheets: reopening a startup spreadsheet failed', error));
  void startUpdateCheck(context).catch(() => undefined);

  return { tools: TOOL_DEFINITIONS, executeTool };
}

export function deactivate(): void {}

const SKIPPED_VERSION_KEY = 'sheets.skippedUpdateVersion';

/** Shows a notification when GitHub Releases has a newer version than the installed one. */
async function startUpdateCheck(context: vscode.ExtensionContext): Promise<void> {
  if (context.extensionMode === vscode.ExtensionMode.Test) return;
  if (!vscode.workspace.getConfiguration('sheets').get<boolean>('checkForUpdates', true)) return;
  const manifest = context.extension.packageJSON as { version: string; repository?: { url?: string } };
  const de = vscode.env.language.toLowerCase().startsWith('de');
  await checkForUpdates({
    current: manifest.version,
    repositoryUrl: manifest.repository?.url ?? '',
    fetch,
    skipped: context.globalState.get<string>(SKIPPED_VERSION_KEY),
    async notify(release) {
      const install = de ? 'Installieren' : 'Install';
      const open = de ? 'Release öffnen' : 'Open Release';
      const skip = de ? 'Diese Version überspringen' : 'Skip This Version';
      const message = de
        ? `Sheets ${release.version} ist verfügbar (installiert: ${manifest.version}).`
        : `Sheets ${release.version} is available (installed: ${manifest.version}).`;
      const choice = await vscode.window.showInformationMessage(message, ...(release.vsix ? [install] : []), open, skip);
      return choice === install ? 'install' : choice === open ? 'open' : choice === skip ? 'skip' : undefined;
    },
    async install(release) {
      try {
        await vscode.window.withProgress({
          location: vscode.ProgressLocation.Notification,
          title: de ? `Sheets ${release.version} wird installiert…` : `Installing Sheets ${release.version}…`,
        }, async () => {
          const bytes = await downloadVsix(release.vsix, fetch);
          // installFromVSIX only accepts file: URIs; globalStorageUri may use vscode-userdata:.
          const dir = await mkdtemp(join(tmpdir(), 'sheets-update-'));
          try {
            const target = join(dir, basename(release.vsix.name));
            await writeFile(target, bytes);
            await vscode.commands.executeCommand('workbench.extensions.command.installFromVSIX', [vscode.Uri.file(target)]);
          } finally {
            await rm(dir, { recursive: true, force: true }).catch(() => undefined);
          }
        });
      } catch (error) {
        const open = de ? 'Release öffnen' : 'Open Release';
        const reason = error instanceof Error ? error.message : String(error);
        const choice = await vscode.window.showErrorMessage(
          de ? `Sheets ${release.version} konnte nicht installiert werden: ${reason}` : `Could not install Sheets ${release.version}: ${reason}`, open);
        if (choice === open) void vscode.env.openExternal(vscode.Uri.parse(release.url, true));
      }
      // VS Code itself offers to restart extensions after installing the VSIX.
    },
    open: (url) => void vscode.env.openExternal(vscode.Uri.parse(url, true)),
    skip: (version) => void context.globalState.update(SKIPPED_VERSION_KEY, version),
  });
}

/** Exports the active sheet of the active Sheets editor as CSV or TSV. */
async function exportCsv(): Promise<void> {
  const doc = provider.activeDocument;
  const de = vscode.env.language.toLowerCase().startsWith('de');
  if (!doc?.panel) {
    void vscode.window.showInformationMessage(de ? 'Bitte zuerst eine Tabelle im Sheets-Editor öffnen.' : 'Open a spreadsheet in the Sheets editor first.');
    return;
  }
  const choices = [
    { label: de ? 'Semikolon (;)' : 'Semicolon (;)', delimiter: ';', ext: 'csv' },
    { label: de ? 'Komma (,)' : 'Comma (,)', delimiter: ',', ext: 'csv' },
    { label: de ? 'Tabulator' : 'Tab', delimiter: '\t', ext: 'tsv' },
  ];
  if (!de) choices.unshift(choices.splice(1, 1)[0]!);
  const choice = await vscode.window.showQuickPick(choices, { placeHolder: de ? 'Trennzeichen für den Export' : 'Delimiter for the export' });
  if (!choice) return;
  const mode = await vscode.window.showQuickPick([
    {
      label: de ? 'Für Excel (Textschutz)' : 'For Excel (text protection)',
      description: de ? 'Stellt Formeltext einen Tabulator voran; verändert exportierte Textwerte.' : 'Prefixes formula-like text with a tab; changes exported text values.',
      escapeFormulas: true,
    },
    {
      label: de ? 'Unveränderte Werte' : 'Unchanged values',
      description: de ? 'Für Datenimporte. Fremder Text kann in Excel als Formel ausgeführt werden.' : 'For data imports. Untrusted text may execute as a formula in Excel.',
      escapeFormulas: false,
    },
  ], { placeHolder: de ? 'Verwendung der exportierten Datei' : 'How the exported file will be used' });
  if (!mode) return;
  let result: CsvExport;
  try {
    result = await doc.request<CsvExport>('exportCsv', { delimiter: choice.delimiter, escapeFormulas: mode.escapeFormulas });
  } catch (error) {
    const reason = error instanceof Error ? error.message : String(error);
    void vscode.window.showErrorMessage(de ? `CSV-Export fehlgeschlagen: ${reason}` : `CSV export failed: ${reason}`);
    return;
  }
  const base = doc.uri.path.split('/').pop()!.replace(/\.[^.]+$/, '');
  const name = `${base}-${result.sheet}.${choice.ext}`.replace(/[\\/:*?"<>|]/g, '_');
  const target = await vscode.window.showSaveDialog({
    defaultUri: vscode.Uri.joinPath(doc.uri, '..', name),
    filters: choice.ext === 'tsv' ? { TSV: ['tsv'] } : { CSV: ['csv'] },
  });
  if (!target) return;
  await vscode.workspace.fs.writeFile(target, decodeBytes(result.bytes));
}

function resolveUri(file: string | vscode.Uri | undefined): vscode.Uri | undefined {
  if (!file) return provider.activeDocument?.uri;
  if (typeof file !== 'string') return file;
  const folder = vscode.workspace.workspaceFolders?.[0];
  return resolveFileUri(file, folder?.uri);
}

async function headlessAccessDenied(uri: vscode.Uri): Promise<string | undefined> {
  const folders = vscode.workspace.workspaceFolders?.map((folder) => folder.uri) ?? [];
  const error = headlessAccessError(uri, folders, uri.scheme === 'file' && (platform() === 'win32' || platform() === 'darwin'));
  if (error) return error;
  const outside = 'Spreadsheets reached through a symbolic link that leads outside the workspace cannot be used without opening them.';
  if (uri.scheme === 'file') {
    const roots = folders.filter((folder) => folder.scheme === 'file').map((folder) => folder.fsPath);
    return (await escapesFolders(uri.fsPath, roots)) ? outside : undefined;
  }
  // Other file systems cannot resolve links; refuse a link instead.
  return await hasProviderSymlink(uri, folders, (part) => vscode.workspace.fs.stat(part)) ? outside : undefined;
}

async function executeTool(name: string, input: Record<string, unknown>, file?: string | vscode.Uri): Promise<ToolResult> {
  const uri = resolveUri(file);
  if (!uri) return { ok: false, error: 'No spreadsheet is open. Pass "file" with the path of an .xlsx or .csv file.' };
  const def = TOOL_DEFINITIONS.find((t) => t.name === name);
  if (!def) return { ok: false, error: `Unknown tool "${name}"` };
  if (!def.readOnly && vscode.workspace.fs.isWritableFileSystem(uri.scheme) === false) {
    return { ok: false, error: 'This spreadsheet file system is read-only.' };
  }

  // Live: the document is open in an editor.
  const doc = provider.documents.get(uri.toString());
  if (doc?.panel) {
    try {
      return await doc.request<ToolResult>('executeTool', { name, input });
    } catch (e) {
      return { ok: false, error: (e as Error).message };
    }
  }

  // Headless: load, run, save. Agents reach only spreadsheets in the workspace.
  let denied: string | undefined;
  try { denied = await headlessAccessDenied(uri); }
  catch { return { ok: false, error: 'Cannot verify safe access to this spreadsheet.' }; }
  if (denied) return { ok: false, error: denied };
  if (!def.readOnly && !vscode.workspace.isTrusted) {
    return { ok: false, error: 'Trust this workspace before allowing agents to modify unopened spreadsheet files.' };
  }
  if (!def.readOnly && !vscode.workspace.getConfiguration('sheets').get<boolean>('agentTools.headlessWrites', true)) {
    return { ok: false, error: 'Modifying files that are not open is disabled (setting sheets.agentTools.headlessWrites). Open the file in the Sheets editor first.' };
  }
  let wb: Workbook;
  let original: Uint8Array | undefined;
  try {
    original = await vscode.workspace.fs.readFile(uri);
    wb = openWorkbook(original, uri.path);
  } catch (e) {
    if (name === 'write_range' && (e as { code?: string }).code === 'FileNotFound') wb = new Workbook();
    else return { ok: false, error: `Cannot open ${uri.fsPath}: ${(e as Error).message}` };
  }
  const result = new SpreadsheetAgent(wb).execute(name, input);
  if (result.ok && !def.readOnly) {
    try {
      const bytes = saveWorkbook(wb, formatFromFileName(uri.path));
      const current = await Promise.resolve(vscode.workspace.fs.readFile(uri)).catch((error: { code?: string }) => {
        if (error.code === 'FileNotFound') return undefined;
        throw error;
      });
      if (original ? !current || current.length !== original.length || current.some((value, index) => value !== original![index]) : current !== undefined) {
        return { ok: false, error: 'Sheets: Die Datei wurde während der Bearbeitung extern verändert und nicht überschrieben.' };
      }
      await vscode.workspace.fs.writeFile(uri, bytes);
    } catch (error) {
      return { ok: false, error: error instanceof Error ? error.message : String(error) };
    }
  }
  return result;
}
