import * as vscode from 'vscode';
import { decodeBytes, encodeBytes } from './binary.js';
import type { HostToWebview, WebviewToHost } from './protocol.js';
import { fill, shellStrings } from './webview/strings.js';

/** A spreadsheet file opened in the custom editor. The live model lives in the webview. */
export class SpreadsheetDocument implements vscode.CustomDocument {
  readonly uri: vscode.Uri;
  bytes: Uint8Array;
  /** Last observed on-disk version; null means an older backup cannot prove it. */
  diskBytes: Uint8Array | null | undefined;
  panel: vscode.WebviewPanel | null = null;
  /** Set while the webview could not open the file; such a document must never be saved. */
  loadError: string | undefined;
  private nextId = 1;
  private pending = new Map<number, { resolve: (v: unknown) => void; reject: (e: Error) => void; timer: ReturnType<typeof setTimeout> }>();
  private initialized = false;
  private queued = new Map<number, HostToWebview>();
  private readonly onDisposeEmitter = new vscode.EventEmitter<void>();
  readonly onDidDispose = this.onDisposeEmitter.event;

  constructor(uri: vscode.Uri, bytes: Uint8Array, public backup = false) {
    this.uri = uri;
    this.bytes = bytes;
    this.diskBytes = backup ? null : bytes;
  }

  post(message: HostToWebview): void {
    void this.panel?.webview.postMessage(message);
  }

  /** Sends a request to the webview and waits for the response. */
  request<T>(method: 'getBackupData'): Promise<T>;
  request<T>(method: 'getFileData', params?: { fileName: string }): Promise<T>;
  request<T>(method: 'executeTool', params: { name: string; input: unknown }): Promise<T>;
  request<T>(method: 'exportCsv', params: { delimiter: string }): Promise<T>;
  request<T>(method: 'getFileData' | 'getBackupData' | 'executeTool' | 'exportCsv', params?: { name: string; input: unknown } | { delimiter: string } | { fileName: string }): Promise<T> {
    if (!this.panel) return Promise.reject(new Error('Spreadsheet editor is not open'));
    const id = this.nextId++;
    return new Promise<T>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.queued.delete(id);
        if (this.pending.delete(id)) reject(new Error(`Webview request "${method}" timed out`));
      }, 60_000);
      this.pending.set(id, { resolve: resolve as (v: unknown) => void, reject, timer });
      const message: HostToWebview = method === 'executeTool' ? { type: 'request', id, method, params: params as { name: string; input: unknown } }
        : method === 'exportCsv' ? { type: 'request', id, method, params: params as { delimiter: string } }
          : method === 'getFileData' ? { type: 'request', id, method, ...(params ? { params: params as { fileName: string } } : {}) }
            : { type: 'request', id, method };
      if (this.initialized) this.post(message);
      else this.queued.set(id, message);
    });
  }

  markInitialized(): void {
    this.initialized = true;
    for (const message of this.queued.values()) this.post(message);
    this.queued.clear();
  }

  reset(error = 'Spreadsheet editor reloading'): void {
    this.initialized = false;
    for (const p of this.pending.values()) { clearTimeout(p.timer); p.reject(new Error(error)); }
    this.pending.clear();
    this.queued.clear();
  }

  handleResponse(id: number, result: unknown, error?: string): void {
    const p = this.pending.get(id);
    if (!p) return;
    this.pending.delete(id);
    this.queued.delete(id);
    clearTimeout(p.timer);
    if (error) p.reject(new Error(error));
    else p.resolve(result);
  }

  dispose(): void {
    this.reset('Document closed');
    this.onDisposeEmitter.fire();
    this.onDisposeEmitter.dispose();
  }
}

/**
 * Custom editor for .xlsx/.csv files. Undo/redo, dirty state, save, revert
 * and hot exit are integrated with VS Code; edits happen in the webview.
 */
export class SpreadsheetEditorProvider implements vscode.CustomEditorProvider<SpreadsheetDocument> {
  static readonly viewTypes = ['sheets.spreadsheet', 'sheets.csv'];
  readonly documents = new Map<string, SpreadsheetDocument>();
  activeDocument: SpreadsheetDocument | undefined;

  private readonly changeEmitter = new vscode.EventEmitter<vscode.CustomDocumentEditEvent<SpreadsheetDocument>>();
  readonly onDidChangeCustomDocument = this.changeEmitter.event;

  constructor(private readonly context: vscode.ExtensionContext) {}

  static register(context: vscode.ExtensionContext): SpreadsheetEditorProvider {
    const provider = new SpreadsheetEditorProvider(context);
    for (const viewType of SpreadsheetEditorProvider.viewTypes) {
      // VS Code subscribes separately for each view type. Broadcasting every edit
      // to both registrations makes the other one report "No custom document found".
      const documents = new WeakSet<SpreadsheetDocument>();
      const onDidChangeCustomDocument: vscode.Event<vscode.CustomDocumentEditEvent<SpreadsheetDocument>> =
        (listener, thisArgs, disposables) => provider.onDidChangeCustomDocument((event) => {
          if (documents.has(event.document)) listener.call(thisArgs, event);
        }, undefined, disposables);
      const editor: vscode.CustomEditorProvider<SpreadsheetDocument> = {
        onDidChangeCustomDocument,
        openCustomDocument: async (uri, openContext) => {
          const document = await provider.openCustomDocument(uri, openContext);
          documents.add(document);
          return document;
        },
        resolveCustomEditor: provider.resolveCustomEditor.bind(provider),
        saveCustomDocument: provider.saveCustomDocument.bind(provider),
        saveCustomDocumentAs: provider.saveCustomDocumentAs.bind(provider),
        revertCustomDocument: provider.revertCustomDocument.bind(provider),
        backupCustomDocument: provider.backupCustomDocument.bind(provider),
      };
      context.subscriptions.push(
        vscode.window.registerCustomEditorProvider(viewType, editor, {
          webviewOptions: { retainContextWhenHidden: true },
          supportsMultipleEditorsPerDocument: false,
        }),
      );
    }
    return provider;
  }

  async openCustomDocument(uri: vscode.Uri, openContext: vscode.CustomDocumentOpenContext): Promise<SpreadsheetDocument> {
    const source = openContext.backupId ? vscode.Uri.parse(openContext.backupId) : uri;
    const bytes = openContext.untitledDocumentData ?? (await vscode.workspace.fs.readFile(source));
    const doc = new SpreadsheetDocument(uri, bytes, !!openContext.backupId);
    if (openContext.untitledDocumentData) doc.diskBytes = undefined;
    if (openContext.backupId) {
      const text = new TextDecoder().decode(bytes);
      if (text.startsWith('{"type":"sheets-backup-v1",')) {
        const snapshot = JSON.parse(text);
        if (typeof snapshot.diskBytes === 'string') doc.diskBytes = decodeBytes(snapshot.diskBytes);
      }
    }
    this.documents.set(uri.toString(), doc);
    doc.onDidDispose(() => {
      if (this.documents.get(uri.toString()) === doc) this.documents.delete(uri.toString());
      if (this.activeDocument === doc) this.activeDocument = undefined;
    });
    return doc;
  }

  async resolveCustomEditor(document: SpreadsheetDocument, panel: vscode.WebviewPanel): Promise<void> {
    document.reset();
    document.panel = panel;
    this.activeDocument = document;
    panel.webview.options = { enableScripts: true, localResourceRoots: [vscode.Uri.joinPath(this.context.extensionUri, 'dist')] };
    panel.webview.html = this.html(panel.webview, document.uri.path.split('/').pop() ?? '');
    panel.onDidChangeViewState(() => {
      if (panel.active) this.activeDocument = document;
    });
    panel.onDidDispose(() => {
      if (document.panel === panel) { document.panel = null; document.reset('Spreadsheet editor closed'); }
    });
    panel.webview.onDidReceiveMessage((msg: WebviewToHost) => {
      switch (msg.type) {
        case 'initialized':
          document.loadError = undefined;
          document.markInitialized();
          break;
        case 'ready':
          document.post({
            type: 'init',
            bytes: encodeBytes(document.bytes),
            fileName: document.uri.path.split('/').pop() ?? 'Book1.xlsx',
            language: vscode.env.language,
            readOnly: vscode.workspace.fs.isWritableFileSystem(document.uri.scheme) === false,
            backup: document.backup,
          });
          break;
        case 'edit':
          this.changeEmitter.fire({
            document,
            label: msg.label,
            undo: () => document.post({ type: 'undo' }),
            redo: () => document.post({ type: 'redo' }),
          });
          break;
        case 'response':
          document.handleResponse(msg.id, msg.result, msg.error);
          break;
        case 'openLink':
          // Only web and mail links; VS Code asks before opening untrusted domains.
          if (/^(https?:|mailto:)/i.test(msg.url)) void vscode.env.openExternal(vscode.Uri.parse(msg.url, true));
          break;
        case 'error':
          document.loadError = msg.message;
          document.reset(msg.message);
          void vscode.window.showErrorMessage(`Sheets: ${msg.message}`);
          break;
        case 'warning':
          void vscode.window.showWarningMessage(msg.message);
          break;
      }
    });
  }

  async saveCustomDocument(document: SpreadsheetDocument): Promise<void> {
    await this.saveCustomDocumentAs(document, document.uri);
  }

  async saveCustomDocumentAs(document: SpreadsheetDocument, destination: vscode.Uri): Promise<void> {
    // Refuse at once instead of waiting for a webview that has no workbook (and would time out).
    if (document.loadError !== undefined) throw new Error(fill(shellStrings(vscode.env.language).saveRefusedNotOpened, document.uri.path.split('/').pop() ?? ''));
    const bytes = await document.request<string>('getFileData', { fileName: destination.path.split('/').pop() ?? '' });
    const data = decodeBytes(bytes);
    if (destination.toString() === document.uri.toString()) {
      if (document.diskBytes === null) throw new Error('Sheets: Die Dateiversion dieser älteren Sicherung ist unbekannt. Bitte mit „Speichern unter“ in eine neue Datei speichern.');
      if (document.diskBytes !== undefined) {
        const current = await vscode.workspace.fs.readFile(destination);
        if (current.length !== document.diskBytes.length || current.some((value, index) => value !== document.diskBytes![index])) {
          throw new Error('Sheets: Die Datei wurde außerhalb des Editors verändert. Sie wurde nicht überschrieben. Bitte neu laden oder mit „Speichern unter“ eine Kopie erstellen.');
        }
      }
    }
    await vscode.workspace.fs.writeFile(destination, data);
    if (destination.toString() === document.uri.toString()) { document.bytes = data; document.diskBytes = data; document.backup = false; }
  }

  async revertCustomDocument(document: SpreadsheetDocument): Promise<void> {
    document.bytes = await vscode.workspace.fs.readFile(document.uri);
    document.diskBytes = document.bytes;
    document.backup = false;
    document.reset();
    document.post({
      type: 'init',
      bytes: encodeBytes(document.bytes),
      fileName: document.uri.path.split('/').pop() ?? 'Book1.xlsx',
      language: vscode.env.language,
      readOnly: false,
    });
  }

  async backupCustomDocument(document: SpreadsheetDocument, context: vscode.CustomDocumentBackupContext): Promise<vscode.CustomDocumentBackup> {
    if (document.loadError !== undefined) throw new Error(document.loadError);
    const snapshot = await document.request<string>('getBackupData');
    const data = JSON.parse(new TextDecoder().decode(decodeBytes(snapshot)));
    data.diskBytes = document.diskBytes === undefined || document.diskBytes === null ? null : encodeBytes(document.diskBytes);
    await vscode.workspace.fs.writeFile(context.destination, new TextEncoder().encode(JSON.stringify(data)));
    return {
      id: context.destination.toString(),
      delete: () => void vscode.workspace.fs.delete(context.destination).then(undefined, () => undefined),
    };
  }

  private html(webview: vscode.Webview, fileName: string): string {
    // Shown until the webview has parsed the file, which takes a few seconds for large files.
    const strings = shellStrings(vscode.env.language);
    const escape = (text: string) => text.replace(/[&<>"]/g, (c) => `&#${c.charCodeAt(0)};`);
    const script = webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'dist', 'webview.js'));
    const nonce = Array.from({ length: 32 }, () => Math.floor(Math.random() * 36).toString(36)).join('');
    return `<!DOCTYPE html>
<html lang="${vscode.env.language}">
<head>
<meta charset="UTF-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src ${webview.cspSource} data:; style-src ${webview.cspSource} 'unsafe-inline'; font-src ${webview.cspSource}; script-src 'nonce-${nonce}';">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>html, body { margin: 0; padding: 0; height: 100%; overflow: hidden; } #app { height: 100vh; }
.sheets-status { box-sizing: border-box; height: 100%; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 8px; padding: 24px; text-align: center; white-space: pre-line; font-family: var(--vscode-font-family, sans-serif); font-size: var(--vscode-font-size, 13px); color: var(--vscode-foreground, #333); background: var(--vscode-editor-background, #fff); }
.sheets-status-loading strong::before { content: ''; display: inline-block; width: 12px; height: 12px; margin-right: 8px; vertical-align: -1px; border: 2px solid currentColor; border-right-color: transparent; border-radius: 50%; animation: sheets-spin 0.8s linear infinite; }
.sheets-status-error strong { color: var(--vscode-errorForeground, #c00); }
.sheets-status div { opacity: 0.8; max-width: 640px; }
@keyframes sheets-spin { to { transform: rotate(360deg); } }
@media (prefers-reduced-motion: reduce) { .sheets-status-loading strong::before { animation: none; } }</style>
</head>
<body>
<div id="app"><div class="sheets-status sheets-status-loading" role="status"><strong>${escape(fill(strings.loading, fileName))}</strong><div>${escape(strings.loadingDetail)}</div></div></div>
<script nonce="${nonce}" src="${script}"></script>
</body>
</html>`;
  }
}
