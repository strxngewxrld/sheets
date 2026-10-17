import * as vscode from 'vscode';
import { decodeBytes, encodeBytes } from './binary.js';
import type { HostToWebview, WebviewToHost } from './protocol.js';

/** A spreadsheet file opened in the custom editor. The live model lives in the webview. */
export class SpreadsheetDocument implements vscode.CustomDocument {
  readonly uri: vscode.Uri;
  bytes: Uint8Array;
  /** Last observed on-disk version; null means an older backup cannot prove it. */
  diskBytes: Uint8Array | null | undefined;
  panel: vscode.WebviewPanel | null = null;
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
  request<T>(method: 'getFileData' | 'getBackupData'): Promise<T>;
  request<T>(method: 'executeTool', params: { name: string; input: unknown }): Promise<T>;
  request<T>(method: 'exportCsv', params: { delimiter: string }): Promise<T>;
  request<T>(method: 'getFileData' | 'getBackupData' | 'executeTool' | 'exportCsv', params?: { name: string; input: unknown } | { delimiter: string }): Promise<T> {
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
      context.subscriptions.push(
        vscode.window.registerCustomEditorProvider(viewType, provider, {
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
      this.documents.delete(uri.toString());
      if (this.activeDocument === doc) this.activeDocument = undefined;
    });
    return doc;
  }

  async resolveCustomEditor(document: SpreadsheetDocument, panel: vscode.WebviewPanel): Promise<void> {
    document.reset();
    document.panel = panel;
    this.activeDocument = document;
    panel.webview.options = { enableScripts: true, localResourceRoots: [vscode.Uri.joinPath(this.context.extensionUri, 'dist')] };
    panel.webview.html = this.html(panel.webview);
    panel.onDidChangeViewState(() => {
      if (panel.active) this.activeDocument = document;
    });
    panel.onDidDispose(() => {
      if (document.panel === panel) { document.panel = null; document.reset('Spreadsheet editor closed'); }
    });
    panel.webview.onDidReceiveMessage((msg: WebviewToHost) => {
      switch (msg.type) {
        case 'initialized':
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
          document.reset(msg.message);
          void vscode.window.showErrorMessage(`Sheets: ${msg.message}`);
          break;
      }
    });
  }

  async saveCustomDocument(document: SpreadsheetDocument): Promise<void> {
    await this.saveCustomDocumentAs(document, document.uri);
  }

  async saveCustomDocumentAs(document: SpreadsheetDocument, destination: vscode.Uri): Promise<void> {
    const bytes = await document.request<string>('getFileData');
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
    const snapshot = await document.request<string>('getBackupData');
    const data = JSON.parse(new TextDecoder().decode(decodeBytes(snapshot)));
    data.diskBytes = document.diskBytes === undefined || document.diskBytes === null ? null : encodeBytes(document.diskBytes);
    await vscode.workspace.fs.writeFile(context.destination, new TextEncoder().encode(JSON.stringify(data)));
    return {
      id: context.destination.toString(),
      delete: () => void vscode.workspace.fs.delete(context.destination).then(undefined, () => undefined),
    };
  }

  private html(webview: vscode.Webview): string {
    const script = webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'dist', 'webview.js'));
    const nonce = Array.from({ length: 32 }, () => Math.floor(Math.random() * 36).toString(36)).join('');
    return `<!DOCTYPE html>
<html lang="${vscode.env.language}">
<head>
<meta charset="UTF-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src ${webview.cspSource} data:; style-src ${webview.cspSource} 'unsafe-inline'; font-src ${webview.cspSource}; script-src 'nonce-${nonce}';">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>html, body { margin: 0; padding: 0; height: 100%; overflow: hidden; } #app { height: 100vh; }</style>
</head>
<body>
<div id="app"></div>
<script nonce="${nonce}" src="${script}"></script>
</body>
</html>`;
  }
}
