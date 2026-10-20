import { describe, expect, it, vi } from 'vitest';

vi.mock('vscode', () => ({
  EventEmitter: class {
    listeners = new Set<Function>();
    event = (listener: Function, thisArgs?: unknown, disposables?: { dispose(): void }[]) => {
      const bound = listener.bind(thisArgs);
      this.listeners.add(bound);
      const disposable = { dispose: () => this.listeners.delete(bound) };
      disposables?.push(disposable);
      return disposable;
    };
    fire(value: unknown) { for (const listener of this.listeners) listener(value); }
    dispose() { this.listeners.clear(); }
  },
  window: { registerCustomEditorProvider: vi.fn(() => ({ dispose() {} })) },
  workspace: { fs: { readFile: vi.fn(async () => new Uint8Array()) } },
}));
import * as vscode from 'vscode';
import { SpreadsheetEditorProvider } from '../src/provider.js';

describe('custom editor lifecycle', () => {
  it('routes edit events only to the registration that owns the document', async () => {
    vi.mocked(vscode.window.registerCustomEditorProvider).mockClear();
    const provider = SpreadsheetEditorProvider.register({ subscriptions: [] } as unknown as vscode.ExtensionContext);
    const registrations = vi.mocked(vscode.window.registerCustomEditorProvider).mock.calls;
    const xlsx = registrations[0]![1] as vscode.CustomEditorProvider;
    const csv = registrations[1]![1] as vscode.CustomEditorProvider;
    const uri = (name: string) => ({ toString: () => name }) as vscode.Uri;
    const token = {} as vscode.CancellationToken;
    const xlsxDoc = await xlsx.openCustomDocument(uri('book.xlsx'), {}, token);
    const csvDoc = await csv.openCustomDocument(uri('book.csv'), {}, token);
    const onXlsx = vi.fn(), onCsv = vi.fn();
    xlsx.onDidChangeCustomDocument(onXlsx);
    csv.onDidChangeCustomDocument(onCsv);
    const emit = (document: vscode.CustomDocument) =>
      (provider as any).changeEmitter.fire({ document, undo() {}, redo() {} });
    emit(xlsxDoc);
    expect(onXlsx).toHaveBeenCalledTimes(1);
    expect(onCsv).not.toHaveBeenCalled();
    emit(csvDoc);
    expect(onCsv).toHaveBeenCalledTimes(1);
    expect(onXlsx).toHaveBeenCalledTimes(1);
  });

  it('late disposal of an old editor does not remove its replacement', async () => {
    const provider = new SpreadsheetEditorProvider({} as vscode.ExtensionContext);
    const uri = { toString: () => 'file:///C:/book.xlsx' } as vscode.Uri;
    const old = await provider.openCustomDocument(uri, {});
    const current = await provider.openCustomDocument(uri, {});
    old.dispose();
    expect(provider.documents.get(uri.toString())).toBe(current);
    current.dispose();
    expect(provider.documents.size).toBe(0);
  });
});
