import { beforeEach, describe, expect, it, vi } from 'vitest';
vi.mock('vscode', () => ({
  TabInputText: class { constructor(public uri: unknown) {} },
  window: { tabGroups: { all: [] as unknown[], close: vi.fn(async () => true) } },
  workspace: { getConfiguration: vi.fn(() => ({ inspect: vi.fn(() => ({})) })) },
  commands: { executeCommand: vi.fn(async () => undefined) },
}));
import * as vscode from 'vscode';
import { reopenStartupSpreadsheets } from '../src/startup-editors.js';

function tab(path: string, dirty = false) {
  return { input: new vscode.TabInputText({ path } as vscode.Uri), isDirty: dirty, isActive: true, isPreview: true };
}
function setup(tabs: unknown[], overrides = {}, active = true) {
  (vscode.window.tabGroups as any).all = [{ tabs, isActive: active, viewColumn: 2 }];
  vi.mocked(vscode.workspace.getConfiguration).mockReturnValue({ inspect: () => overrides } as any);
}
beforeEach(() => vi.clearAllMocks());

describe('startup spreadsheet editor recovery', () => {
  it.each(['/C:/Grüße #100%.xlsx', '/Users/name/Documents/Grüße #100%.xlsx', '/Users/name/Documents/Makros.XLSM'])(
    'reopens %s in its group and closes only the original text tab', async path => {
    const original = tab(path);
    setup([original], {}, false);
    await reopenStartupSpreadsheets();
    expect(vscode.commands.executeCommand).toHaveBeenCalledWith('vscode.openWith', original.input.uri, 'sheets.spreadsheet', {
      viewColumn: 2, preview: true, preserveFocus: true,
    });
    expect(vscode.window.tabGroups.close).toHaveBeenCalledWith(original, true);
  });
  it('leaves dirty tabs, CSV and other editors alone', async () => {
    setup([tab('/C:/dirty.xlsx', true), tab('/C:/sample.csv'), { input: {}, isDirty: false }]);
    await reopenStartupSpreadsheets();
    expect(vscode.commands.executeCommand).not.toHaveBeenCalled();
  });
  it.each(['globalValue', 'workspaceValue', 'workspaceFolderValue'])('respects explicit %s editor associations', async scope => {
    setup([tab('/C:/book.xlsx')], { [scope]: { '*.xlsx': 'default' } });
    await reopenStartupSpreadsheets();
    expect(vscode.commands.executeCommand).not.toHaveBeenCalled();
  });
  it('never closes the original tab when opening Sheets fails', async () => {
    setup([tab('/C:/book.xlsx')]);
    vi.mocked(vscode.commands.executeCommand).mockRejectedValueOnce(new Error('Cannot open'));
    await expect(reopenStartupSpreadsheets()).rejects.toThrow('Cannot open');
    expect(vscode.window.tabGroups.close).not.toHaveBeenCalled();
  });
});
