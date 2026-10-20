import * as vscode from 'vscode';

/** Recover binary spreadsheet tabs resolved before custom editors became available. */
export async function reopenStartupSpreadsheets(): Promise<void> {
  for (const group of vscode.window.tabGroups.all) {
    for (const tab of [...group.tabs]) {
      if (!(tab.input instanceof vscode.TabInputText) || tab.isDirty || !/\.(xlsx|xlsm)$/i.test(tab.input.uri.path)) continue;
      const uri = tab.input.uri;
      const associations = vscode.workspace.getConfiguration('workbench', uri).inspect<Record<string, string>>('editorAssociations');
      // A user/workspace mapping may deliberately select a different editor.
      // Be conservative for arbitrary glob patterns: leave every explicit mapping alone.
      if ([associations?.globalValue, associations?.workspaceValue, associations?.workspaceFolderValue]
        .some(value => value && Object.keys(value).length > 0)) continue;
      if (!group.tabs.includes(tab) || tab.isDirty) continue;
      await vscode.commands.executeCommand('vscode.openWith', uri, 'sheets.spreadsheet', {
        viewColumn: group.viewColumn, preview: tab.isPreview, preserveFocus: !group.isActive || !tab.isActive,
      });
      if (group.tabs.includes(tab) && !tab.isDirty) await vscode.window.tabGroups.close(tab, true);
    }
  }
}
