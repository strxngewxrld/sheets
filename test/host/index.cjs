const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vscode = require('vscode');

exports.run = async () => {
  const checks = [];
  const report = (data) => {
    const file = process.env.SHEETS_HOST_RESULT;
    fs.writeFileSync(file + '.tmp', JSON.stringify(data));
    fs.renameSync(file + '.tmp', file);
  };
  try {
    const extension = vscode.extensions.getExtension('Tom-Knueppel.sheets-vscode');
    assert(extension, 'Sheets extension is missing');
    const api = await extension.activate();
    const startupDeadline = Date.now() + 15_000;
    while (!(vscode.window.tabGroups.activeTabGroup.activeTab?.input instanceof vscode.TabInputCustom) && Date.now() < startupDeadline) {
      await new Promise(resolve => setTimeout(resolve, 100));
    }
    const startupTab = vscode.window.tabGroups.activeTabGroup.activeTab;
    assert(startupTab?.input instanceof vscode.TabInputCustom && startupTab.input.viewType === 'sheets.spreadsheet',
      `Startup XLSX did not select Sheets: ${JSON.stringify({ input: startupTab?.input, associations: vscode.workspace.getConfiguration('workbench').inspect('editorAssociations') })}`);
    checks.push('an XLSX passed on the command line opens directly in Sheets at startup');
    const expected = extension.packageJSON.contributes.languageModelTools;
    const available = new Set(vscode.lm.tools.map((tool) => tool.name));
    for (const tool of expected) assert(available.has(tool.name), `Tool not visible in VS Code registry: ${tool.name}`);
    checks.push(`${expected.length} tools visible to the VS Code LM API`);

    // This fixture is created only inside the isolated test directory.
    const file = path.join(process.env.SHEETS_HOST_WORKDIR, 'Migration Grüße #100%.xlsx');
    const created = await api.executeTool('write_range', {
      start: 'A1', parse: false, values: [['Legacy No.', 'Country', 'Amount'], ['0004711', 'DE', 12.5], ['0004712', 'AT', 20]],
    }, file);
    assert.equal(created.ok, true, JSON.stringify(created));
    const table = await api.executeTool('format_table', { range: 'A1:C3', name: 'Migration' }, file);
    assert.equal(table.ok, true, JSON.stringify(table));

    const invoke = async (name, input) => {
      const cancellation = new vscode.CancellationTokenSource();
      const timer = setTimeout(() => cancellation.cancel(), 45_000);
      try {
        const result = await vscode.lm.invokeTool(name, { input, toolInvocationToken: undefined }, cancellation.token);
        return result.content.filter((part) => part instanceof vscode.LanguageModelTextPart).map((part) => part.value).join('\n');
      } finally { clearTimeout(timer); cancellation.dispose(); }
    };
    assert.match(await invoke('sheets_read_range', { file, range: 'A1:C3' }), /0004711/);
    checks.push('lm.invokeTool reads unopened XLSX without losing leading zeroes');
    assert.match(await invoke('sheets_read_range', { file: vscode.Uri.file(file).toString(), range: 'A1:C3' }), /0004711/);
    checks.push('native paths and escaped file URIs handle spaces, Unicode, # and %');

    // An agent edit of an unopened CSV must keep everything but the edited value:
    // delimiter, decimal comma, CRLF, missing final line break and the Windows-1252 umlaut.
    const csvFile = path.join(process.env.SHEETS_HOST_WORKDIR, 'export.csv');
    const latin = (text) => Buffer.from(text, 'latin1');
    fs.writeFileSync(csvFile, latin('Name;Wert\r\nMüller;1,5\r\nB;2'));
    const csvEdit = await api.executeTool('write_range', { start: 'B3', values: [[3]] }, csvFile);
    assert.equal(csvEdit.ok, true, JSON.stringify(csvEdit));
    assert.deepEqual(fs.readFileSync(csvFile), latin('Name;Wert\r\nMüller;1,5\r\nB;3'), 'Headless CSV write changed the file format');
    checks.push('headless agent writes keep the CSV format and Windows-1252 encoding');
    await vscode.commands.executeCommand('vscode.open', vscode.Uri.file(file));
    const tab = vscode.window.tabGroups.activeTabGroup.activeTab;
    assert(tab?.input instanceof vscode.TabInputCustom, 'XLSX did not open in a custom editor');
    assert.equal(tab.input.viewType, 'sheets.spreadsheet', 'XLSX opened in the wrong editor');
    checks.push('opening XLSX normally selects Sheets without an explicit editor override');
    assert.match(await invoke('sheets_list_tables', { file }), /Migration/);
    checks.push('lm.invokeTool reaches the live editor immediately after opening');
    await assert.rejects(() => invoke('sheets_read_range', { file }), /range|required|invalid/i);
    checks.push('invalid tool input is rejected');
    const beforeSave = fs.readFileSync(file);
    const edited = await api.executeTool('write_range', { start: 'C2', values: [[33.5]] }, file);
    assert.equal(edited.ok, true, JSON.stringify(edited));
    await vscode.commands.executeCommand('workbench.action.files.save');
    assert.notDeepEqual(fs.readFileSync(file), beforeSave, 'Live editor did not save its change');
    await vscode.commands.executeCommand('workbench.action.closeAllEditors');
    const saved = JSON.stringify(await api.executeTool('read_range', { range: 'A2:C2' }, file));
    assert.match(saved, /0004711/);
    assert.match(saved, /33\.5/);
    assert.match(JSON.stringify(await api.executeTool('list_tables', {}, file)), /Migration/);
    checks.push('live editor saves XLSX and reopening retains values and native tables');

    // A file that cannot be opened must never be overwritten, and saving must not wait for a timeout.
    const corrupt = path.join(process.env.SHEETS_HOST_WORKDIR, 'corrupt.xlsx');
    fs.writeFileSync(corrupt, 'this is not a zip archive');
    await vscode.commands.executeCommand('vscode.open', vscode.Uri.file(corrupt));
    await new Promise((resolve) => setTimeout(resolve, 3000));
    const started = Date.now();
    await vscode.commands.executeCommand('workbench.action.files.save').then(undefined, () => undefined);
    assert(Date.now() - started < 20_000, 'Saving an unopenable file waited for a webview timeout');
    assert.equal(fs.readFileSync(corrupt, 'utf8'), 'this is not a zip archive', 'An unopenable file was overwritten');
    await vscode.commands.executeCommand('workbench.action.closeAllEditors');
    checks.push('a file that cannot be opened is never overwritten by save');
    report({ status: 'passed', checks });
  } catch (error) {
    report({ status: 'failed', checks, error: error.stack ?? String(error) });
    throw error;
  } finally {
    // This is a separate application instance with an isolated profile.
    void vscode.commands.executeCommand('workbench.action.quit');
  }
};
