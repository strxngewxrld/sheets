// Runs only in the isolated profile created by scripts/test-ui.mjs.
const vscode = require('vscode');
const fs = require('node:fs');
const path = require('node:path');
exports.run = async () => {
  const dir = process.env.SHEETS_UI_WORKDIR;
  if (!dir) throw Error('Missing isolated UI test directory');
  const api = await vscode.extensions.getExtension('Tom-Knueppel.sheets-vscode').activate();
  const file = path.join(dir, 'ui.xlsx');
  const created = await api.executeTool('write_range', { start: 'A1', parse: false, values: [
    ['Artikel', 'Land', 'Menge', 'Preis', 'Umsatz'],
    ['Müller & Söhne', 'DE', 3, 12.5, '=C2*D2'],
    ['Bürobedarf', 'AT', 8, 9.99, '=C3*D3'],
    ['東京', 'JP', 2, 199, '=C4*D4'],
    ['0004711', 'CH', 12, 3.5, '=C5*D5'],
    ['Gesamt', '', '', '', '=SUM(E2:E5)'],
  ] }, file);
  if (!created.ok) throw Error(JSON.stringify(created));
  await vscode.commands.executeCommand('vscode.open', vscode.Uri.file(file));
  await vscode.commands.executeCommand('workbench.action.closeAuxiliaryBar');
  await vscode.commands.executeCommand('notifications.clearAll');
  fs.writeFileSync(path.join(dir, 'ready.json'), JSON.stringify({ file }));
  let last = '';
  await new Promise(() => setInterval(async () => {
    const input = path.join(dir, 'command.json');
    if (!fs.existsSync(input)) return;
    const raw = fs.readFileSync(input, 'utf8');
    if (last === raw) return;
    last = raw;
    const c = JSON.parse(raw);
    try {
      let result;
      if (c.state) result = vscode.window.tabGroups.all.flatMap(g => g.tabs.map(t => ({ label: t.label, dirty: t.isDirty, active: t.isActive })));
      else if (c.config) await vscode.workspace.getConfiguration().update(c.config, c.value, vscode.ConfigurationTarget.Global);
      else if (c.open) await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(path.join(dir, c.open)), c.editor ?? 'sheets.spreadsheet', { preview: false });
      else if (c.tool) result = await api.executeTool(c.tool, c.input ?? {}, file);
      else result = await vscode.commands.executeCommand(c.command, ...(c.args ?? []));
      fs.writeFileSync(path.join(dir, 'response.json'), JSON.stringify({ id: c.id, result }));
    } catch (e) {
      fs.writeFileSync(path.join(dir, 'response.json'), JSON.stringify({ id: c.id, error: e.stack }));
    }
  }, 100));
};
