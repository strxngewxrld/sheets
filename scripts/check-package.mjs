import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { resolve } from 'node:path';
import { unzipSync, strFromU8, strToU8, zipSync } from 'fflate';
import { runInNewContext } from 'node:vm';
import { createRequire, isBuiltin } from 'node:module';

const hostRequire = createRequire(import.meta.url);

export async function checkPackage(path) {
  const files = unzipSync(await readFile(path));
  const names = Object.keys(files);
  const json = (name) => JSON.parse(strFromU8(files[`extension/${name}`] ?? assert.fail(`Missing ${name}`)));
  const pkg = json('package.json');
  for (const name of ['extension.vsixmanifest', '[Content_Types].xml', 'extension/LICENSE.txt', 'extension/readme.md', 'extension/changelog.md', 'extension/package.nls.json', 'extension/package.nls.de.json', `extension/${pkg.main.replace(/^\.\//, '')}`, 'extension/dist/webview.js', `extension/${pkg.icon}`]) assert(files[name], `Missing ${name}`);
  assert(!names.some((n) => /(?:\/node_modules\/|\/src\/|\/test\/|\.map$|\.ts$|\.vsix$)/.test(n)), 'Package contains development files');
  assert.equal(pkg.main, './dist/extension.js');
  assert(files['extension/THIRD_PARTY_NOTICES.md'], 'Missing bundled dependency licenses');
  assert.equal(pkg.name, 'sheets-vscode');
  // The extension ID (publisher.name) is used by the VS Code host test.
  assert.equal(pkg.publisher, 'Tom-Knueppel');
  assert.equal(pkg.contributes.configurationDefaults['workbench.editorAssociations']['*.xlsx'], 'sheets.spreadsheet');
  assert.equal(pkg.contributes.customEditors.find((editor) => editor.viewType === 'sheets.spreadsheet').priority, 'default');
  const png = files[`extension/${pkg.icon}`];
  assert.equal(png[0], 137, 'Icon must be PNG');
  const view = new DataView(png.buffer, png.byteOffset, png.byteLength);
  assert(view.getUint32(16) >= 128 && view.getUint32(20) >= 128, 'Icon is too small');
  const en = json('package.nls.json'), de = json('package.nls.de.json');
  for (const match of JSON.stringify(pkg).matchAll(/%([\w.]+)%/g)) assert(en[match[1]] && de[match[1]], `Missing translation: ${match[1]}`);
  // Load the actual packaged CommonJS entry point with only the VS Code API
  // mocked: this catches wrong package type, missing bundles and external deps.
  const commands = new Map(), providers = new Map(), tools = new Map();
  let disk = new Uint8Array(), writes = 0;
  const disposable = () => ({ dispose() {} });
  const vscode = {
    EventEmitter: class { event = () => disposable(); fire() {} dispose() {} },
    MarkdownString: class {
      value = '';
      appendText(value) { this.value += value; return this; }
      appendCodeblock(value) { this.value += value; return this; }
    },
    CancellationError: class extends Error {},
    LanguageModelTextPart: class { constructor(value) { this.value = value; } },
    LanguageModelToolResult: class { constructor(content) { this.content = content; } },
    window: {
      registerCustomEditorProvider(name, provider) { providers.set(name, provider); return disposable(); },
      showErrorMessage() {},
    },
    commands: { registerCommand(name, action) { commands.set(name, action); return disposable(); } },
    lm: { registerTool(name, tool) { tools.set(name, tool); return disposable(); } },
    Uri: {
      parse(value) { return { path: '/test.xlsx', toString: () => value }; },
      joinPath(uri, ...segments) { return { path: [uri.path, ...segments].join('/') }; },
    },
    env: { language: 'en' },
    workspace: { isTrusted: false, getConfiguration() { return { get() { return true; } }; }, fs: {
      isWritableFileSystem() { return true; },
      async readFile() { return disk; },
      async writeFile(_uri, value) { disk = value; writes++; },
    } },
  };
  const module = { exports: {} };
  runInNewContext(strFromU8(files['extension/dist/extension.js']), {
    module, exports: module.exports,
    require(name) {
      if (isBuiltin(name)) return hostRequire(name);
      assert.equal(name, 'vscode', `Unbundled dependency: ${name}`);
      return vscode;
    },
    console, setTimeout, clearTimeout, TextEncoder, TextDecoder, Uint8Array, Buffer, btoa, atob, structuredClone,
  }, { filename: 'extension/dist/extension.js', timeout: 10_000 });
  const context = { subscriptions: [], extensionUri: { path: '/extension' } };
  const api = module.exports.activate(context);
  for (const command of pkg.contributes.commands) assert(commands.has(command.command), `Unregistered command ${command.command}`);
  for (const editor of pkg.contributes.customEditors) assert(providers.has(editor.viewType), `Unregistered editor ${editor.viewType}`);
  for (const tool of pkg.contributes.languageModelTools) assert(tools.has(tool.name), `Unregistered tool ${tool.name}`);
  assert.equal(api.tools.length, pkg.contributes.languageModelTools.length);
  const readTool = tools.get('sheets_read_range'), writeTool = tools.get('sheets_write_range');
  assert.equal(readTool.prepareInvocation({ input: {} }).confirmationMessages, undefined);
  assert(writeTool.prepareInvocation({ input: { file: 'fixture.xlsx' } }).confirmationMessages);
  await assert.rejects(() => writeTool.invoke({ input: {} }, { isCancellationRequested: true }), vscode.CancellationError);
  const restricted = await api.executeTool('write_range', { start: 'A1', values: [[1]] }, 'file:///test.xlsx');
  assert.equal(restricted.ok, false);
  assert.match(restricted.error, /Trust this workspace/);
  const provider = providers.get('sheets.spreadsheet');
  const uri = { path: '/fixture.xlsx', scheme: 'file', toString: () => 'file:///fixture.xlsx' };
  const doc = await provider.openCustomDocument(uri, { untitledDocumentData: new Uint8Array() });
  const sent = [];
  let receive;
  await provider.resolveCustomEditor(doc, {
    webview: {
      asWebviewUri(uri) { return uri.path; },
      postMessage(message) { sent.push(message); return Promise.resolve(true); },
      onDidReceiveMessage(callback) { receive = callback; },
    },
    onDidChangeViewState() {}, onDidDispose() {},
  });
  const response = doc.request('getFileData');
  assert.equal(sent.length, 0, 'Requests must wait for the webview to load');
  receive({ type: 'ready' });
  assert.deepEqual(sent.map((m) => m.type), ['init'], 'Script readiness is not workbook readiness');
  receive({ type: 'initialized' });
  assert.equal(sent[1].method, 'getFileData');
  receive({ type: 'response', id: sent[1].id, result: 'saved-file' });
  assert.equal(await response, 'saved-file');
  const invocation = readTool.invoke({ input: { file: 'file:///fixture.xlsx', range: 'A1' } }, { isCancellationRequested: false });
  const toolRequest = sent.at(-1);
  assert.equal(toolRequest.method, 'executeTool');
  assert.equal(toolRequest.params.name, 'read_range');
  receive({ type: 'response', id: toolRequest.id, result: { ok: true, result: { values: [['Copilot tool path verified']] } } });
  assert.match((await invocation).content[0].value, /Copilot tool path verified/);
  const rejectedInvocation = readTool.invoke({ input: { file: 'file:///fixture.xlsx', range: 'A1' } }, { isCancellationRequested: false });
  const rejected = assert.rejects(rejectedInvocation, /Unknown sheet/);
  receive({ type: 'response', id: sent.at(-1).id, result: { ok: false, error: 'Unknown sheet' } });
  await rejected;
  const blockedSave = assert.rejects(provider.saveCustomDocumentAs(doc, uri), /XLSX-Schutz/);
  receive({ type: 'response', id: sent.at(-1).id, error: 'XLSX-Schutz: unsupported edit' });
  await blockedSave;
  assert.equal(writes, 0, 'A rejected serialization must never write to disk');
  doc.diskBytes = Uint8Array.of(1); disk = Uint8Array.of(2);
  const conflict = assert.rejects(provider.saveCustomDocumentAs(doc, uri), /außerhalb des Editors/);
  receive({ type: 'response', id: sent.at(-1).id, result: 'Aw==' });
  await conflict;
  assert.equal(writes, 0, 'External changes must never be overwritten');
  disk = Uint8Array.of(1);
  const save = provider.saveCustomDocumentAs(doc, uri);
  receive({ type: 'response', id: sent.at(-1).id, result: 'Aw==' });
  await save;
  assert.equal(writes, 1);
  assert.equal(doc.diskBytes[0], 3);
  const backup = provider.backupCustomDocument(doc, { destination: { path: '/backup', toString: () => 'file:///backup' } });
  assert.equal(sent.at(-1).method, 'getBackupData');
  receive({ type: 'response', id: sent.at(-1).id, result: Buffer.from(JSON.stringify({ type: 'sheets-backup-v1', workbook: { sheets: [] } })).toString('base64') });
  await backup;
  const restored = await provider.openCustomDocument(uri, { backupId: 'file:///backup' });
  assert.equal(restored.backup, true);
  assert.equal(restored.diskBytes[0], 3, 'Hot-exit backup must retain the original disk version');
  restored.dispose();
  const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const fixture = {
    '_rels/.rels': strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    'xl/workbook.xml': strToU8(`<workbook xmlns="${ns}" xmlns:r="${rel}"><sheets><sheet name="Sheet1" sheetId="1" r:id="r1"/></sheets></workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(`<worksheet xmlns="${ns}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Original</t></is></c></row></sheetData></worksheet>`),
    'customXml/opaque.bin': Uint8Array.of(0, 1, 255),
  };
  disk = zipSync(fixture);
  const originalDisk = disk, originalWrites = writes;
  vscode.workspace.isTrusted = true;
  const unsafe = await api.executeTool('format_table', { range: 'A1:A2' }, 'file:///headless.xlsx');
  assert.equal(unsafe.ok, false);
  assert.match(unsafe.error, /XLSX-Schutz/);
  assert.equal(writes, originalWrites);
  assert.deepEqual(disk, originalDisk);
  const safe = await api.executeTool('write_range', { start: 'A1', values: [['Changed']], parse: false }, 'file:///headless.xlsx');
  assert.equal(safe.ok, true, safe.error);
  assert.deepEqual(unzipSync(disk)['customXml/opaque.bin'], fixture['customXml/opaque.bin']);
  doc.reset();
  const failed = doc.request('getFileData').catch((error) => error.message);
  receive({ type: 'error', message: 'Cannot open workbook' });
  assert.equal(await failed, 'Cannot open workbook');
  doc.dispose();
  console.log(`VSIX verified: ${names.length} files, ${commands.size} commands, ${providers.size} editors, ${tools.size} AI tools; packaged extension activates.`);
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const root = new URL('..', import.meta.url);
  const pkg = JSON.parse(await readFile(new URL('package.json', root), 'utf8'));
  await checkPackage(process.argv[2] ?? fileURLToPath(new URL(`artifacts/${pkg.name}-${pkg.version}.vsix`, root)));
}
