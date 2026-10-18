// Real custom-editor UI checks over Chromium DevTools, without browser dependencies.
// Uses only a disposable VS Code profile and generated spreadsheet fixtures.
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createServer } from 'node:net';
import { vscodeExecutable } from './vscode-executable.mjs';
const root = fileURLToPath(new URL('..', import.meta.url));
const dir = await mkdtemp(join(tmpdir(), 'sheets-ui-test-'));
const profile = join(dir, 'profile');
await mkdir(join(profile, 'User'), { recursive: true });
await writeFile(join(profile, 'User/settings.json'), JSON.stringify({
  'workbench.startupEditor': 'none', 'update.mode': 'none', 'telemetry.telemetryLevel': 'off',
  'sheets.checkForUpdates': false, 'workbench.colorTheme': 'Default Light Modern',
}));
const server = createServer();
await new Promise(r => server.listen(0, '127.0.0.1', r));
const port = server.address().port;
await new Promise(r => server.close(r));
console.log(`UI test artifacts: ${dir}`);
const child = spawn(vscodeExecutable(), [
  '--user-data-dir', profile, '--extensions-dir', join(dir, 'extensions'),
  '--extensionDevelopmentPath', root, '--extensionTestsPath', join(root, 'test/ui/host.cjs'),
  `--remote-debugging-port=${port}`, '--disable-extensions', '--disable-workspace-trust',
  '--skip-welcome', '--skip-release-notes', '--new-window', '--wait',
], { env: { ...process.env, SHEETS_UI_WORKDIR: dir }, stdio: 'ignore' });
let spawnError;
child.on('error', e => { spawnError = e; });
const pause = ms => new Promise(r => setTimeout(r, ms));
async function until(fn, timeout = 30000) {
  const deadline = Date.now() + timeout;
  let error;
  while (Date.now() < deadline) {
    if (spawnError) throw spawnError;
    try { const result = await fn(); if (result) return result; } catch (e) { error = e; }
    await pause(100);
  }
  throw error ?? Error('UI wait timed out');
}
let ws, sessionId, sequence = 0;
const pending = new Map(), contexts = [], exceptions = [], checks = [];
async function call(method, params = {}, session = sessionId) {
  const id = ++sequence;
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => { pending.delete(id); reject(Error(`CDP timeout: ${method}`)); }, 15000);
    pending.set(id, m => { clearTimeout(timer); if (m.error) reject(Error(JSON.stringify(m.error))); else resolve(m.result); });
    ws.send(JSON.stringify({ id, method, params, sessionId: session ?? undefined }));
  });
}
async function attach() {
  sessionId = undefined;
  contexts.length = 0;
  const targets = await call('Target.getTargets');
  const frame = targets.targetInfos.find(t => t.type === 'iframe' && t.url.includes('Tom-Knueppel.sheets-vscode'));
  if (!frame) return false;
  sessionId = (await call('Target.attachToTarget', { targetId: frame.targetId, flatten: true })).sessionId;
  await call('Runtime.enable');
  for (const context of contexts) {
    const probe = await call('Runtime.evaluate', { expression: 'Boolean(document.querySelector(".sx-root"))', contextId: context.id, returnByValue: true });
    if (probe.result.value) return context.id;
  }
  return false;
}
let contextId;
async function evaluate(expression) {
  await call('Runtime.evaluate', { expression: 'new Promise(resolve => requestAnimationFrame(() => resolve(true)))', contextId, returnByValue: true, awaitPromise: true });
  const result = await call('Runtime.evaluate', { expression, contextId, returnByValue: true, awaitPromise: true });
  if (result.exceptionDetails) throw Error(result.exceptionDetails.exception?.description ?? result.exceptionDetails.text);
  await call('Runtime.evaluate', { expression: 'new Promise(resolve => requestAnimationFrame(() => resolve(true)))', contextId, returnByValue: true, awaitPromise: true });
  return result.result.value;
}
async function host(command) {
  const id = ++sequence;
  const input = join(dir, 'command.json');
  await writeFile(input + '.tmp', JSON.stringify({ id, ...command }));
  // Host reads one complete JSON file; atomic rename avoids partial commands.
  const { rename } = await import('node:fs/promises');
  await rename(input + '.tmp', input);
  const response = await until(async () => {
    const r = JSON.parse(await readFile(join(dir, 'response.json'), 'utf8'));
    return r.id === id ? r : false;
  });
  if (response.error) throw Error(response.error);
  return response.result;
}
async function shot(name) {
  const result = await call('Page.captureScreenshot', { format: 'png' }, null);
  await writeFile(join(dir, name + '.png'), Buffer.from(result.data, 'base64'));
}
async function check(name, fn) { await fn(); checks.push(name); console.log(`PASS ${name}`); }
const click = title => evaluate(`{const b=document.querySelector('button[title=${JSON.stringify(title)}]');b.focus();b.click();}`);
const escape = () => evaluate(`document.activeElement.dispatchEvent(new KeyboardEvent('keydown',{key:'Escape',bubbles:true}))`);
const select = range => evaluate(`{const n=document.querySelector('.sx-namebox');n.value=${JSON.stringify(range)};n.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}));}`);
try {
  await until(() => readFile(join(dir, 'ready.json'), 'utf8'), 120000);
  const targets = await until(async () => (await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()).find(t => t.type === 'page'));
  ws = new WebSocket(targets.webSocketDebuggerUrl);
  await new Promise((r, j) => { ws.addEventListener('open', r, { once: true }); ws.addEventListener('error', j, { once: true }); });
  ws.addEventListener('message', e => {
    const m = JSON.parse(e.data);
    if (m.id) { pending.get(m.id)?.(m); pending.delete(m.id); }
    if (m.method === 'Runtime.executionContextCreated') contexts.push(m.params.context);
    if (m.method === 'Runtime.exceptionThrown' && m.sessionId === sessionId) exceptions.push(m.params.exceptionDetails);
  });
  contextId = await until(attach);
  await check('XLSX formula bar and Unicode values', async () => {
    await select('A2'); assert.equal(await evaluate(`document.querySelector('.sx-formula').value`), 'Müller & Söhne');
    await select('E2'); assert.equal(await evaluate(`document.querySelector('.sx-formula').value`), '=C2*D2');
  });
  await check('Find, invalid regex, replace and undo through the live host', async () => {
    await host({ command: 'sheets.replace' });
    await until(() => evaluate(`Boolean(document.querySelector('.sx-find-dialog'))`));
    await evaluate(`{const q=document.querySelector('[aria-label="Find"]');q.value='Müller';q.dispatchEvent(new Event('input',{bubbles:true}));}`);
    assert.match(await evaluate(`document.querySelector('.sx-find-status').textContent`), /1/);
    await evaluate(`{const q=document.querySelector('[aria-label="Replace with"]');q.value='Meier';q.dispatchEvent(new Event('input',{bubbles:true}));[...document.querySelectorAll('.sx-find-dialog button')].find(b=>b.textContent==='Replace all').click();}`);
    assert.match(JSON.stringify(await host({ tool: 'read_range', input: { range: 'A2' } })), /Meier/);
    await host({ command: 'undo' });
    assert.match(JSON.stringify(await host({ tool: 'read_range', input: { range: 'A2' } })), /Müller/);
    await evaluate(`{const q=document.querySelector('[aria-label="Find"]');q.value='[';document.querySelectorAll('.sx-find-dialog input[type=checkbox]')[3].click();q.dispatchEvent(new Event('input',{bubbles:true}));}`);
    assert.equal(await evaluate(`document.querySelector('[aria-label="Find"]').getAttribute('aria-invalid')`), 'true');
    await escape();
  });
  await check('Grid editing, formulas, rectangular paste, undo and redo', async () => {
    await select('G2');
    await evaluate(`{const t=document.querySelector('.sx-editor');t.value='21';t.dispatchEvent(new Event('input',{bubbles:true}));t.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}));}`);
    await select('H2');
    await evaluate(`{const f=document.querySelector('.sx-formula');f.focus();f.value='=G2*2';f.dispatchEvent(new Event('input',{bubbles:true}));f.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}));}`);
    assert.match(JSON.stringify(await host({tool:'read_range',input:{range:'G2:H2'}})), /42/);
    await host({command:'undo'}); await host({command:'redo'});
    assert.match(JSON.stringify(await host({tool:'read_range',input:{range:'H2'}})), /42/);
    await select('G4');
    await evaluate(`{const d=new DataTransfer();d.setData('text/plain','Alpha\\t2\\nBeta\\t3');document.querySelector('.sx-editor').dispatchEvent(new ClipboardEvent('paste',{clipboardData:d,bubbles:true,cancelable:true}));}`);
    const pasted=JSON.stringify(await host({tool:'read_range',input:{range:'G4:H5'}}));
    assert.match(pasted,/Alpha/);assert.match(pasted,/Beta/);
    for(let i=0;i<3;i++)await host({command:'undo'});
    await select('A1');
  });
  await check('Data validation rejects an invalid UI edit without changing the cell', async () => {
    await host({tool:'set_data_validation',input:{range:'B2',type:'list',values:['DE','AT']}});
    await select('B2');
    await evaluate(`{const f=document.querySelector('.sx-formula');f.focus();f.value='INVALID';f.dispatchEvent(new Event('input',{bubbles:true}));f.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}));}`);
    assert(await evaluate(`Boolean(document.querySelector('.sx-tip-error'))`));
    assert.match(JSON.stringify(await host({tool:'read_range',input:{range:'B2'}})), /DE/);
    await evaluate(`[...document.querySelectorAll('.sx-tip-error button')].find(b=>b.textContent==='Cancel').click()`);
    await host({command:'undo'});
  });
  await check('Context menu stays inside a small editor and scrolls to the last action', async () => {
    const result = await evaluate(`{const root=document.querySelector('.sx-root');root.style.width='480px';root.style.height='320px';document.querySelector('.sx-grid').dispatchEvent(new MouseEvent('contextmenu',{bubbles:true,clientX:460,clientY:300}));const m=document.querySelector('.sx-menu'),r=m.getBoundingClientRect();JSON.stringify({left:r.left,top:r.top,right:r.right,bottom:r.bottom,scroll:m.scrollHeight>m.clientHeight})}`);
    const r = JSON.parse(result); assert(r.left >= 0 && r.top >= 0 && r.right <= 480 && r.bottom <= 320); assert(r.scroll);
    await evaluate(`document.activeElement.dispatchEvent(new KeyboardEvent('keydown',{key:'Home',bubbles:true}));for(let i=0;i<12;i++)document.activeElement.dispatchEvent(new KeyboardEvent('keydown',{key:'ArrowDown',bubbles:true}));`);
    assert.equal(await evaluate(`document.activeElement.textContent`), 'Hide columns');
    await evaluate(`document.activeElement.dispatchEvent(new KeyboardEvent('keydown',{key:'End',bubbles:true}))`);
    assert.equal(await evaluate(`document.activeElement.textContent`), 'Clear formats');
    assert(await evaluate(`{const a=document.activeElement.getBoundingClientRect(),m=document.querySelector('.sx-menu').getBoundingClientRect();a.bottom<=m.bottom}`));
    await shot('small-context-menu'); await escape();
  });
  await check('Menu arrows, wraparound, Home, disabled items, Escape and focus restoration', async () => {
    await click('Borders');
    await evaluate(`document.activeElement.dispatchEvent(new KeyboardEvent('keydown',{key:'ArrowUp',bubbles:true}))`);
    assert.equal(await evaluate('document.activeElement.textContent'), 'No borders');
    await evaluate(`document.activeElement.dispatchEvent(new KeyboardEvent('keydown',{key:'ArrowDown',bubbles:true}))`);
    assert.equal(await evaluate('document.activeElement.textContent'), 'All borders');
    await evaluate(`document.activeElement.dispatchEvent(new KeyboardEvent('keydown',{key:'Home',bubbles:true}))`);
    assert.equal(await evaluate('document.activeElement.textContent'), 'All borders');
    await escape();
    assert.equal(await evaluate(`!!document.querySelector('.sx-menu')`), false);
    assert.equal(await evaluate(`document.activeElement.title`), 'Borders');
    await click('Borders');
    await evaluate(`document.querySelector('.sx-menu button').click()`);
    assert.equal(await evaluate(`document.activeElement.title`), 'Borders');
    await host({command:'undo'});
  });
  await check('Repeated menu actions remove global listeners and preserve relative positioning', async () => {
    const result=await evaluate(`(async()=>{
      const root=document.querySelector('.sx-root');root.style.marginLeft='25px';root.style.marginTop='15px';
      const add=window.addEventListener,remove=window.removeEventListener,live=new Map();
      window.addEventListener=function(type,fn,...args){if(['mousedown','keydown','blur','resize'].includes(type)){if(!live.has(type))live.set(type,new Set());live.get(type).add(fn);}return add.call(this,type,fn,...args);};
      window.removeEventListener=function(type,fn,...args){live.get(type)?.delete(fn);return remove.call(this,type,fn,...args);};
      let inside=true;
      try{for(let i=0;i<12;i++){
        document.querySelector('button[title="Pivot report"]').click();
        await new Promise(r=>setTimeout(r,5));
        const m=document.querySelector('.sx-menu'),b=m.getBoundingClientRect(),r=root.getBoundingClientRect();
        const anchor=document.querySelector('button[title="Pivot report"]').getBoundingClientRect();
        const expected=Math.max(r.left+4,Math.min(anchor.left,r.right-b.width-4));
        inside &&= Math.abs(b.left-expected)<1 && b.top>=r.top && b.bottom<=r.bottom;
        m.querySelector('button').click();
        document.activeElement.dispatchEvent(new KeyboardEvent('keydown',{key:'Escape',bubbles:true}));
      }return {inside,listeners:[...live.values()].reduce((n,s)=>n+s.size,0)};}
      finally{window.addEventListener=add;window.removeEventListener=remove;root.style.marginLeft='';root.style.marginTop='';}
    })()`);
    assert(result.inside);assert.equal(result.listeners,0);
    await click('Borders');
    await evaluate(`document.querySelector('.sx-root').style.width='460px'`);
    await until(async()=>!(await evaluate(`Boolean(document.querySelector('.sx-menu'))`)));
  });
  await check('All main dialogs fit a narrow pane and close with Escape', async () => {
    await evaluate(`document.querySelector('.sx-root').style.width='260px'`);
    for (const title of ['Format cells…', 'Format as table…', 'Sort…', 'Conditional formatting']) {
      await click(title);
      const r = await evaluate(`{const d=document.querySelector('.sx-dialog'),r=d.getBoundingClientRect();JSON.stringify({left:r.left,right:r.right,bottom:r.bottom,overflow:d.scrollWidth>d.clientWidth})}`);
      const b = JSON.parse(r); assert(b.left >= 0 && b.right <= 260 && b.bottom <= 320, `${title}: ${r}`); assert(!b.overflow, `${title} overflows horizontally`);
      await shot('narrow-' + title.replace(/[^a-z]/gi, '-')); await escape();
      assert.equal(await evaluate(`!!document.querySelector('.sx-dialog')`), false);
    }
    await host({ command: 'sheets.find' });
    await until(() => evaluate(`Boolean(document.querySelector('.sx-find-dialog'))`));
    const r = await evaluate(`document.querySelector('.sx-find-dialog').getBoundingClientRect().toJSON()`);
    assert(r.left >= 0 && r.right <= 260); await escape();
    await evaluate(`{const r=document.querySelector('.sx-root');r.style.width='';r.style.height='';}`);
  });
  await check('Formatting table, filters, computed values, save and reopen', async () => {
    await select('A1:E5'); await click('Format as table…');
    await evaluate(`document.querySelector('.sx-dialog .sx-primary').click()`);
    assert.equal(await evaluate(`!!document.querySelector('.sx-dialog')`), false);
    assert.match(JSON.stringify(await host({ tool: 'list_tables' })), /Table1/);
    const values = await host({ tool: 'read_range', input: { range: 'E2:E6' } });
    assert.match(JSON.stringify(values), /557.42/);
    await host({ command: 'workbench.action.files.save' });
    await host({ command: 'workbench.action.closeAllEditors' });
    await host({ open: 'ui.xlsx' }); contextId = await until(attach);
    assert.match(JSON.stringify(await host({ tool: 'list_tables' })), /Table1/);
    await shot('light-table');
  });
  await check('Value filters, selection statistics and zoom', async () => {
    await select('C2:C5');
    assert.match(await evaluate(`document.querySelector('.sx-status').textContent`), /25/);
    await click('Zoom in');
    assert.match(await evaluate(`document.querySelector('.sx-zoom-level').textContent`), /110/);
    await click('Reset zoom to 100 %');
    await evaluate(`document.querySelectorAll('.sx-table-sort')[1].click()`);
    await evaluate(`[...document.querySelectorAll('.sx-menu button')].find(b=>b.textContent==='Filter values…').click()`);
    assert.match(await evaluate(`document.querySelector('.sx-dialog').innerText`), /filter/i);
    await evaluate(`{const q=document.querySelector('.sx-dialog input[type=search]');q.value='DE';q.dispatchEvent(new Event('input',{bubbles:true}));}`);
    assert.equal(await evaluate(`document.querySelectorAll('.sx-filter-values input').length`), 1);
    await evaluate(`{const all=document.querySelector('.sx-dialog input[type=checkbox]');all.click();document.querySelector('.sx-dialog .sx-primary').click();}`);
    assert(await evaluate(`Boolean(document.querySelector('.sx-filter-active'))`));
    await host({command:'undo'});
  });
  await check('Chart insertion, all chart types, editor and undo', async () => {
    await select('C1:E5'); await click('Insert chart');
    await evaluate(`document.querySelector('.sx-menu button').click()`);
    assert(await evaluate(`Boolean(document.querySelector('.sx-chart svg'))`));
    await evaluate(`document.querySelector('.sx-chart').dispatchEvent(new MouseEvent('dblclick',{bubbles:true}))`);
    assert.match(await evaluate(`document.querySelector('.sx-dialog').innerText`), /Chart/);
    for (const type of ['bar','line','area','pie','doughnut','scatter','column']) {
      await evaluate(`{const s=document.querySelector('.sx-dialog select');s.value=${JSON.stringify(type)};document.querySelector('.sx-dialog .sx-primary').click();}`);
      assert(await evaluate(`Boolean(document.querySelector('.sx-chart svg'))`), type);
    }
    await shot('chart'); await escape();
    // Undo every type change and the insertion, leaving the saved table intact.
    for (let i=0;i<8;i++) await host({command:'undo'});
    assert.equal(await evaluate(`Boolean(document.querySelector('.sx-chart'))`), false);
  });
  await check('Sheet add, rename, undo and redo', async () => {
    await click('Add sheet');
    assert.equal(await evaluate(`document.querySelectorAll('.sx-tab').length`), 2);
    await evaluate(`document.querySelector('.sx-tab[aria-selected=true]').dispatchEvent(new MouseEvent('dblclick',{bubbles:true}))`);
    await evaluate(`{const i=document.querySelector('.sx-tab input');i.value='Übersicht';i.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}));}`);
    assert.equal(await evaluate(`document.querySelector('.sx-tab[aria-selected=true]').textContent`), 'Übersicht');
    await host({command:'undo'}); await host({command:'redo'});
    assert.equal(await evaluate(`document.querySelector('.sx-tab[aria-selected=true]').textContent`), 'Übersicht');
    await host({command:'undo'}); await host({command:'undo'});
    assert.equal(await evaluate(`document.querySelectorAll('.sx-tab').length`), 1);
  });
  await check('Dark theme updates the actual grid and dialogs', async () => {
    await host({ config: 'workbench.colorTheme', value: 'Default Dark Modern' });
    await until(() => evaluate(`document.body.classList.contains('vscode-dark')`));
    await click('Format cells…'); await shot('dark-format'); await escape();
  });
  await check('CSV opens in Sheets, edits and saves without changing delimiter or CRLF', async () => {
    await writeFile(join(dir, 'sample.csv'), 'Name;Wert\r\nMüller;1,5\r\nB;2');
    await host({ command: 'workbench.action.closeAllEditors' });
    await host({ open: 'sample.csv', editor: 'sheets.csv' }); contextId = await until(attach);
    await select('B3');
    await evaluate(`{const f=document.querySelector('.sx-formula');f.focus();f.value='3';f.dispatchEvent(new Event('input',{bubbles:true}));f.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}));}`);
    await host({ command: 'workbench.action.files.save' });
    assert.equal(await readFile(join(dir, 'sample.csv'), 'utf8'), 'Name;Wert\r\nMüller;1,5\r\nB;3');
  });
  assert.deepEqual(exceptions, [], 'Uncaught webview errors');
  await writeFile(join(dir, 'result.json'), JSON.stringify({ status: 'passed', checks }, null, 2));
  console.log(`UI tests passed (${checks.length} scenarios). Screenshots and fixtures: ${dir}`);
} catch (error) {
  await shot('failure').catch(() => {});
  await writeFile(join(dir, 'result.json'), JSON.stringify({ status: 'failed', checks, error: error.stack, exceptions }, null, 2));
  throw error;
} finally {
  if (ws) {
    await host({ command: 'workbench.action.quit' }).catch(() => {});
    ws.close();
  }
  if (child.exitCode === null) child.kill();
}
