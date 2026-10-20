import { spawn } from 'node:child_process';
import { access, copyFile, mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { vscodeExecutable, vscodeEnvironment } from './vscode-executable.mjs';

const root = fileURLToPath(new URL('..', import.meta.url));
const executable = vscodeExecutable();
if (/[\\/]/.test(executable)) await access(executable);
const dir = await mkdtemp(join(tmpdir(), 'sheets-host-test-'));
const profile = join(dir, 'user-data');
const resultPath = join(dir, 'result.json');
const startupFile = join(dir, 'startup.xlsx');
await copyFile(join(root, 'packages/core/test/fixtures/excel-features.xlsx'), startupFile);
await mkdir(join(profile, 'User'), { recursive: true });
await writeFile(join(profile, 'User/settings.json'), JSON.stringify({
  'workbench.startupEditor': 'none', 'update.mode': 'none', 'telemetry.telemetryLevel': 'off',
  'sheets.checkForUpdates': false,
}));
console.log(`Testing the VS Code tool registry and custom editor in an isolated profile: ${dir}`);
const child = spawn(executable, [
  '--user-data-dir', profile, '--extensions-dir', join(dir, 'extensions'),
  '--extensionDevelopmentPath', process.env.SHEETS_EXTENSION_PATH ?? root, '--extensionTestsPath', join(root, 'test/host/index.cjs'),
  '--disable-extensions', '--disable-gpu', '--disable-workspace-trust',
  '--skip-welcome', '--skip-release-notes', '--new-window', '--wait',
  startupFile,
], { stdio: 'inherit', env: vscodeEnvironment({ SHEETS_HOST_RESULT: resultPath, SHEETS_HOST_WORKDIR: dir }) });
let spawnError;
child.on('error', (error) => { spawnError = error; });
const closed = new Promise((resolve) => child.once('exit', resolve));
const deadline = Date.now() + 180_000;
try {
  // On macOS the CLI can remain waiting after tests finish. The atomic report,
  // rather than the GUI process lifetime, determines the test outcome.
  let result;
  while (!result) {
    if (spawnError) throw spawnError;
    const data = await readFile(resultPath, 'utf8').catch((error) => { if (error.code === 'ENOENT') return null; throw error; });
    if (data) result = JSON.parse(data);
    else {
      // With a startup file, the macOS/Linux CLI's --wait ends when that file
      // closes, while the GUI host continues testing other documents. A clean
      // CLI exit is not proof that the host has finished; await its report.
      if (child.exitCode !== null && child.exitCode !== 0) {
        throw new Error(`VS Code exited before reporting results (code ${child.exitCode}). Logs: ${dir}`);
      }
      if (Date.now() >= deadline) throw new Error(`VS Code test timed out. Logs: ${dir}`);
      await new Promise((resolve) => setTimeout(resolve, 250));
    }
  }
  if (result.status !== 'passed') throw new Error(result.error);
  console.log(`Host tests passed: ${result.checks.join('; ')}.`);
  await Promise.race([closed, new Promise((resolve) => setTimeout(resolve, 2000))]);
  console.log(`Host test report and fixtures: ${dir}`);
} finally {
  if (child.exitCode === null) child.kill();
}
