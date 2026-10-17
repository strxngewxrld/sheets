import { spawn } from 'node:child_process';
import { access, mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('..', import.meta.url));
const executable = process.env.VSCODE_EXECUTABLE ?? (process.platform === 'darwin'
  ? '/Applications/Visual Studio Code.app/Contents/Resources/app/bin/code' : 'code');
if (executable.includes('/')) await access(executable);
const dir = await mkdtemp(join(tmpdir(), 'sheets-host-test-'));
const profile = join(dir, 'user-data');
const resultPath = join(dir, 'result.json');
await mkdir(join(profile, 'User'), { recursive: true });
await writeFile(join(profile, 'User/settings.json'), JSON.stringify({
  'workbench.startupEditor': 'none', 'update.mode': 'none', 'telemetry.telemetryLevel': 'off',
}));
console.log(`Testing the VS Code tool registry and custom editor in an isolated profile: ${dir}`);
const child = spawn(executable, [
  '--user-data-dir', profile, '--extensions-dir', join(dir, 'extensions'),
  '--extensionDevelopmentPath', root, '--extensionTestsPath', join(root, 'test/host/index.cjs'),
  '--disable-extensions', '--disable-gpu', '--disable-workspace-trust',
  '--skip-welcome', '--skip-release-notes', '--new-window', '--wait',
], { stdio: 'inherit', env: { ...process.env, SHEETS_HOST_RESULT: resultPath, SHEETS_HOST_WORKDIR: dir } });
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
      if (Date.now() >= deadline) throw new Error(`VS Code test timed out. Logs: ${dir}`);
      await new Promise((resolve) => setTimeout(resolve, 250));
    }
  }
  if (result.status !== 'passed') throw new Error(result.error);
  console.log(`Host tests passed: ${result.checks.join('; ')}.`);
  await Promise.race([closed, new Promise((resolve) => setTimeout(resolve, 2000))]);
  await rm(dir, { recursive: true, force: true });
} finally {
  if (child.exitCode === null) child.kill();
}
