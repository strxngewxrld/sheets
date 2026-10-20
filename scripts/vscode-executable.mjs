import { execFileSync } from 'node:child_process';
import { win32 } from 'node:path';

/** VS Code terminals and agents can inherit Electron's Node-only mode. */
export function vscodeEnvironment(extra = {}, inherited = process.env) {
  const env = { ...inherited, ...extra };
  for (const key of Object.keys(env)) {
    if (['ELECTRON_RUN_AS_NODE', 'VSCODE_DEV'].includes(key.toUpperCase())) delete env[key];
  }
  return env;
}

/** Windows batch launchers cannot be spawned directly; use the adjacent GUI binary. */
export function vscodeExecutable(platform = process.platform, override = process.env.VSCODE_EXECUTABLE, locate = () =>
  execFileSync('where.exe', ['code.cmd'], { encoding: 'utf8' }).trim().split(/\r?\n/)[0]) {
  if (platform === 'win32') {
    const cli = override ?? locate();
    if (/\.(cmd|bat)$/i.test(cli)) return win32.resolve(win32.dirname(cli), '..', 'Code.exe');
    return cli;
  }
  return override ?? (platform === 'darwin'
    ? '/Applications/Visual Studio Code.app/Contents/Resources/app/bin/code' : 'code');
}
