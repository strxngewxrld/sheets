import { describe, expect, it } from 'vitest';
import { vscodeExecutable } from '../scripts/vscode-executable.mjs';

describe('host test launcher', () => {
  it('resolves the Windows CLI without a shell, including spaces', () => {
    expect(vscodeExecutable('win32', undefined, () => 'C:\\Program Files\\Microsoft VS Code\\bin\\code.cmd'))
      .toBe('C:\\Program Files\\Microsoft VS Code\\Code.exe');
  });

  it('accepts explicit Windows CLI and executable paths', () => {
    expect(vscodeExecutable('win32', 'D:\\VS Code\\bin\\code.cmd')).toBe('D:\\VS Code\\Code.exe');
    expect(vscodeExecutable('win32', 'D:\\VS Code\\Code.exe')).toBe('D:\\VS Code\\Code.exe');
  });

  it('uses the native launcher on macOS and Linux and respects overrides', () => {
    expect(vscodeExecutable('darwin', undefined)).toContain('Visual Studio Code.app');
    expect(vscodeExecutable('linux', undefined)).toBe('code');
    expect(vscodeExecutable('linux', '/opt/code/bin/code')).toBe('/opt/code/bin/code');
  });
});
