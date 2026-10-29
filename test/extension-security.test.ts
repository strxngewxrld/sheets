import { beforeEach, describe, expect, it, vi } from 'vitest';

const host = vi.hoisted(() => ({
  documents: new Map<string, unknown>(),
  stat: vi.fn(), readFile: vi.fn(), writeFile: vi.fn(), request: vi.fn(),
  writable: true, trusted: true,
}));

vi.mock('vscode', () => ({
  FileType: { SymbolicLink: 64 },
  ExtensionMode: { Test: 3 },
  commands: { registerCommand: () => ({ dispose() {} }) },
  workspace: {
    get isTrusted() { return host.trusted; },
    workspaceFolders: [{ uri: { scheme: 'vscode-remote', authority: 'ssh-remote+host', path: '/workspace' } }],
    getConfiguration: () => ({ get: () => true }),
    fs: { stat: host.stat, readFile: host.readFile, writeFile: host.writeFile, isWritableFileSystem: () => host.writable },
  },
}));
vi.mock('../src/provider.js', () => ({
  SpreadsheetEditorProvider: { register: () => ({ documents: host.documents }) },
}));
vi.mock('../src/copilot.js', () => ({ registerCopilotTools() {} }));
vi.mock('../src/startup-editors.js', () => ({ reopenStartupSpreadsheets: async () => {} }));

import { activate } from '../src/extension.js';
import type { ExtensionContext, Uri } from 'vscode';

const uri = (path: string): Uri => ({
  scheme: 'vscode-remote', authority: 'ssh-remote+host', path, fsPath: path,
  with: (change: { path: string }) => uri(change.path),
  toString: () => `vscode-remote://ssh-remote+host${path}`,
}) as Uri;
const file = uri('/workspace/linked/data.csv');
const write = { start: 'A1', values: [['changed']] };

beforeEach(() => {
  vi.clearAllMocks();
  host.documents.clear();
  host.writable = true;
  host.trusted = true;
  host.stat.mockResolvedValue({ type: 1 });
  host.readFile.mockResolvedValue(new TextEncoder().encode('original\n'));
  host.writeFile.mockResolvedValue(undefined);
  host.request.mockResolvedValue({ ok: true });
});

const api = () => activate({ subscriptions: [], extensionMode: 3 } as unknown as ExtensionContext);

describe('extension agent security boundary', () => {
  it('rejects a remote parent symlink before reading or writing any bytes', async () => {
    host.stat.mockImplementation(async (part: Uri) => ({ type: part.path.endsWith('/linked') ? 64 : 1 }));
    expect(await api().executeTool('write_range', write, file)).toMatchObject({ ok: false, error: expect.stringMatching(/symbolic link/) });
    expect(host.readFile).not.toHaveBeenCalled();
    expect(host.writeFile).not.toHaveBeenCalled();
  });

  it('fails closed when a provider cannot verify a directory', async () => {
    host.stat.mockRejectedValue(new Error('Permission denied'));
    expect(await api().executeTool('read_range', { range: 'A1' }, file)).toMatchObject({ ok: false });
    expect(host.readFile).not.toHaveBeenCalled();
  });

  it('rejects writes to open read-only files before contacting their webview', async () => {
    host.writable = false;
    host.documents.set(file.toString(), { panel: {}, request: host.request });
    expect(await api().executeTool('write_range', write, file)).toMatchObject({ ok: false, error: expect.stringMatching(/read-only/) });
    expect(host.request).not.toHaveBeenCalled();
    expect(await api().executeTool('read_range', { range: 'A1' }, file)).toMatchObject({ ok: true });
  });

  it('retains workspace trust restrictions on unopened files', async () => {
    host.trusted = false;
    expect(await api().executeTool('write_range', write, file)).toMatchObject({ ok: false, error: expect.stringMatching(/Trust this workspace/) });
    expect(host.writeFile).not.toHaveBeenCalled();
  });

  it('allows writes to regular files after verifying their ancestors', async () => {
    expect(await api().executeTool('write_range', write, file)).toMatchObject({ ok: true });
    expect(host.stat.mock.calls.map(([part]) => part.path)).toEqual(['/workspace/linked/data.csv', '/workspace/linked', '/workspace']);
    expect(host.writeFile).toHaveBeenCalledTimes(1);
  });
});
