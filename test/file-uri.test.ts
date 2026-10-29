import { mkdir, mkdtemp, rm, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it, vi } from 'vitest';

vi.mock('vscode', () => ({
  FileType: { SymbolicLink: 64 },
  Uri: {
    file: (path: string) => ({ kind: 'file', path }),
    parse: (value: string) => ({ kind: 'uri', value }),
    joinPath: (folder: unknown, path: string) => ({ kind: 'relative', folder, path }),
  },
}));
import { escapesFolders, hasProviderSymlink, headlessAccessError, isInsideFolder, resolveFileUri } from '../src/file-uri.js';

describe('spreadsheet tool file paths', () => {
  const folder = { scheme: 'vscode-remote', path: '/workspace' } as import('vscode').Uri;

  it.each(['C:\\Users\\Name\\Sales #1.xlsx', 'D:/Data/ä.xlsx', '\\\\server\\share\\Sales.xlsx', '/home/user/Sales #1.xlsx'])(
    'preserves an absolute native path: %s', (path) => {
      expect(resolveFileUri(path, folder)).toEqual({ kind: 'file', path });
    });

  it.each(['file:///C:/Data/Sales%20%231.xlsx', 'file://server/share/Sales.xlsx', 'vscode-remote://ssh-remote+host/home/user/Sales.xlsx'])(
    'preserves URI escaping and authority: %s', (value) => {
      expect(resolveFileUri(value, folder)).toEqual({ kind: 'uri', value });
    });

  it.each(['data\\Sales #1.xlsx', 'data/Sales #1.xlsx'])(
    'resolves relative paths in the workspace: %s', (path) => {
      expect(resolveFileUri(path, folder)).toEqual({ kind: 'relative', folder, path: 'data/Sales #1.xlsx' });
    });

  it('keeps native path resolution when no workspace is open', () => {
    expect(resolveFileUri('Sales.xlsx')).toEqual({ kind: 'file', path: 'Sales.xlsx' });
  });
});

describe('agent access without an open editor', () => {
  const folder = { scheme: 'file', authority: '', path: '/work/project' };
  const uri = (path: string, scheme = 'file', authority = '') => ({ scheme, authority, path });

  it('allows spreadsheets inside a workspace folder', () => {
    expect(headlessAccessError(uri('/work/project/data/Sales.xlsx'), [folder])).toBeUndefined();
    expect(headlessAccessError(uri('/work/project/a.CSV'), [folder])).toBeUndefined();
  });

  it.each(['/work/project/../secret/a.xlsx', '/work/project-other/a.xlsx', '/home/user/a.csv', '/work/project.xlsx'])(
    'refuses a path outside the workspace: %s', (path) => {
      expect(headlessAccessError(uri(path), [folder])).toMatch(/inside the workspace/);
    });

  it.each(['/work/project/.bashrc', '/work/project/notes.txt', '/work/project/.vscode/tasks.json'])(
    'refuses files that are not spreadsheets: %s', (path) => {
      expect(headlessAccessError(uri(path), [folder])).toMatch(/Only \.xlsx/);
    });

  it('requires the same scheme and authority', () => {
    expect(isInsideFolder(uri('/work/project/a.xlsx', 'vscode-remote', 'ssh-remote+host'), folder)).toBe(false);
    const remote = uri('/work/project', 'vscode-remote', 'ssh-remote+Host');
    expect(isInsideFolder(uri('/work/project/a.xlsx', 'vscode-remote', 'ssh-remote+host'), remote)).toBe(true);
  });

  it('compares case-insensitively only when asked', () => {
    expect(isInsideFolder(uri('/Work/Project/a.xlsx'), folder)).toBe(false);
    expect(isInsideFolder(uri('/Work/Project/a.xlsx'), folder, true)).toBe(true);
  });

  it('detects symbolic links leading outside the workspace', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'sheets-links-'));
    try {
      const root = join(dir, 'root');
      await mkdir(join(root, 'sub'), { recursive: true });
      await mkdir(join(dir, 'outside'));
      await writeFile(join(dir, 'outside', 'rc'), 'x');
      await writeFile(join(root, 'sub', 'ok.csv'), 'x');
      await symlink(join(dir, 'outside', 'rc'), join(root, 'file.csv'));
      await symlink(join(dir, 'outside'), join(root, 'dir'));
      await symlink(join(dir, 'outside', 'missing'), join(root, 'dangling.csv'));
      expect(await escapesFolders(join(root, 'sub', 'ok.csv'), [root])).toBe(false);
      expect(await escapesFolders(join(root, 'new.xlsx'), [root])).toBe(false);
      expect(await escapesFolders(join(root, 'file.csv'), [root])).toBe(true);
      expect(await escapesFolders(join(root, 'dir', 'new.xlsx'), [root])).toBe(true);
      expect(await escapesFolders(join(root, 'dangling.csv'), [root])).toBe(true);
    } finally {
      await rm(dir, { recursive: true, force: true });
    }
  });
});

describe('remote provider access', () => {
  const folder = { scheme: 'vscode-remote', authority: 'ssh-remote+host', path: '/work/project' };
  const uri = (path: string): import('vscode').Uri => ({ ...folder, path, with: (change: { path: string }) => uri(change.path) }) as import('vscode').Uri;

  it('rejects symlinked parent directories even when the file itself is regular', async () => {
    const stat = vi.fn(async (part: import('vscode').Uri) => ({ type: part.path === '/work/project/linked' ? 64 : 1 }));
    expect(await hasProviderSymlink(uri('/work/project/linked/data.csv'), [folder], stat)).toBe(true);
    expect(stat.mock.calls.map(([part]) => part.path)).toContain('/work/project/linked');
  });

  it('allows missing new files but still checks their ancestors', async () => {
    const stat = vi.fn(async (part: import('vscode').Uri) => {
      if (part.path.endsWith('.csv')) throw { code: 'FileNotFound' };
      return { type: 2 };
    });
    expect(await hasProviderSymlink(uri('/work/project/data.csv'), [folder], stat)).toBe(false);
    expect(stat.mock.calls.map(([part]) => part.path)).toEqual(['/work/project/data.csv', '/work/project']);
  });

  it('does not treat provider errors or remote case differences as safe', async () => {
    await expect(hasProviderSymlink(uri('/work/project/data.csv'), [folder], async () => { throw new Error('Permission denied'); })).rejects.toThrow('Permission denied');
    const stat = vi.fn(async () => ({ type: 1 }));
    expect(await hasProviderSymlink(uri('/Work/Project/data.csv'), [folder], stat)).toBe(true);
    expect(stat).not.toHaveBeenCalled();
  });

  it('rejects ambiguous Windows separators before querying the provider', async () => {
    const stat = vi.fn(async () => ({ type: 1 }));
    expect(await hasProviderSymlink(uri('/work/project/sub/..\\..\\secret.csv'), [folder], stat)).toBe(true);
    expect(stat).not.toHaveBeenCalled();
  });
});
