import { describe, expect, it, vi } from 'vitest';

vi.mock('vscode', () => ({
  Uri: {
    file: (path: string) => ({ kind: 'file', path }),
    parse: (value: string) => ({ kind: 'uri', value }),
    joinPath: (folder: unknown, path: string) => ({ kind: 'relative', folder, path }),
  },
}));
import { resolveFileUri } from '../src/file-uri.js';

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
