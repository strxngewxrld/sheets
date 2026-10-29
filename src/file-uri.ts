import { lstat, realpath } from 'node:fs/promises';
import { basename, dirname, isAbsolute, join, posix, relative, sep } from 'node:path';
import * as vscode from 'vscode';

/** Resolve native paths before URI parsing, preserving URI escaping and authorities. */
export function resolveFileUri(file: string, folder?: vscode.Uri): vscode.Uri {
  if (/^[a-z]:[\\/]/i.test(file) || file.startsWith('\\\\')) return vscode.Uri.file(file);
  if (/^[a-z][a-z0-9+.-]*:\/\//i.test(file)) return vscode.Uri.parse(file);
  if (file.startsWith('/')) return vscode.Uri.file(file);
  // joinPath expects URI separators, including for remote Windows workspaces.
  return folder ? vscode.Uri.joinPath(folder, file.replace(/\\/g, '/')) : vscode.Uri.file(file);
}

const SPREADSHEET_FILE = /\.(xlsx|xlsm|csv|tsv)$/i;

type UriParts = Pick<vscode.Uri, 'scheme' | 'authority' | 'path'>;

/** True when `uri` lies in `folder` (same scheme and authority, ".." segments resolved). */
export function isInsideFolder(uri: UriParts, folder: UriParts, ignoreCase = false): boolean {
  const fold = (s: string) => (ignoreCase ? s.toLowerCase() : s);
  if (uri.scheme !== folder.scheme || uri.authority.toLowerCase() !== folder.authority.toLowerCase()) return false;
  const base = fold(posix.normalize(folder.path).replace(/\/+$/, ''));
  const path = fold(posix.normalize(uri.path));
  return path.startsWith(`${base}/`);
}

/**
 * Why an agent may not reach `uri` without an open Sheets editor, or undefined.
 * Only spreadsheets in the workspace are reachable, so a prompt-injected agent
 * cannot read or overwrite arbitrary files of the user.
 */
export function headlessAccessError(uri: UriParts, folders: readonly UriParts[], ignoreCase = false): string | undefined {
  if (!SPREADSHEET_FILE.test(uri.path)) return 'Only .xlsx, .xlsm, .csv and .tsv files are supported.';
  if (!folders.some((folder) => isInsideFolder(uri, folder, ignoreCase))) {
    return 'Only spreadsheets inside the workspace can be used without opening them. Open the file in the Sheets editor first.';
  }
  return undefined;
}

/** Providers cannot realpath remote files. Check every component and fail closed on I/O errors. */
export async function hasProviderSymlink(
  uri: vscode.Uri,
  folders: readonly UriParts[],
  stat: (uri: vscode.Uri) => Thenable<{ type: number }>,
): Promise<boolean> {
  // Windows providers can interpret backslashes as separators even though URI paths use '/'.
  if (uri.path.includes('\\')) return true;
  const folder = folders.filter((root) => isInsideFolder(uri, root)).sort((a, b) => b.path.length - a.path.length)[0];
  if (!folder) return true;
  const root = posix.normalize(folder.path).replace(/\/+$/, '') || '/';
  let path = posix.normalize(uri.path);
  for (;;) {
    try {
      if ((await stat(uri.with({ path }))).type & vscode.FileType.SymbolicLink) return true;
    } catch (error) {
      // New spreadsheets may not exist yet; permission and provider errors are not proof of safety.
      if ((error as { code?: string }).code !== 'FileNotFound') throw error;
    }
    if (path === root) return false;
    const parent = posix.dirname(path);
    if (parent === path) return true;
    path = parent;
  }
}

/** True when the native `path` resolves (through symbolic links) to a location outside every folder. */
export async function escapesFolders(path: string, folders: readonly string[]): Promise<boolean> {
  const real = await realPathOfNearest(path);
  if (real === undefined) return true;
  const roots = await Promise.all(folders.map((folder) => realpath(folder).catch(() => folder)));
  return !roots.some((root) => {
    const rel = relative(root, real);
    return rel !== '' && rel !== '..' && !rel.startsWith(`..${sep}`) && !isAbsolute(rel);
  });
}

/** Real path of `path`, or of its nearest existing ancestor for a file yet to be created; undefined for a dangling link. */
async function realPathOfNearest(path: string): Promise<string | undefined> {
  try {
    return await realpath(path);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error;
    // writeFile would create the target of a dangling symlink.
    if (await lstat(path).then(() => true, () => false)) return undefined;
    const parent = dirname(path);
    if (parent === path) return path;
    const real = await realPathOfNearest(parent);
    return real === undefined ? undefined : join(real, basename(path));
  }
}
