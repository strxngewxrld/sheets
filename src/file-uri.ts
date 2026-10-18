import * as vscode from 'vscode';

/** Resolve native paths before URI parsing, preserving URI escaping and authorities. */
export function resolveFileUri(file: string, folder?: vscode.Uri): vscode.Uri {
  if (/^[a-z]:[\\/]/i.test(file) || file.startsWith('\\\\')) return vscode.Uri.file(file);
  if (/^[a-z][a-z0-9+.-]*:\/\//i.test(file)) return vscode.Uri.parse(file);
  if (file.startsWith('/')) return vscode.Uri.file(file);
  // joinPath expects URI separators, including for remote Windows workspaces.
  return folder ? vscode.Uri.joinPath(folder, file.replace(/\\/g, '/')) : vscode.Uri.file(file);
}
