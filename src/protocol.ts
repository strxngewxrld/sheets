/** Messages between the extension host and the spreadsheet webview. */

export interface InitMessage {
  type: 'init';
  /** Base64 file bytes; typed arrays are not preserved by every VS Code transport. */
  bytes: string;
  fileName: string;
  language: string;
  readOnly: boolean;
  /** A hot-exit snapshot containing the model and the original XLSX package. */
  backup?: boolean;
}

/** Editor commands bound to VS Code keybindings (see package.json). */
export type ViewCommand = 'find' | 'replace' | 'hideRows' | 'unhideRows' | 'hideCols' | 'unhideCols';

export type HostToWebview =
  | InitMessage
  | { type: 'undo' }
  | { type: 'redo' }
  | { type: 'command'; command: ViewCommand }
  | { type: 'request'; id: number; method: 'getFileData' | 'getBackupData' }
  | { type: 'request'; id: number; method: 'executeTool'; params: { name: string; input: unknown } };

export type WebviewToHost =
  | { type: 'ready' }
  | { type: 'initialized' }
  | { type: 'edit'; label: string }
  | { type: 'response'; id: number; result?: unknown; error?: string }
  | { type: 'error'; message: string };
