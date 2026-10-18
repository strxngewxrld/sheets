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
export type ViewCommand = 'find' | 'replace' | 'hideRows' | 'unhideRows' | 'hideCols' | 'unhideCols' | 'goTo' | 'zoomIn' | 'zoomOut' | 'zoomReset';

export type HostToWebview =
  | InitMessage
  | { type: 'undo' }
  | { type: 'redo' }
  | { type: 'command'; command: ViewCommand }
  | { type: 'request'; id: number; method: 'getBackupData' }
  /** `fileName` is the save target; its extension selects the written format ("Save As"). */
  | { type: 'request'; id: number; method: 'getFileData'; params?: { fileName: string } }
  | { type: 'request'; id: number; method: 'executeTool'; params: { name: string; input: unknown } }
  | { type: 'request'; id: number; method: 'exportCsv'; params: { delimiter: string } };

/** Result of an exportCsv request: the active sheet as CSV. */
export interface CsvExport {
  sheet: string;
  /** Base64 file bytes. */
  bytes: string;
}

export type WebviewToHost =
  | { type: 'ready' }
  | { type: 'initialized' }
  | { type: 'edit'; label: string }
  | { type: 'openLink'; url: string }
  | { type: 'response'; id: number; result?: unknown; error?: string }
  | { type: 'error'; message: string }
  /** A non-blocking notice for the user, e.g. content a CSV save does not keep. */
  | { type: 'warning'; message: string };
