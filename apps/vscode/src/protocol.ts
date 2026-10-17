/** Messages between the extension host and the spreadsheet webview. */

export interface InitMessage {
  type: 'init';
  bytes: Uint8Array;
  fileName: string;
  language: string;
  readOnly: boolean;
}

export type HostToWebview =
  | InitMessage
  | { type: 'undo' }
  | { type: 'redo' }
  | { type: 'request'; id: number; method: 'getFileData' }
  | { type: 'request'; id: number; method: 'executeTool'; params: { name: string; input: unknown } };

export type WebviewToHost =
  | { type: 'ready' }
  | { type: 'edit'; label: string }
  | { type: 'response'; id: number; result?: unknown; error?: string }
  | { type: 'error'; message: string };
