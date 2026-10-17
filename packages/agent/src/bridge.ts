import type { ChangeEvent } from '@sheets/core';
import type { SpreadsheetAgent, ToolResult } from './agent.js';
import type { ToolDefinition } from './tools.js';

/**
 * JSON-RPC 2.0 bridge so a spreadsheet running in an iframe, webview or
 * another window can be controlled remotely (e.g. by an agent running in
 * the host page or a VS Code extension).
 *
 * Methods: "tools/list", "tools/call" { name, arguments }, "workbook/snapshot".
 * Notifications from the server: "workbook/changed" { origin, structural, ops }.
 */

export interface MessagePort {
  postMessage(message: unknown): void;
  addEventListener(type: 'message', listener: (e: { data: unknown }) => void): void;
  removeEventListener(type: 'message', listener: (e: { data: unknown }) => void): void;
}

interface RpcRequest {
  jsonrpc: '2.0';
  id?: number | string;
  method: string;
  params?: Record<string, unknown>;
  channel?: string;
}

interface RpcResponse {
  jsonrpc: '2.0';
  id: number | string;
  result?: unknown;
  error?: { code: number; message: string };
  channel?: string;
}

const CHANNEL = 'sheets-agent';

export interface BridgeServerOptions {
  /** Forward workbook change notifications to the client (default true). */
  notifyChanges?: boolean;
  /** Accept only messages for which this returns true (e.g. origin checks). */
  accept?: (data: unknown) => boolean;
}

/** Serves an agent over a message port. Returns a dispose function. */
export function serveAgent(agent: SpreadsheetAgent, port: MessagePort, options: BridgeServerOptions = {}): () => void {
  const onMessage = (e: { data: unknown }) => {
    const msg = e.data as RpcRequest;
    if (!msg || msg.jsonrpc !== '2.0' || msg.channel !== CHANNEL || typeof msg.method !== 'string' || msg.id === undefined) return;
    if (options.accept && !options.accept(e.data)) return;
    const reply = (body: Omit<RpcResponse, 'jsonrpc' | 'id' | 'channel'>) => port.postMessage({ jsonrpc: '2.0', id: msg.id!, channel: CHANNEL, ...body });
    switch (msg.method) {
      case 'tools/list':
        return reply({ result: { tools: agent.tools } });
      case 'tools/call': {
        const name = String(msg.params?.name ?? '');
        return reply({ result: agent.execute(name, msg.params?.arguments ?? {}) });
      }
      case 'workbook/snapshot':
        return reply({ result: agent.workbook.toJSON() });
      default:
        return reply({ error: { code: -32601, message: `Method not found: ${msg.method}` } });
    }
  };
  port.addEventListener('message', onMessage);
  const unsubscribe =
    options.notifyChanges === false
      ? () => {}
      : agent.onChange((ev: ChangeEvent) =>
          port.postMessage({ jsonrpc: '2.0', channel: CHANNEL, method: 'workbook/changed', params: { origin: ev.origin, structural: ev.structural, ops: ev.ops } }),
        );
  return () => {
    port.removeEventListener('message', onMessage);
    unsubscribe();
  };
}

export interface AgentClient {
  listTools(): Promise<ToolDefinition[]>;
  call(name: string, args?: Record<string, unknown>): Promise<ToolResult>;
  snapshot(): Promise<unknown>;
  onChange(listener: (e: { origin: string; structural: boolean; ops: unknown[] }) => void): () => void;
  dispose(): void;
}

/** Connects to an agent served with `serveAgent` on the other side of a port. */
export function connectAgent(port: MessagePort, timeoutMs = 30_000): AgentClient {
  let nextId = 1;
  const pending = new Map<number | string, { resolve: (v: unknown) => void; reject: (e: Error) => void; timer: ReturnType<typeof setTimeout> }>();
  const listeners = new Set<(e: { origin: string; structural: boolean; ops: unknown[] }) => void>();
  const onMessage = (e: { data: unknown }) => {
    const msg = e.data as RpcResponse & { method?: string; params?: { origin: string; structural: boolean; ops: unknown[] } };
    if (!msg || msg.jsonrpc !== '2.0' || msg.channel !== CHANNEL) return;
    if (msg.method === 'workbook/changed' && msg.params) {
      listeners.forEach((l) => l(msg.params!));
      return;
    }
    const p = msg.id !== undefined ? pending.get(msg.id) : undefined;
    if (!p) return;
    pending.delete(msg.id);
    clearTimeout(p.timer);
    if (msg.error) p.reject(new Error(msg.error.message));
    else p.resolve(msg.result);
  };
  port.addEventListener('message', onMessage);
  const request = <T>(method: string, params?: Record<string, unknown>) =>
    new Promise<T>((resolve, reject) => {
      const id = nextId++;
      const timer = setTimeout(() => {
        pending.delete(id);
        reject(new Error(`Request ${method} timed out`));
      }, timeoutMs);
      pending.set(id, { resolve: resolve as (v: unknown) => void, reject, timer });
      port.postMessage({ jsonrpc: '2.0', id, method, params, channel: CHANNEL });
    });
  return {
    listTools: async () => (await request<{ tools: ToolDefinition[] }>('tools/list')).tools,
    call: (name, args = {}) => request<ToolResult>('tools/call', { name, arguments: args }),
    snapshot: () => request('workbook/snapshot'),
    onChange: (l) => {
      listeners.add(l);
      return () => listeners.delete(l);
    },
    dispose: () => {
      port.removeEventListener('message', onMessage);
      for (const p of pending.values()) clearTimeout(p.timer);
      pending.clear();
    },
  };
}
