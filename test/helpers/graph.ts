import { vi } from 'vitest';
import type { AuthSession, LoginState } from '../../src/auth/session.js';
import { parseConfig } from '../../src/config.js';
import { GraphClient } from '../../src/graph/client.js';
import { allTools, type ToolContext } from '../../src/tools/index.js';

export interface FakeRequest {
  method: string;
  /** Path (and query) relative to the fake base URL, or the absolute URL for nextLinks. */
  path: string;
  url: URL;
  body: any;
  headers: Record<string, string>;
}

export interface FakeReply {
  status?: number;
  json?: unknown;
}

export type Route = (req: FakeRequest) => FakeReply | undefined | void;

export const BASE = 'https://graph.test/v1.0';

/** A real GraphClient wired to an in-memory fetch. Routes are tried in order; the first to answer wins. */
export function fakeGraph(...routes: Route[]) {
  const calls: FakeRequest[] = [];
  const fetchFn = vi.fn(async (input: any, init: any = {}) => {
    const url = new URL(String(input));
    const req: FakeRequest = {
      method: init.method ?? 'GET',
      path: url.toString().startsWith(BASE) ? url.toString().slice(BASE.length) : url.toString(),
      url,
      body: init.body ? JSON.parse(init.body) : undefined,
      headers: init.headers ?? {},
    };
    calls.push(req);
    // JSON batching: answer each GET sub-request through the same routes.
    if (req.method === 'POST' && req.path === '/$batch') {
      const responses = (req.body.requests as { id: string; method: string; url: string }[]).map((sub) => {
        const subUrl = new URL(BASE + sub.url);
        const subReq: FakeRequest = { method: sub.method, path: sub.url, url: subUrl, body: undefined, headers: {} };
        calls.push(subReq);
        for (const r of routes) {
          const out = r(subReq);
          if (out) return { id: sub.id, status: out.status ?? 200, body: out.json };
        }
        return { id: sub.id, status: 500, body: { error: { code: 'NoRoute', message: `no route for ${sub.method} ${sub.url}` } } };
      });
      return new Response(JSON.stringify({ responses }), { status: 200 });
    }
    for (const r of routes) {
      const out = r(req);
      if (out) {
        const status = out.status ?? 200;
        return new Response(status === 204 || out.json === undefined ? null : JSON.stringify(out.json), { status });
      }
    }
    return new Response(JSON.stringify({ error: { code: 'NoRoute', message: `no route for ${req.method} ${req.path}` } }), { status: 500 });
  });
  const graph = new GraphClient(async () => 'token', BASE, { fetch: fetchFn as any, maxRetries: 0, sleep: async () => {} });
  return { graph, calls, fetchFn };
}

/** Matches a method and path prefix (the query string is ignored). */
export const on =
  (method: string, pathPrefix: string, reply: FakeReply | ((req: FakeRequest) => FakeReply)): Route =>
  (req) =>
    req.method === method && req.path.split('?')[0] === pathPrefix ? (typeof reply === 'function' ? reply(req) : reply) : undefined;

export const graphError = (status: number, code: string, message = code): FakeReply => ({ status, json: { error: { code, message } } });

export function fakeAuth(overrides: Partial<Record<'startLogin' | 'status' | 'logout', any>> = {}) {
  const auth = {
    startLogin: vi.fn(async (_force?: boolean): Promise<LoginState> => ({ state: 'signed_in', username: 'me@x.com' })),
    status: vi.fn(async (): Promise<LoginState> => ({ state: 'signed_in', username: 'me@x.com' })),
    logout: vi.fn(async () => ({ removedAccounts: 1 })),
    ...Object.fromEntries(Object.entries(overrides).map(([k, v]) => [k, vi.fn(v)])),
  };
  return auth as typeof auth & AuthSession;
}

export function makeCtx(graph: GraphClient, env: Record<string, string> = {}, auth: AuthSession = fakeAuth()): ToolContext {
  return { graph, cfg: parseConfig({ ...env }), auth };
}

export function tool(name: string) {
  const t = allTools().find((x) => x.name === name);
  if (!t) throw new Error(`unknown tool ${name}`);
  return t;
}

export const run = (name: string, args: unknown, ctx: ToolContext): Promise<any> => tool(name).handler(args, ctx);
