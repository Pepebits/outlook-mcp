import { log } from '../log.js';
import { GraphError } from './errors.js';

export interface RequestOptions {
  query?: Record<string, string | number | boolean | undefined>;
  body?: unknown;
  headers?: Record<string, string>;
  /** Return the raw Response (caller handles body); errors are still thrown. */
  raw?: boolean;
}

export interface GraphClientOptions {
  sleep?: (ms: number) => Promise<void>;
  fetch?: typeof fetch;
  maxRetries?: number;
  /** Per-attempt timeout in milliseconds (default 60 s). */
  timeoutMs?: number;
  /** Reject every non-GET request before it is sent (read-only mode, ADR 0001). */
  readOnly?: boolean;
}

/** Throttling is always safe to retry: Graph did not process the request. */
const THROTTLED = 429;
/** Gateway errors may hide a request that was processed, so only idempotent methods retry them. */
const RETRY_IF_IDEMPOTENT = new Set([500, 502, 503, 504]);
const IDEMPOTENT = new Set(['GET', 'HEAD', 'PUT', 'DELETE', 'PATCH']);
export const MAX_RETRY_AFTER_MS = 60_000;
/** Graph accepts at most 20 requests per JSON batch. */
export const BATCH_LIMIT = 20;

export class GraphClient {
  private readonly sleep: (ms: number) => Promise<void>;
  private readonly fetchFn: typeof fetch;
  private readonly maxRetries: number;
  private readonly timeoutMs: number;
  private readonly readOnly: boolean;
  private readonly origin: string;
  private readonly basePath: string;

  constructor(
    private readonly getToken: () => Promise<string>,
    private readonly base: string,
    opts: GraphClientOptions = {},
  ) {
    this.sleep = opts.sleep ?? ((ms) => new Promise((r) => setTimeout(r, ms)));
    this.fetchFn = opts.fetch ?? ((...a) => fetch(...a));
    this.maxRetries = opts.maxRetries ?? 4;
    this.timeoutMs = opts.timeoutMs ?? 60_000;
    this.readOnly = opts.readOnly ?? false;
    const b = new URL(base);
    this.origin = b.origin;
    this.basePath = b.pathname.replace(/\/+$/, '');
  }

  /**
   * Relative paths are placed under the base URL. Absolute URLs (nextLinks) must stay on the
   * same origin and API version, otherwise the bearer token would be sent to another host.
   */
  private buildUrl(pathOrUrl: string, query?: RequestOptions['query']): string {
    const isAbsolute = /^[a-z][a-z0-9+.-]*:/i.test(pathOrUrl);
    const url = new URL(isAbsolute ? pathOrUrl : `${this.base}${pathOrUrl.startsWith('/') ? '' : '/'}${pathOrUrl}`);
    if (isAbsolute && (url.origin !== this.origin || !(url.pathname === this.basePath || url.pathname.startsWith(`${this.basePath}/`)))) {
      throw new Error(`Refusing to send the access token to ${url.origin}${url.pathname}: only ${this.origin}${this.basePath} is allowed. Pass the nextLink exactly as returned.`);
    }
    if (query) {
      for (const [k, v] of Object.entries(query)) if (v !== undefined) url.searchParams.set(k, String(v));
    }
    return url.toString();
  }

  private backoffMs(attempt: number, res?: Response): number {
    const ra = Number(res?.headers.get('Retry-After'));
    if (Number.isFinite(ra) && ra > 0) return Math.min(ra * 1000, MAX_RETRY_AFTER_MS);
    return Math.min(30_000, 500 * 2 ** attempt) + Math.floor(Math.random() * 250);
  }

  async request<T = unknown>(method: string, pathOrUrl: string, opts: RequestOptions = {}): Promise<T> {
    return this.send<T>(method.toUpperCase(), pathOrUrl, opts, false);
  }

  /**
   * Runs GET requests (paths relative to the API version) through JSON batching, 20 per $batch call.
   * A $batch of GETs only reads, so it is allowed in read-only mode and retried like a GET. Returns one
   * entry per path, in order; failed sub-requests come back with their status for the caller to handle.
   */
  async batchGet<T = unknown>(paths: string[]): Promise<{ status: number; body?: T }[]> {
    const out: { status: number; body?: T }[] = [];
    for (let start = 0; start < paths.length; start += BATCH_LIMIT) {
      const chunk = paths.slice(start, start + BATCH_LIMIT);
      const requests = chunk.map((url, i) => ({ id: String(i), method: 'GET', url: url.startsWith('/') ? url : `/${url}` }));
      const res = await this.send<{ responses?: { id: string; status: number; body?: T }[] }>('POST', '/$batch', { body: { requests } }, true);
      const byId = new Map((res?.responses ?? []).map((r) => [r.id, r]));
      for (let i = 0; i < chunk.length; i++) {
        const r = byId.get(String(i));
        out.push(r ? { status: r.status, body: r.body } : { status: 0 });
      }
    }
    return out;
  }

  private async send<T>(m: string, pathOrUrl: string, opts: RequestOptions, readsOnly: boolean): Promise<T> {
    if (this.readOnly && m !== 'GET' && !readsOnly) {
      throw new Error(`Read-only mode: refusing ${m} ${pathOrUrl.split('?')[0]}.`);
    }
    const url = this.buildUrl(pathOrUrl, opts.query);
    const idempotent = readsOnly || IDEMPOTENT.has(m);
    for (let attempt = 0; ; attempt++) {
      const headers: Record<string, string> = {
        Authorization: `Bearer ${await this.getToken()}`,
        Accept: 'application/json',
        ...opts.headers,
      };
      let body: string | undefined;
      if (opts.body !== undefined) {
        headers['Content-Type'] = 'application/json';
        body = JSON.stringify(opts.body);
      }
      let res: Response;
      try {
        res = await this.fetchFn(url, { method: m, headers, body, signal: AbortSignal.timeout(this.timeoutMs) });
      } catch (err) {
        // A network failure on a non-idempotent request may still have reached Graph: never resend it.
        if (idempotent && attempt < this.maxRetries) {
          const wait = this.backoffMs(attempt);
          log.warn(`Graph request failed (${describeNetworkError(err)}); retrying in ${wait}ms (attempt ${attempt + 1}/${this.maxRetries})`);
          await this.sleep(wait);
          continue;
        }
        throw new NetworkError(err, m);
      }
      const retryable = res.status === THROTTLED || (idempotent && RETRY_IF_IDEMPOTENT.has(res.status));
      if (retryable && attempt < this.maxRetries) {
        const wait = this.backoffMs(attempt, res);
        log.warn(`Graph returned ${res.status}; retrying in ${wait}ms (attempt ${attempt + 1}/${this.maxRetries})`);
        await this.sleep(wait);
        continue;
      }
      if (!res.ok) throw await GraphError.from(res, pathOrUrl);
      if (opts.raw) return res as unknown as T;
      if (res.status === 204) return undefined as T;
      const text = await res.text();
      return (text ? JSON.parse(text) : undefined) as T;
    }
  }

  get<T = unknown>(path: string, opts?: RequestOptions) {
    return this.request<T>('GET', path, opts);
  }
  post<T = unknown>(path: string, body?: unknown, opts?: RequestOptions) {
    return this.request<T>('POST', path, { ...opts, body });
  }
  patch<T = unknown>(path: string, body?: unknown, opts?: RequestOptions) {
    return this.request<T>('PATCH', path, { ...opts, body });
  }
  delete(path: string, opts?: RequestOptions) {
    return this.request<void>('DELETE', path, opts);
  }
  /** GET returning the raw Response (for binary content such as attachments). */
  getStream(path: string, opts?: RequestOptions): Promise<Response> {
    return this.request<Response>('GET', path, { ...opts, raw: true, headers: { Accept: '*/*', ...opts?.headers } });
  }
}

function describeNetworkError(err: unknown): string {
  if (err instanceof Error && (err.name === 'TimeoutError' || err.name === 'AbortError')) return 'timed out';
  const cause = (err as { cause?: { code?: string; message?: string } })?.cause;
  return cause?.code ?? cause?.message ?? (err instanceof Error ? err.message : String(err));
}

/** Microsoft Graph could not be reached (offline, DNS, reset, timeout). */
export class NetworkError extends Error {
  constructor(cause: unknown, method: string) {
    const reason = describeNetworkError(cause);
    const hint = method === 'GET' ? '' : ' The request may or may not have been processed; check the mailbox before retrying.';
    super(`Could not reach Microsoft Graph (${reason}).${hint}`);
    this.name = 'NetworkError';
  }
}
