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
}

const RETRY_STATUS = new Set([429, 503, 504]);

export class GraphClient {
  private readonly sleep: (ms: number) => Promise<void>;
  private readonly fetchFn: typeof fetch;
  private readonly maxRetries: number;

  constructor(
    private readonly getToken: () => Promise<string>,
    private readonly base: string,
    opts: GraphClientOptions = {},
  ) {
    this.sleep = opts.sleep ?? ((ms) => new Promise((r) => setTimeout(r, ms)));
    this.fetchFn = opts.fetch ?? ((...a) => fetch(...a));
    this.maxRetries = opts.maxRetries ?? 4;
  }

  private buildUrl(pathOrUrl: string, query?: RequestOptions['query']): string {
    const isAbsolute = /^https?:\/\//i.test(pathOrUrl);
    const url = new URL(isAbsolute ? pathOrUrl : `${this.base}${pathOrUrl.startsWith('/') ? '' : '/'}${pathOrUrl}`);
    if (query) {
      for (const [k, v] of Object.entries(query)) if (v !== undefined) url.searchParams.set(k, String(v));
    }
    return url.toString();
  }

  private backoffMs(attempt: number, res: Response): number {
    const ra = Number(res.headers.get('Retry-After'));
    if (Number.isFinite(ra) && ra > 0) return ra * 1000;
    return Math.min(30_000, 500 * 2 ** attempt) + Math.floor(Math.random() * 250);
  }

  async request<T = unknown>(method: string, pathOrUrl: string, opts: RequestOptions = {}): Promise<T> {
    const url = this.buildUrl(pathOrUrl, opts.query);
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
      const res = await this.fetchFn(url, { method, headers, body });
      if (RETRY_STATUS.has(res.status) && attempt < this.maxRetries) {
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
