import type { ICachePlugin, TokenCacheContext } from '@azure/msal-node';
import { execFile, type ExecFileException } from 'node:child_process';
import { createHash, randomBytes } from 'node:crypto';
import fs from 'node:fs';
import fsp from 'node:fs/promises';
import path from 'node:path';
import { log } from '../log.js';
import { FileCachePlugin } from './cache.js';

export const KEYCHAIN_SERVICE = 'outlook-mcp';

export type TokenStoreMode = 'auto' | 'keychain' | 'file';

/** A place the serialized MSAL cache can live. Values are opaque strings. */
export interface SecretBackend {
  readonly name: string;
  get(): Promise<string | undefined>;
  set(value: string): Promise<void>;
  delete(): Promise<void>;
  /** Like get(), plus when the value was last written (epoch ms) if the backend records it. */
  getEntry?(): Promise<{ value: string; updatedAt?: number } | undefined>;
}

export interface ExecResult {
  code: number | null;
  stdout: string;
  stderr: string;
  /** Set when the binary could not be started (e.g. ENOENT). */
  spawnError?: string;
}

export type Exec = (cmd: string, args: string[], input?: string) => Promise<ExecResult>;

/** Runs a binary without a shell; the secret (if any) goes through stdin, never argv. */
export const defaultExec: Exec = (cmd, args, input) =>
  new Promise((resolve) => {
    const child = execFile(cmd, args, { encoding: 'utf8', maxBuffer: 16 * 1024 * 1024 }, (err, stdout, stderr) => {
      const e = err as (ExecFileException & { code?: number | string }) | null;
      if (e && typeof e.code === 'string') return resolve({ code: null, stdout, stderr, spawnError: e.code });
      resolve({ code: e ? ((e.code as number | undefined) ?? 1) : 0, stdout, stderr });
    });
    child.stdin?.on('error', () => {});
    child.stdin?.end(input ?? '');
  });

/** Account name for the keychain entry: the cache path, or a hash of it when it holds characters (`"`, `'`, `\`, newlines) `security -i` cannot quote. */
export function keychainAccount(cachePath: string): string {
  if (/^[^"\\\r\n']+$/.test(cachePath)) return cachePath;
  return `token-cache-${createHash('sha256').update(cachePath).digest('hex').slice(0, 16)}`;
}

// `security -i` reads commands from stdin and truncates lines at about 4 KB, so values are stored base64
// encoded in chunks that stay well under that limit.
const MAC_CHUNK = 3000;
// Upper bound when sweeping legacy (`<account>#<i>`) chunk slots; a token cache needs only a few.
const MAC_LEGACY_SCAN_CAP = 256;
// Consecutive missing legacy slots (past the manifest count) after which the sweep stops.
const MAC_LEGACY_GAP = 4;
const NOT_FOUND = 44; // errSecItemNotFound

interface Manifest {
  /** Generation id; undefined for the legacy `chunks:N` layout. */
  gen?: string;
  n: number;
  /** Epoch ms of the write; undefined for the legacy layout. */
  ts?: number;
}

function parseManifest(raw: string): Manifest {
  const legacy = /^chunks:(\d+)$/.exec(raw);
  if (legacy) return { n: Number(legacy[1]) };
  const m = /^chunks:([0-9a-f]+):(\d+):(\d+)$/.exec(raw);
  if (m) return { gen: m[1], n: Number(m[2]), ts: Number(m[3]) };
  throw new Error('Token cache manifest in the keychain is not recognized.');
}

/**
 * macOS Keychain through the `security` CLI.
 *
 * Secret handling: `security add-generic-password` only accepts the password as an argument (-w/-X), which
 * would expose the whole token cache in `ps` output to other local processes. Instead the command is fed to
 * `security -i` over stdin, so the secret never appears in argv. The cost is a ~4 KB line limit, handled by
 * splitting the (base64) value across several items.
 *
 * Layout: chunks live under `<account>#<gen>#<i>` with a fresh generation id per write, and the manifest
 * `<account>` = `chunks:<gen>:<n>:<epochMs>` is written last, so a concurrent reader sees either the old or
 * the new complete value. The previous generation is deleted afterwards. The older `chunks:N` /
 * `<account>#<i>` layout is still readable and is replaced on the next write.
 */
export class MacKeychainBackend implements SecretBackend {
  readonly name = 'macOS Keychain';
  private readonly account: string;

  constructor(
    cachePath: string,
    private readonly exec: Exec = defaultExec,
  ) {
    this.account = keychainAccount(cachePath);
  }

  /** Value of an item, or undefined when it does not exist (exit 44). Any other failure throws. */
  private async lookup(account: string): Promise<string | undefined> {
    const r = await this.exec('security', ['find-generic-password', '-s', KEYCHAIN_SERVICE, '-a', account, '-w']);
    if (r.spawnError) throw new Error(`security CLI unavailable (${r.spawnError})`);
    if (r.code === 0) return r.stdout.trim();
    if (r.code === NOT_FOUND) return undefined;
    throw new Error(`security find-generic-password failed (exit ${r.code}): ${r.stderr.trim() || 'no details'}`);
  }

  private async add(account: string, value: string): Promise<void> {
    const line = `add-generic-password -U -s ${KEYCHAIN_SERVICE} -a "${account}" -w ${value}\n`;
    const r = await this.exec('security', ['-i'], line);
    if (r.spawnError) throw new Error(`security CLI unavailable (${r.spawnError})`);
    if (r.code !== 0) throw new Error(`security add-generic-password failed (exit ${r.code}): ${r.stderr.trim() || 'no details'}`);
  }

  /** Deletes an item; returns false when it did not exist. Any other failure throws. */
  private async remove(account: string): Promise<boolean> {
    const r = await this.exec('security', ['delete-generic-password', '-s', KEYCHAIN_SERVICE, '-a', account]);
    if (r.spawnError) throw new Error(`security CLI unavailable (${r.spawnError})`);
    if (r.code === 0) return true;
    if (r.code === NOT_FOUND) return false;
    throw new Error(`security delete-generic-password failed (exit ${r.code}): ${r.stderr.trim() || 'no details'}`);
  }

  private chunkAccount(m: Manifest, i: number): string {
    return m.gen ? `${this.account}#${m.gen}#${i}` : `${this.account}#${i}`;
  }

  private async readManifest(): Promise<Manifest | undefined> {
    const raw = await this.lookup(this.account);
    return raw === undefined ? undefined : parseManifest(raw);
  }

  async getEntry(): Promise<{ value: string; updatedAt?: number } | undefined> {
    // A concurrent writer may replace (and delete) the generation we are reading: re-read the manifest once.
    for (let attempt = 0; ; attempt++) {
      const m = await this.readManifest();
      if (!m) return undefined;
      let b64 = '';
      let complete = true;
      for (let i = 0; i < m.n; i++) {
        const part = await this.lookup(this.chunkAccount(m, i));
        if (part === undefined) {
          complete = false;
          break;
        }
        b64 += part;
      }
      if (complete) return { value: Buffer.from(b64, 'base64').toString('utf8'), updatedAt: m.ts };
      if (attempt >= 1) throw new Error('Token cache in the keychain is incomplete.');
    }
  }

  async get(): Promise<string | undefined> {
    return (await this.getEntry())?.value;
  }

  async set(value: string): Promise<void> {
    const b64 = Buffer.from(value, 'utf8').toString('base64');
    const chunks = b64.match(new RegExp(`.{1,${MAC_CHUNK}}`, 'g')) ?? [''];
    let previous: Manifest | undefined;
    try {
      previous = await this.readManifest();
    } catch {
      previous = undefined; // unreadable manifest: it is about to be replaced anyway
    }
    const next: Manifest = { gen: randomBytes(6).toString('hex'), n: chunks.length, ts: Date.now() };
    for (let i = 0; i < chunks.length; i++) await this.add(this.chunkAccount(next, i), chunks[i] || '=');
    await this.add(this.account, `chunks:${next.gen}:${next.n}:${next.ts}`);
    if ((await this.get()) !== value) throw new Error('Could not write the token cache to the macOS Keychain.');
    if (previous) {
      try {
        await this.removeChunks(previous);
      } catch (err) {
        log.warn(`could not remove the previous token cache chunks from the keychain: ${(err as Error).message}`);
      }
    }
  }

  async delete(): Promise<void> {
    let current: Manifest | undefined;
    try {
      current = await this.readManifest();
    } catch {
      current = undefined;
    }
    if (current) await this.removeChunks(current);
    // Sweep legacy `<account>#<i>` slots (orphans of the old layout): stop after a short run of missing slots past the manifest count.
    const legacyCount = current && !current.gen ? current.n : 0;
    let misses = 0;
    for (let i = 0; i < MAC_LEGACY_SCAN_CAP && misses < MAC_LEGACY_GAP; i++) {
      const existed = await this.remove(`${this.account}#${i}`);
      misses = existed || i < legacyCount ? 0 : misses + 1;
    }
    await this.remove(this.account);
  }

  private async removeChunks(m: Manifest): Promise<void> {
    for (let i = 0; i < m.n; i++) await this.remove(this.chunkAccount(m, i));
  }
}

/** Linux Secret Service (GNOME Keyring, KWallet, ...) through `secret-tool`; the secret is passed on stdin. */
export class SecretToolBackend implements SecretBackend {
  readonly name = 'Secret Service (secret-tool)';
  private readonly attrs: string[];

  constructor(
    cachePath: string,
    private readonly exec: Exec = defaultExec,
  ) {
    this.attrs = ['service', KEYCHAIN_SERVICE, 'account', cachePath];
  }

  async get(): Promise<string | undefined> {
    const r = await this.exec('secret-tool', ['lookup', ...this.attrs]);
    if (r.spawnError) throw new Error(`secret-tool unavailable (${r.spawnError})`);
    if (r.code === 0) return r.stdout === '' ? undefined : r.stdout;
    // lookup exits 1 with no output at all when nothing matches; anything else is a real failure.
    if (r.code === 1 && r.stdout === '' && !r.stderr.trim()) return undefined;
    throw new Error(`secret-tool lookup failed (exit ${r.code}): ${r.stderr.trim() || 'no details'}`);
  }

  async set(value: string): Promise<void> {
    const r = await this.exec('secret-tool', ['store', '--label=outlook-mcp token cache', ...this.attrs], value);
    if (r.spawnError) throw new Error(`secret-tool unavailable (${r.spawnError})`);
    if (r.code !== 0) throw new Error(`secret-tool store failed: ${r.stderr.trim() || `exit ${r.code}`}`);
  }

  async delete(): Promise<void> {
    const r = await this.exec('secret-tool', ['clear', ...this.attrs]);
    if (r.spawnError) throw new Error(`secret-tool unavailable (${r.spawnError})`);
    if (r.code === 0 || (r.code === 1 && !r.stdout && !r.stderr.trim())) return;
    throw new Error(`secret-tool clear failed (exit ${r.code}): ${r.stderr.trim() || 'no details'}`);
  }
}

/** True when `name` is an executable file somewhere on PATH. */
export function hasExecutable(name: string, envPath = process.env.PATH ?? ''): boolean {
  for (const dir of envPath.split(path.delimiter).filter(Boolean)) {
    try {
      fs.accessSync(path.join(dir, name), fs.constants.X_OK);
      return true;
    } catch {
      // keep looking
    }
  }
  return false;
}

export interface StoreEnv {
  platform: NodeJS.Platform;
  hasSecretTool: boolean;
}

export const currentStoreEnv = (): StoreEnv => ({ platform: process.platform, hasSecretTool: hasExecutable('secret-tool') });

/** Which OS keychain this machine supports, if any (Windows and others fall back to the file store). */
export function keychainKind(env: StoreEnv): 'macos' | 'secret-tool' | undefined {
  if (env.platform === 'darwin') return 'macos';
  if (env.platform === 'linux' && env.hasSecretTool) return 'secret-tool';
  return undefined;
}

export type StoreChoice = { kind: 'file' } | { kind: 'keychain'; backend: SecretBackend; strict: boolean };

/**
 * Pure selection of the token store. `auto` uses the keychain when this platform has one and
 * silently falls back to the file; `keychain` is strict and fails when none is available.
 */
export function selectTokenStore(
  mode: TokenStoreMode,
  cachePath: string,
  env: StoreEnv = currentStoreEnv(),
  exec: Exec = defaultExec,
): StoreChoice {
  if (mode === 'file') return { kind: 'file' };
  const kind = keychainKind(env);
  if (!kind) {
    if (mode === 'keychain') {
      throw new Error(
        'OUTLOOK_TOKEN_STORE=keychain is not available here: it needs macOS, or Linux with `secret-tool` installed. Use "auto" or "file".',
      );
    }
    return { kind: 'file' };
  }
  const backend = kind === 'macos' ? new MacKeychainBackend(cachePath, exec) : new SecretToolBackend(cachePath, exec);
  return { kind: 'keychain', backend, strict: mode === 'keychain' };
}

/**
 * MSAL cache plugin backed by an OS keychain. A cache file is imported into the keychain (and deleted)
 * when the keychain is empty, or when the file is newer than the keychain value. In non-strict (auto) mode,
 * a keychain failure degrades that access to the file store instead of breaking sign-in, but the keychain
 * is retried on every later access, and a successful keychain write removes the fallback file.
 */
export class KeychainCachePlugin implements ICachePlugin {
  private readonly file: FileCachePlugin;
  private warned = false;

  constructor(
    private readonly backend: SecretBackend,
    private readonly filePath: string,
    private readonly strict = false,
  ) {
    this.file = new FileCachePlugin(filePath);
  }

  /** Reads the keychain, reconciling it with a cache file left by a legacy install or a fallback write. */
  private async load(): Promise<string | undefined> {
    const entry = this.backend.getEntry
      ? await this.backend.getEntry()
      : await this.backend.get().then((value) => (value === undefined ? undefined : { value, updatedAt: undefined }));
    let fileValue: string;
    let mtimeMs: number;
    try {
      fileValue = await fsp.readFile(this.filePath, 'utf8');
      mtimeMs = (await fsp.stat(this.filePath)).mtimeMs;
    } catch (err) {
      if ((err as NodeJS.ErrnoException).code === 'ENOENT') return entry?.value;
      throw err;
    }
    // Keychain without a recorded time (e.g. Secret Service): a leftover file can only come from a write
    // made while the keychain was unavailable, so treat it as newer.
    if (entry && entry.updatedAt !== undefined && mtimeMs <= entry.updatedAt) {
      log.warn(`removing stale token cache file ${this.filePath}; the ${this.backend.name} has a newer copy`);
      await fsp.rm(this.filePath, { force: true });
      return entry.value;
    }
    await this.backend.set(fileValue);
    await fsp.rm(this.filePath, { force: true });
    log.info(`token cache moved from ${this.filePath} to the ${this.backend.name}`);
    return fileValue;
  }

  private fail(err: unknown): void {
    if (this.strict) throw err;
    if (!this.warned) log.warn(`${this.backend.name} unavailable (${(err as Error).message}); using the token cache file instead`);
    this.warned = true;
  }

  async beforeCacheAccess(ctx: TokenCacheContext): Promise<void> {
    try {
      const value = await this.load();
      if (value !== undefined) ctx.tokenCache.deserialize(value);
      return;
    } catch (err) {
      this.fail(err);
    }
    await this.file.beforeCacheAccess(ctx);
  }

  async afterCacheAccess(ctx: TokenCacheContext): Promise<void> {
    if (!ctx.cacheHasChanged) return;
    try {
      await this.backend.set(ctx.tokenCache.serialize());
      this.warned = false;
      await fsp.rm(this.filePath, { force: true });
      return;
    } catch (err) {
      this.fail(err);
    }
    await this.file.afterCacheAccess(ctx);
  }
}
