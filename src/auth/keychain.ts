import type { ICachePlugin, TokenCacheContext } from '@azure/msal-node';
import { execFile, type ExecFileException } from 'node:child_process';
import { createHash } from 'node:crypto';
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

/** Account name for the keychain entry: the cache path, or a hash of it when it holds characters `security -i` cannot quote. */
export function keychainAccount(cachePath: string): string {
  if (/^[^"\\\r\n]+$/.test(cachePath)) return cachePath;
  return `token-cache-${createHash('sha256').update(cachePath).digest('hex').slice(0, 16)}`;
}

// `security -i` reads commands from stdin and truncates lines at about 4 KB, so values are stored base64
// encoded in chunks that stay well under that limit.
const MAC_CHUNK = 3000;

/**
 * macOS Keychain through the `security` CLI.
 *
 * Secret handling: `security add-generic-password` only accepts the password as an argument (-w/-X), which
 * would expose the whole token cache in `ps` output to other local processes. Instead the command is fed to
 * `security -i` over stdin, so the secret never appears in argv. The cost is a ~4 KB line limit, handled by
 * splitting the (base64) value across several items.
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

  private find(account: string): Promise<ExecResult> {
    return this.exec('security', ['find-generic-password', '-s', KEYCHAIN_SERVICE, '-a', account, '-w']);
  }

  private async add(account: string, value: string): Promise<void> {
    const line = `add-generic-password -U -s ${KEYCHAIN_SERVICE} -a "${account}" -w ${value}\n`;
    const r = await this.exec('security', ['-i'], line);
    if (r.spawnError) throw new Error(`security CLI unavailable (${r.spawnError})`);
  }

  private async remove(account: string): Promise<void> {
    await this.exec('security', ['delete-generic-password', '-s', KEYCHAIN_SERVICE, '-a', account]);
  }

  private async readCount(): Promise<number | undefined> {
    const r = await this.find(this.account);
    if (r.spawnError) throw new Error(`security CLI unavailable (${r.spawnError})`);
    if (r.code !== 0) return undefined;
    const m = /^chunks:(\d+)$/.exec(r.stdout.trim());
    return m ? Number(m[1]) : undefined;
  }

  async get(): Promise<string | undefined> {
    const count = await this.readCount();
    if (count === undefined) return undefined;
    let b64 = '';
    for (let i = 0; i < count; i++) {
      const r = await this.find(`${this.account}#${i}`);
      if (r.code !== 0) throw new Error('Token cache in the keychain is incomplete.');
      b64 += r.stdout.trim();
    }
    return Buffer.from(b64, 'base64').toString('utf8');
  }

  async set(value: string): Promise<void> {
    const old = (await this.readCount()) ?? 0;
    const b64 = Buffer.from(value, 'utf8').toString('base64');
    const chunks = b64.match(new RegExp(`.{1,${MAC_CHUNK}}`, 'g')) ?? [''];
    for (let i = 0; i < chunks.length; i++) await this.add(`${this.account}#${i}`, chunks[i] || '=');
    await this.add(this.account, `chunks:${chunks.length}`);
    for (let i = chunks.length; i < old; i++) await this.remove(`${this.account}#${i}`);
    if ((await this.get()) !== value) throw new Error('Could not write the token cache to the macOS Keychain.');
  }

  async delete(): Promise<void> {
    const count = (await this.readCount()) ?? 0;
    for (let i = 0; i < count; i++) await this.remove(`${this.account}#${i}`);
    await this.remove(this.account);
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
    // lookup exits 1 with empty output when nothing matches; real failures also write to stderr.
    if (r.code !== 0) {
      if (!r.stderr.trim()) return undefined;
      throw new Error(`secret-tool lookup failed: ${r.stderr.trim()}`);
    }
    return r.stdout === '' ? undefined : r.stdout;
  }

  async set(value: string): Promise<void> {
    const r = await this.exec('secret-tool', ['store', '--label=outlook-mcp token cache', ...this.attrs], value);
    if (r.spawnError) throw new Error(`secret-tool unavailable (${r.spawnError})`);
    if (r.code !== 0) throw new Error(`secret-tool store failed: ${r.stderr.trim() || `exit ${r.code}`}`);
  }

  async delete(): Promise<void> {
    await this.exec('secret-tool', ['clear', ...this.attrs]);
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
 * MSAL cache plugin backed by an OS keychain. A legacy cache file is imported into the keychain (and
 * deleted) the first time the keychain is empty. In non-strict (auto) mode, a keychain failure at runtime
 * degrades to the file store instead of breaking sign-in.
 */
export class KeychainCachePlugin implements ICachePlugin {
  private readonly file: FileCachePlugin;
  private degraded = false;
  private migrated = false;

  constructor(
    private readonly backend: SecretBackend,
    private readonly filePath: string,
    private readonly strict = false,
  ) {
    this.file = new FileCachePlugin(filePath);
  }

  /** Imports the legacy file into the keychain when the keychain has nothing yet. */
  private async load(): Promise<string | undefined> {
    const existing = await this.backend.get();
    if (existing !== undefined || this.migrated) return existing;
    let legacy: string;
    try {
      legacy = await fsp.readFile(this.filePath, 'utf8');
    } catch (err) {
      if ((err as NodeJS.ErrnoException).code === 'ENOENT') return undefined;
      throw err;
    }
    await this.backend.set(legacy);
    await fsp.rm(this.filePath, { force: true });
    this.migrated = true;
    log.info(`token cache moved from ${this.filePath} to the ${this.backend.name}`);
    return legacy;
  }

  private fail(err: unknown): void {
    if (this.strict) throw err;
    if (!this.degraded) log.warn(`${this.backend.name} unavailable (${(err as Error).message}); using the token cache file instead`);
    this.degraded = true;
  }

  async beforeCacheAccess(ctx: TokenCacheContext): Promise<void> {
    if (!this.degraded) {
      try {
        const value = await this.load();
        if (value !== undefined) ctx.tokenCache.deserialize(value);
        return;
      } catch (err) {
        this.fail(err);
      }
    }
    await this.file.beforeCacheAccess(ctx);
  }

  async afterCacheAccess(ctx: TokenCacheContext): Promise<void> {
    if (!ctx.cacheHasChanged) return;
    if (!this.degraded) {
      try {
        await this.backend.set(ctx.tokenCache.serialize());
        return;
      } catch (err) {
        this.fail(err);
      }
    }
    await this.file.afterCacheAccess(ctx);
  }
}
