import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import {
  KeychainCachePlugin,
  MacKeychainBackend,
  SecretToolBackend,
  keychainAccount,
  selectTokenStore,
  type Exec,
  type SecretBackend,
} from '../src/auth/keychain.js';
import { clearTokenStore } from '../src/auth/msal.js';
import { parseConfig } from '../src/config.js';

class FakeBackend implements SecretBackend {
  name = 'fake';
  value: string | undefined;
  updatedAt: number | undefined;
  failing = false;
  async getEntry() {
    if (this.failing) throw new Error('locked');
    return this.value === undefined ? undefined : { value: this.value, updatedAt: this.updatedAt };
  }
  async get() {
    if (this.failing) throw new Error('locked');
    return this.value;
  }
  async set(v: string) {
    if (this.failing) throw new Error('locked');
    this.value = v;
  }
  async delete() {
    this.value = undefined;
  }
}
const exists = (f: string) => fs.stat(f).then(() => true, () => false);

function fakeCtx(initial = '', changed = false) {
  let data = initial;
  const tokenCache = {
    deserialize: (s: string) => {
      data = s;
    },
    serialize: () => data,
  };
  return { ctx: { tokenCache, cacheHasChanged: changed } as any, get data() { return data; } };
}

describe('selectTokenStore', () => {
  const mac = { platform: 'darwin' as const, hasSecretTool: false };
  const linux = { platform: 'linux' as const, hasSecretTool: true };
  const linuxBare = { platform: 'linux' as const, hasSecretTool: false };
  const win = { platform: 'win32' as const, hasSecretTool: false };

  it('file mode never uses a keychain', () => {
    expect(selectTokenStore('file', '/c', mac).kind).toBe('file');
  });
  it('auto uses the keychain where available and the file elsewhere', () => {
    const a = selectTokenStore('auto', '/c', mac);
    expect(a.kind === 'keychain' && a.backend instanceof MacKeychainBackend && !a.strict).toBe(true);
    const l = selectTokenStore('auto', '/c', linux);
    expect(l.kind === 'keychain' && l.backend instanceof SecretToolBackend).toBe(true);
    expect(selectTokenStore('auto', '/c', linuxBare).kind).toBe('file');
    expect(selectTokenStore('auto', '/c', win).kind).toBe('file');
  });
  it('keychain mode is strict and errors when unsupported', () => {
    const k = selectTokenStore('keychain', '/c', mac);
    expect(k.kind === 'keychain' && k.strict).toBe(true);
    expect(() => selectTokenStore('keychain', '/c', win)).toThrow(/not available/);
    expect(() => selectTokenStore('keychain', '/c', linuxBare)).toThrow(/secret-tool/);
  });
});

describe('keychainAccount', () => {
  it('uses the path, or a hash when it cannot be quoted', () => {
    expect(keychainAccount('/Users/a b/token.json')).toBe('/Users/a b/token.json');
    expect(keychainAccount('/x/"y"')).toMatch(/^token-cache-[0-9a-f]{16}$/);
    expect(keychainAccount("/Users/o'brien/t.json")).toMatch(/^token-cache-[0-9a-f]{16}$/);
  });
});

describe('KeychainCachePlugin', () => {
  let dir: string;
  let file: string;
  beforeEach(async () => {
    dir = await fs.mkdtemp(path.join(os.tmpdir(), 'outlook-mcp-kc-'));
    file = path.join(dir, 'token-cache.json');
  });
  afterEach(() => fs.rm(dir, { recursive: true, force: true }));

  it('loads and saves through the backend', async () => {
    const b = new FakeBackend();
    b.value = '{"a":1}';
    const plugin = new KeychainCachePlugin(b, file);
    const r = fakeCtx();
    await plugin.beforeCacheAccess(r.ctx);
    expect(r.data).toBe('{"a":1}');
    const w = fakeCtx('{"a":2}', true);
    await plugin.afterCacheAccess(w.ctx);
    expect(b.value).toBe('{"a":2}');
    await plugin.afterCacheAccess(fakeCtx('{"a":3}', false).ctx);
    expect(b.value).toBe('{"a":2}');
  });

  it('imports a legacy file into an empty keychain and deletes the file', async () => {
    await fs.writeFile(file, '{"legacy":true}');
    const b = new FakeBackend();
    const r = fakeCtx();
    await new KeychainCachePlugin(b, file).beforeCacheAccess(r.ctx);
    expect(b.value).toBe('{"legacy":true}');
    expect(r.data).toBe('{"legacy":true}');
    await expect(fs.stat(file)).rejects.toThrow();
  });

  it('prefers a newer cache file over the keychain, imports it and deletes the file', async () => {
    await fs.writeFile(file, '{"file":true}');
    const b = new FakeBackend();
    b.value = '{"kc":true}';
    b.updatedAt = Date.now() - 60_000;
    const r = fakeCtx();
    await new KeychainCachePlugin(b, file).beforeCacheAccess(r.ctx);
    expect(r.data).toBe('{"file":true}');
    expect(b.value).toBe('{"file":true}');
    expect(await exists(file)).toBe(false);
  });

  it('prefers a newer keychain value and removes the stale plaintext file', async () => {
    await fs.writeFile(file, '{"file":true}');
    const b = new FakeBackend();
    b.value = '{"kc":true}';
    b.updatedAt = Date.now() + 60_000;
    const r = fakeCtx();
    await new KeychainCachePlugin(b, file).beforeCacheAccess(r.ctx);
    expect(r.data).toBe('{"kc":true}');
    expect(b.value).toBe('{"kc":true}');
    expect(await exists(file)).toBe(false);
  });

  it('treats a file as newer when the keychain records no time', async () => {
    await fs.writeFile(file, '{"file":true}');
    const b = new FakeBackend();
    b.value = '{"kc":true}';
    await new KeychainCachePlugin(b, file).beforeCacheAccess(fakeCtx().ctx);
    expect(b.value).toBe('{"file":true}');
    expect(await exists(file)).toBe(false);
  });

  it('auto mode retries the keychain after a failure and removes the fallback file on success', async () => {
    const b = new FakeBackend();
    b.failing = true;
    const plugin = new KeychainCachePlugin(b, file, false);
    await plugin.afterCacheAccess(fakeCtx('{"x":1}', true).ctx);
    expect(await exists(file)).toBe(true);
    b.failing = false;
    await plugin.afterCacheAccess(fakeCtx('{"x":2}', true).ctx);
    expect(b.value).toBe('{"x":2}');
    expect(await exists(file)).toBe(false);
  });

  it('does not delete the file when the import fails', async () => {
    await fs.writeFile(file, '{"legacy":true}');
    const b = new FakeBackend();
    b.set = async () => {
      throw new Error('denied');
    };
    await expect(new KeychainCachePlugin(b, file, true).beforeCacheAccess(fakeCtx().ctx)).rejects.toThrow('denied');
    await expect(fs.stat(file)).resolves.toBeTruthy();
  });

  it('auto mode degrades to the file store when the keychain fails', async () => {
    const b = new FakeBackend();
    b.failing = true;
    const plugin = new KeychainCachePlugin(b, file, false);
    await plugin.afterCacheAccess(fakeCtx('{"x":1}', true).ctx);
    expect(await fs.readFile(file, 'utf8')).toBe('{"x":1}');
    const r = fakeCtx();
    await plugin.beforeCacheAccess(r.ctx);
    expect(r.data).toBe('{"x":1}');
  });

  it('strict mode surfaces keychain errors', async () => {
    const b = new FakeBackend();
    b.failing = true;
    await expect(new KeychainCachePlugin(b, file, true).beforeCacheAccess(fakeCtx().ctx)).rejects.toThrow('locked');
  });
});

type ExecResultLike = { code: number | null; stdout: string; stderr: string };

/** Emulates the subset of the macOS `security` CLI used by MacKeychainBackend. */
function fakeSecurity(
  store = new Map<string, string>(),
  lineLimit = 4096,
): { exec: Exec; store: Map<string, string>; argvs: string[][]; hook: { fn?: (args: string[]) => ExecResultLike | undefined } } {
  const argvs: string[][] = [];
  const hook: { fn?: (args: string[]) => ExecResultLike | undefined } = {};
  const exec: Exec = async (cmd, args, input) => {
    argvs.push(args);
    const h = hook.fn?.(args);
    if (h) return h;
    const flag = (f: string) => args[args.indexOf(f) + 1];
    if (args[0] === '-i') {
      for (const line of (input ?? '').split('\n').filter(Boolean)) {
        if (line.length > lineLimit) return { code: 0, stdout: '', stderr: 'line too long' };
        const m = /^add-generic-password -U -s (\S+) -a "([^"]*)" -w (\S+)$/.exec(line);
        if (m) store.set(`${m[1]}|${m[2]}`, m[3]);
      }
      return { code: 0, stdout: '', stderr: '' };
    }
    const key = `${flag('-s')}|${flag('-a')}`;
    if (args[0] === 'find-generic-password') {
      return store.has(key) ? { code: 0, stdout: `${store.get(key)}\n`, stderr: '' } : { code: 44, stdout: '', stderr: 'not found' };
    }
    if (args[0] === 'delete-generic-password') {
      return { code: store.delete(key) ? 0 : 44, stdout: '', stderr: '' };
    }
    return { code: 1, stdout: '', stderr: 'unknown' };
  };
  return { exec, store, argvs, hook };
}

describe('MacKeychainBackend', () => {
  it('round-trips large values in chunks without putting the secret in argv', async () => {
    const f = fakeSecurity();
    const b = new MacKeychainBackend('/tmp/t.json', f.exec);
    expect(await b.get()).toBeUndefined();
    const big = JSON.stringify({ token: 'é'.repeat(9000), n: 1 });
    await b.set(big);
    expect(await b.get()).toBe(big);
    expect(f.store.size).toBeGreaterThan(2);
    expect(f.argvs.flat().some((a) => a.includes(big.slice(10, 40)))).toBe(false);
    const small = '{"a":1}';
    await b.set(small);
    expect(await b.get()).toBe(small);
    expect(f.store.size).toBe(2);
    expect([...f.store.keys()].some((k) => /#[0-9a-f]+#0$/.test(k))).toBe(true);
    await b.delete();
    expect(f.store.size).toBe(0);
    expect(await b.get()).toBeUndefined();
  });

  it('removes chunks orphaned without a manifest', async () => {
    const f = fakeSecurity();
    const b = new MacKeychainBackend('/tmp/t.json', f.exec);
    await b.set('{"a":1}');
    f.store.set('outlook-mcp|/tmp/t.json#2', 'stale');
    await b.delete();
    expect(f.store.size).toBe(0);
  });

  it('replaces the previous generation and writes the manifest last', async () => {
    const f = fakeSecurity();
    const order: string[] = [];
    f.hook.fn = (args) => {
      if (args[0] === '-i') order.push('add');
      return undefined;
    };
    const b = new MacKeychainBackend('/tmp/t.json', f.exec);
    await b.set('{"a":1}');
    const first = [...f.store.keys()].sort();
    await b.set('{"a":2}');
    const second = [...f.store.keys()].sort();
    expect(second).toHaveLength(2);
    expect(second.filter((k) => k.includes('#'))[0]).not.toBe(first.filter((k) => k.includes('#'))[0]);
    expect(f.store.get('outlook-mcp|/tmp/t.json')).toMatch(/^chunks:[0-9a-f]+:1:\d+$/);
    expect(await b.get()).toBe('{"a":2}');
  });

  it('retries once when a concurrent write replaces the generation being read', async () => {
    const f = fakeSecurity();
    const b = new MacKeychainBackend('/tmp/t.json', f.exec);
    await b.set('{"v":1}');
    const oldKeys = new Map(f.store);
    let tripped = false;
    f.hook.fn = (args) => {
      // After the first chunk lookup, simulate another process finishing a newer write.
      if (!tripped && args[0] === 'find-generic-password' && args[args.indexOf('-a') + 1].includes('#')) {
        tripped = true;
        for (const k of oldKeys.keys()) f.store.delete(k);
        const gen = 'abc123';
        f.store.set('outlook-mcp|/tmp/t.json#abc123#0', Buffer.from('{"v":2}').toString('base64'));
        f.store.set('outlook-mcp|/tmp/t.json', `chunks:${gen}:1:${Date.now()}`);
      }
      return undefined;
    };
    expect(await b.get()).toBe('{"v":2}');
  });

  it('reads the legacy layout and migrates it on the next write', async () => {
    const f = fakeSecurity();
    f.store.set('outlook-mcp|/tmp/t.json', 'chunks:2');
    const full = Buffer.from('{"legacy":1}').toString('base64');
    f.store.set('outlook-mcp|/tmp/t.json#0', full.slice(0, 8));
    f.store.set('outlook-mcp|/tmp/t.json#1', full.slice(8));
    const b = new MacKeychainBackend('/tmp/t.json', f.exec);
    expect(await b.get()).toBe('{"legacy":1}');
    await b.set('{"new":1}');
    expect(f.store.has('outlook-mcp|/tmp/t.json#0')).toBe(false);
    expect(f.store.has('outlook-mcp|/tmp/t.json#1')).toBe(false);
    expect(await b.get()).toBe('{"new":1}');
  });

  it('delete removes orphans beyond 16 slots without a limit of 16', async () => {
    const f = fakeSecurity();
    const b = new MacKeychainBackend('/tmp/t.json', f.exec);
    f.store.set('outlook-mcp|/tmp/t.json', 'chunks:40');
    for (let i = 0; i < 40; i++) f.store.set(`outlook-mcp|/tmp/t.json#${i}`, 'x');
    await b.delete();
    expect(f.store.size).toBe(0);
  });

  it('treats only exit 44 as not found; other failures throw without leaking secrets', async () => {
    const f = fakeSecurity();
    const b = new MacKeychainBackend('/tmp/t.json', f.exec);
    f.hook.fn = (args) => (args[0] === 'find-generic-password' ? { code: 36, stdout: 'SECRET', stderr: 'User interaction is not allowed.' } : undefined);
    const err = await b.get().catch((e: Error) => e);
    expect(err).toBeInstanceOf(Error);
    expect((err as Error).message).toContain('User interaction is not allowed.');
    expect((err as Error).message).not.toContain('SECRET');
  });

  it('surfaces keychain read errors to auto mode instead of treating them as empty', async () => {
    const f = fakeSecurity();
    f.hook.fn = (args) => (args[0] === 'find-generic-password' ? { code: 36, stdout: '', stderr: 'locked' } : undefined);
    const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'outlook-mcp-kc-'));
    try {
      const file = path.join(dir, 't.json');
      await fs.writeFile(file, '{"f":1}');
      const r = fakeCtx();
      await new KeychainCachePlugin(new MacKeychainBackend('/tmp/t.json', f.exec), file, false).beforeCacheAccess(r.ctx);
      expect(r.data).toBe('{"f":1}');
      await expect(new KeychainCachePlugin(new MacKeychainBackend('/tmp/t.json', f.exec), file, true).beforeCacheAccess(fakeCtx().ctx)).rejects.toThrow('locked');
    } finally {
      await fs.rm(dir, { recursive: true, force: true });
    }
  });

  it('delete throws on failures other than not-found', async () => {
    const f = fakeSecurity();
    const b = new MacKeychainBackend('/tmp/t.json', f.exec);
    await b.set('{"a":1}');
    f.hook.fn = (args) => (args[0] === 'delete-generic-password' ? { code: 1, stdout: '', stderr: 'denied' } : undefined);
    await expect(b.delete()).rejects.toThrow('denied');
  });

  it('reports a failed write', async () => {
    const f = fakeSecurity(new Map(), 100);
    await expect(new MacKeychainBackend('/tmp/t.json', f.exec).set('x'.repeat(500))).rejects.toThrow();
  });
});

describe('SecretToolBackend', () => {
  it('passes the secret on stdin and treats a silent exit 1 as empty', async () => {
    const calls: { args: string[]; input?: string }[] = [];
    let stored: string | undefined;
    const exec: Exec = async (_c, args, input) => {
      calls.push({ args, input });
      if (args[0] === 'store') {
        stored = input;
        return { code: 0, stdout: '', stderr: '' };
      }
      if (args[0] === 'lookup') return stored === undefined ? { code: 1, stdout: '', stderr: '' } : { code: 0, stdout: stored, stderr: '' };
      stored = undefined;
      return { code: 0, stdout: '', stderr: '' };
    };
    const b = new SecretToolBackend('/c.json', exec);
    expect(await b.get()).toBeUndefined();
    await b.set('{"s":1}');
    expect(calls.find((c) => c.args[0] === 'store')?.args.join(' ')).not.toContain('{"s":1}');
    expect(calls.find((c) => c.args[0] === 'store')?.args).toContain('outlook-mcp');
    expect(await b.get()).toBe('{"s":1}');
    await b.delete();
    expect(await b.get()).toBeUndefined();
  });
});

describe('SecretToolBackend errors', () => {
  const mk = (r: { code: number | null; stdout: string; stderr: string }) => new SecretToolBackend('/c.json', async () => r);
  it('only a silent exit 1 means not found', async () => {
    expect(await mk({ code: 1, stdout: '', stderr: '' }).get()).toBeUndefined();
    await expect(mk({ code: 1, stdout: '', stderr: 'dbus down' }).get()).rejects.toThrow('dbus down');
    await expect(mk({ code: 2, stdout: '', stderr: '' }).get()).rejects.toThrow(/exit 2/);
  });
  it('clear failures throw', async () => {
    await expect(mk({ code: 1, stdout: '', stderr: 'locked' }).delete()).rejects.toThrow('locked');
  });
});

describe('clearTokenStore', () => {
  it('deletes the keychain entry even in file mode, and reports real failures', async () => {
    const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'outlook-mcp-kc-'));
    try {
      const cfg = parseConfig({ OUTLOOK_TOKEN_STORE: 'file', OUTLOOK_TOKEN_CACHE: path.join(dir, 't.json') });
      await fs.writeFile(cfg.tokenCachePath, '{}');
      const f = fakeSecurity();
      f.store.set(`outlook-mcp|${cfg.tokenCachePath}`, 'chunks:1');
      f.store.set(`outlook-mcp|${cfg.tokenCachePath}#0`, 'eA==');
      await clearTokenStore(cfg, { platform: 'darwin', hasSecretTool: false }, f.exec);
      expect(f.store.size).toBe(0);
      expect(await exists(cfg.tokenCachePath)).toBe(false);

      f.hook.fn = (args) => (args[0] === 'delete-generic-password' ? { code: 1, stdout: '', stderr: 'denied' } : undefined);
      f.store.set(`outlook-mcp|${cfg.tokenCachePath}`, 'chunks:1');
      await expect(clearTokenStore(cfg, { platform: 'darwin', hasSecretTool: false }, f.exec)).rejects.toThrow('denied');
      // Platforms without a keychain just clear the file.
      await clearTokenStore(cfg, { platform: 'win32', hasSecretTool: false }, f.exec);
    } finally {
      await fs.rm(dir, { recursive: true, force: true });
    }
  });
});
