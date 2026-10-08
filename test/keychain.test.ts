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

class FakeBackend implements SecretBackend {
  name = 'fake';
  value: string | undefined;
  failing = false;
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

  it('keeps the file when the keychain already has data', async () => {
    await fs.writeFile(file, '{"legacy":true}');
    const b = new FakeBackend();
    b.value = '{"kc":true}';
    const r = fakeCtx();
    await new KeychainCachePlugin(b, file).beforeCacheAccess(r.ctx);
    expect(r.data).toBe('{"kc":true}');
    await expect(fs.stat(file)).resolves.toBeTruthy();
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

/** Emulates the subset of the macOS `security` CLI used by MacKeychainBackend. */
function fakeSecurity(store = new Map<string, string>(), lineLimit = 4096): { exec: Exec; store: Map<string, string>; argvs: string[][] } {
  const argvs: string[][] = [];
  const exec: Exec = async (cmd, args, input) => {
    argvs.push(args);
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
  return { exec, store, argvs };
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
