import { describe, expect, it, vi } from 'vitest';
import { AuthSession } from '../src/auth/session.js';
import { isNetworkFailure } from '../src/auth/msal.js';
import { parseConfig } from '../src/config.js';

const cfg = parseConfig({ OUTLOOK_TOKEN_STORE: 'file', OUTLOOK_TOKEN_CACHE: '/nonexistent/outlook-mcp-test/cache.json' });

describe('AuthSession.getToken', () => {
  it('shares one refresh between concurrent callers and reuses the token until near expiry', async () => {
    let now = 0;
    let n = 0;
    const fetchToken = vi.fn(async () => ({ token: `t${++n}`, expiresAt: 60 * 60 * 1000 }));
    const s = new AuthSession(cfg, () => now, fetchToken);
    expect(await Promise.all([s.getToken(), s.getToken(), s.getToken()])).toEqual(['t1', 't1', 't1']);
    expect(fetchToken).toHaveBeenCalledTimes(1);
    now = 50 * 60 * 1000;
    expect(await s.getToken()).toBe('t1');
    now = 56 * 60 * 1000; // inside the 5-minute refresh margin
    expect(await s.getToken()).toBe('t2');
    expect(fetchToken).toHaveBeenCalledTimes(2);
  });

  it('does not cache failures', async () => {
    const fetchToken = vi.fn().mockRejectedValueOnce(new Error('nope')).mockResolvedValueOnce({ token: 'ok', expiresAt: Date.now() + 3_600_000 });
    const s = new AuthSession(cfg, Date.now, fetchToken);
    await expect(s.getToken()).rejects.toThrow('nope');
    expect(await s.getToken()).toBe('ok');
  });
});

describe('isNetworkFailure', () => {
  it('tells network errors from expired sessions', () => {
    expect(isNetworkFailure({ errorCode: 'network_error', message: 'x' })).toBe(true);
    expect(isNetworkFailure(new TypeError('fetch failed'))).toBe(true);
    expect(isNetworkFailure({ errorCode: 'invalid_grant', message: 'AADSTS70000: token expired' })).toBe(false);
  });
});
