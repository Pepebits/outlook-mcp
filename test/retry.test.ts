import { describe, expect, it, vi } from 'vitest';
import { AuthRequiredError } from '../src/auth/msal.js';
import { GraphClient } from '../src/graph/client.js';
import { GraphError, mapGraphError } from '../src/graph/errors.js';

const json = (status: number, body: unknown, headers: Record<string, string> = {}) =>
  new Response(JSON.stringify(body), { status, headers });

function client(fetchMock: ReturnType<typeof vi.fn>, sleep = vi.fn(async () => {})) {
  return { sleep, c: new GraphClient(async () => 'tok', 'https://g.test/v1.0', { fetch: fetchMock as any, sleep }) };
}

describe('GraphClient', () => {
  it('honors Retry-After then succeeds', async () => {
    const f = vi.fn().mockResolvedValueOnce(json(429, {}, { 'Retry-After': '2' })).mockResolvedValueOnce(json(200, { ok: 1 }));
    const { c, sleep } = client(f);
    expect(await c.get('/me')).toEqual({ ok: 1 });
    expect(sleep).toHaveBeenCalledWith(2000);
  });
  it('gives up after 4 retries', async () => {
    const f = vi.fn().mockImplementation(async () => json(429, { error: { code: 'TooManyRequests', message: 'slow' } }, { 'Retry-After': '1' }));
    const { c } = client(f);
    await expect(c.get('/me')).rejects.toMatchObject({ status: 429, code: 'TooManyRequests' });
    expect(f).toHaveBeenCalledTimes(5);
  });
  it('maps 401 to AuthRequiredError', async () => {
    const f = vi.fn().mockResolvedValue(json(401, { error: { code: 'InvalidAuthenticationToken', message: 'x' } }));
    const { c } = client(f);
    const err = await c.get('/me').catch((e) => e);
    expect(mapGraphError(err)).toBeInstanceOf(AuthRequiredError);
  });
  it('passes a same-origin nextLink through unchanged and returns undefined on 204', async () => {
    const f = vi.fn().mockResolvedValueOnce(json(200, { value: [] })).mockResolvedValueOnce(new Response(null, { status: 204 }));
    const { c } = client(f);
    await c.get('https://g.test/v1.0/me/messages?$skip=10');
    expect(f.mock.calls[0][0]).toBe('https://g.test/v1.0/me/messages?$skip=10');
    expect(await c.delete('/me/messages/1')).toBeUndefined();
  });
  it('never sends the token to another origin, scheme or API path', async () => {
    const f = vi.fn();
    const { c } = client(f);
    for (const link of ['https://evil.test/v1.0/me', 'http://g.test/v1.0/me', 'https://g.test/beta/me', 'https://g.test/v1.0evil/me', 'https://g.test.evil.test/v1.0/me']) {
      await expect(c.get(link)).rejects.toThrow(/Refusing to send the access token/);
    }
    expect(f).not.toHaveBeenCalled();
  });
  it('retries POST only on 429, never on 5xx (the request may have been processed)', async () => {
    const f = vi.fn().mockResolvedValue(json(504, { error: { code: 'GatewayTimeout', message: 'x' } }));
    const { c } = client(f);
    await expect(c.post('/me/sendMail', {})).rejects.toMatchObject({ status: 504 });
    expect(f).toHaveBeenCalledTimes(1);
    const g = vi.fn().mockResolvedValueOnce(json(429, {}, { 'Retry-After': '1' })).mockResolvedValueOnce(new Response(null, { status: 202 }));
    await client(g).c.post('/me/sendMail', {});
    expect(g).toHaveBeenCalledTimes(2);
  });
  it('retries GET on 503 and on network errors; POST network errors are not retried', async () => {
    const f = vi.fn().mockResolvedValueOnce(json(503, {})).mockRejectedValueOnce(new TypeError('fetch failed')).mockResolvedValueOnce(json(200, { ok: 1 }));
    expect(await client(f).c.get('/me')).toEqual({ ok: 1 });
    expect(f).toHaveBeenCalledTimes(3);
    const g = vi.fn().mockRejectedValue(Object.assign(new TypeError('fetch failed'), { cause: { code: 'ECONNRESET' } }));
    await expect(client(g).c.post('/me/sendMail', {})).rejects.toThrow(/ECONNRESET.*may or may not have been processed/);
    expect(g).toHaveBeenCalledTimes(1);
  });
  it('caps Retry-After at 60 s and passes a timeout signal', async () => {
    const f = vi.fn().mockResolvedValueOnce(json(429, {}, { 'Retry-After': '600' })).mockResolvedValueOnce(json(200, {}));
    const { c, sleep } = client(f);
    await c.get('/me');
    expect(sleep).toHaveBeenCalledWith(60_000);
    expect(f.mock.calls[0][1].signal).toBeInstanceOf(AbortSignal);
  });
  it('read-only client refuses every non-GET request before sending it', async () => {
    const f = vi.fn().mockResolvedValue(json(200, {}));
    const c = new GraphClient(async () => 'tok', 'https://g.test/v1.0', { fetch: f as any, readOnly: true });
    await c.get('/me');
    for (const call of [() => c.post('/me/sendMail', {}), () => c.patch('/me/messages/1', {}), () => c.delete('/me/messages/1')]) {
      await expect(call()).rejects.toThrow(/Read-only mode/);
    }
    expect(f).toHaveBeenCalledTimes(1);
  });
});

describe('mapGraphError', () => {
  const e = (status: number, code: string, p: string, message = code) => mapGraphError(new GraphError(message, status, code, p)).message;
  it('words 404s after what was requested', () => {
    expect(e(404, 'ErrorItemNotFound', '/me/messages/1')).toMatch(/^Message not found/);
    expect(e(404, 'ErrorItemNotFound', '/me/messages/1/attachments/2')).toMatch(/^Attachment not found/);
    expect(e(404, 'ErrorItemNotFound', '/me/mailFolders/inbox/messageRules/9')).toMatch(/^Inbox rule not found/);
    expect(e(404, 'ErrorItemNotFound', '/me/mailFolders/xyz/childFolders')).toMatch(/^Folder not found/);
    expect(e(404, 'ErrorItemNotFound', '/me/mailFolders/inbox/messages/1')).toMatch(/^Message not found/);
  });
  it('explains a rejected From address', () => {
    expect(e(403, 'ErrorSendAsDenied', '/me/sendMail', 'denied')).toMatch(/not allowed to send as that From address/);
  });
});
