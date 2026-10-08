import { describe, expect, it, vi } from 'vitest';
import { AuthRequiredError } from '../src/auth/msal.js';
import { GraphClient } from '../src/graph/client.js';
import { mapGraphError } from '../src/graph/errors.js';

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
  it('passes absolute nextLink through unchanged and returns undefined on 204', async () => {
    const f = vi.fn().mockResolvedValueOnce(json(200, { value: [] })).mockResolvedValueOnce(new Response(null, { status: 204 }));
    const { c } = client(f);
    await c.get('https://other.test/next?$skip=10');
    expect(f.mock.calls[0][0]).toBe('https://other.test/next?$skip=10');
    expect(await c.delete('/me/messages/1')).toBeUndefined();
  });
});
