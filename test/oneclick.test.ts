import { describe, expect, it, vi } from 'vitest';
import { postOneClick } from '../src/tools/unsubscribe.js';

describe('postOneClick', () => {
  it('reports success', async () => {
    const f = vi.fn(async () => new Response(null, { status: 200 }));
    expect(await postOneClick('https://x.example/u', f as unknown as typeof fetch)).toEqual({ ok: true, status: 200 });
    expect(f).toHaveBeenCalledOnce();
  });

  it('falls back to the link when refused', async () => {
    const f = async () => new Response(null, { status: 403 });
    const r = await postOneClick('https://x.example/u', f as unknown as typeof fetch);
    expect(r).toMatchObject({ ok: false });
    expect((r as { note: string }).note).toContain('HTTP 403');
  });

  it('does not throw on network errors', async () => {
    const f = async () => {
      throw new TypeError('fetch failed', { cause: new Error('ECONNREFUSED') });
    };
    const r = await postOneClick('https://x.example/u', f as unknown as typeof fetch);
    expect(r).toMatchObject({ ok: false });
    expect((r as { note: string }).note).toContain('One-click request failed (ECONNREFUSED)');
  });
});
