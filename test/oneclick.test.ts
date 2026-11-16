import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { postOneClick } from '../src/tools/unsubscribe.js';
import { isPrivateAddress, resolver } from '../src/util/netguard.js';

const PUBLIC = '93.184.216.34';
let dnsTable: Record<string, string[]>;
beforeEach(() => {
  dnsTable = { 'x.example': [PUBLIC], 'other.example': [PUBLIC], 'intranet.example': ['10.0.0.5'] };
  vi.spyOn(resolver, 'lookup').mockImplementation(async (h) => dnsTable[h] ?? []);
});
afterEach(() => vi.restoreAllMocks());

describe('postOneClick', () => {
  it('reports success', async () => {
    const f = vi.fn(async () => new Response(null, { status: 200 }));
    expect(await postOneClick('https://x.example/u', f as unknown as typeof fetch)).toEqual({ ok: true, status: 200 });
    expect(f).toHaveBeenCalledOnce();
    expect((f.mock.calls[0] as any[])[1].redirect).toBe('manual');
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

  it('never POSTs to private, loopback or plain-http targets', async () => {
    const f = vi.fn(async () => new Response(null, { status: 200 }));
    for (const url of ['https://10.0.0.5/u', 'https://127.0.0.1/u', 'https://[::1]/u', 'https://localhost/u', 'https://intranet.example/u', 'http://x.example/u']) {
      const r = await postOneClick(url, f as unknown as typeof fetch);
      expect(r, url).toMatchObject({ ok: false });
      expect((r as { note: string }).note).toMatch(/not attempted/);
    }
    expect(f).not.toHaveBeenCalled();
  });

  it('checks every redirect before following it', async () => {
    const redirect = (to: string) => new Response(null, { status: 307, headers: { location: to } });
    const f = vi.fn().mockResolvedValueOnce(redirect('https://other.example/u')).mockResolvedValueOnce(new Response(null, { status: 200 }));
    expect(await postOneClick('https://x.example/u', f as unknown as typeof fetch)).toEqual({ ok: true, status: 200 });
    expect(f.mock.calls[1][0]).toBe('https://other.example/u');

    const g = vi.fn().mockResolvedValueOnce(redirect('http://127.0.0.1/admin'));
    expect(await postOneClick('https://x.example/u', g as unknown as typeof fetch)).toMatchObject({ ok: false });
    expect(g).toHaveBeenCalledOnce();
  });
});

describe('isPrivateAddress', () => {
  it('classifies addresses', () => {
    for (const ip of ['10.1.2.3', '172.16.0.1', '192.168.1.1', '127.0.0.1', '169.254.169.254', '100.64.0.1', '::1', 'fe80::1', 'fd00::1', '::ffff:10.0.0.1', '0.0.0.0']) {
      expect(isPrivateAddress(ip), ip).toBe(true);
    }
    for (const ip of [PUBLIC, '2606:4700::1111', '8.8.8.8']) expect(isPrivateAddress(ip), ip).toBe(false);
  });
});
