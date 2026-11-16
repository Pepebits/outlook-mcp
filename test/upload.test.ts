import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { UPLOAD_CHUNK_BYTES, uploadFileInChunks } from '../src/util/upload.js';
import { fakeGraph, graphError, makeCtx, on, run } from './helpers/graph.js';

const MB = 1024 * 1024;
let dir: string;

beforeEach(async () => {
  dir = await fs.mkdtemp(path.join(os.tmpdir(), 'outlook-mcp-up-'));
});
afterEach(async () => {
  vi.unstubAllGlobals();
  await fs.rm(dir, { recursive: true, force: true });
});

async function tempFile(name: string, size: number) {
  const p = path.join(dir, name);
  await fs.writeFile(p, Buffer.alloc(size, 7));
  return p;
}

/** Stubs global fetch (used for the pre-authenticated PUTs) and records every call. */
function stubUploads(reply: (n: number) => number = () => 202) {
  const puts: { url: string; headers: Record<string, string>; length: number }[] = [];
  vi.stubGlobal(
    'fetch',
    vi.fn(async (url: any, init: any) => {
      puts.push({ url: String(url), headers: init.headers, length: init.body.length });
      return new Response(null, { status: reply(puts.length) });
    }),
  );
  return puts;
}

describe('uploadFileInChunks', () => {
  it('sends a file just over 3 MB as one chunk with a full Content-Range and no Authorization', async () => {
    const size = 3 * MB + 1;
    const puts = stubUploads();
    await uploadFileInChunks('https://upload.test/s1', await tempFile('a.bin', size), size);
    expect(puts).toHaveLength(1);
    expect(puts[0].headers['Content-Range']).toBe(`bytes 0-${size - 1}/${size}`);
    expect(puts[0].headers['Content-Length']).toBe(String(size));
    expect(Object.keys(puts[0].headers).map((h) => h.toLowerCase())).not.toContain('authorization');
  });

  it('splits larger files on 320 KiB-multiple boundaries', () => {
    expect(UPLOAD_CHUNK_BYTES % (320 * 1024)).toBe(0);
  });

  it('spans several chunks with contiguous ranges', async () => {
    const size = 9 * MB;
    const puts = stubUploads();
    await uploadFileInChunks('https://upload.test/s1', await tempFile('b.bin', size), size);
    const c = UPLOAD_CHUNK_BYTES;
    expect(puts.map((p) => p.headers['Content-Range'])).toEqual([
      `bytes 0-${c - 1}/${size}`,
      `bytes ${c}-${2 * c - 1}/${size}`,
      `bytes ${2 * c}-${size - 1}/${size}`,
    ]);
    expect(puts.reduce((n, p) => n + p.length, 0)).toBe(size);
  });

  it('retries a chunk on 503 and then succeeds', async () => {
    const puts = stubUploads((n) => (n === 1 ? 503 : 200));
    const sleep = vi.fn(async () => {});
    await uploadFileInChunks('https://upload.test/s1', await tempFile('c.bin', 10), 10, { sleep });
    expect(puts).toHaveLength(2);
    expect(puts[1].headers['Content-Range']).toBe('bytes 0-9/10');
    expect(sleep).toHaveBeenCalledTimes(1);
  });

  it('gives up after repeated 5xx and fails fast on 4xx', async () => {
    let puts = stubUploads(() => 500);
    await expect(uploadFileInChunks('https://upload.test/s1', await tempFile('d.bin', 10), 10, { sleep: async () => {} })).rejects.toThrow(/HTTP 500/);
    expect(puts).toHaveLength(4);
    puts = stubUploads(() => 403);
    await expect(uploadFileInChunks('https://upload.test/s1', await tempFile('d.bin', 10), 10, { sleep: async () => {} })).rejects.toThrow(/HTTP 403/);
    expect(puts).toHaveLength(1);
  });
});

/** Attachments are only read from OUTLOOK_ATTACH_DIRS (ADR 0002); the temp dir is the allowed root here. */
const ctx = (graph: Parameters<typeof makeCtx>[0]) => makeCtx(graph, { OUTLOOK_ATTACH_DIRS: dir });

describe('large attachments in send_mail / create_draft', () => {
  const args = (files: string[]) => ({ to: ['a@x.com'], subject: 's', body: 'b', attachments: files.map((p) => ({ path: p })) });
  const session = on('POST', '/me/messages/d1/attachments/createUploadSession', { json: { uploadUrl: 'https://upload.test/s1' } });

  it('send_mail: draft, upload session, chunks, then send', async () => {
    const small = await tempFile('small.txt', 100);
    const big = await tempFile('big.bin', 3 * MB + 1);
    const puts = stubUploads();
    const { graph, calls } = fakeGraph(
      on('POST', '/me/messages', { json: { id: 'd1' } }),
      session,
      on('POST', '/me/messages/d1/send', { status: 202 }),
    );
    const out = await run('send_mail', args([small, big]), ctx(graph));
    expect(out).toEqual({ sent: true, to: ['a@x.com'] });
    expect(calls.map((c) => `${c.method} ${c.path}`)).toEqual(['POST /me/messages', 'POST /me/messages/d1/attachments/createUploadSession', 'POST /me/messages/d1/send']);
    expect(calls[0].body.attachments).toHaveLength(1);
    expect(calls[0].body.attachments[0].name).toBe('small.txt');
    expect(calls[1].body).toEqual({ AttachmentItem: { attachmentType: 'file', name: 'big.bin', size: 3 * MB + 1 } });
    expect(puts).toHaveLength(1);
    expect(puts[0].url).toBe('https://upload.test/s1');
  });

  it('send_mail keeps small attachments on the inline sendMail path', async () => {
    const small = await tempFile('small.txt', 100);
    const { graph, calls } = fakeGraph(on('POST', '/me/sendMail', { status: 202 }));
    await run('send_mail', args([small]), ctx(graph));
    expect(calls).toHaveLength(1);
    expect(calls[0].body.message.attachments[0].contentBytes).toBe(Buffer.alloc(100, 7).toString('base64'));
  });

  it('send_mail deletes the draft and reports the error when an upload fails', async () => {
    const big = await tempFile('big.bin', 3 * MB);
    stubUploads(() => 403);
    const { graph, calls } = fakeGraph(
      on('POST', '/me/messages', { json: { id: 'd1' } }),
      session,
      on('DELETE', '/me/messages/d1', { status: 204 }),
    );
    await expect(run('send_mail', args([big]), ctx(graph))).rejects.toThrow(/big\.bin.*HTTP 403/);
    expect(calls.map((c) => `${c.method} ${c.path}`)).toEqual(['POST /me/messages', 'POST /me/messages/d1/attachments/createUploadSession', 'DELETE /me/messages/d1']);
  });

  it('send_mail deletes the draft when the session cannot be created', async () => {
    const big = await tempFile('big.bin', 3 * MB);
    const { graph, calls } = fakeGraph(
      on('POST', '/me/messages', { json: { id: 'd1' } }),
      on('POST', '/me/messages/d1/attachments/createUploadSession', graphError(403, 'ErrorAccessDenied')),
      on('DELETE', '/me/messages/d1', { status: 204 }),
    );
    await expect(run('send_mail', args([big]), ctx(graph))).rejects.toThrow(/big\.bin/);
    expect(calls.at(-1)).toMatchObject({ method: 'DELETE', path: '/me/messages/d1' });
    expect(calls.some((c) => c.path.endsWith('/send'))).toBe(false);
  });

  it('create_draft keeps the draft and lists the attachments that failed', async () => {
    const ok = await tempFile('ok.bin', 3 * MB);
    const bad = await tempFile('bad.bin', 3 * MB);
    let n = 0;
    const puts = stubUploads(() => (++n === 1 ? 200 : 403));
    const { graph, calls } = fakeGraph(on('POST', '/me/messages', { json: { id: 'd1', webLink: 'w' } }), session);
    const out = await run('create_draft', args([ok, bad]), ctx(graph));
    expect(out.draftId).toBe('d1');
    expect(out.attachmentsFailed).toHaveLength(1);
    expect(out.attachmentsFailed[0].name).toBe('bad.bin');
    expect(puts).toHaveLength(2);
    expect(calls.some((c) => c.method === 'DELETE')).toBe(false);
  });

  it('rejects missing files and files over 150 MB before touching the mailbox', async () => {
    const { graph, calls } = fakeGraph();
    await expect(run('create_draft', args([path.join(dir, 'nope')]), ctx(graph))).rejects.toThrow(/not found/);
    const huge = path.join(dir, 'huge.bin');
    await fs.writeFile(huge, '');
    await fs.truncate(huge, 150 * MB + 1);
    await expect(run('create_draft', args([huge]), ctx(graph))).rejects.toThrow(/150 MB/);
    expect(calls).toHaveLength(0);
  });

  it('shares the inline budget: two 2 MB files give one inline attachment and one upload session', async () => {
    const a = await tempFile('a.bin', 2 * MB);
    const b = await tempFile('b.bin', 2 * MB);
    stubUploads();
    const { graph, calls } = fakeGraph(on('POST', '/me/messages', { json: { id: 'd1' } }), session, on('POST', '/me/messages/d1/send', { status: 202 }));
    await run('send_mail', args([a, b]), ctx(graph));
    expect(calls[0].body.attachments).toHaveLength(1);
    expect(calls.filter((c) => c.path.endsWith('/createUploadSession'))).toHaveLength(1);
  });

  it('keeps the draft and warns that the mail may have been sent when /send returns 504', async () => {
    const big = await tempFile('big.bin', 3 * MB);
    stubUploads();
    const { graph, calls } = fakeGraph(on('POST', '/me/messages', { json: { id: 'd1' } }), session, on('POST', '/me/messages/d1/send', graphError(504, 'GatewayTimeout')));
    await expect(run('send_mail', args([big]), ctx(graph))).rejects.toThrow(/may have been sent/);
    expect(calls.filter((c) => c.path.endsWith('/send'))).toHaveLength(1);
    expect(calls.some((c) => c.method === 'DELETE')).toBe(false);
  });
});

describe('attachment roots (ADR 0002)', () => {
  const args = (p: string) => ({ to: ['a@x.com'], subject: 's', body: 'b', attachments: [{ path: p }] });

  it('refuses files outside OUTLOOK_ATTACH_DIRS before touching the mailbox', async () => {
    const other = await fs.mkdtemp(path.join(os.tmpdir(), 'outlook-mcp-other-'));
    try {
      await fs.writeFile(path.join(other, 'secret.txt'), 'x');
      const { graph, calls } = fakeGraph();
      await expect(run('send_mail', args(path.join(other, 'secret.txt')), ctx(graph))).rejects.toThrow(/OUTLOOK_ATTACH_DIRS/);
      expect(calls).toHaveLength(0);
    } finally {
      await fs.rm(other, { recursive: true, force: true });
    }
  });

  it('refuses hidden files and files inside hidden directories', async () => {
    await fs.mkdir(path.join(dir, '.ssh'));
    await fs.writeFile(path.join(dir, '.ssh', 'id_ed25519'), 'k');
    await fs.writeFile(path.join(dir, '.env'), 'k');
    const { graph } = fakeGraph();
    await expect(run('send_mail', args(path.join(dir, '.ssh', 'id_ed25519')), ctx(graph))).rejects.toThrow(/hidden/);
    await expect(run('send_mail', args(path.join(dir, '.env')), ctx(graph))).rejects.toThrow(/hidden/);
  });

  it('follows symlinks before checking, so a link inside the root cannot point outside it', async () => {
    const other = await fs.mkdtemp(path.join(os.tmpdir(), 'outlook-mcp-other-'));
    try {
      await fs.writeFile(path.join(other, 'secret.txt'), 'x');
      await fs.symlink(path.join(other, 'secret.txt'), path.join(dir, 'link.txt'));
      const { graph } = fakeGraph();
      await expect(run('send_mail', args(path.join(dir, 'link.txt')), ctx(graph))).rejects.toThrow(/outside the allowed/);
    } finally {
      await fs.rm(other, { recursive: true, force: true });
    }
  });
});

describe('from / fromName', () => {
  const base = { to: ['a@x.com'], subject: 's', body: 'b' };

  it('sets message.from with the optional display name', async () => {
    const { graph, calls } = fakeGraph(on('POST', '/me/sendMail', { status: 202 }), on('POST', '/me/messages', { json: { id: 'd1' } }));
    await run('send_mail', { ...base, from: 'alias@x.com', fromName: 'Fox' }, ctx(graph));
    expect(calls[0].body.message.from).toEqual({ emailAddress: { address: 'alias@x.com', name: 'Fox' } });
    await run('create_draft', { ...base, from: 'alias@x.com' }, ctx(graph));
    expect(calls[1].body.from).toEqual({ emailAddress: { address: 'alias@x.com' } });
  });

  it('rejects fromName without from', async () => {
    const { graph, calls } = fakeGraph();
    await expect(run('send_mail', { ...base, fromName: 'Fox' }, ctx(graph))).rejects.toThrow(/fromName needs from/);
    expect(calls).toHaveLength(0);
  });
});
