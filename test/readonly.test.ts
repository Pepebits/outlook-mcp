import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { describe, expect, it, vi } from 'vitest';
import { parseConfig } from '../src/config.js';
import { GraphClient } from '../src/graph/client.js';
import { allTools, selectTools, toolAnnotations } from '../src/tools/index.js';
import { fakeAuth } from './helpers/graph.js';

describe('read-only mode', () => {
  const names = (ro: boolean) => selectTools(allTools(), ro).map((t) => t.name);
  it('removes every mutating tool', () => {
    const ro = names(true);
    for (const n of ['send_mail', 'create_draft', 'reply_message', 'forward_message', 'move_message', 'mark_read', 'flag_message', 'delete_message', 'bulk_action', 'create_rule', 'delete_rule', 'block_sender']) {
      expect(ro).not.toContain(n);
      expect(names(false)).toContain(n);
    }
    expect(ro).toEqual(expect.arrayContaining(['list_folders', 'list_messages', 'search_messages', 'get_message', 'list_attachments', 'download_attachment', 'find_newsletters', 'mail_digest', 'get_unsubscribe_info', 'list_rules']));
  });

  it('every non-mutating tool only issues GET requests (ADR 0001)', async () => {
    const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'outlook-mcp-ro-'));
    const fetchFn = vi.fn(async (_url: any, init: any) => {
      // A generic body that satisfies list and single-item shapes well enough to keep handlers going.
      const body = { id: 'x', name: 'x.txt', value: [{ id: 'x', subject: 's', from: { emailAddress: { address: 'a@x.com' } }, internetMessageHeaders: [] }] };
      return init.method === 'GET' ? new Response(JSON.stringify(body), { status: 200 }) : new Response(null, { status: 500 });
    });
    const graph = new GraphClient(async () => 't', 'https://graph.test/v1.0', { fetch: fetchFn as any, readOnly: true, maxRetries: 0 });
    const ctx = { graph, cfg: parseConfig({ OUTLOOK_READ_ONLY: 'true', OUTLOOK_DOWNLOAD_DIR: dir }), auth: fakeAuth() };
    const args = { id: 'x', ids: undefined, messageId: 'x', attachmentId: 'x', query: 'x', path: 'out.txt' };
    try {
      for (const t of selectTools(allTools(), true)) {
        const err = await t.handler(args, ctx).then(() => undefined, (e: Error) => e);
        expect(err?.message ?? '', t.name).not.toMatch(/Read-only mode/);
      }
      expect(fetchFn.mock.calls.length).toBeGreaterThan(0);
      expect(fetchFn.mock.calls.every(([, init]) => (init as RequestInit).method === 'GET')).toBe(true);
    } finally {
      await fs.rm(dir, { recursive: true, force: true });
    }
  });
});

describe('tool annotations', () => {
  const ann = (name: string) => toolAnnotations(allTools().find((t) => t.name === name)!);
  it('derive read-only, destructive and open-world hints from the tool definition', () => {
    expect(ann('list_messages')).toEqual({ readOnlyHint: true, openWorldHint: false });
    expect(ann('delete_message')).toEqual({ readOnlyHint: false, destructiveHint: true, openWorldHint: false });
    expect(ann('mark_read')).toMatchObject({ readOnlyHint: false, destructiveHint: false });
    expect(ann('send_mail')).toMatchObject({ destructiveHint: true, openWorldHint: true });
    expect(ann('unsubscribe')).toMatchObject({ openWorldHint: true });
  });
});
