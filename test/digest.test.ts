import { describe, expect, it } from 'vitest';
import { buildDigest, type DigestMessage } from '../src/util/digest.js';

const m = (id: string, address: string, received: string, extra: Partial<DigestMessage> = {}): DigestMessage => ({
  id,
  subject: `s-${id}`,
  receivedDateTime: received,
  isRead: false,
  importance: 'normal',
  from: { emailAddress: { address, name: address.split('@')[0] } },
  ...extra,
});

const msgs: DigestMessage[] = [
  m('1', 'News@shop.com', '2026-01-01T10:00:00Z'),
  m('2', 'news@shop.com', '2026-01-01T12:00:00Z', { hasAttachments: true }),
  m('3', 'boss@corp.com', '2026-01-01T11:00:00Z', { importance: 'high' }),
  m('4', 'news@shop.com', '2026-01-01T09:00:00Z', { isRead: true, flag: { flagStatus: 'flagged' } }),
  m('5', 'news@shop.com', '2026-01-01T08:00:00Z'),
  { id: '6', receivedDateTime: '2026-01-01T07:00:00Z', from: null },
];

describe('buildDigest', () => {
  const d = buildDigest(msgs);
  it('counts', () => {
    expect(d.counts).toEqual({ total: 6, unread: 4, flagged: 1, highImportance: 1, withAttachments: 1 });
  });
  it('lists high-importance and flagged messages newest first', () => {
    expect(d.priority.map((p) => p.id)).toEqual(['3', '4']);
    expect(d.priority[1].flagged).toBe(true);
    expect(d.priority[0].importance).toBe('high');
  });
  it('groups by lowercased sender with latest subject and capped samples', () => {
    const g = d.groups!;
    expect(g.map((x) => x.address)).toEqual(['news@shop.com', 'boss@corp.com', '(unknown sender)']);
    expect(g[0]).toMatchObject({ count: 4, unread: 3, latestSubject: 's-2', latestReceived: '2026-01-01T12:00:00Z' });
    expect(g[0].sampleIds).toEqual(['2', '1', '4']);
  });
  it('omits groups with groupBy none and respects limits', () => {
    expect(buildDigest(msgs, { groupBy: 'none' }).groups).toBeUndefined();
    expect(buildDigest(msgs, { priorityLimit: 1 }).priority).toHaveLength(1);
  });
  it('handles an empty list', () => {
    expect(buildDigest([])).toEqual({ counts: { total: 0, unread: 0, flagged: 0, highImportance: 0, withAttachments: 0 }, priority: [], groups: [] });
  });
});
