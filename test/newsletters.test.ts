import { describe, expect, it } from 'vitest';
import { groupBySender, matchesDomain, unsubscribeKind, type SenderMessage } from '../src/util/newsletters.js';

const msg = (id: string, address: string, received: string, isRead = true, name = ''): SenderMessage => ({
  id,
  subject: `s-${id}`,
  receivedDateTime: received,
  isRead,
  from: { emailAddress: { address, name } },
});

describe('groupBySender', () => {
  const msgs = [
    msg('1', 'News@Shop.com', '2026-01-01T00:00:00Z', false, 'Shop'),
    msg('2', 'news@shop.com', '2026-01-03T00:00:00Z', false),
    msg('3', 'news@shop.com', '2026-01-02T00:00:00Z', true),
    msg('4', 'bob@friend.org', '2026-01-04T00:00:00Z'),
    { id: '5', from: null },
  ];

  it('groups case-insensitively, counts and picks the latest sample', () => {
    const g = groupBySender(msgs);
    expect(g.map((x) => x.address)).toEqual(['news@shop.com', 'bob@friend.org']);
    expect(g[0]).toMatchObject({ count: 3, unread: 2, lastReceived: '2026-01-03T00:00:00Z', sampleMessageId: '2', sampleSubject: 's-2', name: 'Shop' });
  });

  it('skips excluded domains including subdomains', () => {
    const g = groupBySender([...msgs, msg('6', 'a@mail.shop.com', '2026-01-05T00:00:00Z')], { excludeDomains: ['Shop.com'] });
    expect(g.map((x) => x.address)).toEqual(['bob@friend.org']);
  });
});

describe('matchesDomain', () => {
  it('does not match lookalike suffixes', () => {
    expect(matchesDomain('a@notshop.com', 'shop.com')).toBe(false);
    expect(matchesDomain('a@shop.com', '@shop.com')).toBe(true);
  });
});

describe('unsubscribeKind', () => {
  it('ranks one-click > mailto > link > none', () => {
    expect(unsubscribeKind({ https: ['https://x'], mailto: [], oneClick: true })).toBe('one-click');
    expect(unsubscribeKind({ https: ['https://x'], mailto: [{ address: 'a@b', subject: '', body: '' }], oneClick: false })).toBe('mailto');
    expect(unsubscribeKind({ https: ['https://x'], mailto: [], oneClick: false })).toBe('link');
    expect(unsubscribeKind({ https: [], mailto: [], oneClick: false })).toBe('none');
  });
});
