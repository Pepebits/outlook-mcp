import type { UnsubscribeInfo } from './unsubscribe.js';

export type UnsubscribeKind = 'one-click' | 'mailto' | 'link' | 'none';

export interface SenderMessage {
  id: string;
  subject?: string | null;
  receivedDateTime?: string | null;
  isRead?: boolean | null;
  from?: { emailAddress?: { name?: string | null; address?: string | null } } | null;
}

export interface SenderGroup {
  address: string;
  name: string;
  count: number;
  unread: number;
  lastReceived: string;
  sampleMessageId: string;
  sampleSubject: string;
}

export function unsubscribeKind(info: UnsubscribeInfo): UnsubscribeKind {
  if (info.oneClick) return 'one-click';
  if (info.mailto.length) return 'mailto';
  if (info.https.length) return 'link';
  return 'none';
}

/** True when `address` is `domain` or a subdomain of it (case-insensitive). */
export function matchesDomain(address: string, domain: string): boolean {
  const d = domain.trim().toLowerCase().replace(/^@/, '');
  if (!d) return false;
  const host = address.toLowerCase().split('@')[1] ?? '';
  return host === d || host.endsWith(`.${d}`);
}

/**
 * Groups messages by lowercased sender address, sorted by count (desc, then most
 * recent first). The sample message is the most recent one of each sender.
 */
export function groupBySender(messages: SenderMessage[], opts: { excludeDomains?: string[] } = {}): SenderGroup[] {
  const excluded = opts.excludeDomains ?? [];
  const groups = new Map<string, SenderGroup>();
  for (const m of messages) {
    const address = m.from?.emailAddress?.address?.trim().toLowerCase();
    if (!address) continue;
    if (excluded.some((d) => matchesDomain(address, d))) continue;
    const received = m.receivedDateTime ?? '';
    const g = groups.get(address);
    if (!g) {
      groups.set(address, {
        address,
        name: m.from?.emailAddress?.name ?? '',
        count: 1,
        unread: m.isRead === false ? 1 : 0,
        lastReceived: received,
        sampleMessageId: m.id,
        sampleSubject: m.subject ?? '',
      });
      continue;
    }
    g.count++;
    if (m.isRead === false) g.unread++;
    if (received > g.lastReceived) {
      g.lastReceived = received;
      g.sampleMessageId = m.id;
      g.sampleSubject = m.subject ?? '';
      if (m.from?.emailAddress?.name) g.name = m.from.emailAddress.name;
    }
  }
  return [...groups.values()].sort((a, b) => b.count - a.count || b.lastReceived.localeCompare(a.lastReceived));
}
