export interface DigestMessage {
  id: string;
  subject?: string | null;
  from?: { emailAddress?: { name?: string | null; address?: string | null } } | null;
  receivedDateTime?: string | null;
  isRead?: boolean | null;
  importance?: string | null;
  flag?: { flagStatus?: string | null } | null;
  hasAttachments?: boolean | null;
}

export interface DigestCounts {
  total: number;
  unread: number;
  flagged: number;
  highImportance: number;
  withAttachments: number;
}

export interface DigestPriorityItem {
  id: string;
  subject: string;
  from: string;
  receivedDateTime: string;
  isRead: boolean;
  importance: string;
  flagged: boolean;
}

export interface DigestGroup {
  address: string;
  name: string;
  count: number;
  unread: number;
  latestSubject: string;
  latestReceived: string;
  /** Message ids, newest first. */
  sampleIds: string[];
}

export interface Digest {
  counts: DigestCounts;
  priority: DigestPriorityItem[];
  groups?: DigestGroup[];
}

export interface DigestOptions {
  groupBy?: 'sender' | 'none';
  /** Max high-importance / flagged messages listed (default 20). */
  priorityLimit?: number;
  /** Max sample ids per sender group (default 3). */
  sampleLimit?: number;
}

const isFlagged = (m: DigestMessage) => m.flag?.flagStatus === 'flagged';
const isHigh = (m: DigestMessage) => m.importance === 'high';
const newestFirst = (a: DigestMessage, b: DigestMessage) => (b.receivedDateTime ?? '').localeCompare(a.receivedDateTime ?? '');

/** Pure aggregation of a message list into counts, priority items and per-sender groups. */
export function buildDigest(messages: DigestMessage[], opts: DigestOptions = {}): Digest {
  const sorted = [...messages].sort(newestFirst);
  const counts: DigestCounts = {
    total: sorted.length,
    unread: sorted.filter((m) => m.isRead === false).length,
    flagged: sorted.filter(isFlagged).length,
    highImportance: sorted.filter(isHigh).length,
    withAttachments: sorted.filter((m) => m.hasAttachments === true).length,
  };
  const priority = sorted
    .filter((m) => isHigh(m) || isFlagged(m))
    .slice(0, opts.priorityLimit ?? 20)
    .map((m) => ({
      id: m.id,
      subject: m.subject ?? '(no subject)',
      from: m.from?.emailAddress?.address ?? '',
      receivedDateTime: m.receivedDateTime ?? '',
      isRead: m.isRead !== false,
      importance: m.importance ?? 'normal',
      flagged: isFlagged(m),
    }));
  if ((opts.groupBy ?? 'sender') === 'none') return { counts, priority };

  const sampleLimit = opts.sampleLimit ?? 3;
  const groups = new Map<string, DigestGroup>();
  for (const m of sorted) {
    const address = m.from?.emailAddress?.address?.trim().toLowerCase() || '(unknown sender)';
    const g = groups.get(address);
    if (!g) {
      groups.set(address, {
        address,
        name: m.from?.emailAddress?.name ?? '',
        count: 1,
        unread: m.isRead === false ? 1 : 0,
        latestSubject: m.subject ?? '(no subject)',
        latestReceived: m.receivedDateTime ?? '',
        sampleIds: [m.id],
      });
      continue;
    }
    g.count++;
    if (m.isRead === false) g.unread++;
    if (g.sampleIds.length < sampleLimit) g.sampleIds.push(m.id);
  }
  return {
    counts,
    priority,
    groups: [...groups.values()].sort((a, b) => b.count - a.count || b.latestReceived.localeCompare(a.latestReceived)),
  };
}
