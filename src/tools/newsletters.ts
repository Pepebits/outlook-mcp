import { z } from 'zod';
import type { Message, Page } from '../graph/types.js';
import { groupBySender, unsubscribeKind, type SenderGroup, type UnsubscribeKind } from '../util/newsletters.js';
import { wellKnownFolder } from '../util/odata.js';
import { parseUnsubscribeHeaders } from '../util/unsubscribe.js';
import type { ToolDef } from './index.js';

const DEFAULT_MAX = 500;
const LOOKUP_CONCURRENCY = 4;

async function mapLimit<T, R>(items: T[], limit: number, fn: (item: T) => Promise<R>): Promise<R[]> {
  const out = new Array<R>(items.length);
  let next = 0;
  const worker = async () => {
    while (next < items.length) {
      const i = next++;
      out[i] = await fn(items[i]);
    }
  };
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
  return out;
}

export const newsletterTools: ToolDef[] = [
  {
    name: 'find_newsletters',
    description:
      'Scan recent messages in a folder (default inbox), group them by sender and report which senders are mailing lists (have a List-Unsubscribe header), most frequent first. Each result says how it can be unsubscribed (one-click, mailto, link); pass its sampleMessageId to unsubscribe. Read-only.',
    schema: {
      folderId: z.string().optional().describe('Folder id or well-known name (default: inbox)'),
      maxMessages: z.number().int().min(1).max(2000).optional().describe(`Messages to scan (default ${DEFAULT_MAX}, max 2000)`),
      includeNoUnsubscribe: z.boolean().optional().describe('Also list senders without a List-Unsubscribe header'),
      excludeDomains: z.array(z.string().min(1)).optional().describe('Sender domains to skip, e.g. ["github.com"]'),
    },
    mutating: false,
    async handler(a, { graph }) {
      const max: number = a.maxMessages ?? DEFAULT_MAX;
      const path = `/me/mailFolders/${encodeURIComponent(wellKnownFolder(a.folderId ?? 'inbox'))}/messages`;
      const messages: Message[] = [];
      let page = await graph.get<Page<Message>>(path, {
        query: {
          $top: Math.min(100, max),
          $select: 'id,from,subject,receivedDateTime,isRead',
          $orderby: 'receivedDateTime desc',
        },
      });
      for (;;) {
        messages.push(...page.value);
        const next = page['@odata.nextLink'];
        if (messages.length >= max || !next) break;
        page = await graph.get<Page<Message>>(next);
      }
      const scanned = messages.slice(0, max);
      const groups = groupBySender(scanned, { excludeDomains: a.excludeDomains });

      let lookupErrors = 0;
      const kinds = await mapLimit(groups, LOOKUP_CONCURRENCY, async (g): Promise<UnsubscribeKind> => {
        try {
          const m = await graph.get<{ internetMessageHeaders?: { name: string; value: string }[] }>(
            `/me/messages/${encodeURIComponent(g.sampleMessageId)}`,
            { query: { $select: 'internetMessageHeaders' } },
          );
          return unsubscribeKind(parseUnsubscribeHeaders(m.internetMessageHeaders ?? []));
        } catch {
          lookupErrors++;
          return 'none';
        }
      });

      const senders = groups
        .map((g: SenderGroup, i) => ({ ...g, unsubscribe: kinds[i] }))
        .filter((s) => a.includeNoUnsubscribe || s.unsubscribe !== 'none');
      return {
        scanned: scanned.length,
        sendersTotal: groups.length,
        newsletterSenders: senders.filter((s) => s.unsubscribe !== 'none').length,
        lookupErrors: lookupErrors || undefined,
        senders,
      };
    },
  },
];
