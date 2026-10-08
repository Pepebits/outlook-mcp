import { z } from 'zod';
import type { GraphClient } from '../graph/client.js';
import { parseUnsubscribeHeaders, type UnsubscribeInfo } from '../util/unsubscribe.js';
import { untrustedFields } from '../util/untrusted.js';
import { idOrIds, runBulk } from './bulk.js';
import type { ToolDef } from './index.js';

interface HeaderMessage {
  subject?: string;
  from?: { emailAddress?: { name?: string; address?: string } };
  internetMessageHeaders?: { name: string; value: string }[];
}

async function loadInfo(graph: GraphClient, id: string) {
  const m = await graph.get<HeaderMessage>(`/me/messages/${encodeURIComponent(id)}`, {
    query: { $select: 'subject,from,internetMessageHeaders' },
  });
  const from = m.from?.emailAddress;
  return {
    // Sender-controlled text, kept apart from the structural fields.
    untrusted: untrustedFields({
      from: from ? `${from.name ?? ''} <${from.address ?? ''}>`.trim() : '',
      subject: m.subject ?? '',
    }),
    info: parseUnsubscribeHeaders(m.internetMessageHeaders ?? []),
  };
}

function describe(info: UnsubscribeInfo): string {
  if (info.oneClick) return 'one-click (RFC 8058): unsubscribe can do it automatically';
  if (info.mailto.length) return 'mailto: unsubscribe can send the request email automatically';
  if (info.https.length) return 'link only: open it in a browser to finish';
  return 'none: the sender does not advertise List-Unsubscribe';
}

export type OneClickResult = { ok: true; status: number } | { ok: false; note: string };

/** POSTs an RFC 8058 one-click request. Never throws: failures become a note with the fallback link advice. */
export async function postOneClick(url: string, fetchFn: typeof fetch = fetch): Promise<OneClickResult> {
  try {
    const res = await fetchFn(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: 'List-Unsubscribe=One-Click',
      redirect: 'follow',
      signal: AbortSignal.timeout(15000),
    });
    if (!res.ok) return { ok: false, note: `One-click was refused (HTTP ${res.status}). Open this link in a browser to finish.` };
    return { ok: true, status: res.status };
  } catch (err) {
    const cause = (err as { cause?: { message?: string } }).cause?.message;
    const reason = cause ?? (err instanceof Error ? err.message : String(err));
    return { ok: false, note: `One-click request failed (${reason}). Open this link in a browser to finish.` };
  }
}

export const unsubscribeTools: ToolDef[] = [
  {
    name: 'get_unsubscribe_info',
    description:
      "Read a message's List-Unsubscribe headers and report how to leave that mailing list (one-click, mailto or link). Does not change anything.",
    schema: { id: z.string().min(1) },
    mutating: false,
    async handler(a, { graph }) {
      const { untrusted, info } = await loadInfo(graph, a.id);
      return { untrusted, method: describe(info), ...info };
    },
  },
  {
    name: 'unsubscribe',
    description:
      'Unsubscribe from the mailing list a message came from, using its List-Unsubscribe header. Prefers RFC 8058 one-click (an HTTPS POST), then sends the mailto request from your account. If only a web link exists, returns it for the user to open. Accepts id or ids (up to 50). Never use this on spam or phishing: unsubscribing confirms the address is active.',
    schema: { ...idOrIds },
    mutating: true,
    async handler(a, { graph }) {
      return runBulk(a, async (id) => {
        const { untrusted, info } = await loadInfo(graph, id);
        if (info.oneClick) {
          const url = info.https[0];
          const result = await postOneClick(url);
          if (!result.ok) return { untrusted, done: false, method: 'link', url, note: result.note };
          return { untrusted, done: true, method: 'one-click', status: result.status };
        }
        const mailto = info.mailto[0];
        if (mailto) {
          await graph.post('/me/sendMail', {
            message: {
              subject: mailto.subject,
              body: { contentType: 'Text', content: mailto.body },
              toRecipients: [{ emailAddress: { address: mailto.address } }],
            },
            saveToSentItems: true,
          });
          return { untrusted, done: true, method: 'mailto', sentTo: mailto.address };
        }
        if (info.https.length) {
          return { untrusted, done: false, method: 'link', url: info.https[0], note: 'Open this link in a browser to finish.' };
        }
        return { untrusted, done: false, method: 'none', note: 'No List-Unsubscribe header; use the link in the email body or block the sender.' };
      });
    },
  },
];
