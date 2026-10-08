import { z } from 'zod';
import type { Message, Page } from '../graph/types.js';
import { messageSummary, recipients } from '../util/format.js';
import { assertSearchCompatible, buildFilter, escapeSearchQuery, wellKnownFolder } from '../util/odata.js';
import { htmlToText, truncate } from '../util/sanitize.js';
import type { ToolDef } from './index.js';

const SUMMARY_FIELDS = 'id,subject,from,receivedDateTime,isRead,hasAttachments,importance,bodyPreview';
const DETAIL_FIELDS =
  'id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,isRead,flag,importance,hasAttachments,internetMessageId,webLink,body,parentFolderId';

const messagesPath = (folderId?: string) =>
  `/me/mailFolders/${encodeURIComponent(wellKnownFolder(folderId ?? 'inbox'))}/messages`;

function pageResult(page: Page<Message>) {
  return { messages: page.value.map(messageSummary), nextLink: page['@odata.nextLink'] };
}

export const messageTools: ToolDef[] = [
  {
    name: 'list_messages',
    description:
      'List messages in a folder (default inbox), newest first. Supports filters. Pass nextLink from a previous response to get the next page.',
    schema: {
      folderId: z.string().optional().describe('Folder id or well-known name (default: inbox)'),
      top: z.number().int().min(1).max(100).optional().describe('Page size'),
      unreadOnly: z.boolean().optional(),
      from: z.string().optional().describe('Exact sender email address'),
      since: z.string().optional().describe('Received on/after (ISO 8601)'),
      until: z.string().optional().describe('Received on/before (ISO 8601)'),
      hasAttachments: z.boolean().optional(),
      nextLink: z.string().optional().describe('Opaque nextLink from a previous response'),
    },
    mutating: false,
    async handler(a, { graph, cfg }) {
      if (a.nextLink) return pageResult(await graph.get<Page<Message>>(a.nextLink));
      const query = {
        $top: a.top ?? cfg.defaultTop,
        $select: SUMMARY_FIELDS,
        $orderby: 'receivedDateTime desc',
        $filter: buildFilter(a),
      };
      return pageResult(await graph.get<Page<Message>>(messagesPath(a.folderId), { query }));
    },
  },
  {
    name: 'search_messages',
    description:
      'Full-text search using KQL (e.g. from:alice subject:"report" hasattachments:true). Results are not sorted and Graph caps search at about 250 results. Pass nextLink for more.',
    schema: {
      query: z.string().min(1).describe('KQL search query'),
      folderId: z.string().optional().describe('Restrict to a folder (default: whole mailbox)'),
      top: z.number().int().min(1).max(100).optional(),
      nextLink: z.string().optional(),
    },
    mutating: false,
    async handler(a, { graph, cfg }) {
      const headers = { ConsistencyLevel: 'eventual' };
      if (a.nextLink) return pageResult(await graph.get<Page<Message>>(a.nextLink, { headers }));
      const query = { $search: escapeSearchQuery(a.query), $top: a.top ?? cfg.defaultTop, $select: SUMMARY_FIELDS };
      assertSearchCompatible(query);
      const path = a.folderId ? messagesPath(a.folderId) : '/me/messages';
      return pageResult(await graph.get<Page<Message>>(path, { query, headers }));
    },
  },
  {
    name: 'get_message',
    description: 'Get one message with headers, recipients and body (plain text by default; HTML is sanitized and truncated).',
    schema: {
      id: z.string().min(1),
      format: z.enum(['text', 'html']).optional().describe('Body format (default text)'),
      maxChars: z.number().int().min(100).optional().describe('Max body characters returned'),
    },
    mutating: false,
    async handler(a, { graph, cfg }) {
      const format = a.format ?? 'text';
      const m = await graph.get<Message>(`/me/messages/${encodeURIComponent(a.id)}`, {
        query: { $select: DETAIL_FIELDS },
        headers: format === 'text' ? { Prefer: 'outlook.body-content-type="text"' } : undefined,
      });
      const max = a.maxChars ?? cfg.maxBodyChars;
      let body = m.body?.content ?? '';
      if (m.body?.contentType === 'html' && format === 'text') body = htmlToText(body);
      else if (format === 'text') body = body.replace(/\r\n?/g, '\n').trim();
      return {
        id: m.id,
        subject: m.subject ?? '(no subject)',
        from: messageSummary(m).from,
        to: recipients(m.toRecipients),
        cc: recipients(m.ccRecipients),
        bcc: recipients(m.bccRecipients),
        receivedDateTime: m.receivedDateTime,
        sentDateTime: m.sentDateTime,
        isRead: m.isRead,
        flag: m.flag?.flagStatus,
        importance: m.importance,
        hasAttachments: m.hasAttachments,
        internetMessageId: m.internetMessageId,
        webLink: m.webLink,
        bodyFormat: format,
        body: truncate(body, max),
      };
    },
  },
];
