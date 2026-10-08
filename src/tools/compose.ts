import fs from 'node:fs/promises';
import path from 'node:path';
import { z } from 'zod';
import type { GraphClient } from '../graph/client.js';
import { mapGraphError } from '../graph/errors.js';
import type { Message } from '../graph/types.js';
import { uploadFileInChunks } from '../util/upload.js';
import type { ToolDef } from './index.js';

const INLINE_LIMIT_BYTES = 3 * 1024 * 1024;
const MAX_ATTACHMENT_BYTES = 150 * 1024 * 1024;

const toRecipients = (list?: string[]) => (list ?? []).map((address) => ({ emailAddress: { address } }));

interface LargeAttachment {
  path: string;
  name: string;
  size: number;
}

/** Validates every file; small ones become inline base64 attachments, large ones need an upload session. */
async function buildAttachments(items?: { path: string }[]) {
  const inline = [];
  const large: LargeAttachment[] = [];
  for (const item of items ?? []) {
    const stat = await fs.stat(item.path).catch(() => {
      throw new Error(`Attachment file not found: ${item.path}`);
    });
    if (!stat.isFile()) throw new Error(`Attachment is not a file: ${item.path}`);
    const name = path.basename(item.path);
    if (stat.size > MAX_ATTACHMENT_BYTES) {
      throw new Error(`Attachment "${name}" is ${stat.size} bytes; the maximum is 150 MB.`);
    }
    if (stat.size >= INLINE_LIMIT_BYTES) {
      large.push({ path: item.path, name, size: stat.size });
      continue;
    }
    inline.push({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name,
      contentBytes: (await fs.readFile(item.path)).toString('base64'),
    });
  }
  return { inline, large };
}

const composeShape = {
  to: z.array(z.string().email()).min(1).describe('To recipients'),
  cc: z.array(z.string().email()).optional(),
  bcc: z.array(z.string().email()).optional(),
  subject: z.string(),
  body: z.string(),
  bodyType: z.enum(['text', 'html']).optional().describe('Default text'),
  attachments: z.array(z.object({ path: z.string() })).optional().describe('Local files (up to 150 MB each; files of 3 MB or more are sent through an upload session)'),
};

async function buildMessage(a: any) {
  const message: Record<string, unknown> = {
    subject: a.subject,
    body: { contentType: a.bodyType === 'html' ? 'HTML' : 'Text', content: a.body },
    toRecipients: toRecipients(a.to),
    ccRecipients: toRecipients(a.cc),
    bccRecipients: toRecipients(a.bcc),
  };
  const { inline, large } = await buildAttachments(a.attachments);
  if (inline.length) message.attachments = inline;
  return { message, large };
}

/** Uploads one large file to a draft through an upload session. */
async function uploadLarge(graph: GraphClient, draftId: string, file: LargeAttachment) {
  const session = await graph.post<{ uploadUrl: string }>(`/me/messages/${encodeURIComponent(draftId)}/attachments/createUploadSession`, {
    AttachmentItem: { attachmentType: 'file', name: file.name, size: file.size },
  });
  if (!session?.uploadUrl) throw new Error('Graph did not return an upload URL.');
  await uploadFileInChunks(session.uploadUrl, file.path, file.size);
}

const errText = (err: unknown) => mapGraphError(err).message;

export const composeTools: ToolDef[] = [
  {
    name: 'send_mail',
    description: 'Send an email immediately. This cannot be undone. Attachments of 3 MB or more (up to 150 MB) are uploaded in chunks first; if an upload fails nothing is sent.',
    schema: { ...composeShape, saveToSentItems: z.boolean().optional().describe('Default true') },
    mutating: true,
    async handler(a, { graph }) {
      const { message, large } = await buildMessage(a);
      if (!large.length) {
        await graph.post('/me/sendMail', { message, saveToSentItems: a.saveToSentItems ?? true });
        return { sent: true, to: a.to };
      }
      if (a.saveToSentItems === false) {
        throw new Error('saveToSentItems: false is not supported with attachments of 3 MB or more; the message is always saved to Sent Items.');
      }
      // Large attachments need a draft: create it, upload the files, then send it.
      const draft = await graph.post<Message>('/me/messages', message);
      try {
        for (const file of large) {
          try {
            await uploadLarge(graph, draft.id, file);
          } catch (err) {
            throw new Error(`Could not upload attachment "${file.name}": ${errText(err)}`);
          }
        }
        await graph.post(`/me/messages/${encodeURIComponent(draft.id)}/send`);
      } catch (err) {
        // Leave nothing half-sent behind.
        await graph.delete(`/me/messages/${encodeURIComponent(draft.id)}`).catch(() => undefined);
        throw err;
      }
      return { sent: true, to: a.to };
    },
  },
  {
    name: 'create_draft',
    description: 'Create a draft message in the Drafts folder (not sent). Attachments of 3 MB or more (up to 150 MB) are uploaded in chunks.',
    schema: composeShape,
    mutating: true,
    async handler(a, { graph }) {
      const { message, large } = await buildMessage(a);
      const m = await graph.post<Message>('/me/messages', message);
      const failed: { name: string; error: string }[] = [];
      for (const file of large) {
        try {
          await uploadLarge(graph, m.id, file);
        } catch (err) {
          failed.push({ name: file.name, error: errText(err) });
        }
      }
      return failed.length ? { draftId: m.id, webLink: m.webLink, attachmentsFailed: failed } : { draftId: m.id, webLink: m.webLink };
    },
  },
  {
    name: 'reply_message',
    description: 'Reply (or reply-all) to a message. Sends immediately.',
    schema: { id: z.string().min(1), comment: z.string(), replyAll: z.boolean().optional() },
    mutating: true,
    async handler(a, { graph }) {
      const action = a.replyAll ? 'replyAll' : 'reply';
      await graph.post(`/me/messages/${encodeURIComponent(a.id)}/${action}`, { comment: a.comment });
      return { sent: true, action };
    },
  },
  {
    name: 'forward_message',
    description: 'Forward a message to the given recipients. Sends immediately.',
    schema: {
      id: z.string().min(1),
      to: z.array(z.string().email()).min(1),
      comment: z.string().optional(),
    },
    mutating: true,
    async handler(a, { graph }) {
      await graph.post(`/me/messages/${encodeURIComponent(a.id)}/forward`, {
        comment: a.comment ?? '',
        toRecipients: toRecipients(a.to),
      });
      return { sent: true, to: a.to };
    },
  },
];
