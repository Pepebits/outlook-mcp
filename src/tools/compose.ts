import fs from 'node:fs/promises';
import path from 'node:path';
import { z } from 'zod';
import type { GraphClient } from '../graph/client.js';
import { mapGraphError } from '../graph/errors.js';
import type { Message } from '../graph/types.js';
import { resolveAttachmentPath } from '../util/paths.js';
import { uploadFileInChunks } from '../util/upload.js';
import type { ToolDef } from './index.js';

/**
 * Graph caps a JSON request at about 4 MB; base64 adds a third. Inline attachments share this raw
 * budget, and anything that does not fit goes through an upload session.
 */
const INLINE_BUDGET_BYTES = 3 * 1024 * 1024;
const MAX_ATTACHMENT_BYTES = 150 * 1024 * 1024;

const toRecipients = (list?: string[]) => (list ?? []).map((address) => ({ emailAddress: { address } }));

interface LargeAttachment {
  path: string;
  name: string;
  size: number;
}

/**
 * Validates every file against the allowed roots (ADR 0002). Small files become inline base64
 * attachments while they fit the shared inline budget; the rest need an upload session.
 */
async function buildAttachments(items: { path: string }[] | undefined, roots: string[]) {
  const files: LargeAttachment[] = [];
  for (const item of items ?? []) {
    const real = await resolveAttachmentPath(roots, item.path);
    const stat = await fs.stat(real);
    if (!stat.isFile()) throw new Error(`Attachment is not a file: ${item.path}`);
    const name = path.basename(item.path);
    if (stat.size > MAX_ATTACHMENT_BYTES) {
      throw new Error(`Attachment "${name}" is ${stat.size} bytes; the maximum is 150 MB.`);
    }
    files.push({ path: real, name, size: stat.size });
  }
  const inline = [];
  const large: LargeAttachment[] = [];
  let budget = INLINE_BUDGET_BYTES;
  // Smallest first, so as many files as possible travel inline; output keeps the caller's order.
  const inlineSet = new Set<LargeAttachment>();
  for (const f of [...files].sort((x, y) => x.size - y.size)) {
    if (f.size < budget) {
      inlineSet.add(f);
      budget -= f.size;
    }
  }
  for (const f of files) {
    if (!inlineSet.has(f)) {
      large.push(f);
      continue;
    }
    inline.push({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: f.name,
      contentBytes: (await fs.readFile(f.path)).toString('base64'),
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
  from: z.string().email().optional().describe('From address: must be the signed-in mailbox or one of its aliases / SendAs identities'),
  fromName: z.string().max(100).optional().describe('Display name shown with the From address (e.g. Fox); requires from'),
  attachments: z
    .array(z.object({ path: z.string() }))
    .optional()
    .describe('Local files inside OUTLOOK_ATTACH_DIRS (default: download dir, ~/Documents, ~/Desktop); hidden files are refused. Up to 150 MB each; files that do not fit the ~3 MB inline budget go through an upload session'),
};

async function buildMessage(a: any, roots: string[]) {
  if (a.fromName !== undefined && a.from === undefined) throw new Error('fromName needs from: set the From address as well.');
  const message: Record<string, unknown> = {
    subject: a.subject,
    body: { contentType: a.bodyType === 'html' ? 'HTML' : 'Text', content: a.body },
    toRecipients: toRecipients(a.to),
    ccRecipients: toRecipients(a.cc),
    bccRecipients: toRecipients(a.bcc),
  };
  if (a.from) message.from = { emailAddress: { address: a.from, ...(a.fromName ? { name: a.fromName } : {}) } };
  const { inline, large } = await buildAttachments(a.attachments, roots);
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
    description: 'Send an email immediately. This cannot be undone. Attachments that do not fit the ~3 MB inline budget (up to 150 MB each) are uploaded in chunks first; if an upload fails nothing is sent.',
    schema: { ...composeShape, saveToSentItems: z.boolean().optional().describe('Default true') },
    mutating: true,
    openWorld: true,
    async handler(a, { graph, cfg }) {
      const { message, large } = await buildMessage(a, cfg.attachDirs);
      if (!large.length) {
        await graph.post('/me/sendMail', { message, saveToSentItems: a.saveToSentItems ?? true });
        return { sent: true, to: a.to };
      }
      if (a.saveToSentItems === false) {
        throw new Error('saveToSentItems: false is not supported with attachments of 3 MB or more; the message is always saved to Sent Items.');
      }
      // Large attachments need a draft: create it, upload the files, then send it.
      const draft = await graph.post<Message>('/me/messages', message);
      const draftPath = `/me/messages/${encodeURIComponent(draft.id)}`;
      for (const file of large) {
        try {
          await uploadLarge(graph, draft.id, file);
        } catch (err) {
          // Leave nothing half-sent behind.
          await graph.delete(draftPath).catch(() => undefined);
          throw new Error(`Could not upload attachment "${file.name}": ${errText(err)}`);
        }
      }
      try {
        await graph.post(`${draftPath}/send`);
      } catch (err) {
        const status = (err as { status?: number }).status;
        if (status !== undefined && status < 500) {
          await graph.delete(draftPath).catch(() => undefined);
          throw err;
        }
        // A gateway error or a dropped connection may hide a send that went through, so the
        // draft is kept: if it is still in Drafts, it was not sent.
        throw new Error(`The send request failed (${errText(err)}). The message may have been sent: check Sent Items before retrying; if it was not sent, it is still in Drafts.`);
      }
      return { sent: true, to: a.to };
    },
  },
  {
    name: 'create_draft',
    description: 'Create a draft message in the Drafts folder (not sent). Attachments of 3 MB or more (up to 150 MB) are uploaded in chunks.',
    schema: composeShape,
    mutating: true,
    destructive: false,
    async handler(a, { graph, cfg }) {
      const { message, large } = await buildMessage(a, cfg.attachDirs);
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
    openWorld: true,
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
    openWorld: true,
    async handler(a, { graph }) {
      await graph.post(`/me/messages/${encodeURIComponent(a.id)}/forward`, {
        comment: a.comment ?? '',
        toRecipients: toRecipients(a.to),
      });
      return { sent: true, to: a.to };
    },
  },
];
