import fs from 'node:fs/promises';
import path from 'node:path';
import { z } from 'zod';
import type { Message } from '../graph/types.js';
import type { ToolDef } from './index.js';

const MAX_ATTACHMENT_BYTES = 3 * 1024 * 1024;

const toRecipients = (list?: string[]) => (list ?? []).map((address) => ({ emailAddress: { address } }));

async function buildAttachments(items?: { path: string }[]) {
  const out = [];
  for (const item of items ?? []) {
    const stat = await fs.stat(item.path).catch(() => {
      throw new Error(`Attachment file not found: ${item.path}`);
    });
    if (!stat.isFile()) throw new Error(`Attachment is not a file: ${item.path}`);
    if (stat.size >= MAX_ATTACHMENT_BYTES) {
      throw new Error(
        `Attachment "${path.basename(item.path)}" is ${stat.size} bytes; only files under 3 MB are supported (large-file upload sessions are on the roadmap).`,
      );
    }
    out.push({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: path.basename(item.path),
      contentBytes: (await fs.readFile(item.path)).toString('base64'),
    });
  }
  return out;
}

const composeShape = {
  to: z.array(z.string().email()).min(1).describe('To recipients'),
  cc: z.array(z.string().email()).optional(),
  bcc: z.array(z.string().email()).optional(),
  subject: z.string(),
  body: z.string(),
  bodyType: z.enum(['text', 'html']).optional().describe('Default text'),
  attachments: z.array(z.object({ path: z.string() })).optional().describe('Local files (< 3 MB each)'),
};

async function buildMessage(a: any) {
  const message: Record<string, unknown> = {
    subject: a.subject,
    body: { contentType: a.bodyType === 'html' ? 'HTML' : 'Text', content: a.body },
    toRecipients: toRecipients(a.to),
    ccRecipients: toRecipients(a.cc),
    bccRecipients: toRecipients(a.bcc),
  };
  const attachments = await buildAttachments(a.attachments);
  if (attachments.length) message.attachments = attachments;
  return message;
}

export const composeTools: ToolDef[] = [
  {
    name: 'send_mail',
    description: 'Send an email immediately. This cannot be undone.',
    schema: { ...composeShape, saveToSentItems: z.boolean().optional().describe('Default true') },
    mutating: true,
    async handler(a, { graph }) {
      await graph.post('/me/sendMail', {
        message: await buildMessage(a),
        saveToSentItems: a.saveToSentItems ?? true,
      });
      return { sent: true, to: a.to };
    },
  },
  {
    name: 'create_draft',
    description: 'Create a draft message in the Drafts folder (not sent).',
    schema: composeShape,
    mutating: true,
    async handler(a, { graph }) {
      const m = await graph.post<Message>('/me/messages', await buildMessage(a));
      return { draftId: m.id, webLink: m.webLink };
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
