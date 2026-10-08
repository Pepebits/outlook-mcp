import type { Message, Recipient } from '../graph/types.js';

export function formatRecipient(r?: Recipient): string {
  const a = r?.emailAddress;
  if (!a) return '';
  return a.name && a.address ? `${a.name} <${a.address}>` : (a.address ?? a.name ?? '');
}

export function recipients(list?: Recipient[]): string[] {
  return (list ?? []).map(formatRecipient).filter(Boolean);
}

export function messageSummary(m: Message) {
  return {
    id: m.id,
    subject: m.subject ?? '(no subject)',
    from: formatRecipient(m.from),
    receivedDateTime: m.receivedDateTime,
    isRead: m.isRead,
    hasAttachments: m.hasAttachments,
    importance: m.importance,
    preview: m.bodyPreview,
  };
}

export function toTextResult(data: unknown, isError = false) {
  const text = typeof data === 'string' ? data : JSON.stringify(data, null, 2);
  return { content: [{ type: 'text' as const, text }], ...(isError ? { isError: true } : {}) };
}
