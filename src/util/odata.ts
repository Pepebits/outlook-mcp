export const WELL_KNOWN_FOLDERS: Record<string, string> = {
  inbox: 'inbox',
  drafts: 'drafts',
  sent: 'sentitems',
  sentitems: 'sentitems',
  'sent items': 'sentitems',
  deleted: 'deleteditems',
  deleteditems: 'deleteditems',
  trash: 'deleteditems',
  junk: 'junkemail',
  junkemail: 'junkemail',
  spam: 'junkemail',
  archive: 'archive',
  outbox: 'outbox',
};

/** Map friendly names to Graph well-known folder names; other ids pass through. */
export function wellKnownFolder(idOrName: string): string {
  return WELL_KNOWN_FOLDERS[idOrName.trim().toLowerCase()] ?? idOrName;
}

export function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}

export interface FilterInput {
  unreadOnly?: boolean;
  from?: string;
  since?: string;
  until?: string;
  hasAttachments?: boolean;
}

function isoDate(value: string, label: string): string {
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) throw new Error(`Invalid ${label} date: "${value}". Use ISO 8601, e.g. 2025-01-31.`);
  return d.toISOString();
}

/** Build an OData $filter expression, or undefined when no criteria are given. */
export function buildFilter(f: FilterInput): string | undefined {
  const parts: string[] = [];
  if (f.unreadOnly) parts.push('isRead eq false');
  if (f.from) parts.push(`from/emailAddress/address eq '${escapeODataString(f.from)}'`);
  if (f.since) parts.push(`receivedDateTime ge ${isoDate(f.since, 'since')}`);
  if (f.until) parts.push(`receivedDateTime le ${isoDate(f.until, 'until')}`);
  if (f.hasAttachments !== undefined) parts.push(`hasAttachments eq ${f.hasAttachments}`);
  return parts.length ? parts.join(' and ') : undefined;
}

/** Quote a search query for $search, escaping backslashes and double quotes. */
export function escapeSearchQuery(query: string): string {
  const q = query.trim();
  if (!q) throw new Error('Search query must not be empty.');
  return `"${q.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;
}

/** Graph rejects $search combined with $filter or $orderby on messages. */
export function assertSearchCompatible(params: Record<string, unknown>): void {
  if (params['$search'] !== undefined && (params['$filter'] !== undefined || params['$orderby'] !== undefined)) {
    throw new Error('$search cannot be combined with $filter or $orderby.');
  }
}
