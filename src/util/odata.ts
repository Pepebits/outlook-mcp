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

export function isoDate(value: string, label: string): string {
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) throw new Error(`Invalid ${label} date: "${value}". Use ISO 8601, e.g. 2025-01-31.`);
  return d.toISOString();
}

/** An inclusive upper bound: a date without a time ("2025-01-31") covers that whole day (UTC). */
export function untilDate(value: string): string {
  if (/^\d{4}-\d{2}-\d{2}$/.test(value.trim())) return isoDate(`${value.trim()}T23:59:59.999Z`, 'until');
  return isoDate(value, 'until');
}

/**
 * Builds $filter for a listing sorted by receivedDateTime desc. Exchange rejects
 * such a sort (InefficientFilter) unless receivedDateTime is filtered first, so a
 * receivedDateTime clause always leads, with an open lower bound when none is given.
 */
export function buildFilter(f: FilterInput): string | undefined {
  const parts: string[] = [];
  if (f.unreadOnly) parts.push('isRead eq false');
  if (f.from) parts.push(`from/emailAddress/address eq '${escapeODataString(f.from)}'`);
  if (f.hasAttachments !== undefined) parts.push(`hasAttachments eq ${f.hasAttachments}`);
  const dates: string[] = [];
  if (f.since) dates.push(`receivedDateTime ge ${isoDate(f.since, 'since')}`);
  if (f.until) dates.push(`receivedDateTime le ${untilDate(f.until)}`);
  if (!parts.length) return dates.length ? dates.join(' and ') : undefined;
  if (!f.since) dates.unshift('receivedDateTime ge 1900-01-01T00:00:00Z');
  return [...dates, ...parts].join(' and ');
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
