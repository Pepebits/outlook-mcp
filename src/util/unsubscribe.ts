export interface UnsubscribeInfo {
  /** https links from List-Unsubscribe, in header order. */
  https: string[];
  /** mailto targets from List-Unsubscribe, in header order. */
  mailto: MailtoTarget[];
  /** True when the sender supports RFC 8058 one-click unsubscribe. */
  oneClick: boolean;
}

export interface MailtoTarget {
  address: string;
  subject: string;
  body: string;
}

interface Header {
  name: string;
  value: string;
}

function header(headers: Header[], name: string): string | undefined {
  return headers.find((h) => h.name.toLowerCase() === name.toLowerCase())?.value;
}

export function parseMailto(uri: string): MailtoTarget | null {
  let url: URL;
  try {
    url = new URL(uri);
  } catch {
    return null;
  }
  if (url.protocol !== 'mailto:') return null;
  const address = decodeURIComponent(url.pathname);
  if (!address.includes('@')) return null;
  return {
    address,
    subject: url.searchParams.get('subject') ?? 'unsubscribe',
    body: url.searchParams.get('body') ?? 'unsubscribe',
  };
}

/** Parses the List-Unsubscribe and List-Unsubscribe-Post headers (RFC 2369, RFC 8058). */
export function parseUnsubscribeHeaders(headers: Header[]): UnsubscribeInfo {
  const raw = header(headers, 'List-Unsubscribe') ?? '';
  const post = header(headers, 'List-Unsubscribe-Post') ?? '';
  const uris = [...raw.matchAll(/<([^>]+)>/g)].map((m) => m[1].trim());
  const https = uris.filter((u) => /^https:\/\//i.test(u));
  const mailto = uris.map(parseMailto).filter((m): m is MailtoTarget => m !== null);
  const oneClick = https.length > 0 && /List-Unsubscribe=One-Click/i.test(post);
  return { https, mailto, oneClick };
}
