import { describe, expect, it } from 'vitest';
import { assertSearchCompatible, buildFilter, escapeODataString, escapeSearchQuery, wellKnownFolder } from '../src/util/odata.js';

describe('odata', () => {
  it('escapes single quotes', () => expect(escapeODataString("o'brien")).toBe("o''brien"));
  it('builds combined filters', () => {
    const f = buildFilter({ unreadOnly: true, from: "a'b@x.com", since: '2025-01-01T00:00:00Z', until: '2025-02-01T00:00:00Z', hasAttachments: true });
    expect(f).toBe(
      "isRead eq false and from/emailAddress/address eq 'a''b@x.com' and receivedDateTime ge 2025-01-01T00:00:00.000Z and receivedDateTime le 2025-02-01T00:00:00.000Z and hasAttachments eq true",
    );
  });
  it('returns undefined for empty filter and rejects bad dates', () => {
    expect(buildFilter({})).toBeUndefined();
    expect(() => buildFilter({ since: 'nope' })).toThrow();
  });
  it('maps well-known folders', () => {
    expect(wellKnownFolder('Sent')).toBe('sentitems');
    expect(wellKnownFolder('AAMk123')).toBe('AAMk123');
  });
  it('escapes search queries', () => {
    expect(escapeSearchQuery('subject:"hi"')).toBe('"subject:\\"hi\\""');
    expect(() => escapeSearchQuery('  ')).toThrow();
  });
  it('guards $search combinations', () => {
    expect(() => assertSearchCompatible({ $search: 'x', $orderby: 'a' })).toThrow();
    expect(() => assertSearchCompatible({ $search: 'x', $filter: 'a' })).toThrow();
    expect(() => assertSearchCompatible({ $search: 'x' })).not.toThrow();
  });
});
