import { describe, expect, it } from 'vitest';
import { assertSearchCompatible, buildFilter, escapeODataString, escapeSearchQuery, wellKnownFolder } from '../src/util/odata.js';

describe('odata', () => {
  it('escapes single quotes', () => expect(escapeODataString("o'brien")).toBe("o''brien"));
  it('builds combined filters', () => {
    const f = buildFilter({ unreadOnly: true, from: "a'b@x.com", since: '2025-01-01T00:00:00Z', until: '2025-02-01T00:00:00Z', hasAttachments: true });
    expect(f).toBe(
      "receivedDateTime ge 2025-01-01T00:00:00.000Z and receivedDateTime le 2025-02-01T00:00:00.000Z and isRead eq false and from/emailAddress/address eq 'a''b@x.com' and hasAttachments eq true",
    );
  });
  it('leads with a receivedDateTime bound so the date sort is accepted', () => {
    expect(buildFilter({ from: 'a@x.com' })).toBe(
      "receivedDateTime ge 1900-01-01T00:00:00Z and from/emailAddress/address eq 'a@x.com'",
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
