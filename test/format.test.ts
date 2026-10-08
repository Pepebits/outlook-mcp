import { describe, expect, it } from 'vitest';
import { formatRecipient, messageSummary, recipients, toTextResult } from '../src/util/format.js';

describe('format', () => {
  it('formats recipients', () => {
    expect(formatRecipient({ emailAddress: { name: 'Ann', address: 'a@x.com' } })).toBe('Ann <a@x.com>');
    expect(recipients([{ emailAddress: { address: 'b@x.com' } }])).toEqual(['b@x.com']);
    expect(recipients(undefined)).toEqual([]);
  });
  it('summarizes messages', () => {
    const s = messageSummary({ id: '1', subject: null, from: { emailAddress: { address: 'a@x.com' } } });
    expect(s.subject).toBe('(no subject)');
    expect(s.from).toBe('a@x.com');
  });
  it('builds text results', () => {
    expect(toTextResult({ a: 1 }).content[0].text).toContain('"a": 1');
    expect(toTextResult('boom', true).isError).toBe(true);
  });
});
