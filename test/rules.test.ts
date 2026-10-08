import { describe, expect, it } from 'vitest';
import { GraphError, mapGraphError } from '../src/graph/errors.js';
import { blockRuleName, blockTargets, buildRule, nextSequence } from '../src/util/rules.js';

describe('buildRule', () => {
  it('builds a move rule with conditions', () => {
    const r = buildRule({
      displayName: 'x',
      fromAddresses: ['A@b.com'],
      subjectContains: ['sale'],
      action: 'move',
      destinationFolderId: 'F1',
      sequence: 3,
    });
    expect(r).toEqual({
      displayName: 'x',
      sequence: 3,
      isEnabled: true,
      conditions: { fromAddresses: [{ emailAddress: { address: 'A@b.com' } }], subjectContains: ['sale'] },
      actions: { moveToFolder: 'F1', stopProcessingRules: true },
    });
  });

  it('maps delete, markRead and stopProcessingRules=false', () => {
    const base = { displayName: 'x', senderContains: ['@e.com'], sequence: 1 };
    expect(buildRule({ ...base, action: 'delete' }).actions).toEqual({ delete: true, stopProcessingRules: true });
    expect(buildRule({ ...base, action: 'markRead', stopProcessingRules: false }).actions).toEqual({ markAsRead: true });
  });

  it('rejects rules without conditions or without a destination', () => {
    expect(() => buildRule({ displayName: 'x', action: 'delete', sequence: 1 })).toThrow(/at least one condition/);
    expect(() => buildRule({ displayName: 'x', senderContains: ['a'], action: 'move', sequence: 1 })).toThrow(/destination/);
    expect(() => buildRule({ displayName: 'x', senderContains: [' '], action: 'delete', sequence: 1 })).toThrow();
  });
});

describe('block helpers', () => {
  it('normalises addresses and domains', () => {
    expect(blockTargets([' Spam@X.com '], ['@Ads.example'])).toEqual({ fromAddresses: ['spam@x.com'], senderContains: ['@ads.example'] });
    expect(() => blockTargets([], [])).toThrow();
    expect(() => blockTargets(['nope'], [])).toThrow(/full email/);
  });

  it('names the rule with the prefix and truncates', () => {
    expect(blockRuleName(['a@b.com'], ['c.com'])).toBe('Blocked by outlook-mcp: a@b.com, c.com');
    expect(blockRuleName(['x'.repeat(300) + '@b.com'], []).length).toBe(120);
  });

  it('picks the next sequence', () => {
    expect(nextSequence([])).toBe(1);
    expect(nextSequence([{ sequence: 2 }, { sequence: 5 }, {}])).toBe(6);
  });
});

describe('rules permission error', () => {
  it('tells the user to add MailboxSettings.ReadWrite', () => {
    const e = mapGraphError(new GraphError('denied', 403, 'ErrorAccessDenied', '/me/mailFolders/inbox/messageRules'));
    expect(e.message).toMatch(/MailboxSettings\.ReadWrite/);
    expect(e.message).toMatch(/npm run auth/);
  });
});
