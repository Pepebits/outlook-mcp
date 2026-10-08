import { describe, expect, it } from 'vitest';
import { GraphError } from '../src/graph/errors.js';
import { resolveIds, runBulk } from '../src/tools/bulk.js';

describe('resolveIds', () => {
  it('requires exactly one of id / ids', () => {
    expect(() => resolveIds({})).toThrow(/exactly one/);
    expect(() => resolveIds({ id: 'a', ids: ['b'] })).toThrow(/exactly one/);
    expect(resolveIds({ id: 'a' })).toEqual({ ids: ['a'], single: true });
    expect(resolveIds({ ids: ['a', 'b'] })).toEqual({ ids: ['a', 'b'], single: false });
  });

  it('rejects empty and oversized lists', () => {
    expect(() => resolveIds({ ids: [] })).toThrow();
    expect(() => resolveIds({ ids: Array.from({ length: 51 }, (_, i) => String(i)) })).toThrow(/at most 50/);
  });
});

describe('runBulk', () => {
  it('returns the plain result for a single id and propagates its errors', async () => {
    expect(await runBulk({ id: 'a' }, async (id) => ({ newId: `${id}2` }))).toEqual({ newId: 'a2' });
    await expect(
      runBulk({ id: 'a' }, async () => {
        throw new Error('boom');
      }),
    ).rejects.toThrow('boom');
  });

  it('keeps going after errors and preserves order', async () => {
    const out = (await runBulk({ ids: ['a', 'b', 'c', 'd', 'e'] }, async (id) => {
      if (id === 'b') throw new GraphError('gone', 404, 'ErrorItemNotFound');
      if (id === 'd') throw new Error('bad');
      return { newId: `${id}2` };
    })) as { results: { id: string; ok: boolean; newId?: string; error?: string }[]; succeeded: number; failed: number };
    expect(out.results.map((r) => r.id)).toEqual(['a', 'b', 'c', 'd', 'e']);
    expect(out.results.map((r) => r.ok)).toEqual([true, false, true, false, true]);
    expect(out.results[0].newId).toBe('a2');
    expect(out.results[1].error).toMatch(/not found/i);
    expect(out.results[3].error).toBe('bad');
    expect(out.succeeded).toBe(3);
    expect(out.failed).toBe(2);
  });

  it('limits concurrency', async () => {
    let active = 0;
    let peak = 0;
    await runBulk({ ids: Array.from({ length: 12 }, (_, i) => String(i)) }, async () => {
      active++;
      peak = Math.max(peak, active);
      await new Promise((r) => setTimeout(r, 2));
      active--;
      return {};
    });
    expect(peak).toBeLessThanOrEqual(4);
  });
});
