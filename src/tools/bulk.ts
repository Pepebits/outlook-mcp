import { z } from 'zod';
import { mapGraphError } from '../graph/errors.js';

export const MAX_BULK_IDS = 50;
const BULK_CONCURRENCY = 4;

/** Zod shape fragment for tools that accept one message (`id`) or many (`ids`). */
export const idOrIds = {
  id: z.string().min(1).optional().describe('A single message id'),
  ids: z
    .array(z.string().min(1))
    .min(1)
    .max(MAX_BULK_IDS)
    .optional()
    .describe(`Several message ids (up to ${MAX_BULK_IDS}); give either id or ids, not both`),
};

export interface BulkItem {
  id: string;
  ok: boolean;
  error?: string;
  [key: string]: unknown;
}

export interface BulkResult {
  results: BulkItem[];
  succeeded: number;
  failed: number;
}

/** Validates that exactly one of id / ids is given. */
export function resolveIds(a: { id?: string; ids?: string[] }): { ids: string[]; single: boolean } {
  const hasId = a.id !== undefined;
  const hasIds = a.ids !== undefined;
  if (hasId === hasIds) throw new Error('Provide exactly one of "id" or "ids".');
  if (hasIds) {
    if (!a.ids!.length) throw new Error('"ids" must not be empty.');
    if (a.ids!.length > MAX_BULK_IDS) throw new Error(`"ids" accepts at most ${MAX_BULK_IDS} ids.`);
    return { ids: a.ids!, single: false };
  }
  return { ids: [a.id!], single: true };
}

/**
 * Runs `fn` for one id (returning its result unchanged, errors propagate) or for
 * many ids (never stopping on the first error; results keep the input order).
 */
export async function runBulk(
  a: { id?: string; ids?: string[] },
  fn: (id: string) => Promise<unknown>,
  concurrency = BULK_CONCURRENCY,
): Promise<unknown> {
  const { ids, single } = resolveIds(a);
  if (single) return fn(ids[0]);

  const results = new Array<BulkItem>(ids.length);
  let next = 0;
  const worker = async () => {
    while (next < ids.length) {
      const i = next++;
      try {
        const out = await fn(ids[i]);
        results[i] = { ...(out && typeof out === 'object' ? (out as object) : {}), id: ids[i], ok: true };
      } catch (err) {
        results[i] = { id: ids[i], ok: false, error: mapGraphError(err).message };
      }
    }
  };
  await Promise.all(Array.from({ length: Math.min(concurrency, ids.length) }, worker));
  const succeeded = results.filter((r) => r.ok).length;
  return { results, succeeded, failed: results.length - succeeded } satisfies BulkResult;
}
