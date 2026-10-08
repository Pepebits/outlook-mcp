import os from 'node:os';
import path from 'node:path';

/**
 * Resolve `target` (relative paths are placed inside `baseDir`) and make sure
 * the result stays within `baseDir`. Throws on traversal attempts.
 */
export function resolveSafePath(baseDir: string, target: string): string {
  const base = path.resolve(baseDir);
  let t = target;
  if (t === '~') t = os.homedir();
  else if (t.startsWith('~/')) t = path.join(os.homedir(), t.slice(2));
  const resolved = path.resolve(base, t);
  const rel = path.relative(base, resolved);
  if (rel === '' || rel.startsWith('..') || path.isAbsolute(rel)) {
    throw new Error(`Path "${target}" is outside the allowed download directory (${base}).`);
  }
  return resolved;
}

/** Make an attachment name safe to use as a file name. */
export function safeFileName(name: string): string {
  const cleaned = path.basename(name).replace(/[\x00-\x1f<>:"/\\|?*]/g, '_').trim();
  return cleaned && cleaned !== '.' && cleaned !== '..' ? cleaned : 'attachment';
}
