import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';

function expandTilde(p: string): string {
  if (p === '~') return os.homedir();
  if (p.startsWith('~/')) return path.join(os.homedir(), p.slice(2));
  return p;
}

const isInside = (root: string, target: string) => {
  const rel = path.relative(root, target);
  return rel !== '' && !rel.startsWith('..') && !path.isAbsolute(rel);
};

/** realpath of `p`, or of its nearest existing ancestor joined with the missing tail. */
async function realpathLoose(p: string): Promise<string> {
  const tail: string[] = [];
  let cur = p;
  for (;;) {
    try {
      return path.join(await fs.realpath(cur), ...tail.reverse());
    } catch (err) {
      if ((err as NodeJS.ErrnoException).code !== 'ENOENT') throw err;
      const parent = path.dirname(cur);
      if (parent === cur) return p;
      tail.push(path.basename(cur));
      cur = parent;
    }
  }
}

/**
 * Resolve `target` (relative paths are placed inside `baseDir`) and make sure
 * the result stays within `baseDir`. Lexical check only; see resolveSafeWritePath.
 */
export function resolveSafePath(baseDir: string, target: string): string {
  const base = path.resolve(baseDir);
  const resolved = path.resolve(base, expandTilde(target));
  if (!isInside(base, resolved)) {
    throw new Error(`Path "${target}" is outside the allowed download directory (${base}).`);
  }
  return resolved;
}

/**
 * resolveSafePath plus a real-path check, so a symlinked directory inside the download
 * directory cannot lead the write outside it. Returns the real destination path.
 */
export async function resolveSafeWritePath(baseDir: string, target: string): Promise<string> {
  return assertRealInside(baseDir, resolveSafePath(baseDir, target), target);
}

/** Resolves `p` through symlinks (as far as it exists) and checks it is inside the real `baseDir`. */
export async function assertRealInside(baseDir: string, p: string, label = p): Promise<string> {
  const realBase = await realpathLoose(path.resolve(baseDir));
  const real = await realpathLoose(p);
  if (!isInside(realBase, real)) {
    throw new Error(`Path "${label}" resolves outside the allowed download directory (${realBase}).`);
  }
  return real;
}

/**
 * Validates a local file for sending as an attachment (ADR 0002): its real path must be inside one
 * of `roots` (also resolved), and no component below the root may be a dotfile or dot-directory.
 * Returns the real path, which is the one that must be read.
 */
export async function resolveAttachmentPath(roots: string[], target: string): Promise<string> {
  const lexical = path.resolve(expandTilde(target));
  let real: string;
  try {
    real = await fs.realpath(lexical);
  } catch {
    throw new Error(`Attachment file not found: ${target}`);
  }
  for (const root of roots) {
    const realRoot = await realpathLoose(path.resolve(root));
    if (isInside(realRoot, real)) {
      const below = path.relative(realRoot, real).split(path.sep);
      if (below.some((part) => part.startsWith('.'))) {
        throw new Error(`Attachment "${target}" is a hidden file or inside a hidden directory; those are never sent.`);
      }
      return real;
    }
  }
  throw new Error(
    `Attachment "${target}" is outside the allowed directories (${roots.join(path.delimiter)}). Move the file there or add its directory to OUTLOOK_ATTACH_DIRS.`,
  );
}

/** Make an attachment name safe to use as a file name. */
export function safeFileName(name: string): string {
  const cleaned = path.basename(name).replace(/[\x00-\x1f<>:"/\\|?*]/g, '_').trim();
  return cleaned && cleaned !== '.' && cleaned !== '..' ? cleaned : 'attachment';
}
