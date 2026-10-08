import os from 'node:os';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { resolveSafePath, safeFileName } from '../src/util/paths.js';

describe('paths', () => {
  const base = path.join(os.tmpdir(), 'dl');
  it('resolves relative paths inside base', () => {
    expect(resolveSafePath(base, 'a/b.txt')).toBe(path.join(base, 'a', 'b.txt'));
  });
  it('rejects traversal and outside absolute paths', () => {
    expect(() => resolveSafePath(base, '../evil')).toThrow();
    expect(() => resolveSafePath(base, '/etc/passwd')).toThrow();
    expect(() => resolveSafePath(base, '.')).toThrow();
  });
  it('accepts absolute paths inside base', () => {
    expect(resolveSafePath(base, path.join(base, 'x.pdf'))).toBe(path.join(base, 'x.pdf'));
  });
  it('sanitizes file names', () => {
    expect(safeFileName('../a:b.txt')).toBe('a_b.txt');
    expect(safeFileName('..')).toBe('attachment');
  });
});
