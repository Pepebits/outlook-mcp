import type { ICachePlugin, TokenCacheContext } from '@azure/msal-node';
import fs from 'node:fs/promises';
import path from 'node:path';

/** Persists the MSAL token cache to a file readable only by the current user. */
export class FileCachePlugin implements ICachePlugin {
  constructor(private readonly file: string) {}

  async beforeCacheAccess(ctx: TokenCacheContext): Promise<void> {
    try {
      ctx.tokenCache.deserialize(await fs.readFile(this.file, 'utf8'));
    } catch (err) {
      if ((err as NodeJS.ErrnoException).code !== 'ENOENT') throw err;
    }
  }

  async afterCacheAccess(ctx: TokenCacheContext): Promise<void> {
    if (!ctx.cacheHasChanged) return;
    await fs.mkdir(path.dirname(this.file), { recursive: true, mode: 0o700 });
    const tmp = `${this.file}.${process.pid}.tmp`;
    await fs.writeFile(tmp, ctx.tokenCache.serialize(), { mode: 0o600 });
    await fs.rename(tmp, this.file);
    await fs.chmod(this.file, 0o600);
  }
}
