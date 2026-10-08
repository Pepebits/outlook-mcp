import { config as loadDotenv } from 'dotenv';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { z } from 'zod';

/**
 * Shared public-client app registration for personal Microsoft accounts. It has no
 * secret: tokens are issued to each user and stay on their machine. Set
 * OUTLOOK_CLIENT_ID to use your own app registration instead.
 */
export const DEFAULT_CLIENT_ID = '056314dd-a452-4fee-8971-77459938c7b7';

export interface Config {
  clientId: string;
  tenant: string;
  scopes: string[];
  tokenCachePath: string;
  readOnly: boolean;
  downloadDir: string;
  defaultTop: number;
  maxBodyChars: number;
  graphBaseUrl: string;
}

const schema = z.object({
  OUTLOOK_CLIENT_ID: z.string().trim().optional().transform((v) => v || DEFAULT_CLIENT_ID),
  OUTLOOK_TENANT: z.string().trim().min(1).default('consumers'),
  OUTLOOK_SCOPES: z.string().trim().min(1).default('User.Read Mail.ReadWrite Mail.Send MailboxSettings.ReadWrite offline_access'),
  OUTLOOK_TOKEN_CACHE: z.string().trim().optional(),
  OUTLOOK_READ_ONLY: z.string().trim().default('false'),
  OUTLOOK_DOWNLOAD_DIR: z.string().trim().optional(),
  OUTLOOK_DEFAULT_TOP: z.coerce.number().int().min(1).max(100).default(20),
  OUTLOOK_MAX_BODY_CHARS: z.coerce.number().int().min(100).default(20000),
  OUTLOOK_GRAPH_BASE_URL: z.string().trim().url().default('https://graph.microsoft.com/v1.0'),
});

/** Treat empty strings (e.g. `VAR=` in .env) as unset. */
function clean(env: NodeJS.ProcessEnv): Record<string, string | undefined> {
  const out: Record<string, string | undefined> = {};
  for (const [k, v] of Object.entries(env)) out[k] = v === undefined || v.trim() === '' ? undefined : v;
  return out;
}

export function expandHome(p: string): string {
  if (p === '~') return os.homedir();
  if (p.startsWith('~/')) return path.join(os.homedir(), p.slice(2));
  return p;
}

export function parseConfig(env: NodeJS.ProcessEnv): Config {
  const parsed = schema.safeParse(clean(env));
  if (!parsed.success) {
    const lines = parsed.error.issues.map((i) => `  - ${i.path.join('.') || 'config'}: ${i.message}`);
    throw new Error(`Invalid configuration:\n${lines.join('\n')}`);
  }
  const e = parsed.data;
  return {
    clientId: e.OUTLOOK_CLIENT_ID,
    tenant: e.OUTLOOK_TENANT,
    scopes: e.OUTLOOK_SCOPES.split(/[\s,]+/).filter(Boolean),
    tokenCachePath: path.resolve(
      expandHome(e.OUTLOOK_TOKEN_CACHE ?? path.join(os.homedir(), '.config', 'outlook-mcp', 'token-cache.json')),
    ),
    readOnly: ['1', 'true', 'yes', 'on'].includes(e.OUTLOOK_READ_ONLY.toLowerCase()),
    downloadDir: path.resolve(expandHome(e.OUTLOOK_DOWNLOAD_DIR ?? path.join(os.homedir(), 'Downloads'))),
    defaultTop: e.OUTLOOK_DEFAULT_TOP,
    maxBodyChars: e.OUTLOOK_MAX_BODY_CHARS,
    graphBaseUrl: e.OUTLOOK_GRAPH_BASE_URL.replace(/\/+$/, ''),
  };
}

/**
 * Load .env from the current directory and from the package directory (so that
 * launching with an absolute path from any cwd works). Real env vars win.
 */
export function loadConfig(): Config {
  const pkgDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
  loadDotenv({ path: path.resolve(process.cwd(), '.env'), quiet: true });
  loadDotenv({ path: path.join(pkgDir, '.env'), quiet: true });
  try {
    return parseConfig(process.env);
  } catch (err) {
    process.stderr.write(`[outlook-mcp] ${(err as Error).message}\n`);
    process.exit(1);
  }
}
