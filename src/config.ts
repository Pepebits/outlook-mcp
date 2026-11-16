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
  tokenStore: 'auto' | 'keychain' | 'file';
  readOnly: boolean;
  downloadDir: string;
  /** Real-path roots that send_mail / create_draft may read attachments from (ADR 0002). */
  attachDirs: string[];
  defaultTop: number;
  maxBodyChars: number;
  graphBaseUrl: string;
}

export const READ_WRITE_SCOPES = 'User.Read Mail.ReadWrite Mail.Send MailboxSettings.ReadWrite offline_access';
/** Default scopes in read-only mode, so the cached token cannot write either (ADR 0001). */
export const READ_ONLY_SCOPES = 'User.Read Mail.Read MailboxSettings.Read offline_access';

/** Global and national-cloud Microsoft Graph hosts. The access token is only ever sent to one of these. */
export const GRAPH_HOSTS = new Set(['graph.microsoft.com', 'graph.microsoft.us', 'dod-graph.microsoft.us', 'microsoftgraph.chinacloudapi.cn']);

const truthy = (v: string) => ['1', 'true', 'yes', 'on'].includes(v.toLowerCase());

const schema = z.object({
  OUTLOOK_CLIENT_ID: z.string().trim().optional().transform((v) => v || DEFAULT_CLIENT_ID),
  OUTLOOK_TENANT: z.string().trim().min(1).default('consumers'),
  OUTLOOK_SCOPES: z.string().trim().min(1).optional(),
  OUTLOOK_TOKEN_CACHE: z.string().trim().optional(),
  OUTLOOK_TOKEN_STORE: z.enum(['auto', 'keychain', 'file']).default('auto'),
  OUTLOOK_READ_ONLY: z.string().trim().default('false'),
  OUTLOOK_DOWNLOAD_DIR: z.string().trim().optional(),
  OUTLOOK_ATTACH_DIRS: z.string().trim().optional(),
  OUTLOOK_DEFAULT_TOP: z.coerce.number().int().min(1).max(100).default(20),
  OUTLOOK_MAX_BODY_CHARS: z.coerce.number().int().min(100).default(20000),
  OUTLOOK_GRAPH_BASE_URL: z
    .string()
    .trim()
    .url()
    .default('https://graph.microsoft.com/v1.0')
    .refine((v) => {
      const u = new URL(v);
      return u.protocol === 'https:' && GRAPH_HOSTS.has(u.hostname.toLowerCase());
    }, `must be https on a Microsoft Graph host (${[...GRAPH_HOSTS].join(', ')})`),
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
  const readOnly = truthy(e.OUTLOOK_READ_ONLY);
  const downloadDir = path.resolve(expandHome(e.OUTLOOK_DOWNLOAD_DIR ?? path.join(os.homedir(), 'Downloads')));
  const attachDirs = e.OUTLOOK_ATTACH_DIRS
    ? e.OUTLOOK_ATTACH_DIRS.split(path.delimiter).map((d) => d.trim()).filter(Boolean).map((d) => path.resolve(expandHome(d)))
    : [downloadDir, path.join(os.homedir(), 'Documents'), path.join(os.homedir(), 'Desktop')];
  return {
    clientId: e.OUTLOOK_CLIENT_ID,
    tenant: e.OUTLOOK_TENANT,
    scopes: (e.OUTLOOK_SCOPES ?? (readOnly ? READ_ONLY_SCOPES : READ_WRITE_SCOPES)).split(/[\s,]+/).filter(Boolean),
    tokenCachePath: path.resolve(
      expandHome(e.OUTLOOK_TOKEN_CACHE ?? path.join(os.homedir(), '.config', 'outlook-mcp', 'token-cache.json')),
    ),
    tokenStore: e.OUTLOOK_TOKEN_STORE,
    readOnly,
    downloadDir,
    attachDirs,
    defaultTop: e.OUTLOOK_DEFAULT_TOP,
    maxBodyChars: e.OUTLOOK_MAX_BODY_CHARS,
    graphBaseUrl: e.OUTLOOK_GRAPH_BASE_URL.replace(/\/+$/, ''),
  };
}

/**
 * Load .env from the package directory and from OUTLOOK_ENV_FILE when set. Real env vars win.
 * The current directory is deliberately not read: MCP hosts often start the server inside an
 * arbitrary project, and a .env there could redirect the token cache or downloads.
 */
export function loadConfig(): Config {
  const pkgDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
  const explicit = process.env.OUTLOOK_ENV_FILE?.trim();
  if (explicit) loadDotenv({ path: path.resolve(expandHome(explicit)), quiet: true });
  loadDotenv({ path: path.join(pkgDir, '.env'), quiet: true });
  try {
    return parseConfig(process.env);
  } catch (err) {
    process.stderr.write(`[outlook-mcp] ${(err as Error).message}\n`);
    process.exit(1);
  }
}
