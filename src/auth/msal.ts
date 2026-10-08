import { LogLevel, PublicClientApplication, type AccountInfo } from '@azure/msal-node';
import type { Config } from '../config.js';
import { log } from '../log.js';
import { FileCachePlugin } from './cache.js';

export class AuthRequiredError extends Error {
  constructor(
    message = 'Not signed in or session expired. Run `npm run auth` (or `npx outlook-mcp auth`) in a terminal, then retry.',
  ) {
    super(message);
    this.name = 'AuthRequiredError';
  }
}

export function createPca(cfg: Config): PublicClientApplication {
  return new PublicClientApplication({
    auth: {
      clientId: cfg.clientId,
      authority: `https://login.microsoftonline.com/${cfg.tenant}`,
    },
    cache: { cachePlugin: new FileCachePlugin(cfg.tokenCachePath) },
    system: {
      loggerOptions: {
        piiLoggingEnabled: false,
        logLevel: LogLevel.Warning,
        loggerCallback: (_level, message) => log.debug(`msal: ${message}`),
      },
    },
  });
}

export async function getFirstAccount(pca: PublicClientApplication): Promise<AccountInfo | undefined> {
  return (await pca.getTokenCache().getAllAccounts())[0];
}

export async function getAccessToken(
  pca: PublicClientApplication,
  cfg: Config,
  opts: { interactive: boolean },
): Promise<string> {
  const account = await getFirstAccount(pca);
  if (account) {
    try {
      const res = await pca.acquireTokenSilent({ account, scopes: cfg.scopes });
      if (res?.accessToken) return res.accessToken;
    } catch (err) {
      log.debug(`silent token acquisition failed: ${(err as Error).name}`);
    }
  }
  if (!opts.interactive) throw new AuthRequiredError();

  const res = await pca.acquireTokenByDeviceCode({
    scopes: cfg.scopes,
    deviceCodeCallback: (r) => process.stderr.write(`\n${r.message}\n\n`),
  });
  if (!res?.accessToken) throw new AuthRequiredError('Device code sign-in did not return a token.');
  return res.accessToken;
}
