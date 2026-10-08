import { LogLevel, PublicClientApplication, type AccountInfo } from '@azure/msal-node';
import type { Config } from '../config.js';
import { log } from '../log.js';
import fs from 'node:fs/promises';
import type { ICachePlugin } from '@azure/msal-node';
import { FileCachePlugin } from './cache.js';
import { KeychainCachePlugin, currentStoreEnv, defaultExec, selectTokenStore, type Exec, type StoreEnv } from './keychain.js';

export class AuthRequiredError extends Error {
  constructor(
    message = 'Not signed in or session expired. Call the login tool (or run `npm run auth` in a terminal), then retry.',
  ) {
    super(message);
    this.name = 'AuthRequiredError';
  }
}

export function createCachePlugin(cfg: Config): ICachePlugin {
  const store = selectTokenStore(cfg.tokenStore, cfg.tokenCachePath);
  return store.kind === 'keychain'
    ? new KeychainCachePlugin(store.backend, cfg.tokenCachePath, store.strict)
    : new FileCachePlugin(cfg.tokenCachePath);
}

/**
 * Deletes the token cache file and the keychain entry. The keychain is always tried where the platform has
 * one, whatever OUTLOOK_TOKEN_STORE says, so a switch to `file` cannot leave a token behind. A missing
 * entry is fine; any other keychain failure is thrown so logout does not report success falsely.
 */
export async function clearTokenStore(cfg: Config, env: StoreEnv = currentStoreEnv(), exec: Exec = defaultExec): Promise<void> {
  await fs.rm(cfg.tokenCachePath, { force: true });
  const store = selectTokenStore('auto', cfg.tokenCachePath, env, exec);
  if (store.kind !== 'keychain') return;
  try {
    await store.backend.delete();
  } catch (err) {
    throw new Error(
      `Signed out locally, but the ${store.backend.name} entry could not be removed: ${(err as Error).message}`,
    );
  }
}

export function createPca(cfg: Config): PublicClientApplication {
  return new PublicClientApplication({
    auth: {
      clientId: cfg.clientId,
      authority: `https://login.microsoftonline.com/${cfg.tenant}`,
    },
    cache: { cachePlugin: createCachePlugin(cfg) },
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
