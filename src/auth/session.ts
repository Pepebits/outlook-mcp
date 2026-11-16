import type { PublicClientApplication } from '@azure/msal-node';
import type { Config } from '../config.js';
import { log } from '../log.js';
import { clearTokenStore, createPca, getAccessTokenResult, getFirstAccount, keepOnlyAccount, type AccessToken } from './msal.js';

/** Refresh this long before expiry, so a token never expires mid-request. */
const REFRESH_MARGIN_MS = 5 * 60 * 1000;

export interface PendingLogin {
  verificationUri: string;
  userCode: string;
  message: string;
  expiresAt: string;
}

export type LoginState =
  | { state: 'signed_in'; username: string }
  | { state: 'pending'; login: PendingLogin }
  | { state: 'signed_out'; lastError?: string };

/**
 * Owns the MSAL client for the running server so sign-in and sign-out can be
 * driven from MCP tools: the device code is handed back to the client instead
 * of being printed, and the flow completes in the background.
 */
export class AuthSession {
  private pca: PublicClientApplication;
  private pending?: PendingLogin;
  private lastError?: string;
  private abort?: AbortController;
  private cached?: AccessToken;
  private inflight?: Promise<AccessToken>;

  constructor(
    private cfg: Config,
    private readonly now: () => number = Date.now,
    private readonly fetchToken: (pca: PublicClientApplication) => Promise<AccessToken> = (pca) =>
      getAccessTokenResult(pca, cfg, { interactive: false }),
  ) {
    this.pca = createPca(cfg);
  }

  /**
   * Returns a cached access token until shortly before it expires. Concurrent callers share one
   * refresh, so a bulk run does not hit MSAL (and the keychain) once per request.
   */
  getToken = async (): Promise<string> => {
    const c = this.cached;
    if (c && c.expiresAt !== undefined && c.expiresAt - REFRESH_MARGIN_MS > this.now()) return c.token;
    if (!this.inflight) {
      const p = this.fetchToken(this.pca);
      this.inflight = p;
      p.then(
        (t) => {
          if (this.inflight === p) this.cached = t;
        },
        () => undefined,
      ).finally(() => {
        if (this.inflight === p) this.inflight = undefined;
      });
    }
    return (await this.inflight).token;
  };

  /** Forget the in-memory token (after sign-in or sign-out). */
  private resetToken() {
    this.cached = undefined;
    this.inflight = undefined;
  }

  async status(): Promise<LoginState> {
    if (this.pending) return { state: 'pending', login: this.pending };
    const account = await getFirstAccount(this.pca);
    if (account) {
      try {
        await this.getToken();
        return { state: 'signed_in', username: account.username };
      } catch {
        // Cached account whose refresh token no longer works (or lacks a new scope).
      }
    }
    return this.lastError ? { state: 'signed_out', lastError: this.lastError } : { state: 'signed_out' };
  }

  /** Starts a device code sign-in and resolves as soon as the code is known. */
  async startLogin(force = false): Promise<LoginState> {
    if (this.pending) return { state: 'pending', login: this.pending };
    if (!force) {
      const current = await this.status();
      if (current.state === 'signed_in') return current;
    }
    this.lastError = undefined;
    const abort = new AbortController();
    this.abort = abort;

    return new Promise<LoginState>((resolve, reject) => {
      const request = {
        scopes: this.cfg.scopes,
        cancel: false,
        deviceCodeCallback: (r: { verificationUri: string; userCode: string; message: string; expiresIn: number }) => {
          this.pending = {
            verificationUri: r.verificationUri,
            userCode: r.userCode,
            message: r.message,
            expiresAt: new Date(Date.now() + r.expiresIn * 1000).toISOString(),
          };
          resolve({ state: 'pending', login: this.pending });
        },
      };
      // MSAL polls until request.cancel turns true.
      abort.signal.addEventListener('abort', () => {
        request.cancel = true;
      });
      this.pca
        .acquireTokenByDeviceCode(request)
        .then(async (res) => {
          // A forced sign-in may use another account: drop the old ones so the new one is used.
          if (res?.account) await keepOnlyAccount(this.pca, res.account);
          this.resetToken();
          log.info(`signed in as ${res?.account?.username ?? 'unknown'}`);
        })
        .catch((err: Error) => {
          if (abort.signal.aborted) return;
          this.lastError = err.message;
          log.warn(`device code sign-in failed: ${err.message}`);
          reject(err);
        })
        .finally(() => {
          if (this.abort === abort) {
            this.pending = undefined;
            this.abort = undefined;
          }
        });
    });
  }

  async logout(): Promise<{ removedAccounts: number }> {
    this.abort?.abort();
    this.pending = undefined;
    this.abort = undefined;
    this.resetToken();
    const cache = this.pca.getTokenCache();
    const accounts = await cache.getAllAccounts();
    for (const account of accounts) await cache.removeAccount(account);
    try {
      await clearTokenStore(this.cfg);
    } finally {
      this.pca = createPca(this.cfg);
    }
    return { removedAccounts: accounts.length };
  }
}
