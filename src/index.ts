#!/usr/bin/env node
import { clearTokenStore, createPca, getAccessToken, getFirstAccount } from './auth/msal.js';
import { loadConfig } from './config.js';
import { GraphClient } from './graph/client.js';
import { startServer } from './server.js';

async function main(): Promise<void> {
  const cmd = process.argv[2];
  const cfg = loadConfig();

  switch (cmd) {
    case 'auth': {
      const pca = createPca(cfg);
      await getAccessToken(pca, cfg, { interactive: true });
      const account = await getFirstAccount(pca);
      process.stderr.write(`Signed in as ${account?.username ?? 'unknown'}\n`);
      return;
    }
    case 'whoami': {
      const pca = createPca(cfg);
      const graph = new GraphClient(() => getAccessToken(pca, cfg, { interactive: false }), cfg.graphBaseUrl);
      const me = await graph.get<{ displayName?: string; mail?: string; userPrincipalName?: string }>('/me', {
        query: { $select: 'displayName,mail,userPrincipalName' },
      });
      process.stdout.write(`${me.displayName ?? ''} <${me.mail ?? me.userPrincipalName ?? ''}>\n`);
      return;
    }
    case 'logout': {
      const pca = createPca(cfg);
      const cache = pca.getTokenCache();
      for (const account of await cache.getAllAccounts()) await cache.removeAccount(account);
      await clearTokenStore(cfg);
      process.stderr.write('Signed out; token cache removed.\n');
      return;
    }
    case undefined:
    case 'serve':
      await startServer(cfg);
      return;
    default:
      process.stderr.write(`Unknown command "${cmd}". Usage: outlook-mcp [serve|auth|whoami|logout]\n`);
      process.exit(2);
  }
}

main().catch((err: Error) => {
  process.stderr.write(`[outlook-mcp] ${err.message}\n`);
  process.exit(1);
});
