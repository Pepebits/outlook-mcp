import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { createPca, getAccessToken } from './auth/msal.js';
import type { Config } from './config.js';
import { GraphClient } from './graph/client.js';
import { log } from './log.js';
import { registerAllTools } from './tools/index.js';

export async function startServer(cfg: Config): Promise<void> {
  // stdout carries the MCP protocol; make stray console.log calls harmless.
  console.log = console.error;

  const pca = createPca(cfg);
  const graph = new GraphClient(() => getAccessToken(pca, cfg, { interactive: false }), cfg.graphBaseUrl);
  const server = new McpServer({ name: 'outlook-mcp', version: '0.1.0' });
  const names = registerAllTools(server, { graph, cfg });

  await server.connect(new StdioServerTransport());
  log.info(`ready; ${names.length} tools registered${cfg.readOnly ? ' (read-only mode)' : ''}`);
}
