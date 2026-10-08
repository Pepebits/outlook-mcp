import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { AuthSession } from './auth/session.js';
import type { Config } from './config.js';
import { GraphClient } from './graph/client.js';
import { log } from './log.js';
import { registerAllTools } from './tools/index.js';

export async function startServer(cfg: Config): Promise<void> {
  // stdout carries the MCP protocol; make stray console.log calls harmless.
  console.log = console.error;

  const auth = new AuthSession(cfg);
  const graph = new GraphClient(auth.getToken, cfg.graphBaseUrl);
  const server = new McpServer({ name: 'outlook-mcp', version: '0.3.0' });
  const names = registerAllTools(server, { graph, cfg, auth });

  await server.connect(new StdioServerTransport());
  log.info(`ready; ${names.length} tools registered${cfg.readOnly ? ' (read-only mode)' : ''}`);
}
