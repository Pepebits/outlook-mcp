import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { z } from 'zod';
import type { AuthSession } from '../auth/session.js';
import type { Config } from '../config.js';
import type { GraphClient } from '../graph/client.js';
import { mapGraphError } from '../graph/errors.js';
import { log } from '../log.js';
import { toTextResult } from '../util/format.js';
import { accountTools } from './account.js';
import { attachmentTools } from './attachments.js';
import { bulkActionTools } from './bulkaction.js';
import { composeTools } from './compose.js';
import { digestTools } from './digest.js';
import { folderTools } from './folders.js';
import { manageTools } from './manage.js';
import { messageTools } from './messages.js';
import { newsletterTools } from './newsletters.js';
import { ruleTools } from './rules.js';
import { unsubscribeTools } from './unsubscribe.js';

export interface ToolContext {
  graph: GraphClient;
  cfg: Config;
  auth: AuthSession;
}

export interface ToolDef {
  name: string;
  description: string;
  schema: z.ZodRawShape;
  /** True when the tool changes mailbox state (send, move, delete, ...). */
  mutating: boolean;
  /** Mutating tools are destructive unless marked false (reversible or additive changes). */
  destructive?: false;
  /** True when the tool reaches people or hosts outside the mailbox (sending mail, unsubscribe requests). */
  openWorld?: true;
  handler: (args: any, ctx: ToolContext) => Promise<unknown>;
}

export function allTools(): ToolDef[] {
  return [...accountTools, ...folderTools, ...messageTools, ...attachmentTools, ...composeTools, ...manageTools, ...bulkActionTools, ...unsubscribeTools, ...newsletterTools, ...digestTools, ...ruleTools];
}

/** Pure selection function: drops mutating tools in read-only mode. */
export function selectTools(tools: ToolDef[], readOnly: boolean): ToolDef[] {
  return readOnly ? tools.filter((t) => !t.mutating) : tools;
}

export function withErrorHandling(def: ToolDef, ctx: ToolContext) {
  return async (args: unknown) => {
    try {
      return toTextResult(await def.handler(args, ctx));
    } catch (err) {
      const mapped = mapGraphError(err);
      log.error(`tool ${def.name} failed: ${mapped.message}`);
      return toTextResult(mapped.message, true);
    }
  };
}

/** MCP tool annotations, so clients can tell read-only tools from destructive ones. */
export function toolAnnotations(def: ToolDef) {
  return {
    readOnlyHint: !def.mutating,
    ...(def.mutating ? { destructiveHint: def.destructive ?? true } : {}),
    openWorldHint: def.openWorld ?? false,
  };
}

export function registerAllTools(server: McpServer, ctx: ToolContext): string[] {
  const selected = selectTools(allTools(), ctx.cfg.readOnly);
  for (const def of selected) {
    server.registerTool(
      def.name,
      { description: def.description, inputSchema: def.schema, annotations: toolAnnotations(def) },
      withErrorHandling(def, ctx) as any,
    );
  }
  return selected.map((t) => t.name);
}
