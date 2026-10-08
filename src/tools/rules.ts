import { z } from 'zod';
import type { MailFolder, Page } from '../graph/types.js';
import { WELL_KNOWN_FOLDERS, wellKnownFolder } from '../util/odata.js';
import { blockRuleName, blockTargets, buildRule, nextSequence, type MessageRule } from '../util/rules.js';
import type { ToolContext, ToolDef } from './index.js';

const RULES = '/me/mailFolders/inbox/messageRules';
const WELL_KNOWN_IDS = new Set(Object.values(WELL_KNOWN_FOLDERS));

/** Rules need real folder ids, so well-known names are resolved with a GET. */
async function resolveFolderId(graph: ToolContext['graph'], idOrName: string): Promise<string> {
  const name = wellKnownFolder(idOrName);
  if (!WELL_KNOWN_IDS.has(name)) return idOrName;
  const f = await graph.get<MailFolder>(`/me/mailFolders/${encodeURIComponent(name)}`, { query: { $select: 'id' } });
  return f.id;
}

async function existingRules(graph: ToolContext['graph']): Promise<MessageRule[]> {
  return (await graph.get<Page<MessageRule>>(RULES)).value;
}

function summarize(r: MessageRule) {
  return {
    id: r.id,
    displayName: r.displayName,
    enabled: r.isEnabled,
    sequence: r.sequence,
    conditions: {
      fromAddresses: r.conditions?.fromAddresses?.map((x) => x.emailAddress.address),
      senderContains: r.conditions?.senderContains,
      subjectContains: r.conditions?.subjectContains,
    },
    actions: r.actions,
  };
}

export const ruleTools: ToolDef[] = [
  {
    name: 'list_rules',
    description: 'List the inbox message rules (server-side rules that run on incoming mail). Requires the MailboxSettings.ReadWrite permission.',
    schema: {},
    mutating: false,
    async handler(_a, { graph }) {
      return { rules: (await existingRules(graph)).map(summarize) };
    },
  },
  {
    name: 'create_rule',
    description:
      'Create an inbox rule that acts on future incoming mail matching the conditions (at least one of fromAddresses, senderContains, subjectContains). Actions: move (needs destinationFolderId, an id or a name like archive), delete (to Deleted Items), markRead, junk (to Junk Email). Does not affect mail already received.',
    schema: {
      displayName: z.string().min(1).max(120),
      fromAddresses: z.array(z.string().min(1)).optional().describe('Exact sender email addresses'),
      senderContains: z.array(z.string().min(1)).optional().describe('Text the sender address must contain, e.g. "@example.com"'),
      subjectContains: z.array(z.string().min(1)).optional(),
      action: z.enum(['move', 'delete', 'markRead', 'junk']),
      destinationFolderId: z.string().min(1).optional().describe('Folder id or well-known name; required for action "move"'),
      stopProcessingRules: z.boolean().optional().describe('Stop evaluating later rules after this one (default true)'),
    },
    mutating: true,
    async handler(a, { graph }) {
      let destination: string | undefined;
      if (a.action === 'junk') destination = await resolveFolderId(graph, 'junkemail');
      else if (a.action === 'move') {
        if (!a.destinationFolderId) throw new Error('Action "move" needs destinationFolderId.');
        destination = await resolveFolderId(graph, a.destinationFolderId);
      }
      const rule = buildRule({ ...a, destinationFolderId: destination, sequence: nextSequence(await existingRules(graph)) });
      return { created: summarize(await graph.post<MessageRule>(RULES, rule)) };
    },
  },
  {
    name: 'delete_rule',
    description: 'Delete an inbox rule by id (see list_rules).',
    schema: { id: z.string().min(1) },
    mutating: true,
    async handler(a, { graph }) {
      await graph.delete(`${RULES}/${encodeURIComponent(a.id)}`);
      return { id: a.id, deleted: true };
    },
  },
  {
    name: 'block_sender',
    description:
      'Block senders by creating an inbox rule that moves their future mail to Deleted Items. Give addresses (exact match) and/or domains (matches any sender at that domain). Does not touch mail already received; use delete_message with ids for that. Undo with delete_rule.',
    schema: {
      addresses: z.array(z.string().min(1)).optional().describe('Email addresses to block'),
      domains: z.array(z.string().min(1)).optional().describe('Domains to block, e.g. "example.com"'),
    },
    mutating: true,
    async handler(a, { graph }) {
      const { fromAddresses, senderContains } = blockTargets(a.addresses, a.domains);
      const rule = buildRule({
        displayName: blockRuleName(fromAddresses, senderContains.map((d) => d.slice(1))),
        fromAddresses,
        senderContains,
        action: 'move',
        destinationFolderId: await resolveFolderId(graph, 'deleteditems'),
        stopProcessingRules: true,
        sequence: nextSequence(await existingRules(graph)),
      });
      return { created: summarize(await graph.post<MessageRule>(RULES, rule)) };
    },
  },
];
