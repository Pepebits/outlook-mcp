export const BLOCK_RULE_PREFIX = 'Blocked by outlook-mcp: ';
const MAX_NAME = 120;

export type RuleAction = 'move' | 'delete' | 'markRead' | 'junk';

export interface RuleInput {
  displayName: string;
  fromAddresses?: string[];
  senderContains?: string[];
  subjectContains?: string[];
  action: RuleAction;
  /** Real folder id; required for move and junk. */
  destinationFolderId?: string;
  stopProcessingRules?: boolean;
  sequence: number;
}

export interface MessageRule {
  id?: string;
  displayName: string;
  sequence: number;
  isEnabled: boolean;
  conditions: {
    fromAddresses?: { emailAddress: { address: string } }[];
    senderContains?: string[];
    subjectContains?: string[];
  };
  actions: {
    moveToFolder?: string;
    delete?: boolean;
    markAsRead?: boolean;
    stopProcessingRules?: boolean;
  };
}

const clean = (list?: string[]) => (list ?? []).map((s) => s.trim()).filter(Boolean);

/** Builds the Graph messageRule payload; throws on an unusable combination. */
export function buildRule(input: RuleInput): MessageRule {
  const from = clean(input.fromAddresses);
  const sender = clean(input.senderContains);
  const subject = clean(input.subjectContains);
  if (!from.length && !sender.length && !subject.length) {
    throw new Error('A rule needs at least one condition: fromAddresses, senderContains or subjectContains.');
  }
  const conditions: MessageRule['conditions'] = {};
  if (from.length) conditions.fromAddresses = from.map((address) => ({ emailAddress: { address } }));
  if (sender.length) conditions.senderContains = sender;
  if (subject.length) conditions.subjectContains = subject;

  const actions: MessageRule['actions'] = {};
  switch (input.action) {
    case 'move':
    case 'junk':
      if (!input.destinationFolderId) throw new Error(`Action "${input.action}" needs a destination folder id.`);
      actions.moveToFolder = input.destinationFolderId;
      break;
    case 'delete':
      actions.delete = true;
      break;
    case 'markRead':
      actions.markAsRead = true;
      break;
  }
  if (input.stopProcessingRules ?? true) actions.stopProcessingRules = true;
  return { displayName: input.displayName, sequence: input.sequence, isEnabled: true, conditions, actions };
}

/** Normalises block_sender input: addresses match From exactly, domains match by "@domain" substring. */
export function blockTargets(addresses?: string[], domains?: string[]) {
  const fromAddresses = clean(addresses).map((a) => a.toLowerCase());
  const senderContains = clean(domains).map((d) => `@${d.toLowerCase().replace(/^@/, '')}`);
  if (fromAddresses.some((a) => !a.includes('@'))) throw new Error('addresses must be full email addresses.');
  if (!fromAddresses.length && !senderContains.length) throw new Error('Provide at least one address or domain to block.');
  return { fromAddresses, senderContains };
}

export function blockRuleName(addresses: string[], domains: string[]): string {
  const name = BLOCK_RULE_PREFIX + [...addresses, ...domains].join(', ');
  return name.length > MAX_NAME ? `${name.slice(0, MAX_NAME - 1)}…` : name;
}

export function nextSequence(rules: { sequence?: number }[]): number {
  return rules.reduce((max, r) => Math.max(max, r.sequence ?? 0), 0) + 1;
}
