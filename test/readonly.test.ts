import { describe, expect, it } from 'vitest';
import { allTools, selectTools } from '../src/tools/index.js';

describe('read-only mode', () => {
  const names = (ro: boolean) => selectTools(allTools(), ro).map((t) => t.name);
  it('removes every mutating tool', () => {
    const ro = names(true);
    for (const n of ['send_mail', 'create_draft', 'reply_message', 'forward_message', 'move_message', 'mark_read', 'flag_message', 'delete_message', 'create_rule', 'delete_rule', 'block_sender']) {
      expect(ro).not.toContain(n);
      expect(names(false)).toContain(n);
    }
    expect(ro).toEqual(expect.arrayContaining(['list_folders', 'list_messages', 'search_messages', 'get_message', 'list_attachments', 'download_attachment', 'find_newsletters', 'get_unsubscribe_info', 'list_rules']));
  });
});
