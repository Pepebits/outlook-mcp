import { z } from 'zod';
import type { Message } from '../graph/types.js';
import { idOrIds, runBulk } from './bulk.js';
import type { ToolDef } from './index.js';

const msg = (id: string) => `/me/messages/${encodeURIComponent(id)}`;

export const manageTools: ToolDef[] = [
  {
    name: 'move_message',
    description:
      'Move a message to another folder. Returns the NEW message id (ids change after a move). Accepts id or ids (up to 50). destinationFolderId may be an id or a well-known name such as archive, deleteditems, junkemail.',
    schema: { ...idOrIds, destinationFolderId: z.string().min(1) },
    mutating: true,
    async handler(a, { graph }) {
      return runBulk(a, async (id) => {
        const m = await graph.post<Message>(`${msg(id)}/move`, { destinationId: a.destinationFolderId });
        return { newId: m.id };
      });
    },
  },
  {
    name: 'mark_read',
    description: 'Mark a message as read or unread. Accepts id or ids (up to 50).',
    schema: { ...idOrIds, isRead: z.boolean() },
    mutating: true,
    async handler(a, { graph }) {
      return runBulk(a, async (id) => {
        await graph.patch(msg(id), { isRead: a.isRead });
        return { id, isRead: a.isRead };
      });
    },
  },
  {
    name: 'flag_message',
    description: 'Set the follow-up flag status of a message. Accepts id or ids (up to 50).',
    schema: { ...idOrIds, status: z.enum(['flagged', 'complete', 'notFlagged']) },
    mutating: true,
    async handler(a, { graph }) {
      return runBulk(a, async (id) => {
        await graph.patch(msg(id), { flag: { flagStatus: a.status } });
        return { id, status: a.status };
      });
    },
  },
  {
    name: 'delete_message',
    description:
      'Delete a message. By default moves it to Deleted Items (recoverable). With permanent: true it is deleted irreversibly. Accepts id or ids (up to 50).',
    schema: { ...idOrIds, permanent: z.boolean().optional() },
    mutating: true,
    async handler(a, { graph }) {
      return runBulk(a, async (id) => {
        if (a.permanent) {
          await graph.delete(msg(id));
          return { id, deleted: 'permanently' };
        }
        const m = await graph.post<Message>(`${msg(id)}/move`, { destinationId: 'deleteditems' });
        return { newId: m.id, deleted: 'moved to deleteditems' };
      });
    },
  },
];
