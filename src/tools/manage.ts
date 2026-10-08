import { z } from 'zod';
import type { Message } from '../graph/types.js';
import type { ToolDef } from './index.js';

const msg = (id: string) => `/me/messages/${encodeURIComponent(id)}`;

export const manageTools: ToolDef[] = [
  {
    name: 'move_message',
    description:
      'Move a message to another folder. Returns the NEW message id (ids change after a move). destinationFolderId may be an id or a well-known name such as archive, deleteditems, junkemail.',
    schema: { id: z.string().min(1), destinationFolderId: z.string().min(1) },
    mutating: true,
    async handler(a, { graph }) {
      const m = await graph.post<Message>(`${msg(a.id)}/move`, { destinationId: a.destinationFolderId });
      return { newId: m.id };
    },
  },
  {
    name: 'mark_read',
    description: 'Mark a message as read or unread.',
    schema: { id: z.string().min(1), isRead: z.boolean() },
    mutating: true,
    async handler(a, { graph }) {
      await graph.patch(msg(a.id), { isRead: a.isRead });
      return { id: a.id, isRead: a.isRead };
    },
  },
  {
    name: 'flag_message',
    description: 'Set the follow-up flag status of a message.',
    schema: { id: z.string().min(1), status: z.enum(['flagged', 'complete', 'notFlagged']) },
    mutating: true,
    async handler(a, { graph }) {
      await graph.patch(msg(a.id), { flag: { flagStatus: a.status } });
      return { id: a.id, status: a.status };
    },
  },
  {
    name: 'delete_message',
    description:
      'Delete a message. By default moves it to Deleted Items (recoverable). With permanent: true it is deleted irreversibly.',
    schema: { id: z.string().min(1), permanent: z.boolean().optional() },
    mutating: true,
    async handler(a, { graph }) {
      if (a.permanent) {
        await graph.delete(msg(a.id));
        return { id: a.id, deleted: 'permanently' };
      }
      const m = await graph.post<Message>(`${msg(a.id)}/move`, { destinationId: 'deleteditems' });
      return { newId: m.id, deleted: 'moved to deleteditems' };
    },
  },
];
