import { z } from 'zod';
import type { MailFolder, Page } from '../graph/types.js';
import { wellKnownFolder } from '../util/odata.js';
import type { ToolDef } from './index.js';

export const folderTools: ToolDef[] = [
  {
    name: 'list_folders',
    description:
      'List mail folders. Without parentFolderId returns top-level folders; with it returns child folders. Pass nextLink from a previous response for more. ' +
      'Well-known names usable as folder ids elsewhere: inbox, drafts, sentitems, deleteditems, junkemail, archive.',
    schema: {
      parentFolderId: z.string().optional().describe('Folder id or well-known name to list children of'),
      includeHidden: z.boolean().optional().describe('Include hidden folders'),
      nextLink: z.string().optional().describe('Opaque nextLink from a previous response'),
    },
    mutating: false,
    async handler(args: { parentFolderId?: string; includeHidden?: boolean; nextLink?: string }, { graph }) {
      if (args.nextLink) {
        const next = await graph.get<Page<MailFolder>>(args.nextLink);
        return { folders: next.value, nextLink: next['@odata.nextLink'] };
      }
      const path = args.parentFolderId
        ? `/me/mailFolders/${encodeURIComponent(wellKnownFolder(args.parentFolderId))}/childFolders`
        : '/me/mailFolders';
      const page = await graph.get<Page<MailFolder>>(path, {
        query: {
          $top: 100,
          $select: 'id,displayName,parentFolderId,totalItemCount,unreadItemCount',
          includeHiddenFolders: args.includeHidden ? true : undefined,
        },
      });
      return { folders: page.value, nextLink: page['@odata.nextLink'] };
    },
  },
];
