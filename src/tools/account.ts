import { z } from 'zod';
import type { ToolDef } from './index.js';

export const accountTools: ToolDef[] = [
  {
    name: 'login',
    description:
      'Sign in to the Microsoft account. Returns a URL and a one-time code: show both to the user, who opens the URL, enters the code and accepts. Sign-in then completes in the background; call auth_status to confirm. Use force: true to sign in again (e.g. after adding a permission in Azure).',
    schema: { force: z.boolean().optional() },
    mutating: false,
    async handler(a, { auth }) {
      const s = await auth.startLogin(a.force ?? false);
      if (s.state === 'pending') {
        return {
          state: 'pending',
          url: s.login.verificationUri,
          code: s.login.userCode,
          expiresAt: s.login.expiresAt,
          next: 'Open the URL, enter the code, then call auth_status.',
        };
      }
      return s;
    },
  },
  {
    name: 'auth_status',
    description: 'Report whether the server is signed in, waiting for a device code sign-in, or signed out.',
    schema: {},
    mutating: false,
    async handler(_a, { auth }) {
      return auth.status();
    },
  },
  {
    name: 'logout',
    description: 'Sign out: removes the cached account and deletes the local token cache file.',
    schema: {},
    mutating: false,
    async handler(_a, { auth }) {
      return { state: 'signed_out', ...(await auth.logout()) };
    },
  },
];
