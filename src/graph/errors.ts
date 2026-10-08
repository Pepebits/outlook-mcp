import { AuthRequiredError } from '../auth/msal.js';

export class GraphError extends Error {
  constructor(
    message: string,
    readonly status: number,
    readonly code: string,
    readonly requestPath: string = '',
  ) {
    super(message);
    this.name = 'GraphError';
  }

  static async from(res: Response, requestPath = ''): Promise<GraphError> {
    let code = 'UnknownError';
    let message = res.statusText || `HTTP ${res.status}`;
    try {
      const body = (await res.json()) as { error?: { code?: string; message?: string } };
      if (body?.error) {
        code = body.error.code ?? code;
        message = body.error.message ?? message;
      }
    } catch {
      /* non-JSON error body */
    }
    return new GraphError(message, res.status, code, requestPath);
  }
}

/** Convert any thrown value into a user-friendly message. */
export function mapGraphError(err: unknown): Error {
  if (err instanceof AuthRequiredError) return err;
  if (!(err instanceof GraphError)) return err instanceof Error ? err : new Error(String(err));

  if (err.status === 401 || err.code === 'InvalidAuthenticationToken') return new AuthRequiredError();
  if (err.code === 'ErrorItemNotFound' || err.status === 404) {
    return new Error('Message not found (id may have changed after move). List the folder again to get fresh ids.');
  }
  if ((err.status === 403 || err.code === 'ErrorAccessDenied') && /messageRules/i.test(err.requestPath)) {
    return new Error(
      'Access denied for inbox rules. Add the delegated Microsoft Graph permission MailboxSettings.ReadWrite to your app registration in Azure, include it in OUTLOOK_SCOPES, then re-run `npm run auth`.',
    );
  }
  if (err.status === 403 || err.code === 'ErrorAccessDenied') {
    return new Error(
      `Access denied (${err.code}). The app registration may be missing a delegated permission (Mail.ReadWrite / Mail.Send), or admin consent is required. ${err.message}`,
    );
  }
  if (err.code === 'ErrorInvalidRequest' && /search|\$search/i.test(`${err.requestPath} ${err.message}`)) {
    return new Error(
      `Invalid search query. Use KQL syntax, e.g. from:alice subject:"quarterly report" hasattachments:true. ${err.message}`,
    );
  }
  return new Error(`Graph error ${err.status} ${err.code}: ${err.message}`);
}
