# 0001. Enforce read-only mode below the tool layer

- Status: Accepted
- Date: 2026-11-16

## Context

`OUTLOOK_READ_ONLY=true` is implemented in one place: `selectTools` in `src/tools/index.ts` drops every `ToolDef` whose hand-set `mutating` flag is true. Nothing below that layer knows about the mode:

- `GraphClient` (`src/graph/client.ts`) sends any method to any path.
- The default scopes (`src/config.ts`) still request `Mail.ReadWrite`, `Mail.Send` and `MailboxSettings.ReadWrite`, so the cached token can write even when the server is read-only. The README asks the user to narrow `OUTLOOK_SCOPES` by hand.

The guarantee therefore rests on every future tool author setting one boolean correctly. A read tool that later gains a write side effect (as `unsubscribe` did with its outbound POST and `sendMail`) silently breaks it, and no test can catch a wrong flag because the flag is the only source of truth.

## Decision

Read-only mode is enforced at three layers:

1. **Tool registration** (unchanged): mutating tools are not registered.
2. **Graph client**: `GraphClient` takes a `readOnly` option and rejects every non-GET request before it is sent. No allowlist of "safe POSTs" exists today; one is added only if a read tool ever needs it.
3. **Scopes**: when `OUTLOOK_READ_ONLY=true` and `OUTLOOK_SCOPES` is not set explicitly, the default scope set becomes `User.Read Mail.Read MailboxSettings.Read offline_access`. An explicit `OUTLOOK_SCOPES` always wins.

Non-Graph side effects of read-only tools stay allowed and documented: `download_attachment` writes to the download directory, `login`/`logout` change the local token store.

## Alternatives considered

- **Keep tool-level filtering only.** No work, but the guarantee stays one boolean deep and the token keeps write power.
- **Client guard only, scopes unchanged.** Catches coding mistakes in this server, but a leaked or reused token cache is still write-capable.
- **Scopes only.** Strongest guarantee, but a token cache created earlier with wide scopes keeps working silently, so it cannot be the only layer.
- **A separate read-only binary or package.** Clear separation, but doubles release and documentation work for a single flag.

## Consequences

- A mis-flagged tool fails loudly in read-only mode instead of writing to the mailbox; a test can assert that every non-mutating tool only issues GETs.
- Switching an existing install to read-only does not shrink an already cached token: the user must run `logout` and `auth` again to get the narrow scopes. The README documents this; a startup warning when the cached account has wider scopes than the mode needs is a possible follow-up.
- Switching back to read-write requires a new sign-in (consent for the wider scopes). This is intended friction.
- `list_rules` needs `MailboxSettings.Read` in the read-only default set, otherwise it returns 403.
