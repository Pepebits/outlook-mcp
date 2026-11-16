# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [2.0.0] - 2026-11-16

### Security

- The access token is only sent to the configured Microsoft Graph origin and API version: a `nextLink` pointing to another host, scheme or path is refused.
- `OUTLOOK_GRAPH_BASE_URL` must be https on a Microsoft Graph host (global or national cloud).
- One-click unsubscribe only POSTs to https URLs that resolve to public addresses, and checks each redirect before following it.
- Unsubscribe links and mailto fields, `internetMessageId` and attachment content types are returned under the `untrusted` key.

### Added

- `send_mail` and `create_draft` accept `from` and `fromName` to send from an alias or SendAs identity (thanks @itsrealfortune, #7). `fromName` without `from` is rejected, and a refused From address gets a clear error.
- `OUTLOOK_ATTACH_DIRS`: the directories attachments may be read from (default: the download directory, `~/Documents` and `~/Desktop`).
- `OUTLOOK_ENV_FILE`: path of a `.env` file to load.
- MCP tool annotations (`readOnlyHint`, `destructiveHint`, `openWorldHint`).
- `list_folders` accepts `nextLink`.
- Default download and attachment directories follow the platform: XDG user dirs on Linux (localized names such as `~/Descargas`) and OneDrive-redirected Documents and Desktop on Windows.
- Architecture decision records in `docs/adr`.

### Changed

- **Breaking:** `bulk_action` applies only a previewed selection: the dry run returns a `selectionId`, and `{ selectionId, dryRun: false }` acts on exactly those messages, once, within 10 minutes. `dryRun: false` without a `selectionId`, or with search or filter fields, is refused ([ADR 0003](docs/adr/0003-bind-bulk-action-apply-to-a-dry-run-selection.md)).
- **Breaking:** attachments must be inside `OUTLOOK_ATTACH_DIRS` and must not be hidden files or inside hidden folders; paths are checked after resolving symlinks ([ADR 0002](docs/adr/0002-confine-local-file-access-to-configured-roots.md)).
- **Breaking:** a `.env` in the current working directory is no longer read; use the package directory, `OUTLOOK_ENV_FILE` or your MCP client's env block.
- Read-only mode is also enforced in the Graph client (non-GET requests are refused) and, unless `OUTLOOK_SCOPES` is set, signs in with read-only scopes ([ADR 0001](docs/adr/0001-enforce-read-only-mode-below-the-tool-layer.md)).
- Inline attachments share a single ~3 MB budget; files that do not fit go through upload sessions, so several medium files no longer exceed Graph's request size limit.
- A date without a time in `until` includes that whole day.
- `move_message` accepts folder aliases such as `trash` and `spam`.
- `find_newsletters` and `mail_digest` read List-Unsubscribe headers through JSON batching (`$batch`, 20 per request) instead of one request per sender; failed sub-requests are retried individually. In read-only mode a `$batch` made only of GET requests is the one POST the client allows.

### Fixed

- `POST` requests are no longer retried after a 5xx or a network error, so mail cannot be sent twice; when the final send of a message with large attachments fails that way, the draft is kept and the error says the mail may have been sent.
- Graph requests time out after 60 seconds, network errors on reads are retried, and `Retry-After` is capped at 60 seconds.
- The access token is cached until shortly before it expires, with one shared refresh, instead of reading the keychain on every request.
- Being offline no longer reports "Not signed in".
- `login` with `force: true` and a different account now uses the new account.
- `download_attachment` refuses symlinked directories that lead outside the download directory.
- A failed keychain write no longer leaves partial token cache entries behind.
- 404 errors name what was not found (message, attachment, rule or folder).
- `get_message` describes `format: "html"` correctly (raw HTML, not sanitized).
- The MCP server reports the package version instead of a hard-coded one.

## [1.2.0] - 2026-10-08

### Added

- `bulk_action` tool: apply `delete`, `move`, `archive`, `markRead`, `markUnread`, `flag` or `unflag` to the messages matching a KQL `query` or filters. It runs as a dry run by default (count and a sample of 10), processes up to `max` messages (default 100, hard cap 500) and reports per-message results, including the new id of each moved message.
- `send_mail` and `create_draft` accept attachments of 3 MB up to 150 MB through Graph upload sessions, streamed from disk in chunks. If an upload fails, `send_mail` deletes its draft and sends nothing; `create_draft` keeps the draft and lists the attachments that failed.

## [1.1.0] - 2026-10-08

### Added

- README setup for OpenAI Codex (`codex mcp add` and `config.toml`) and other MCP clients.
- `OUTLOOK_TOKEN_STORE` (`auto`, `keychain`, `file`; default `auto`): the token cache is kept in the macOS Keychain (`security` CLI) or the Linux Secret Service (`secret-tool`) when available, with no native dependencies. An existing cache file is imported into an empty keychain and deleted; `logout` clears both. Windows and systems without a keychain keep using the file.
- `mail_digest` tool (read-only): counts, high-importance and flagged messages, and messages grouped by sender with a newsletter hint for a folder and time window.
- Server-level MCP `instructions` telling clients that email content is untrusted and that sending, deleting, unsubscribing and rule changes should be confirmed with the user.
- Tests for tool handlers using a fake Graph client.

### Changed

- Tool results that include message content now group sender-controlled text (subjects, senders, previews, attachment names) under an `untrusted` key instead of top-level `subject`, `from`, `preview` and `name` fields. This affects `list_messages`, `search_messages`, `get_message`, `list_attachments`, `find_newsletters`, `get_unsubscribe_info` and `unsubscribe`. `download_attachment` now returns the saved path as `untrusted.savedTo` (it contains the sender-chosen file name) instead of a top-level `savedTo`.

### Security

- Prompt-injection hardening: `get_message` bodies are enclosed in `<untrusted_email_content>` markers (with look-alike markers inside the content neutralized), results carry a note that email content is untrusted data, and tool descriptions say so.
- The token cache is stored in the OS keychain by default where one is available, instead of a file. Secrets are passed to `secret-tool` and `security -i` over stdin, never as command-line arguments.

## [1.0.1] - 2026-10-08

### Added

- Listed in the official MCP Registry as `io.github.Pepebits/outlook-mcp` (`mcpName` and `server.json`).

### Changed

- Releases are staged on npm with provenance through Trusted Publishing and approved by a maintainer before they go live.

## [1.0.0] - 2026-10-08

### Added

- Published to npm as `@pepebits/outlook-mcp`: install with `npx -y @pepebits/outlook-mcp`.
- Shared default Azure app for personal Microsoft accounts, so no app registration or `.env` file is needed.

### Changed

- `OUTLOOK_CLIENT_ID` is now optional. Set it (and `OUTLOOK_TENANT`) only to use your own Azure app, for example with work or school accounts.

### Fixed

- `create_rule` and `block_sender` return the rule that was created. On personal mailboxes Graph's response could describe a different, existing rule.

## [0.2.0] - 2026-10-08

### Added

- `login`, `auth_status` and `logout` tools: sign in and out from the MCP client without a terminal. `login` returns the device code URL and code, and sign-in completes in the background.
- `get_unsubscribe_info` and `unsubscribe` tools: read List-Unsubscribe headers and leave a mailing list by RFC 8058 one-click or a mailto request.
- Bulk operations: `delete_message`, `move_message`, `mark_read`, `flag_message` and `unsubscribe` accept `ids` (up to 50) as an alternative to `id`, report per-message results and never stop on the first error.
- `find_newsletters` tool: groups a folder's messages by sender and reports which senders are mailing lists and how to unsubscribe.
- Inbox rule tools `list_rules`, `create_rule`, `delete_rule` and `block_sender`.
- `MailboxSettings.ReadWrite` is now part of the default `OUTLOOK_SCOPES`. Add it in the Azure app registration and run `npm run auth` again.

### Changed

- Node.js 24 (LTS) or newer is now required.
- CI workflow actions updated.

### Fixed

- `list_messages` with `from` no longer fails with `InefficientFilter` when sorted by date.
- `unsubscribe` returns the unsubscribe link when the one-click request is refused, or fails with a network error or timeout.

## [0.1.0]

### Added

- Initial release: folders, message listing, KQL search and reading, attachments, send/draft/reply/forward, move/mark/flag/delete, read-only mode, device code authentication and throttling retries.
