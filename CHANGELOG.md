# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- README setup for OpenAI Codex (`codex mcp add` and `config.toml`) and other MCP clients.
- `OUTLOOK_TOKEN_STORE` (`auto`, `keychain`, `file`; default `auto`): the token cache is kept in the macOS Keychain (`security` CLI) or the Linux Secret Service (`secret-tool`) when available, with no native dependencies. An existing cache file is imported into an empty keychain and deleted; `logout` clears both. Windows and systems without a keychain keep using the file.
- `mail_digest` tool (read-only): counts, high-importance and flagged messages, and messages grouped by sender with a newsletter hint for a folder and time window.
- Server-level MCP `instructions` telling clients that email content is untrusted and that sending, deleting, unsubscribing and rule changes should be confirmed with the user.
- Tests for tool handlers using a fake Graph client.

### Changed

- Tool results that include message content now group sender-controlled text (subjects, senders, previews, attachment names) under an `untrusted` key instead of top-level `subject`, `from`, `preview` and `name` fields. This affects `list_messages`, `search_messages`, `get_message`, `list_attachments` and `find_newsletters`.

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
