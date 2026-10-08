# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
