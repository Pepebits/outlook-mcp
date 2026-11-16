# 0002. Confine local file access to configured roots

- Status: Accepted
- Date: 2026-11-16

## Context

The server touches the local filesystem in two directions, with asymmetric protection:

- **Writes** (`download_attachment`, `src/tools/attachments.ts`) are confined to `OUTLOOK_DOWNLOAD_DIR` by `resolveSafePath` (`src/util/paths.ts`). The check is lexical: it does not resolve symlinks, so a symlinked directory inside the download directory leads writes outside it.
- **Reads** (`send_mail` / `create_draft` attachments, `src/tools/compose.ts`) accept any path the process can read. There is no confinement at all.

Email content is untrusted and reaches the model. A message that persuades the model to call `send_mail` with `attachments: [{ path: "~/.ssh/id_ed25519" }]` turns the server into a file exfiltration channel. The only barrier today is the host's confirmation prompt, which shows a path and a recipient but is routinely approved.

## Decision

Both directions go through one path policy module:

1. **Attachment reads** are allowed only under configured roots, `OUTLOOK_ATTACH_DIRS` (path-list, `~` expanded). Default: the download directory plus `~/Documents` and `~/Desktop`. Dotfiles and dot-directories anywhere in the resolved path are rejected regardless of root.
2. **Containment is checked on real paths**: the target (or, for a file that does not exist yet, its nearest existing parent) is resolved with `fs.realpath` and must stay inside the realpath of an allowed root. This applies to reads and to download writes.
3. The validated real path is the one that is read and uploaded, so a symlink cannot be swapped between the check and the read of a different name. (A file replaced in place between the check and the read is out of scope.)
4. Setting `OUTLOOK_ATTACH_DIRS=/` is the explicit opt-out for location; hidden files and directories are still refused.

## Alternatives considered

- **Leave reads unrestricted and rely on host confirmation.** No breaking change, but the server's own threat model (untrusted content steering the model) says confirmation is a mitigation, not a control.
- **Denylist of sensitive paths (`~/.ssh`, `~/.aws`, keychains, browser profiles).** Non-breaking, but denylists are never complete and differ per platform.
- **Attach only from the download directory.** Simplest and tightest, but breaks the common "send the report in ~/Documents" case and pushes users to copy files around.
- **Drop path-based attachments; accept base64 content from the client.** Removes filesystem reads entirely, but large files would pass through the model context, which is what upload sessions were built to avoid.

## Consequences

- Breaking change for users who attach files from other locations: they must set `OUTLOOK_ATTACH_DIRS`. Ships in a minor release with a CHANGELOG entry and an error message that names the variable.
- The exfiltration path from "prompt-injected model" to "arbitrary local file" is closed by default; what remains is limited to documents the user already keeps in ordinary document folders.
- One module owns path policy, so `resolveSafePath` gains realpath checks and its tests cover symlinks for both directions.
- The default roots follow the platform: XDG user dirs on Linux (localized folder names) and, on Windows, the OneDrive-redirected Documents and Desktop in addition to the local ones.
