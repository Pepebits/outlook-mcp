<p align="center"><img src="https://raw.githubusercontent.com/Pepebits/outlook-mcp/main/docs/banner.png" alt="outlook-mcp" width="720"></p>

A [Model Context Protocol](https://modelcontextprotocol.io) server that lets Claude, Codex and any other MCP client read, search, send and organize your **Outlook / Microsoft 365 / Outlook.com** mail through the Microsoft Graph API.

---

## 🔭 Overview

`outlook-mcp` runs locally over stdio. It signs in with the **OAuth device code flow** (no client secret, no redirect URI), stores the refresh token in your OS keychain (or a local file only you can read), and talks to Microsoft Graph with plain `fetch`.

## ✨ Features

- 📂 Browse folders, list and search messages (KQL), read bodies as sanitized plain text
- 📎 List and download attachments safely into a configured directory
- ✉️ Send mail, create drafts, reply, reply-all and forward (local attachments under 3 MB)
- 🗂️ Move, mark read/unread, flag and delete messages, one at a time or in bulk (`ids`)
- 📰 Find newsletters, unsubscribe from them and block unwanted senders with inbox rules
- 🔒 Read-only mode that removes every mutating tool
- 🔁 Automatic retry with `Retry-After` / exponential backoff for throttling (429/503/504)
- 🧾 Logs only to stderr; never logs tokens or message bodies

## 📋 Requirements

- Node.js **24 (LTS) or newer**
- A Microsoft account (personal works out of the box; work or school needs [your own Azure app](#️-use-your-own-azure-app-optional))

## 🚀 Quick start

No terminal sign-in and no Azure setup needed. You only need Node.js 24+.

### Claude Code

```bash
claude mcp add outlook --scope user -- npx -y @pepebits/outlook-mcp
```

### Claude Desktop

Add to `claude_desktop_config.json` and restart Claude Desktop:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "npx",
      "args": ["-y", "@pepebits/outlook-mcp"]
    }
  }
}
```

### Codex (OpenAI)

```bash
codex mcp add outlook -- npx -y @pepebits/outlook-mcp
```

Or add it to `~/.codex/config.toml` (shared by the Codex CLI and IDE extension):

```toml
[mcp_servers.outlook]
command = "npx"
args = ["-y", "@pepebits/outlook-mcp"]
# env = { OUTLOOK_READ_ONLY = "true" }
```

Run `/mcp` inside Codex to check that it is connected.

### Other MCP clients

Any client that can launch a stdio MCP server works: run `npx -y @pepebits/outlook-mcp` as the server command.

### Sign in

Ask your assistant to *"log in to Outlook"*. The `login` tool returns a URL and a one-time code: open the URL, enter the code and accept. Prefer a terminal? Run `npx -y @pepebits/outlook-mcp auth` instead.

By default the server uses the shared **outlook-mcp** Azure app, which supports **personal Microsoft accounts** (outlook.com, hotmail, live). The consent screen may show an *unverified publisher* warning. No data passes through any server of ours: the server talks to Microsoft Graph directly from your machine and your tokens stay in the local token cache. For work or school accounts, or to use your own app, see [Use your own Azure app](#️-use-your-own-azure-app-optional).

## ⚙️ Configuration (`.env`)

The `.env` file is optional. Variables are read from `.env` in the current directory and in the package directory, and real environment variables override it. You can also pass them through your MCP client's env block, e.g. `claude mcp add outlook --scope user -e OUTLOOK_TENANT=organizations -e OUTLOOK_CLIENT_ID=your-client-id -- npx -y @pepebits/outlook-mcp`.

| Variable | Default | Description |
| --- | --- | --- |
| `OUTLOOK_CLIENT_ID` | shared outlook-mcp app | Application (client) ID of your own Azure app registration (optional). |
| `OUTLOOK_TENANT` | `consumers` | Authority tenant: `consumers` (personal), `organizations` (work/school), `common` (both) or a tenant GUID. |
| `OUTLOOK_SCOPES` | `User.Read Mail.ReadWrite Mail.Send MailboxSettings.ReadWrite offline_access` | Space-separated delegated Graph scopes requested at sign-in. |
| `OUTLOOK_TOKEN_STORE` | `auto` | Where the token cache lives: `auto` (OS keychain when available, else file), `keychain` (fails if no keychain) or `file`. See [Token storage](#️-token-storage). |
| `OUTLOOK_TOKEN_CACHE` | `~/.config/outlook-mcp/token-cache.json` | Token cache file (mode `0600`) for the `file` store; also identifies the keychain entry. |
| `OUTLOOK_READ_ONLY` | `false` | When `true`, send/move/delete/flag/mark tools are not registered at all. |
| `OUTLOOK_DOWNLOAD_DIR` | `~/Downloads` | The only directory attachments may be written to. |
| `OUTLOOK_DEFAULT_TOP` | `20` | Default page size for list/search tools (1-100). |
| `OUTLOOK_MAX_BODY_CHARS` | `20000` | Maximum body characters returned by `get_message`. |
| `OUTLOOK_GRAPH_BASE_URL` | `https://graph.microsoft.com/v1.0` | Graph endpoint (change for national clouds). |
| `LOG_LEVEL` | `info` | `debug`, `info`, `warn` or `error` (written to stderr). |

## 🔐 Authentication

```bash
npx -y @pepebits/outlook-mcp auth      # device code sign-in; prints a URL and a code
npx -y @pepebits/outlook-mcp whoami    # silent token + GET /me
npx -y @pepebits/outlook-mcp logout    # removes accounts and deletes the token cache (keychain entry and file)
```

You can also sign in without a terminal. Just ask your assistant to *"log in to Outlook"*:

- 🔑 `login` returns a URL and a one-time code. Open the URL, enter the code and accept; sign-in finishes in the background.
- ✅ `auth_status` tells you whether you are signed in, still waiting, or signed out.
- 🚪 `logout` removes the cached account and the token cache (keychain entry and file).

Use `login` with `force: true` to sign in again, for example after adding a permission in Azure. If the session is missing or expired, tools return *"Not signed in or session expired. Call the login tool ..."*.

## 🗝️ Token storage

The refresh token is the most sensitive thing this server holds, so by default (`OUTLOOK_TOKEN_STORE=auto`) it goes into the operating system keychain instead of a plain file. No native npm dependencies are used; the server calls the OS tools directly:

| Platform | Backend | Notes |
| --- | --- | --- |
| macOS | Keychain via the `security` CLI | Service `outlook-mcp`, account = the cache path. |
| Linux | Secret Service via `secret-tool` (libsecret) | Used only if `secret-tool` is installed and a keyring is running; otherwise the file store is used. |
| Windows and others | File | `~/.config/outlook-mcp/token-cache.json`, mode `0600`. |

- **Migration:** if a cache file exists and the keychain is still empty, it is imported into the keychain and the file is deleted.
- **Fallback:** in `auto` mode, if the keychain cannot be used at runtime the server falls back to the file and logs a warning. `OUTLOOK_TOKEN_STORE=keychain` never falls back and fails instead; `file` never touches the keychain.
- **Secrets stay out of process listings:** on Linux `secret-tool` reads the secret from stdin. `security add-generic-password` only accepts the password as a command-line argument (visible in `ps` to other local processes), so on macOS the commands are fed to `security -i` over stdin instead. That interface truncates lines at about 4 KB, so the cache is stored base64 encoded across a few small keychain items (`<account>#0`, `#1`, ...).
- `logout` (tool and CLI) removes the keychain entry as well as the file.
- On macOS the first access from a different binary may show a Keychain prompt; click *Always Allow*.

## 🛠️ Use your own Azure app (optional)

By default `outlook-mcp` uses the shared **outlook-mcp** Azure app, which works with personal Microsoft accounts only. Register your own app if you need **work or school** accounts or simply prefer to use your own. Then set `OUTLOOK_CLIENT_ID` (and `OUTLOOK_TENANT`, see below).

### Step by step

1. Open [portal.azure.com](https://portal.azure.com) and go to **Microsoft Entra ID** -> **App registrations** -> **New registration**.
2. Give it a name (for example `outlook-mcp`).
3. Under **Supported account types** choose:
   - **Personal Microsoft accounts only** (Outlook.com, Hotmail, Live), or
   - **Accounts in any organizational directory and personal Microsoft accounts** (both).
   - For work/school only, pick an organizational option.
4. Leave **Redirect URI** empty and click **Register**.
5. In the left menu go to **Manage** -> **Authentication** -> **Settings** tab, set **Allow public client flows** to **Yes** and save.
6. Go to **Manage** -> **API permissions**. `User.Read` is already granted by default. Click **Add a permission** -> **Microsoft Graph** -> **Delegated permissions** and add:
   - 📬 Expand the **Mail** group and tick `Mail.ReadWrite` and `Mail.Send`.
   - ⚙️ Expand the **MailboxSettings** group and tick `MailboxSettings.ReadWrite` (needed for inbox rules and `block_sender`).
   - 🔑 Expand the **OpenId permissions** group and tick `offline_access`.

   Then click **Add permissions**. No admin consent is needed for personal accounts.

| Scope | Used for |
| --- | --- |
| `User.Read` | Sign-in and `whoami` |
| `Mail.ReadWrite` | Reading, searching, moving, flagging and deleting messages |
| `Mail.Send` | `send_mail`, replies, forwards and mailto unsubscribe |
| `MailboxSettings.ReadWrite` | Inbox rules: `list_rules`, `create_rule`, `delete_rule`, `block_sender` |
| `offline_access` | Refresh token, so you only sign in once |

> 💡 If you upgrade from 0.1.0, add `MailboxSettings.ReadWrite` in Azure and run `login` with `force: true` (or `npx -y @pepebits/outlook-mcp auth`) again.

> 💡 The Azure portal may be shown in your language, so labels can differ slightly (e.g. *Administrar* -> *Autenticación* -> *Configuración*, *Permisos de OpenId*).
7. From the **Overview** page copy the **Application (client) ID**. This is your `OUTLOOK_CLIENT_ID`.

> 🏢 **Work or school accounts:** set `OUTLOOK_CLIENT_ID` to your app and `OUTLOOK_TENANT=organizations` (or your tenant GUID). Your organization may require **admin consent** for the mail permissions.

## 🧰 Tools reference

| Tool | Mutating | Description |
| --- | :---: | --- |
| `login` | no | Start a device code sign-in; returns the URL and code (`force` to sign in again). |
| `auth_status` | no | Signed in, pending sign-in or signed out. |
| `logout` | no | Sign out and delete the local token cache. |
| `list_folders` | no | List mail folders or child folders (`parentFolderId`, `includeHidden`). |
| `list_messages` | no | List messages, newest first, with filters (`unreadOnly`, `from`, `since`, `until`, `hasAttachments`) and `nextLink` paging. |
| `search_messages` | no | KQL full-text search with `nextLink` paging. |
| `get_message` | no | Full message with recipients, flags and body (`format` text/html, `maxChars`). |
| `list_attachments` | no | Attachment metadata for a message. |
| `download_attachment` | no | Save an attachment into `OUTLOOK_DOWNLOAD_DIR` (writes locally only; no overwrite unless `overwrite: true`). |
| `send_mail` | yes | Send an email, optionally with local attachments under 3 MB. |
| `create_draft` | yes | Create a draft without sending. |
| `reply_message` | yes | Reply or reply-all (`replyAll`). |
| `forward_message` | yes | Forward to new recipients. |
| `move_message` | yes | Move to a folder; returns the new id. Accepts `id` or `ids` (up to 50). |
| `mark_read` | yes | Mark read or unread. Accepts `id` or `ids` (up to 50). |
| `flag_message` | yes | `flagged`, `complete` or `notFlagged`. Accepts `id` or `ids` (up to 50). |
| `delete_message` | yes | Move to Deleted Items, or `permanent: true` to delete irreversibly. Accepts `id` or `ids` (up to 50). |
| `get_unsubscribe_info` | no | Show how to leave the mailing list a message came from (List-Unsubscribe). |
| `unsubscribe` | yes | Leave a mailing list: RFC 8058 one-click, else a mailto request, else returns the link. Accepts `id` or `ids` (up to 50). |
| `find_newsletters` | no | Scan a folder (`folderId`, `maxMessages` up to 2000, `excludeDomains`, `includeNoUnsubscribe`), group by sender and report mailing lists with their unsubscribe method, most frequent first. |
| `list_rules` | no | List inbox rules. |
| `create_rule` | yes | Create an inbox rule from `fromAddresses` / `senderContains` / `subjectContains` with action `move`, `delete`, `markRead` or `junk`. |
| `delete_rule` | yes | Delete an inbox rule by id. |
| `block_sender` | yes | Block `addresses` and/or `domains` with a rule that moves their future mail to Deleted Items. |

**Bulk operations:** the tools marked "Accepts `id` or `ids`" take exactly one of the two. With `ids` they process up to 50 messages (a few at a time), never stop at the first error and return `{ results: [{ id, ok, ... | error }], succeeded, failed }`. With `id` the response is unchanged.

Well-known folder names accepted anywhere a folder id is expected: `inbox`, `drafts`, `sentitems`, `deleteditems`, `junkemail`, `archive`.

## 🧹 Cleaning up your inbox

Ask Claude something like *"Find the newsletters cluttering my inbox, unsubscribe from the ones I don't read and block the rest"*. Behind the scenes:

1. `find_newsletters` scans the inbox and lists senders with a `List-Unsubscribe` header, most frequent first, each with a `sampleMessageId` and an `unsubscribe` method (`one-click`, `mailto` or `link`).
2. `unsubscribe` with `ids` set to the chosen sample message ids leaves those lists. Senders that only offer a link come back with the URL to open in a browser.
3. `block_sender` with `addresses` or `domains` creates an inbox rule that sends their future mail to Deleted Items. Use `delete_message` with `ids` to clear what is already in the inbox, and `list_rules` / `delete_rule` to review or undo a block.

Inbox rules need the `MailboxSettings.ReadWrite` permission (see the Azure steps above). Do not unsubscribe from spam or phishing: it confirms your address is active. Block those senders instead.

## 👁️ Read-only mode

Set `OUTLOOK_READ_ONLY=true` to register only the non-mutating tools. The mutating tools do not exist for the client, so they cannot be called at all. Combine it with a token cache created using only `User.Read Mail.Read` (set `OUTLOOK_SCOPES` accordingly and re-run `npx -y @pepebits/outlook-mcp auth`) for defense in depth.

## 🛡️ Security

- 🔑 The token cache is stored in the OS keychain when available (macOS Keychain, Linux Secret Service; see [Token storage](#️-token-storage)), otherwise in a local file with mode `0600` in a `0700` directory, written atomically.
- 🙅 No client secret exists: this is a public client using device code flow.
- 🤐 Tokens and message bodies are never logged; MSAL PII logging is disabled.
- 📥 Attachment downloads are confined to `OUTLOOK_DOWNLOAD_DIR`; path traversal is rejected.
- ⚠️ `delete_message` with `permanent: true` is **irreversible**, and `send_mail`, `reply_message` and `forward_message` send immediately. Prefer read-only mode or `create_draft` when in doubt.
- 🧠 Email content is untrusted input: a malicious message may try to instruct the model (prompt injection). To help the model tell data from instructions, results from `get_message`, `list_messages`, `search_messages`, `list_attachments` and `find_newsletters` group sender-controlled text (subjects, senders, previews, attachment names) under an `untrusted` key and carry a `note` saying so. The `get_message` body is enclosed in `<untrusted_email_content>` markers, and any marker found inside the content is neutralized so it cannot close the wrapper early.
- 🤝 The server also sends MCP `instructions` telling clients that email content is untrusted and that sending, deleting, unsubscribing and rule changes should be confirmed with the user. These are mitigations, not guarantees: keep a human in the loop for those actions or use read-only mode.

## 🩺 Troubleshooting

| Symptom | Fix |
| --- | --- |
| `AADSTS7000218` (client assertion / secret required) | Enable **Allow public client flows** in Azure -> Manage -> Authentication -> Settings. |
| `AADSTS50020` or wrong tenant | The account type does not match `OUTLOOK_TENANT`. Use `consumers` for personal, `organizations` for work/school, `common` for both, and make sure the app registration supports that account type. |
| "Not signed in or session expired" | Ask the assistant to call `login`, or run `npx -y @pepebits/outlook-mcp auth`. |
| Search returns "Invalid search query" | `search_messages` uses KQL, e.g. `from:alice subject:"report" hasattachments:true`. Graph returns at most about 250 results per search, and results are not sorted. |
| "Message not found" after a move | Message ids change when a message is moved. Use the `newId` returned by `move_message` or list the folder again. |
| "Access denied for inbox rules" | Add `MailboxSettings.ReadWrite` in Azure -> API permissions, make sure it is in `OUTLOOK_SCOPES`, then run `npx -y @pepebits/outlook-mcp auth` again. |
| Access denied / consent required | Add the missing delegated permission, or ask an admin for consent. |
| Server does not start in a client | Check stderr logs; if you use your own app, `OUTLOOK_CLIENT_ID` must be set via `.env` or the client's `env` block. |

## 🧑‍💻 Development

```bash
git clone https://github.com/Pepebits/outlook-mcp.git
cd outlook-mcp
npm install
cp .env.example .env        # optional
npm run build
npm test
```

Run it from source in your MCP client:

```bash
claude mcp add outlook --scope user -- node /absolute/path/outlook-mcp/dist/index.js
```

Use `npm run inspect` to try the tools in the MCP Inspector.

### 🚢 Releasing (maintainers)

1. Move the `[Unreleased]` notes in `CHANGELOG.md` under the new version, bump the version in `package.json` and `server.json` (top level and package), commit and push.
2. Tag it: `git tag -a vX.Y.Z -m "Short summary" && git push origin vX.Y.Z`. The Release workflow runs the tests, **stages** the version on npm (Trusted Publishing, with provenance) and creates the GitHub release.
3. Approve the staged version on npmjs.com (or `npm stage approve <stage-id>`), which needs your 2FA.
4. The workflow waits for the approval, lists the version in the MCP Registry and announces the release on Telegram. If you approve more than ~6 hours later, the wait times out: run the Release workflow by hand (**Actions → Release → Run workflow**) with the tag to announce it.

## 🗺️ Roadmap

- 📅 Calendar support (`Calendars.ReadWrite`)
- 📦 Large attachments through upload sessions
- 👥 Contacts

## 🤝 Contributing

Issues and pull requests are welcome.

```bash
npm install
npm run typecheck
npm test
npm run build
```

Please keep code, comments and docs in English, add tests for new behavior, and never log secrets or message content.

## 📄 License

[MIT](LICENSE)
