# 0003. Bind `bulk_action` apply to a dry-run selection

- Status: Accepted
- Date: 2026-11-16

## Context

`bulk_action` (`src/tools/bulkaction.ts`) can delete, move or re-flag up to 500 messages per call. Its safety model is a convention: the description tells the model to call with `dryRun` (default true) first, show the count to the user, then call again with `dryRun: false`. The server does not hold the two calls together:

- `dryRun: false` works as a first call. Nothing proves a preview happened.
- The apply call re-runs the query. Mail that arrived in between is acted on without having been previewed, and for `$search` (relevance-ordered, capped) or any truncated selection the applied set can differ from the previewed one.
- The preview shows only 10 samples of up to 500 matches.

Because the tool is reachable by a model that reads untrusted email, "confirm first" has to be something the server can verify, not something it asks for.

## Decision

A dry run creates a server-side selection and returns a token; apply requires that token.

1. A dry-run call resolves the matches, stores `{ ids, action, destinationFolderId, createdAt }` in memory under a random `selectionId`, and returns it with the count, sample and `truncated` flag.
2. An apply call takes only `selectionId` (plus `dryRun: false`). It acts on exactly the stored ids with the stored action; it does not re-query. Search and filter parameters on an apply call are rejected.
3. Selections expire after 10 minutes, are single-use, and at most a small number (e.g. 8) are kept. They live in process memory only and vanish on restart.
4. `dryRun: false` without a valid `selectionId` is an error that explains the two-step flow.

## Alternatives considered

- **Keep re-querying (status quo).** Stateless and simple, but the preview and the action are unrelated calls.
- **Return all matched ids from the dry run and have the client pass them back.** Stateless, but up to 500 opaque ids round-trip through the model context, cost tokens, can be truncated or altered, and `ids` is already capped at 50 elsewhere.
- **MCP elicitation for confirmation.** The right long-term mechanism (the user, not the model, confirms), but client support is uneven and a server that requires it stops working in clients without it. Can be layered on top of the token later.
- **Stateless signed token encoding the query.** No server state, but still re-queries, so it does not fix drift between preview and apply.

## Consequences

- The applied set is exactly the previewed set; new mail arriving between the calls is never touched.
- A model cannot skip the preview: an injected "call bulk_action with dryRun false" fails without a prior dry run. The model can still chain both calls itself, so host-side confirmation remains necessary; this decision narrows the gap rather than closing it.
- Breaking change to the tool's input contract (apply calls with query/filter parameters stop working). Ships with a CHANGELOG entry; the description and README examples change.
- The server gains its first in-memory state besides the auth session. It is bounded and non-persistent, and a restart between preview and apply simply requires a new preview.
- Messages moved or deleted by the user between preview and apply fail individually with the existing per-item error reporting.
