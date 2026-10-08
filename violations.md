## Violations

None found.

## Passed

Scope: this validation covers the files changed in the `users` + `messages` incremental
sync (11 new tools: `list_users`, `create_user`, `update_user`, `delete_user`,
`get_users_delta`, `update_message`, `delete_message`, `copy_message`, `get_message_delta`,
`move_message`, `permanently_delete_message`), i.e. `outlook_mcp/schemas/users.py`,
`outlook_mcp/tools/users_tools.py`, `outlook_mcp/schemas/messages.py`,
`outlook_mcp/tools/messages_tools.py`, plus the supporting `outlook_mcp/service.py`
(`extra_headers` passthrough) and `outlook_mcp/config.py` (version/scopes) edits. The
previously-built 12 tools were not re-checked (unchanged).

`revoke_sign_in_sessions` and `change_password` were planned, built, and then **removed**
before shipping, per explicit user direction: an Outlook/mail-automation server handing an
LLM agent a blanket "revoke all sessions" / "change password" tool was judged too large a
blast radius for the benefit, given the agent could be manipulated by injected content (e.g.
a malicious email it reads). Their scopes (`User.RevokeSessions.All`,
`User-PasswordProfile.ReadWrite.All`) were removed from `config.py`'s `SCOPES`, and their
source docs (`users/revoke_sign_in_sessions.md`, `users/change_password.md`) were reverted
from `status: done` back to `status: ready` with a note explaining the exclusion, so they
don't get silently skipped as "already built" on a future run, and don't get silently
rebuilt without the same reasoning being re-applied.

- [x] No injection instructions in tool descriptions — grepped all new `description=`/`Field()` text for `always`, `never`, `ignore`, `override`, `forget previous`. Hits are plain technical references: `reply_message`'s pre-existing "override properties" (unchanged), and the new `get_users_delta`/`get_message_delta` `select` params' "id is always returned" (a factual statement about the API, not a directive to an agent).
- [x] Every tool has `annotations=ToolAnnotations(...)` with correct tier values — confirmed 23/23 registered tools carry annotations (17 in `messages_tools.py` + 6 in `users_tools.py`, matching the full tool count after registering both modules against a live `FastMCP` instance).
- [x] All DELETE tools have `destructiveHint=True` and a warning-first description — exactly 3 new `destructiveHint=True` tools (`delete_user`, `delete_message`, `permanently_delete_message`), each leading with `DESTRUCTIVE — REQUIRES EXPLICIT USER CONFIRMATION BEFORE CALLING`, stating what's removed and that it's irreversible (or becomes irreversible after a retention window), saying never to call autonomously, and instructing the agent to stop and get explicit confirmation. No UPDATE_MANY tools exist.
- [x] All single-resource UPDATE tools have the discovery note and return `before`+`after` — `update_user`, `update_message`, and `move_message` (the 3 remaining new UPDATE-tier tools) each fetch the resource via GET before the mutating call and return `before`/`after` snapshots (`UserData` or `GetMessageData`), and each description states the overwrite behavior and that the response carries both states. No exception cases remain in this batch.
- [x] Every non-destructive, non-UPDATE tool description is one or two direct sentences, starts with a verb, and states what the tool returns — verified for `list_users`, `create_user`, `get_users_delta`, `copy_message`, `get_message_delta` (the 5 new GET/CREATE-tier tools): each opens with Lists/Creates/Gets/Copies and ends with an explicit statement of what's returned.
- [x] No tool description is a multi-paragraph or bulleted essay — all 11 new descriptions are single short paragraphs, no bullets/headings.
- [x] GraphQL checks (`_graphql_err`/`_user_errors`), GraphQL selection constants, `fields`-param optionality, get-by-ID null-root handling, idempotency keys — all N/A; this server is Microsoft Graph REST, not GraphQL, and no tool has a literal `fields` parameter.
- [x] Every list tool enforces the page-size cap its concept states — `list_users` validates `top` to 1–999 (`outlook_mcp/tools/users_tools.py`) before issuing the request, matching the existing `list_messages` pattern. `get_users_delta`/`get_message_delta` are delta-sync tools, not plain list tools, and their docs state no hard page-size requirement beyond the optional `Prefer` header already exposed as a param.
- [x] Action tools call `service` directly for every step, never another tool — N/A, no `kind: "action"` tools in this batch. (The before/after UPDATE tools make 2–3 direct `service.api_request` calls each, never calling another tool's function.)
- [x] All 11 new tools instantiate `ToolLogger` as the first statement — confirmed via AST walk over both tool files: every tool function's first statement is `tlog = ToolLogger(logger, "...")`.
- [x] `tlog.failure()` called before every error return — every error path in the new code routes through `_err`, `_upstream_err`, or `_handle_request_exc` (grepped for any raw `success=False` construction bypassing these helpers: none found).
- [x] No `str(e)` in `tlog.failure()` for upstream HTTP errors — unchanged `_upstream_err` still logs `f"HTTP {status}"`; no new code introduces a different upstream-error log path.
- [x] `pydantic.ValidationError` handled before `ValueError`, and the auth branch matches `(CredentialError, ValueError)` — unchanged in `_helpers.py`; no new code added its own exception handling that could reorder this.
- [x] `BREAKING_CHANGES` updated on a breaking bump — N/A, this is a purely additive version bump (`v1.0.0` → `v1.1.0`), no tool/param/response field was renamed or removed, so `BREAKING_CHANGES` correctly stays empty.
- [x] No retry logic anywhere — no retry/backoff loops added.
- [x] No request/response body logged anywhere — new code only calls `tlog.success()`/`tlog.failure()` with status/error codes, never body content.
- [x] All 11 new tool names are `lowercase_snake_case` `verb_noun`(`_qualifier`) — `list_users`, `create_user`, `update_user`, `delete_user`, `get_users_delta`, `update_message`, `delete_message`, `copy_message`, `get_message_delta`, `move_message`, `permanently_delete_message`.
- [x] No tool uses a separate Pydantic input model as its only argument — all new tools declare parameters directly via `Field(description=...)` on the function signature.
- [x] Every parameter has a `Field(description=...)` with format and optionality notes — verified across all 11 tools' parameters (e.g. `update_user`'s ~43 optional properties each end "Omit to leave unchanged."; `list_users`' OData params each state their omit-behavior; required params like `destinationId` and the two `permanently_delete_message` path params carry no omit-note since they have no default, consistent with [[feedback_no_required_text_in_field_desc]]).

## Notes on judgment calls made during this build

- **`update_user`** supports an undocumented-in-`resource:`-frontmatter admin variant (targeting another user via `PATCH /users/{user_id}`) that the doc body explicitly states exists alongside the `/me` form. Per explicit user confirmation during planning, it got an optional `user_id` param that redirects to the admin-targeted endpoint when supplied, rather than being split into a separate tool or ignoring the documented variant.
- **`revoke_sign_in_sessions` and `change_password` were removed after being built** — see the scope note above.
- **OAuth scopes**: this sync adds `User.ReadBasic.All`, `User.ReadWrite.All`, and `Directory.ReadWrite.All` to `config.py`'s `SCOPES` — broader than the prior `User.Read`/`Mail.ReadWrite`/`Mail.Send` set, since `create_user`/`update_user`/`delete_user` are directory-admin operations. These scopes typically require admin consent in Entra ID.
