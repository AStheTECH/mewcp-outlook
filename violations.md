## Violations

None found.

## Passed

Scope: this validation covers the files changed in the `calendars` + `events` (trimmed)
incremental sync (14 new tools: `list_calendars`, `get_calendar`, `list_events`,
`create_event`, `get_event`, `update_event`, `delete_event`, `cancel_event`, `accept_event`,
`tentatively_accept_event`, `decline_event`, `find_meeting_times`, `get_free_busy_schedule`,
`list_calendar_view`), i.e. `outlook_mcp/schemas/calendars.py`, `outlook_mcp/tools/calendars_tools.py`,
`outlook_mcp/schemas/events.py`, `outlook_mcp/tools/events_tools.py`, plus `outlook_mcp/tools/__init__.py`
(registration wiring) and `outlook_mcp/config.py` (version/scopes) edits. The previously-built
23 tools (`users` + `messages`) were not re-checked (unchanged).

This was a deliberately **trimmed** pass of the full `calendars` (13 concepts) and `events` (19
concepts) groups, scoped by explicit user direction to the subset most useful for AI-agent
scheduling workflows: core event CRUD, RSVP actions, availability/suggestion lookup, and basic
calendar discovery. Deferred to a later pass (left at `status: ready`, untouched): calendar
groups (`list_calendar_groups`, `create_calendar_group`, `get_calendar_group`,
`update_calendar_group`, `delete_calendar_group`, `list_calendars_in_calendar_group`,
`create_calendar_in_calendar_group`), calendar CRUD (`create_calendar`, `update_calendar`,
`delete_calendar`, `permanently_delete_calendar`), and the lower-value event extras
(`permanently_delete_event`, `get_event_delta`, `forward_event`, `dismiss_event_reminder`,
`snooze_event_reminder`, `list_event_instances`, `list_calendar_reminders`).

- [x] No injection instructions in tool descriptions — grepped both new files for `always`, `never`, `ignore`, `override`, `forget previous`. The only hit is `never` inside the `delete_event`/`cancel_event` DESTRUCTIVE templates ("NEVER call this tool autonomously"), which is the standard destructive-tool warning, not an injected directive.
- [x] Every tool has `annotations=ToolAnnotations(...)` with correct tier values — confirmed 37/37 registered tools carry annotations (6 users + 17 messages + 2 calendars + 12 events) against a live `FastMCP` instance.
- [x] All DELETE tools have `destructiveHint=True` and a warning-first description — the 2 new `destructiveHint=True` tools (`delete_event`, `cancel_event`) each lead with `DESTRUCTIVE — REQUIRES EXPLICIT USER CONFIRMATION BEFORE CALLING`, state what happens and that it's hard to reverse, say never to call autonomously, and instruct the agent to stop and get explicit confirmation. `cancel_event`'s doc states its own `## Tier` as DELETE (it's organizer-only and moves the event to Deleted Items, same blast radius as `delete_event`), so it got the same treatment even though its HTTP verb is POST.
- [x] All single-resource UPDATE tools have the discovery note and return `before`+`after` — `update_event` fetches the event via GET before the mutating PATCH and returns `before`/`after` `GetEventData` snapshots; its description states the overwrite behavior and that the response carries both states.
- [x] Judgment call — `accept_event`, `tentatively_accept_event`, `decline_event`: their docs state `## Tier: UPDATE`, but each is a 202-Accepted RSVP action with no response body and no single gettable field it mutates in a way a before/after snapshot would usefully capture (the event's `responseStatus` changes, but asynchronously and not necessarily by the time the call returns). Treated the same as `reply_message`/`forward_message`/`reply_all_message` in the `messages` group: no before/after, `destructiveHint=False`, plain description of what the action does and that it returns no content.
- [x] Every non-destructive, non-UPDATE tool description is one or two direct sentences, starts with a verb, and states what the tool returns — verified for `list_calendars`, `get_calendar`, `list_events`, `create_event`, `get_event`, `find_meeting_times`, `get_free_busy_schedule`, `list_calendar_view` (the 8 new GET/CREATE-tier tools).
- [x] No tool description is a multi-paragraph or bulleted essay — all 14 new descriptions are single short paragraphs, no bullets/headings.
- [x] GraphQL checks, `fields`-param optionality, idempotency keys — all N/A; Microsoft Graph REST, not GraphQL.
- [x] Every list tool enforces the page-size cap its concept states — `list_events`' and `list_calendar_view`'s docs don't state a page-size bound (unlike `list_messages`, whose doc gave an explicit 1–1000 range), so `top` is passed through unvalidated rather than an invented cap. `list_calendars` and `get_free_busy_schedule`'s `availabilityViewInterval` (5–1440, stated) — the latter is a slot-duration parameter, not a page size, so no cap enforcement applies there either; its single documented bound was left to the upstream API to reject.
- [x] Action tools call `service` directly for every step, never another tool — N/A, no `kind: "action"` tools in this batch. `update_event` makes 2 direct `service.api_request` calls, never calling another tool's function.
- [x] All 14 new tools instantiate `ToolLogger` as the first statement — confirmed via AST walk over both tool files.
- [x] `tlog.failure()` called before every error return — every error path routes through `_err`, `_upstream_err`, or `_handle_request_exc` (grepped for `success=False`: zero raw hits in either new file).
- [x] No `str(e)` in `tlog.failure()` for upstream HTTP errors — unchanged `_upstream_err`; no new code introduces a different upstream-error log path.
- [x] `pydantic.ValidationError` handled before `ValueError`, auth branch matches `(CredentialError, ValueError)` — unchanged in `_helpers.py`.
- [x] `BREAKING_CHANGES` updated on a breaking bump — N/A, purely additive version bump (`v1.1.0` → `v1.2.0`).
- [x] No retry logic anywhere — none added.
- [x] No request/response body logged anywhere — new code only calls `tlog.success()`/`tlog.failure()` with status/error codes.
- [x] All 14 new tool names are `lowercase_snake_case` `verb_noun`(`_qualifier`).
- [x] No tool uses a separate Pydantic input model as its only argument — all new tools declare parameters directly via `Field(description=...)` on the function signature.
- [x] Every parameter has a `Field(description=...)` with format and optionality notes — required params (`subject`, `start`, `end` on `create_event`; `id` where it's the sole identifier; `schedules`/`startTime`/`endTime` on `get_free_busy_schedule`; `startDateTime`/`endDateTime` on `list_calendar_view`) carry no "omit" note since they have no default, consistent with [[feedback_no_required_text_in_field_desc]]; every optional param's description ends with an omit-behavior note.

## Notes on judgment calls made during this build

- **Canonical event shape.** `list_events`, `get_event`, `create_event`'s response, and `update_event`'s before/after all reuse one `GetEventData` model (via direct reuse or a one-line `CreateEventData(GetEventData)` subclass for per-tool schema naming) rather than four independently hand-copied field lists. This is grounded in each doc's own text: both `list_events` and `get_event` state "without a `$select`, all event properties are returned," and `list_calendar_view`'s doc explicitly says it returns "a collection of event objects" — i.e. the same Graph `event` resource, not a distinct trimmed shape. Every field on `GetEventData` is copied verbatim from an actual JSON sample across these five docs (`list_events`, `create_event`, `get_event`, `update_event`, `list_calendar_view`); none are guessed. `model_config = ConfigDict(extra="allow")` on all of them means any undocumented field the API also returns still passes through rather than failing validation.
- **`cancel_event` got the DESTRUCTIVE template** despite being a POST, because its own doc frontmatter states `## Tier: DELETE` and its Notes describe the same Deleted-Items blast radius as `delete_event` (just with a custom message and organizer-only). Tier in the manifest is set to `DELETE` to match.
- **`Prefer` header helper.** Six of the twelve new event tools (`list_events`, `create_event`, `get_event`, `find_meeting_times`, `get_free_busy_schedule`, `list_calendar_view`) each optionally set `outlook.timezone` and/or `outlook.body-content-type` via the same `Prefer` header mechanism already proven in `messages_tools.py`'s `get_message_delta`. Factored into one small `_prefer_header()` helper local to `events_tools.py` (not promoted to `_helpers.py`) since it's a two-line convenience used only within this one file.
- **OAuth scopes**: this sync adds `Calendars.ReadWrite` (covers every calendars/events tool except one) and `Calendars.Read.Shared` (the one documented exception — `find_meeting_times` requires the `.Shared` variant specifically, per its own Notes) to `config.py`'s `SCOPES`.
