"""Events group: list_events, create_event, get_event, update_event, delete_event,
cancel_event, accept_event, tentatively_accept_event, decline_event, find_meeting_times,
get_free_busy_schedule, list_calendar_view."""

import logging
from typing import Any

from fastmcp import FastMCP
from mcp.types import ToolAnnotations
from pydantic import Field

from .. import service
from ..config import CONNECT_TIMEOUT, READ_TIMEOUT
from ..logging_utils import ToolLogger
from ..schemas.events import (
    AcceptEventData,
    AcceptEventResult,
    CancelEventData,
    CancelEventResult,
    CreateEventData,
    CreateEventResult,
    DeclineEventData,
    DeclineEventResult,
    DeleteEventData,
    DeleteEventResult,
    FindMeetingTimesData,
    FindMeetingTimesResult,
    GetEventData,
    GetEventResult,
    GetFreeBusyScheduleData,
    GetFreeBusyScheduleResult,
    ListCalendarViewData,
    ListCalendarViewResult,
    ListEventsData,
    ListEventsResult,
    TentativelyAcceptEventData,
    TentativelyAcceptEventResult,
    UpdateEventData,
    UpdateEventResult,
)
from ._helpers import _err, _handle_request_exc, _upstream_err

logger = logging.getLogger("outlook-mcp.tools.events")


def _prefer_header(timezone: str | None, body_content_type: str | None = None) -> dict[str, str] | None:
    parts = []
    if timezone is not None:
        parts.append(f'outlook.timezone="{timezone}"')
    if body_content_type is not None:
        parts.append(f'outlook.body-content-type="{body_content_type}"')
    return {"Prefer": ", ".join(parts)} if parts else None


def register_events_tools(mcp: FastMCP) -> None:

    @mcp.tool(
        name="list_events",
        description=(
            "Lists single-instance meetings and series masters in the signed-in user's "
            "mailbox — not expanded recurring occurrences, use list_calendar_view for that. "
            "Returns a page of events; follow the response's `odata_next_link` to page "
            "through further results."
        ),
        annotations=ToolAnnotations(readOnlyHint=True, destructiveHint=False, openWorldHint=True),
    )
    def list_events(
        select: str | None = Field(
            default=None,
            description="Comma-separated list of properties to return, e.g. 'subject,start,end'. Omit to return all event properties.",
        ),
        filter: str | None = Field(
            default=None,
            description="OData filter expression to restrict the events returned. Can't be used on the `recurrence` property. Omit to apply no filter.",
        ),
        orderby: str | None = Field(
            default=None,
            description="Property to sort the results by, e.g. 'start/dateTime'. Omit to use the API's default ordering.",
        ),
        top: int | None = Field(
            default=None,
            description="Maximum number of events to return per page. Omit to use the API default.",
        ),
        skip: int | None = Field(
            default=None,
            description=(
                "Number of results to skip. Omit to start from the first result. Prefer "
                "following the returned `odata_next_link` for paging instead of setting this manually."
            ),
        ),
        timezone: str | None = Field(
            default=None,
            description="Time zone name for the returned start/end values, sent as the Prefer: outlook.timezone header, e.g. 'Pacific Standard Time'. Omit to receive times in UTC.",
        ),
        body_content_type: str | None = Field(
            default=None,
            description="Set to 'text' to receive the event body as plain text instead of HTML. Omit to receive HTML.",
        ),
    ) -> ListEventsResult:
        tlog = ToolLogger(logger, "list_events")

        params: dict[str, Any] = {}
        if select is not None:
            params["$select"] = select
        if filter is not None:
            params["$filter"] = filter
        if orderby is not None:
            params["$orderby"] = orderby
        if top is not None:
            params["$top"] = top
        if skip is not None:
            params["$skip"] = skip

        try:
            data, status, retry_after = service.api_request(
                "GET", "/me/events", params=params or None,
                extra_headers=_prefer_header(timezone, body_content_type),
                timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return ListEventsResult(success=True, statusCode=status, data=ListEventsData(**data))
            return _upstream_err(ListEventsResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(ListEventsResult, tlog, exc)

    @mcp.tool(
        name="create_event",
        description=(
            "Creates a new event on the signed-in user's default calendar. Including "
            "`attendees` sends meeting invitations immediately — this can't be suppressed. "
            "Returns the created event."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=True),
    )
    def create_event(
        subject: str = Field(description="Subject line of the event."),
        start: dict[str, Any] = Field(description="Start time as {dateTime: <string>, timeZone: <string>}."),
        end: dict[str, Any] = Field(description="End time as {dateTime: <string>, timeZone: <string>}."),
        body: dict[str, Any] | None = Field(
            default=None,
            description="Event body as {contentType: 'HTML' or 'Text', content: <string>}. Omit to create the event with an empty body.",
        ),
        location: dict[str, Any] | None = Field(
            default=None, description="Location as {displayName: <string>, ...}. Omit for no location.",
        ),
        attendees: list[dict[str, Any]] | None = Field(
            default=None,
            description=(
                "Attendees, each as {emailAddress: {address, name}, type: 'required'|'optional'|'resource'}. "
                "Omit to create the event with no attendees. Sends invitations immediately if supplied."
            ),
        ),
        allowNewTimeProposals: bool | None = Field(
            default=None, description="Whether invitees are allowed to propose a new time. Defaults to true if omitted.",
        ),
        transactionId: str | None = Field(
            default=None, description="Client-generated ID for tracking or idempotency across clients. Omit if not needed.",
        ),
        isOnlineMeeting: bool | None = Field(
            default=None, description="Whether to make this an online meeting. Omit to leave the API default (false).",
        ),
        onlineMeetingProvider: str | None = Field(
            default=None, description="Online meeting provider, e.g. 'teamsForBusiness'. Required if isOnlineMeeting is true. Omit otherwise.",
        ),
        recurrence: dict[str, Any] | None = Field(
            default=None, description="Recurrence pattern as {pattern: {...}, range: {...}}. Omit to create a single, non-recurring event.",
        ),
        sensitivity: str | None = Field(
            default=None, description="normal, personal, private, or confidential. Omit to use the default (normal).",
        ),
        showAs: str | None = Field(
            default=None, description="free, tentative, busy, oof, workingElsewhere, or unknown. Omit to use the default (busy).",
        ),
        reminderMinutesBeforeStart: int | None = Field(
            default=None, description="Minutes before the event start to show a reminder. Omit to use the mailbox default.",
        ),
        isReminderOn: bool | None = Field(
            default=None, description="Whether a reminder is set for the event. Omit to use the mailbox default.",
        ),
        categories: list[str] | None = Field(
            default=None, description="Category names to assign. Omit to assign no categories.",
        ),
        hideAttendees: bool | None = Field(
            default=None, description="When true, only the organizer can see attendees. Omit to leave disabled.",
        ),
        responseRequested: bool | None = Field(
            default=None, description="Whether invitees are asked to send a response. Omit to use the API default (true).",
        ),
        timezone: str | None = Field(
            default=None,
            description="Time zone name to interpret this request's start/end in, sent as the Prefer: outlook.timezone header. Omit to use UTC.",
        ),
    ) -> CreateEventResult:
        tlog = ToolLogger(logger, "create_event")

        payload: dict[str, Any] = {"subject": subject, "start": start, "end": end}
        optional_fields = {
            "body": body, "location": location, "attendees": attendees,
            "allowNewTimeProposals": allowNewTimeProposals, "transactionId": transactionId,
            "isOnlineMeeting": isOnlineMeeting, "onlineMeetingProvider": onlineMeetingProvider,
            "recurrence": recurrence, "sensitivity": sensitivity, "showAs": showAs,
            "reminderMinutesBeforeStart": reminderMinutesBeforeStart, "isReminderOn": isReminderOn,
            "categories": categories, "hideAttendees": hideAttendees,
            "responseRequested": responseRequested,
        }
        payload.update({k: v for k, v in optional_fields.items() if v is not None})

        try:
            data, status, retry_after = service.api_request(
                "POST", "/me/events", body=payload, extra_headers=_prefer_header(timezone),
                timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return CreateEventResult(success=True, statusCode=status, data=CreateEventData(**data))
            return _upstream_err(CreateEventResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(CreateEventResult, tlog, exc)

    @mcp.tool(
        name="get_event",
        description=(
            "Retrieves the full properties of a single event by its id. Use `select` to "
            "request only specific properties."
        ),
        annotations=ToolAnnotations(readOnlyHint=True, destructiveHint=False, openWorldHint=True),
    )
    def get_event(
        id: str = Field(description="Unique identifier of the event."),
        select: str | None = Field(
            default=None,
            description="Comma-separated list of properties to return, e.g. 'subject,start,end'. Omit to return all event properties.",
        ),
        expand: str | None = Field(
            default=None, description="Related resources to expand inline. Omit to expand nothing.",
        ),
        timezone: str | None = Field(
            default=None,
            description="Time zone name for the returned start/end values, sent as the Prefer: outlook.timezone header. Omit to receive times in UTC.",
        ),
        body_content_type: str | None = Field(
            default=None,
            description="Set to 'text' to receive the event body as plain text instead of HTML. Omit to receive HTML.",
        ),
    ) -> GetEventResult:
        tlog = ToolLogger(logger, "get_event")

        params: dict[str, Any] = {}
        if select is not None:
            params["$select"] = select
        if expand is not None:
            params["$expand"] = expand

        try:
            data, status, retry_after = service.api_request(
                "GET", f"/me/events/{id}", params=params or None,
                extra_headers=_prefer_header(timezone, body_content_type),
                timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return GetEventResult(success=True, statusCode=status, data=GetEventData(**data))
            return _upstream_err(GetEventResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(GetEventResult, tlog, exc)

    @mcp.tool(
        name="update_event",
        description=(
            "Updates fields on an existing event. Supply only the fields to change; others "
            "keep their current value. An update that includes only `attendees` sends a "
            "meeting update to just the attendees that changed. NOTE: this overwrites the "
            "current field values and the prior state isn't stored by the API after the "
            "call; the response includes both the before and after state so you have a full "
            "record of what changed."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=True),
    )
    def update_event(
        id: str = Field(description="ID of the event to update."),
        subject: str | None = Field(default=None, description="Subject line of the event. Omit to leave unchanged."),
        body: dict[str, Any] | None = Field(
            default=None,
            description=(
                "Event body as {contentType, content}. If the event is an online meeting, GET the "
                "current body first and preserve the meeting blob when editing — removing it disables "
                "the online meeting. Omit to leave unchanged."
            ),
        ),
        start: dict[str, Any] | None = Field(
            default=None,
            description="Start time as {dateTime, timeZone}. Use only time zones configured for the user's mailbox. Omit to leave unchanged.",
        ),
        end: dict[str, Any] | None = Field(default=None, description="End time as {dateTime, timeZone}. Omit to leave unchanged."),
        location: dict[str, Any] | None = Field(default=None, description="Location as {displayName, ...}. Omit to leave unchanged."),
        locations: list[dict[str, Any]] | None = Field(default=None, description="Multiple locations, each as {displayName, ...}. Omit to leave unchanged."),
        attendees: list[dict[str, Any]] | None = Field(
            default=None,
            description="Attendees, each as {emailAddress, type}. An update with only this field notifies just the changed attendees. Omit to leave unchanged.",
        ),
        categories: list[str] | None = Field(default=None, description="Category names. Omit to leave unchanged."),
        hideAttendees: bool | None = Field(default=None, description="When true, only the organizer can see attendees. Omit to leave unchanged."),
        importance: str | None = Field(default=None, description="low, normal, or high. Omit to leave unchanged."),
        isAllDay: bool | None = Field(default=None, description="Whether the event lasts all day. Omit to leave unchanged."),
        isOnlineMeeting: bool | None = Field(default=None, description="Whether this is an online meeting. Omit to leave unchanged."),
        isReminderOn: bool | None = Field(default=None, description="Whether a reminder is set. Omit to leave unchanged."),
        onlineMeetingProvider: str | None = Field(default=None, description="Online meeting provider, e.g. 'teamsForBusiness'. Omit to leave unchanged."),
        recurrence: dict[str, Any] | None = Field(default=None, description="Recurrence pattern as {pattern, range}. Omit to leave unchanged."),
        reminderMinutesBeforeStart: int | None = Field(default=None, description="Minutes before start to show a reminder. Omit to leave unchanged."),
        responseRequested: bool | None = Field(default=None, description="Whether invitees are asked to respond. Omit to leave unchanged."),
        sensitivity: str | None = Field(default=None, description="normal, personal, private, or confidential. Omit to leave unchanged."),
        showAs: str | None = Field(default=None, description="free, tentative, busy, oof, workingElsewhere, or unknown. Omit to leave unchanged."),
    ) -> UpdateEventResult:
        tlog = ToolLogger(logger, "update_event")

        fields = {
            "subject": subject, "body": body, "start": start, "end": end, "location": location,
            "locations": locations, "attendees": attendees, "categories": categories,
            "hideAttendees": hideAttendees, "importance": importance, "isAllDay": isAllDay,
            "isOnlineMeeting": isOnlineMeeting, "isReminderOn": isReminderOn,
            "onlineMeetingProvider": onlineMeetingProvider, "recurrence": recurrence,
            "reminderMinutesBeforeStart": reminderMinutesBeforeStart,
            "responseRequested": responseRequested, "sensitivity": sensitivity, "showAs": showAs,
        }
        payload = {k: v for k, v in fields.items() if v is not None}
        if not payload:
            return _err(UpdateEventResult, tlog, "VALIDATION_ERROR",
                        "at least one field to change must be provided", 400)

        try:
            before_data, before_status, before_retry = service.api_request(
                "GET", f"/me/events/{id}", timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if not (200 <= before_status < 300):
                return _upstream_err(UpdateEventResult, tlog, before_status, before_data, before_retry)
            before = GetEventData(**before_data)

            after_data, after_status, after_retry = service.api_request(
                "PATCH", f"/me/events/{id}", body=payload, timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if not (200 <= after_status < 300):
                return _upstream_err(UpdateEventResult, tlog, after_status, after_data, after_retry)
            after = GetEventData(**after_data)

            tlog.success()
            return UpdateEventResult(success=True, statusCode=200, data=UpdateEventData(before=before, after=after))
        except Exception as exc:
            return _handle_request_exc(UpdateEventResult, tlog, exc)

    @mcp.tool(
        name="delete_event",
        description=(
            "DESTRUCTIVE — REQUIRES EXPLICIT USER CONFIRMATION BEFORE CALLING. "
            "Deletes an event by moving it to Deleted Items. If the event is a meeting, this "
            "sends a generic cancellation message to attendees — to send a custom "
            "cancellation message instead, use cancel_event (organizer only). NEVER call "
            "this tool autonomously or as part of an automated flow. You MUST stop, tell the "
            "user exactly which event will be deleted, and wait for their explicit written "
            "confirmation before proceeding."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=True, openWorldHint=True),
    )
    def delete_event(
        id: str = Field(description="ID of the event to delete."),
    ) -> DeleteEventResult:
        tlog = ToolLogger(logger, "delete_event")

        try:
            data, status, retry_after = service.api_request(
                "DELETE", f"/me/events/{id}", timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return DeleteEventResult(success=True, statusCode=status, data=DeleteEventData())
            return _upstream_err(DeleteEventResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(DeleteEventResult, tlog, exc)

    @mcp.tool(
        name="cancel_event",
        description=(
            "DESTRUCTIVE — REQUIRES EXPLICIT USER CONFIRMATION BEFORE CALLING. "
            "Sends a cancellation message from the organizer to all attendees and cancels "
            "the meeting, moving it to Deleted Items. Organizer-only — an attendee gets a "
            "400 error. NEVER call this tool autonomously or as part of an automated flow. "
            "You MUST stop, tell the user exactly which meeting will be cancelled and what "
            "message will be sent, and wait for their explicit written confirmation before "
            "proceeding."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=True, openWorldHint=True),
    )
    def cancel_event(
        id: str = Field(
            description="ID of the event to cancel. The organizer can cancel a single occurrence of a recurring meeting by passing that occurrence's event ID.",
        ),
        comment: str | None = Field(default=None, description="Message to send to attendees explaining the cancellation. Omit to send no comment."),
    ) -> CancelEventResult:
        tlog = ToolLogger(logger, "cancel_event")

        payload: dict[str, Any] = {}
        if comment is not None:
            payload["comment"] = comment

        try:
            data, status, retry_after = service.api_request(
                "POST", f"/me/events/{id}/cancel", body=payload, timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return CancelEventResult(success=True, statusCode=status, data=CancelEventData())
            return _upstream_err(CancelEventResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(CancelEventResult, tlog, exc)

    @mcp.tool(
        name="accept_event",
        description=(
            "Accepts an event invitation on the signed-in user's calendar. Applies to user "
            "calendars only, not group calendars. Returns no content on success (202 Accepted)."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=True),
    )
    def accept_event(
        id: str = Field(description="ID of the event to accept."),
        comment: str | None = Field(default=None, description="Message to send with the response. Omit to send no comment."),
        sendResponse: bool | None = Field(default=None, description="Whether to send a response to the organizer. Defaults to true if omitted."),
    ) -> AcceptEventResult:
        tlog = ToolLogger(logger, "accept_event")

        payload: dict[str, Any] = {}
        if comment is not None:
            payload["comment"] = comment
        if sendResponse is not None:
            payload["sendResponse"] = sendResponse

        try:
            data, status, retry_after = service.api_request(
                "POST", f"/me/events/{id}/accept", body=payload, timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return AcceptEventResult(success=True, statusCode=status, data=AcceptEventData())
            return _upstream_err(AcceptEventResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(AcceptEventResult, tlog, exc)

    @mcp.tool(
        name="tentatively_accept_event",
        description=(
            "Tentatively accepts an event invitation on the signed-in user's calendar, "
            "optionally proposing a new time. Applies to user calendars only, not group "
            "calendars. Returns no content on success (202 Accepted)."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=True),
    )
    def tentatively_accept_event(
        id: str = Field(description="ID of the event to tentatively accept."),
        comment: str | None = Field(default=None, description="Message to send with the response. Omit to send no comment."),
        sendResponse: bool | None = Field(
            default=None,
            description="Whether to send a response to the organizer. Defaults to true if omitted. Must be true if proposedNewTime is supplied.",
        ),
        proposedNewTime: dict[str, Any] | None = Field(
            default=None,
            description=(
                "A new time to propose, as {start: {dateTime, timeZone}, end: {dateTime, timeZone}}. "
                "Only valid for events that allow new time proposals. Omit to propose no new time."
            ),
        ),
    ) -> TentativelyAcceptEventResult:
        tlog = ToolLogger(logger, "tentatively_accept_event")

        payload: dict[str, Any] = {}
        if comment is not None:
            payload["comment"] = comment
        if sendResponse is not None:
            payload["sendResponse"] = sendResponse
        if proposedNewTime is not None:
            payload["proposedNewTime"] = proposedNewTime

        try:
            data, status, retry_after = service.api_request(
                "POST", f"/me/events/{id}/tentativelyAccept", body=payload, timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return TentativelyAcceptEventResult(success=True, statusCode=status, data=TentativelyAcceptEventData())
            return _upstream_err(TentativelyAcceptEventResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(TentativelyAcceptEventResult, tlog, exc)

    @mcp.tool(
        name="decline_event",
        description=(
            "Declines an event invitation on the signed-in user's calendar, optionally "
            "proposing a new time. Applies to user calendars only, not group calendars. "
            "Returns no content on success (202 Accepted)."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=True),
    )
    def decline_event(
        id: str = Field(description="ID of the event to decline."),
        comment: str | None = Field(default=None, description="Message to send with the response. Omit to send no comment."),
        sendResponse: bool | None = Field(
            default=None,
            description="Whether to send a response to the organizer. Defaults to true if omitted. Must be true if proposedNewTime is supplied.",
        ),
        proposedNewTime: dict[str, Any] | None = Field(
            default=None,
            description=(
                "A new time to propose, as {start: {dateTime, timeZone}, end: {dateTime, timeZone}}. "
                "Only valid for events that allow new time proposals. Omit to propose no new time."
            ),
        ),
    ) -> DeclineEventResult:
        tlog = ToolLogger(logger, "decline_event")

        payload: dict[str, Any] = {}
        if comment is not None:
            payload["comment"] = comment
        if sendResponse is not None:
            payload["sendResponse"] = sendResponse
        if proposedNewTime is not None:
            payload["proposedNewTime"] = proposedNewTime

        try:
            data, status, retry_after = service.api_request(
                "POST", f"/me/events/{id}/decline", body=payload, timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return DeclineEventResult(success=True, statusCode=status, data=DeclineEventData())
            return _upstream_err(DeclineEventResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(DeclineEventResult, tlog, exc)

    @mcp.tool(
        name="find_meeting_times",
        description=(
            "Finds candidate meeting times based on attendee free/busy availability, "
            "location, and time constraints, without creating an event. Once a time is "
            "picked, create the meeting with create_event. Returns a ranked list of "
            "suggestions, or an empty list with `emptySuggestionsReason` explaining why none "
            "could be found."
        ),
        annotations=ToolAnnotations(readOnlyHint=True, destructiveHint=False, openWorldHint=True),
    )
    def find_meeting_times(
        attendees: list[dict[str, Any]] | None = Field(
            default=None,
            description=(
                "Attendees to check availability for, each as {type: 'required'|'optional', "
                "emailAddress: {name, address}}. Omit (or pass an empty list) to look only for "
                "the organizer's free time."
            ),
        ),
        isOrganizerOptional: bool | None = Field(default=None, description="Whether the organizer's availability can be ignored. Omit to use the default (false)."),
        locationConstraint: dict[str, Any] | None = Field(
            default=None,
            description="Location requirements as {isRequired, suggestLocation, locations: [...]}. Omit to apply no location constraint.",
        ),
        maxCandidates: int | None = Field(default=None, description="Maximum number of suggestions to return. Omit to use the API default (10)."),
        meetingDuration: str | None = Field(default=None, description="Meeting duration as an ISO 8601 duration, e.g. 'PT1H'. Omit to use the API default."),
        minimumAttendeePercentage: float | None = Field(
            default=None, description="Minimum percentage of attendees that must be available for a suggestion to qualify. Omit to apply no minimum.",
        ),
        returnSuggestionReasons: bool | None = Field(default=None, description="Whether to include a human-readable suggestionReason on each result. Omit to use the API default."),
        timeConstraint: dict[str, Any] | None = Field(
            default=None,
            description="Time window(s) to search within as {activityDomain, timeSlots: [{start, end}]}. Omit to search using the API default window and activityDomain 'work'.",
        ),
        timezone: str | None = Field(
            default=None,
            description="Time zone name for the returned suggestion times, sent as the Prefer: outlook.timezone header. Omit to receive times in UTC.",
        ),
    ) -> FindMeetingTimesResult:
        tlog = ToolLogger(logger, "find_meeting_times")

        payload: dict[str, Any] = {}
        fields = {
            "attendees": attendees, "isOrganizerOptional": isOrganizerOptional,
            "locationConstraint": locationConstraint, "maxCandidates": maxCandidates,
            "meetingDuration": meetingDuration, "minimumAttendeePercentage": minimumAttendeePercentage,
            "returnSuggestionReasons": returnSuggestionReasons, "timeConstraint": timeConstraint,
        }
        payload.update({k: v for k, v in fields.items() if v is not None})

        try:
            data, status, retry_after = service.api_request(
                "POST", "/me/findMeetingTimes", body=payload, extra_headers=_prefer_header(timezone),
                timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return FindMeetingTimesResult(success=True, statusCode=status, data=FindMeetingTimesData(**data))
            return _upstream_err(FindMeetingTimesResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(FindMeetingTimesResult, tlog, exc)

    @mcp.tool(
        name="get_free_busy_schedule",
        description=(
            "Gets free/busy availability for a set of users, distribution lists, or "
            "resources over a specified time period. To get ranked meeting-time suggestions "
            "instead of raw availability, use find_meeting_times, which consults the same "
            "schedules."
        ),
        annotations=ToolAnnotations(readOnlyHint=True, destructiveHint=False, openWorldHint=True),
    )
    def get_free_busy_schedule(
        schedules: list[str] = Field(description="Email addresses of the users, distribution lists, or resources to check."),
        startTime: dict[str, Any] = Field(description="Start of the period as {dateTime, timeZone}."),
        endTime: dict[str, Any] = Field(description="End of the period as {dateTime, timeZone}."),
        availabilityViewInterval: int | None = Field(
            default=None,
            description="Duration in minutes of each slot in the returned availabilityView digit string, 5-1440. Omit to use the default (30).",
        ),
        timezone: str | None = Field(
            default=None,
            description="Time zone name for the returned schedule times, sent as the Prefer: outlook.timezone header. Omit to receive times in UTC.",
        ),
    ) -> GetFreeBusyScheduleResult:
        tlog = ToolLogger(logger, "get_free_busy_schedule")

        if not schedules:
            return _err(GetFreeBusyScheduleResult, tlog, "VALIDATION_ERROR",
                        "schedules must contain at least one address", 400)

        payload: dict[str, Any] = {"schedules": schedules, "startTime": startTime, "endTime": endTime}
        if availabilityViewInterval is not None:
            payload["availabilityViewInterval"] = availabilityViewInterval

        try:
            data, status, retry_after = service.api_request(
                "POST", "/me/calendar/getSchedule", body=payload, extra_headers=_prefer_header(timezone),
                timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return GetFreeBusyScheduleResult(success=True, statusCode=status, data=GetFreeBusyScheduleData(**data))
            return _upstream_err(GetFreeBusyScheduleResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(GetFreeBusyScheduleResult, tlog, exc)

    @mcp.tool(
        name="list_calendar_view",
        description=(
            "Gets events for a date range from the default calendar, expanding recurring "
            "series into individual occurrence instances — use this for \"what's actually on "
            "the calendar\" over a range, unlike list_events which returns series masters "
            "unexpanded. Omit `select` to also get `createdDateTime`/`lastModifiedDateTime` "
            "on the results, which can't be requested via select."
        ),
        annotations=ToolAnnotations(readOnlyHint=True, destructiveHint=False, openWorldHint=True),
    )
    def list_calendar_view(
        startDateTime: str = Field(
            description="Start of the time range, ISO 8601 (e.g. '2019-11-08T19:00:00-08:00'). Interpreted using the offset in the value itself; UTC if no offset given.",
        ),
        endDateTime: str = Field(description="End of the time range, ISO 8601, interpreted the same way as startDateTime."),
        select: str | None = Field(
            default=None,
            description="Comma-separated list of properties to return. Omit to return all event properties, including createdDateTime/lastModifiedDateTime.",
        ),
        timezone: str | None = Field(
            default=None,
            description="Time zone name for the returned start/end values, sent as the Prefer: outlook.timezone header. Does not affect how startDateTime/endDateTime are interpreted. Omit to receive times in UTC.",
        ),
        body_content_type: str | None = Field(
            default=None,
            description="Set to 'text' to receive the event body as plain text instead of HTML. Omit to receive HTML.",
        ),
    ) -> ListCalendarViewResult:
        tlog = ToolLogger(logger, "list_calendar_view")

        params: dict[str, Any] = {"startDateTime": startDateTime, "endDateTime": endDateTime}
        if select is not None:
            params["$select"] = select

        try:
            data, status, retry_after = service.api_request(
                "GET", "/me/calendar/calendarView", params=params,
                extra_headers=_prefer_header(timezone, body_content_type),
                timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return ListCalendarViewResult(success=True, statusCode=status, data=ListCalendarViewData(**data))
            return _upstream_err(ListCalendarViewResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(ListCalendarViewResult, tlog, exc)
