from typing import Any

from pydantic import BaseModel, ConfigDict, Field

from ._base import ToolResult


class GetEventData(BaseModel):
    model_config = ConfigDict(extra="allow")

    id: str = Field(
        description=(
            "Unique identifier of the event, consumed by update_event, delete_event, "
            "cancel_event, accept_event, tentatively_accept_event, and decline_event."
        ),
    )
    subject: str | None = None
    bodyPreview: str | None = None
    body: dict[str, Any] | None = None
    hideAttendees: bool | None = None
    start: dict[str, Any] | None = None
    end: dict[str, Any] | None = None
    location: dict[str, Any] | None = None
    locations: list[dict[str, Any]] | None = None
    attendees: list[dict[str, Any]] | None = None
    organizer: dict[str, Any] | None = None
    createdDateTime: str | None = None
    lastModifiedDateTime: str | None = None
    changeKey: str | None = None
    categories: list[str] | None = None
    originalStartTimeZone: str | None = None
    originalEndTimeZone: str | None = None
    iCalUId: str | None = None
    reminderMinutesBeforeStart: int | None = None
    isReminderOn: bool | None = None
    hasAttachments: bool | None = None
    importance: str | None = None
    sensitivity: str | None = None
    isAllDay: bool | None = None
    isCancelled: bool | None = None
    isDraft: bool | None = None
    isOrganizer: bool | None = None
    responseRequested: bool | None = None
    seriesMasterId: str | None = None
    transactionId: str | None = None
    showAs: str | None = None
    type: str | None = None
    webLink: str | None = None
    onlineMeetingUrl: str | None = None
    isOnlineMeeting: bool | None = None
    onlineMeetingProvider: str | None = None
    onlineMeeting: dict[str, Any] | None = None
    allowNewTimeProposals: bool | None = None
    responseStatus: dict[str, Any] | None = None
    recurrence: dict[str, Any] | None = None


class GetEventResult(ToolResult):
    data: GetEventData | None = None


class ListEventsData(BaseModel):
    model_config = ConfigDict(extra="allow", populate_by_name=True)

    value: list[GetEventData] = Field(default_factory=list)
    odata_next_link: str | None = Field(
        default=None,
        alias="@odata.nextLink",
        description="URL to fetch the next page when present; pass it through as-is instead of constructing `skip` manually.",
    )


class ListEventsResult(ToolResult):
    data: ListEventsData | None = None


class CreateEventData(GetEventData):
    pass


class CreateEventResult(ToolResult):
    data: CreateEventData | None = None


class UpdateEventData(BaseModel):
    model_config = ConfigDict(extra="allow")

    before: GetEventData
    after: GetEventData


class UpdateEventResult(ToolResult):
    data: UpdateEventData | None = None


class DeleteEventData(BaseModel):
    """Graph returns 204 No Content for this endpoint."""

    model_config = ConfigDict(extra="allow")


class DeleteEventResult(ToolResult):
    data: DeleteEventData | None = None


class CancelEventData(BaseModel):
    """Graph returns 202 Accepted with no response body for this endpoint."""

    model_config = ConfigDict(extra="allow")


class CancelEventResult(ToolResult):
    data: CancelEventData | None = None


class AcceptEventData(BaseModel):
    """Graph returns 202 Accepted with no response body for this endpoint."""

    model_config = ConfigDict(extra="allow")


class AcceptEventResult(ToolResult):
    data: AcceptEventData | None = None


class TentativelyAcceptEventData(BaseModel):
    """Graph returns 202 Accepted with no response body for this endpoint."""

    model_config = ConfigDict(extra="allow")


class TentativelyAcceptEventResult(ToolResult):
    data: TentativelyAcceptEventData | None = None


class DeclineEventData(BaseModel):
    """Graph returns 202 Accepted with no response body for this endpoint."""

    model_config = ConfigDict(extra="allow")


class DeclineEventResult(ToolResult):
    data: DeclineEventData | None = None


class MeetingTimeSuggestion(BaseModel):
    model_config = ConfigDict(extra="allow")

    confidence: float | None = None
    order: int | None = None
    organizerAvailability: str | None = None
    suggestionReason: str | None = None
    attendeeAvailability: list[dict[str, Any]] | None = None
    locations: list[dict[str, Any]] | None = None
    meetingTimeSlot: dict[str, Any] | None = None


class FindMeetingTimesData(BaseModel):
    model_config = ConfigDict(extra="allow")

    emptySuggestionsReason: str | None = None
    meetingTimeSuggestions: list[MeetingTimeSuggestion] = Field(default_factory=list)


class FindMeetingTimesResult(ToolResult):
    data: FindMeetingTimesData | None = None


class ScheduleInformation(BaseModel):
    model_config = ConfigDict(extra="allow")

    scheduleId: str | None = None
    availabilityView: str | None = Field(
        default=None,
        description="Digit string, one digit per interval: 0 free/workingElsewhere, 1 tentative, 2 busy, 3 out-of-office.",
    )
    scheduleItems: list[dict[str, Any]] | None = None
    workingHours: dict[str, Any] | None = None


class GetFreeBusyScheduleData(BaseModel):
    model_config = ConfigDict(extra="allow")

    value: list[ScheduleInformation] = Field(default_factory=list)


class GetFreeBusyScheduleResult(ToolResult):
    data: GetFreeBusyScheduleData | None = None


class ListCalendarViewData(BaseModel):
    model_config = ConfigDict(extra="allow", populate_by_name=True)

    value: list[GetEventData] = Field(default_factory=list)
    odata_next_link: str | None = Field(
        default=None,
        alias="@odata.nextLink",
        description="URL to fetch the next page when present, for date ranges large enough to be paged.",
    )


class ListCalendarViewResult(ToolResult):
    data: ListCalendarViewData | None = None
