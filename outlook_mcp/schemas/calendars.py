from typing import Any

from pydantic import BaseModel, ConfigDict, Field

from ._base import ToolResult


class CalendarOwner(BaseModel):
    model_config = ConfigDict(extra="allow")

    name: str | None = None
    address: str | None = None


class CalendarListItem(BaseModel):
    model_config = ConfigDict(extra="allow")

    id: str = Field(
        description=(
            "Unique identifier of the calendar, consumed by get_calendar and every events/ "
            "endpoint that targets a specific calendar."
        ),
    )
    name: str | None = None
    color: str | None = None
    hexColor: str | None = None
    changeKey: str | None = None
    canShare: bool | None = None
    canViewPrivateItems: bool | None = None
    canEdit: bool | None = None
    allowedOnlineMeetingProviders: list[str] | None = None
    defaultOnlineMeetingProvider: str | None = None
    isTallyingResponses: bool | None = None
    isRemovable: bool | None = None
    owner: CalendarOwner | None = None


class ListCalendarsData(BaseModel):
    model_config = ConfigDict(extra="allow")

    value: list[CalendarListItem] = Field(default_factory=list)


class ListCalendarsResult(ToolResult):
    data: ListCalendarsData | None = None


class GetCalendarData(BaseModel):
    model_config = ConfigDict(extra="allow")

    id: str = Field(description="Unique identifier of the calendar.")
    name: str | None = None
    color: str | None = None
    isDefaultCalendar: bool | None = None
    hexColor: str | None = None
    changeKey: str | None = None
    canShare: bool | None = None
    canViewPrivateItems: bool | None = None
    canEdit: bool | None = None
    allowedOnlineMeetingProviders: list[str] | None = None
    defaultOnlineMeetingProvider: str | None = None
    isTallyingResponses: bool | None = None
    isRemovable: bool | None = None
    owner: CalendarOwner | None = None


class GetCalendarResult(ToolResult):
    data: GetCalendarData | None = None
