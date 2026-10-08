"""Calendars group: list_calendars, get_calendar."""

import logging
from typing import Any

from fastmcp import FastMCP
from mcp.types import ToolAnnotations
from pydantic import Field

from .. import service
from ..config import CONNECT_TIMEOUT, READ_TIMEOUT
from ..logging_utils import ToolLogger
from ..schemas.calendars import (
    GetCalendarData,
    GetCalendarResult,
    ListCalendarsData,
    ListCalendarsResult,
)
from ._helpers import _handle_request_exc, _upstream_err

logger = logging.getLogger("outlook-mcp.tools.calendars")


def register_calendars_tools(mcp: FastMCP) -> None:

    @mcp.tool(
        name="list_calendars",
        description=(
            "Lists all the signed-in user's calendars, or the calendars in a specific calendar "
            "group. Returns the full set in one response. The `id` of each calendar is used by "
            "get_calendar and by any events/ tool that targets a specific calendar."
        ),
        annotations=ToolAnnotations(readOnlyHint=True, destructiveHint=False, openWorldHint=True),
    )
    def list_calendars(
        select: str | None = Field(
            default=None,
            description="Comma-separated list of properties to return, e.g. 'name,color'. Omit to return the default calendar properties.",
        ),
        calendar_group_id: str | None = Field(
            default=None,
            description="ID of a calendar group to list calendars from instead of all of the user's calendars. Omit to list all calendars.",
        ),
    ) -> ListCalendarsResult:
        tlog = ToolLogger(logger, "list_calendars")

        params: dict[str, Any] = {}
        if select is not None:
            params["$select"] = select

        endpoint = (
            f"/me/calendarGroups/{calendar_group_id}/calendars"
            if calendar_group_id else "/me/calendars"
        )

        try:
            data, status, retry_after = service.api_request(
                "GET", endpoint, params=params or None,
                timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return ListCalendarsResult(success=True, statusCode=status, data=ListCalendarsData(**data))
            return _upstream_err(ListCalendarsResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(ListCalendarsResult, tlog, exc)

    @mcp.tool(
        name="get_calendar",
        description=(
            "Gets the properties of a calendar — the signed-in user's default calendar if `id` "
            "is omitted, or a specific one of their other calendars otherwise."
        ),
        annotations=ToolAnnotations(readOnlyHint=True, destructiveHint=False, openWorldHint=True),
    )
    def get_calendar(
        id: str | None = Field(
            default=None,
            description="ID of a specific calendar to get, from list_calendars. Omit to get the signed-in user's default calendar.",
        ),
        select: str | None = Field(
            default=None,
            description="Comma-separated list of properties to return, e.g. 'name,color'. Omit to return the default calendar properties.",
        ),
    ) -> GetCalendarResult:
        tlog = ToolLogger(logger, "get_calendar")

        params: dict[str, Any] = {}
        if select is not None:
            params["$select"] = select

        endpoint = f"/me/calendars/{id}" if id else "/me/calendar"

        try:
            data, status, retry_after = service.api_request(
                "GET", endpoint, params=params or None,
                timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return GetCalendarResult(success=True, statusCode=status, data=GetCalendarData(**data))
            return _upstream_err(GetCalendarResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(GetCalendarResult, tlog, exc)
