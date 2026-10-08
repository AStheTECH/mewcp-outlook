"""MewCP Outlook tool registration."""

from fastmcp import FastMCP

from .users_tools import register_users_tools
from .messages_tools import register_messages_tools
from .calendars_tools import register_calendars_tools
from .events_tools import register_events_tools


def register_tools(mcp: FastMCP) -> None:
    register_users_tools(mcp)
    register_messages_tools(mcp)
    register_calendars_tools(mcp)
    register_events_tools(mcp)
