"""MewCP Outlook tool registration."""

from fastmcp import FastMCP

from .users_tools import register_users_tools
from .messages_tools import register_messages_tools


def register_tools(mcp: FastMCP) -> None:
    register_users_tools(mcp)
    register_messages_tools(mcp)
