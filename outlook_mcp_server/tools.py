import logging
import requests
from typing import List, Annotated
from pydantic import Field
from fastmcp import FastMCP
from .schema import ApiObjectResponse 
from .service import get_outlook_headers

logger = logging.getLogger("tasks-mcp-server")

class _ToolCollector:
    """MewCP's enterprise tool registration pattern."""
    def __init__(self):
        self.items = []

    def tool(self, *args, **kwargs):
        def decorator(func):
            self.items.append((args, kwargs, func))
            return func
        return decorator

mcp = _ToolCollector()

def register_tools(real_mcp: FastMCP) -> None:
    """Transfers tools from the collector to the real FastMCP server."""
    for args, kwargs, func in mcp.items:
        real_mcp.tool(*args, **kwargs)(func)
