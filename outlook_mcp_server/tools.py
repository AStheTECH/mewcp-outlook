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

@mcp.tool(name="list_emails", description="List recent emails from the signed in user's mailbox")
def list_emails(
    max_results: Annotated[int, Field(default=10, ge=1, le=50, description="Number of emails to return.")]
) -> ApiObjectResponse:
    """Retrieve recent emails"""
    logger.info("Executing list_emails")
    try:
        headers = get_outlook_headers()
        url = f"https://graph.microsoft.com/v1.0/me/messages?$top={max_results}&$select=subject,from,receivedDateTime,bodyPreview"
        response = requests.get(url, headers=headers)
        response.raise_for_status()

        data=response.json()
        return {"message":"Success","emails":data.get('value',[])}
    except requests.exceptions.RequestException as e:
        logger.error(f"Error in listing emails: {e}")
        return {"error":str(e)}