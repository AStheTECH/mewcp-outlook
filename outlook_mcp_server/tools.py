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

@mcp.tool(name="send_email",description="Send a plain text email to a specific recipient")
def send_email(
    subject: Annotated[str,Field(description="The subject line of the email")],
    body: Annotated[str,Field(description="The main text content of the email")],
    to_email: Annotated[str,Field(description="The recipient's email address")]
) -> ApiObjectResponse:
    """Send a new email"""
    try:
        headers=get_outlook_headers()
        url = "https://graph.microsoft.com/v1.0/me/sendMail"

        payload = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "Text",
                    "content": body
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": to_email
                        }
                    }
                ]
            },
            "saveToSentItems": "true"
        }
        response = requests.post(url,headers=headers,json=payload)
        response.raise_for_status()
        return {"message":"Email sent successfully"}
    except requests.exceptions.RequestException as e:
        logger.error(f"Error whilst sending email {e}")
        return {"error":str(e)}

@mcp.tool(name="create_draft", description="Create a new email draft and add it to the drafts folder without sending it.")
def create_draft(
    subject: Annotated[str,Field(description="The subject line of the email")],
    body: Annotated[str,Field(description="The text content of the email")],
    to_email: Annotated[str,Field(description="The recipient's email address")]
) -> ApiObjectResponse:
    """Create a message draft"""
    try:
        headers=get_outlook_headers()
        url = "https://graph.microsoft.com/v1.0/me/messages"

        payload = {
            "subject":subject,
            "body":{
                "contentType":"Text",
                "content":body
            },
            "toRecipients": [
                {"emailAddress":{"address":to_email}}
            ]
        }

        response = requests.post(url, headers=headers, json=payload)
        response.raise_for_status()
        return {"message":"Draft created successfully", "draft":response.json()}
    except requests.exceptions.RequestException as e:
        logger.error(f"Error in creating draft: {e}")
        return {"error":str(e)}