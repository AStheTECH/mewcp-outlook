import logging

from fastmcp_credentials import get_credentials

logger = logging.getLogger("outlook-mcp-server")

def get_outlook_headers():
    """Gets headers pre-configured with the bearer token from the gateway."""
    cred = get_credentials()
    if not cred.access_token:
        raise ValueError("No OAuth access token available in credentials")
        
    return {
        "Authorization": f"Bearer {cred.access_token}",
        "Content-Type": "application/json"
    }