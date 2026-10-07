"""Configuration for MewCP Outlook MCP Server."""

import logging
import os

SERVER_VERSION = "v1.0.0"
BREAKING_CHANGES: list[dict] = []

SCOPES = [
    "https://graph.microsoft.com/User.Read",       # users.get_user
    "https://graph.microsoft.com/Mail.ReadWrite",   # messages: list/get/create draft/forward/reply
    "https://graph.microsoft.com/Mail.Send",        # messages: send_mail, send_draft_message, forward/reply
]

OUTLOOK_API_BASE = "https://graph.microsoft.com/v1.0"

CONNECT_TIMEOUT = 5    # TCP connection — fixed across all servers
READ_TIMEOUT = 30      # Graph mail endpoints are simple CRUD/send calls, well within a 30s SLA


def configure_logging() -> None:
    log_level = os.environ.get("LOG_LEVEL", "INFO").upper()
    try:
        from pythonjsonlogger import jsonlogger
        handler = logging.StreamHandler()
        handler.setFormatter(
            jsonlogger.JsonFormatter(fmt="%(asctime)s %(name)s %(levelname)s %(message)s")
        )
    except ImportError:
        handler = logging.StreamHandler()
    root = logging.getLogger()
    root.handlers.clear()
    root.addHandler(handler)
    root.setLevel(log_level)
