"""Upstream API client for MewCP Outlook MCP Server."""

import logging
from typing import Any

import requests
from fastmcp_credentials import get_credentials

from .config import OUTLOOK_API_BASE, CONNECT_TIMEOUT, READ_TIMEOUT

logger = logging.getLogger("outlook-mcp.service")


def _get_credential() -> str:
    cred = get_credentials()
    if not cred.access_token:
        raise ValueError("No OAuth access token available in credentials")
    return cred.access_token


def _auth_headers() -> dict[str, str]:
    return {
        "Authorization": f"Bearer {_get_credential()}",
        "Content-Type": "application/json",
    }


def _parse_retry_after(header: str | None) -> int | None:
    if not header:
        return None
    try:
        return int(header)
    except ValueError:
        return None


def api_request(
    method: str,
    endpoint: str,
    body: dict[str, Any] | None = None,
    params: dict[str, Any] | None = None,
    extra_headers: dict[str, str] | None = None,
    timeout: tuple[int, int] | None = None,
) -> tuple[dict[str, Any], int, int | None]:
    if timeout is None:
        timeout = (CONNECT_TIMEOUT, READ_TIMEOUT)
    url = f"{OUTLOOK_API_BASE}{endpoint}"
    headers = _auth_headers()
    if extra_headers:
        headers.update(extra_headers)
    resp = requests.request(method=method, url=url, headers=headers,
                            json=body, params=params, timeout=timeout)
    try:
        data = resp.json()
    except Exception:
        data = {"error": resp.text or "Empty response body"}
    return data, resp.status_code, _parse_retry_after(resp.headers.get("Retry-After"))
