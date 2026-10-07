"""Users group: get_user."""

import logging

from fastmcp import FastMCP
from mcp.types import ToolAnnotations
from pydantic import Field

from .. import service
from ..config import CONNECT_TIMEOUT, READ_TIMEOUT
from ..logging_utils import ToolLogger
from ..schemas.users import UserData, UserResult
from ._helpers import _handle_request_exc, _upstream_err

logger = logging.getLogger("outlook-mcp.tools.users")


def register_users_tools(mcp: FastMCP) -> None:

    @mcp.tool(
        name="get_user",
        description=(
            "Gets the signed-in user's own profile, returning properties such as "
            "displayName, mail, jobTitle, userPrincipalName, and id from the default "
            "property set unless `select` narrows it. Only works with a delegated, "
            "signed-in-user token — it returns the caller's own profile, not an "
            "arbitrary user's."
        ),
        annotations=ToolAnnotations(readOnlyHint=True, destructiveHint=False, openWorldHint=True),
    )
    def get_user(
        select: str | None = Field(
            default=None,
            description=(
                "Comma-separated list of properties to return instead of the default set "
                "(businessPhones, displayName, givenName, id, jobTitle, mail, mobilePhone, "
                "officeLocation, preferredLanguage, surname, userPrincipalName). Sent as the "
                "API's `$select` query parameter, e.g. 'displayName,givenName,postalCode'. "
                "Omit to use the default set."
            ),
        ),
    ) -> UserResult:
        tlog = ToolLogger(logger, "get_user")

        try:
            data, status, retry_after = service.api_request(
                "GET", "/me",
                params={"$select": select} if select else None,
                timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return UserResult(success=True, statusCode=status, data=UserData(**data))
            return _upstream_err(UserResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(UserResult, tlog, exc)
