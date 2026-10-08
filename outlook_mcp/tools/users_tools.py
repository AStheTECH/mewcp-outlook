"""Users group: get_user, list_users, create_user, update_user, delete_user,
revoke_sign_in_sessions, get_users_delta, change_password."""

import logging
from typing import Any

from fastmcp import FastMCP
from mcp.types import ToolAnnotations
from pydantic import Field

from .. import service
from ..config import CONNECT_TIMEOUT, READ_TIMEOUT
from ..logging_utils import ToolLogger
from ..schemas.users import (
    ChangePasswordData,
    ChangePasswordResult,
    CreateUserData,
    CreateUserResult,
    DeleteUserData,
    DeleteUserResult,
    GetUsersDeltaData,
    GetUsersDeltaResult,
    ListUsersData,
    ListUsersResult,
    RevokeSignInSessionsData,
    RevokeSignInSessionsResult,
    UpdateUserData,
    UpdateUserResult,
    UserData,
    UserResult,
)
from ._helpers import _err, _handle_request_exc, _upstream_err

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

    @mcp.tool(
        name="list_users",
        description=(
            "Lists user objects in the directory, returning the default property set "
            "(businessPhones, displayName, givenName, id, jobTitle, mail, mobilePhone, "
            "officeLocation, preferredLanguage, surname, userPrincipalName) per user unless "
            "`select` requests others. Not supported for personal Microsoft accounts."
        ),
        annotations=ToolAnnotations(readOnlyHint=True, destructiveHint=False, openWorldHint=True),
    )
    def list_users(
        select: str | None = Field(
            default=None,
            description=(
                "Comma-separated list of properties to return instead of the default set, e.g. "
                "'displayName,givenName,postalCode'. Omit to use the default set."
            ),
        ),
        expand: str | None = Field(
            default=None,
            description="Related entities to expand inline. Omit to expand nothing.",
        ),
        filter: str | None = Field(
            default=None,
            description="OData filter expression to restrict the users returned. Omit to apply no filter.",
        ),
        orderby: str | None = Field(
            default=None,
            description="Property to sort the results by. Omit to use the API's default ordering.",
        ),
        search: str | None = Field(
            default=None,
            description="OData search expression to match against users. Omit to apply no search filtering.",
        ),
        count: bool | None = Field(
            default=None,
            description=(
                "Whether to include a count of matching users and send the ConsistencyLevel: eventual "
                "header, required together with `search` or some `filter` expressions. Omit for a plain request."
            ),
        ),
        top: int | None = Field(
            default=None,
            description=(
                "Maximum number of users to return per page. Omit to use the default (100, max 999; "
                "max 500 if select or filter is used on signInActivity)."
            ),
        ),
    ) -> ListUsersResult:
        tlog = ToolLogger(logger, "list_users")

        if top is not None and (top < 1 or top > 999):
            return _err(ListUsersResult, tlog, "VALIDATION_ERROR", "top must be between 1 and 999", 400)

        params: dict[str, Any] = {}
        if select is not None:
            params["$select"] = select
        if expand is not None:
            params["$expand"] = expand
        if filter is not None:
            params["$filter"] = filter
        if orderby is not None:
            params["$orderby"] = orderby
        if search is not None:
            params["$search"] = search
        if count is not None:
            params["$count"] = "true" if count else "false"
        if top is not None:
            params["$top"] = top

        headers = {"ConsistencyLevel": "eventual"} if count else None

        try:
            data, status, retry_after = service.api_request(
                "GET", "/users", params=params or None, extra_headers=headers,
                timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return ListUsersResult(success=True, statusCode=status, data=ListUsersData(**data))
            return _upstream_err(ListUsersResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(ListUsersResult, tlog, exc)

    @mcp.tool(
        name="create_user",
        description=(
            "Creates a new user object in the directory and returns the created user's default "
            "property set, including the `id` needed by get_user, update_user, and delete_user. "
            "Not supported for personal Microsoft accounts."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=True),
    )
    def create_user(
        accountEnabled: bool = Field(description="Whether the account is enabled."),
        displayName: str = Field(description="The name displayed in the address book."),
        mailNickname: str = Field(description="The mail alias for the user."),
        userPrincipalName: str = Field(
            description="The user's principal name (UPN), usually their sign-in email.",
        ),
        passwordProfile: dict[str, Any] = Field(
            description="Initial password settings: {forceChangePasswordNextSignIn: <bool>, password: <string>}.",
        ),
        onPremisesImmutableId: str | None = Field(
            default=None,
            description=(
                "Associates an on-prem AD account with the Entra user object. Required only when "
                "userPrincipalName uses a federated domain; omit otherwise."
            ),
        ),
    ) -> CreateUserResult:
        tlog = ToolLogger(logger, "create_user")

        payload: dict[str, Any] = {
            "accountEnabled": accountEnabled,
            "displayName": displayName,
            "mailNickname": mailNickname,
            "userPrincipalName": userPrincipalName,
            "passwordProfile": passwordProfile,
        }
        if onPremisesImmutableId is not None:
            payload["onPremisesImmutableId"] = onPremisesImmutableId

        try:
            data, status, retry_after = service.api_request(
                "POST", "/users", body=payload,
                timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return CreateUserResult(success=True, statusCode=status, data=CreateUserData(**data))
            return _upstream_err(CreateUserResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(CreateUserResult, tlog, exc)

    @mcp.tool(
        name="update_user",
        description=(
            "Updates properties on the signed-in user's own profile, or on another user's profile "
            "when `user_id` is given (requires higher-privileged admin permissions). Supply only the "
            "properties to change — others keep their current value. NOTE: this overwrites the "
            "current field values and the prior state isn't stored by the API after the call; the "
            "response includes both the before and after state so you have a full record of what changed."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=True),
    )
    def update_user(
        user_id: str | None = Field(
            default=None,
            description=(
                "id (GUID) or userPrincipalName of another user to update instead of the signed-in "
                "user. Requires higher-privileged permissions than updating /me. Omit to update your "
                "own profile."
            ),
        ),
        aboutMe: str | None = Field(default=None, description="Freeform self-description. Must be updated alone in its own request; not updatable with application-only permissions. Omit to leave unchanged."),
        accountEnabled: bool | None = Field(default=None, description="true if the account is enabled. Requires User.EnableDisableAccount.All + User.Read.All. Omit to leave unchanged."),
        ageGroup: str | None = Field(default=None, description="null, Minor, NotAdult, or Adult. Omit to leave unchanged."),
        birthday: str | None = Field(default=None, description="ISO 8601 UTC. Must be updated alone in its own request. Omit to leave unchanged."),
        businessPhones: list[str] | None = Field(default=None, description="Only one number can actually be set, despite being a collection. Omit to leave unchanged."),
        city: str | None = Field(default=None, description="The user's city. Omit to leave unchanged."),
        companyName: str | None = Field(default=None, description="Max 64 characters. Omit to leave unchanged."),
        consentProvidedForMinor: str | None = Field(default=None, description="null, Granted, Denied, or NotRequired. Omit to leave unchanged."),
        country: str | None = Field(default=None, description="e.g. US, UK. Omit to leave unchanged."),
        customSecurityAttributes: dict[str, Any] | None = Field(default=None, description="Requires the Attribute Assignment Administrator role + CustomSecAttributeAssignment.ReadWrite.All. Omit to leave unchanged."),
        department: str | None = Field(default=None, description="Department name. Omit to leave unchanged."),
        displayName: str | None = Field(default=None, description="The name displayed in the address book; can't be cleared on update. Omit to leave unchanged."),
        employeeHireDate: str | None = Field(default=None, description="ISO 8601 UTC. Omit to leave unchanged."),
        employeeId: str | None = Field(default=None, description="Max 16 characters. Omit to leave unchanged."),
        employeeLeaveDateTime: str | None = Field(default=None, description="Requires the Global Administrator role + User-LifeCycleInfo.ReadWrite.All and User.Read.All. Omit to leave unchanged."),
        employeeOrgData: dict[str, Any] | None = Field(default=None, description="Include both division and costCenter — any you omit are set to null. Omit to leave unchanged."),
        employeeType: str | None = Field(default=None, description="e.g. Employee, Contractor, Consultant, Vendor. Omit to leave unchanged."),
        givenName: str | None = Field(default=None, description="First name. Omit to leave unchanged."),
        identities: list[dict[str, Any]] | None = Field(default=None, description="Replaces the entire collection; must include the userPrincipalName signInType identity. Requires User.ManageIdentities.All. Omit to leave unchanged."),
        interests: list[str] | None = Field(default=None, description="Must be updated alone in its own request. Omit to leave unchanged."),
        jobTitle: str | None = Field(default=None, description="Job title. Omit to leave unchanged."),
        mail: str | None = Field(default=None, description="SMTP address; also updates proxyAddresses. Can't be set to null. Omit to leave unchanged."),
        mailNickname: str | None = Field(default=None, description="Mail alias. Omit to leave unchanged."),
        mobilePhone: str | None = Field(default=None, description="Primary cell number. Requires User-Phone.ReadWrite.All. Omit to leave unchanged."),
        mySite: str | None = Field(default=None, description="URL of personal site. Must be updated alone in its own request. Omit to leave unchanged."),
        officeLocation: str | None = Field(default=None, description="Office location. Omit to leave unchanged."),
        onPremisesExtensionAttributes: dict[str, Any] | None = Field(default=None, description="extensionAttributes 1-15; read-only when synced from on-premises. Omit to leave unchanged."),
        onPremisesImmutableId: str | None = Field(default=None, description="Associates an on-prem AD account with the Entra user object; $ and _ aren't allowed. Omit to leave unchanged."),
        otherMails: list[str] | None = Field(default=None, description="Pass the full set you want — existing values are overwritten, not merged. Max 250 values, 250 chars each. Omit to leave unchanged."),
        passwordPolicies: str | None = Field(default=None, description="e.g. DisableStrongPassword, DisablePasswordExpiration (can combine). Omit to leave unchanged."),
        passwordProfile: dict[str, Any] | None = Field(default=None, description="Resets the user's password; set forceChangePasswordNextSignIn to true as best practice. Not usable for federated users. Omit to leave unchanged."),
        pastProjects: list[str] | None = Field(default=None, description="Must be updated alone in its own request. Omit to leave unchanged."),
        postalCode: str | None = Field(default=None, description="Postal/ZIP code. Omit to leave unchanged."),
        preferredLanguage: str | None = Field(default=None, description="ISO 639-1 code, e.g. en-US. Omit to leave unchanged."),
        responsibilities: list[str] | None = Field(default=None, description="Must be updated alone in its own request. Omit to leave unchanged."),
        schools: list[str] | None = Field(default=None, description="Must be updated alone in its own request. Omit to leave unchanged."),
        skills: list[str] | None = Field(default=None, description="Must be updated alone in its own request. Omit to leave unchanged."),
        state: str | None = Field(default=None, description="State/province. Omit to leave unchanged."),
        streetAddress: str | None = Field(default=None, description="Street address. Omit to leave unchanged."),
        surname: str | None = Field(default=None, description="Last name. Omit to leave unchanged."),
        usageLocation: str | None = Field(default=None, description="Two-letter ISO 3166 country code; required before license assignment. Omit to leave unchanged."),
        userPrincipalName: str | None = Field(default=None, description="UPN / sign-in name. Only A-Z a-z 0-9 ' . - _ ! # ^ ~ allowed, no accent characters. Omit to leave unchanged."),
        userType: str | None = Field(default=None, description="e.g. Member, Guest. Omit to leave unchanged."),
    ) -> UpdateUserResult:
        tlog = ToolLogger(logger, "update_user")

        target = f"/users/{user_id}" if user_id else "/me"

        fields = {
            "aboutMe": aboutMe, "accountEnabled": accountEnabled, "ageGroup": ageGroup,
            "birthday": birthday, "businessPhones": businessPhones, "city": city,
            "companyName": companyName, "consentProvidedForMinor": consentProvidedForMinor,
            "country": country, "customSecurityAttributes": customSecurityAttributes,
            "department": department, "displayName": displayName, "employeeHireDate": employeeHireDate,
            "employeeId": employeeId, "employeeLeaveDateTime": employeeLeaveDateTime,
            "employeeOrgData": employeeOrgData, "employeeType": employeeType, "givenName": givenName,
            "identities": identities, "interests": interests, "jobTitle": jobTitle, "mail": mail,
            "mailNickname": mailNickname, "mobilePhone": mobilePhone, "mySite": mySite,
            "officeLocation": officeLocation, "onPremisesExtensionAttributes": onPremisesExtensionAttributes,
            "onPremisesImmutableId": onPremisesImmutableId, "otherMails": otherMails,
            "passwordPolicies": passwordPolicies, "passwordProfile": passwordProfile,
            "pastProjects": pastProjects, "postalCode": postalCode, "preferredLanguage": preferredLanguage,
            "responsibilities": responsibilities, "schools": schools, "skills": skills, "state": state,
            "streetAddress": streetAddress, "surname": surname, "usageLocation": usageLocation,
            "userPrincipalName": userPrincipalName, "userType": userType,
        }
        payload = {k: v for k, v in fields.items() if v is not None}
        if not payload:
            return _err(UpdateUserResult, tlog, "VALIDATION_ERROR",
                        "at least one property to change must be provided", 400)

        try:
            before_data, before_status, before_retry = service.api_request(
                "GET", target, timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if not (200 <= before_status < 300):
                return _upstream_err(UpdateUserResult, tlog, before_status, before_data, before_retry)
            before = UserData(**before_data)

            patch_data, patch_status, patch_retry = service.api_request(
                "PATCH", target, body=payload, timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if not (200 <= patch_status < 300):
                return _upstream_err(UpdateUserResult, tlog, patch_status, patch_data, patch_retry)

            after_data, after_status, after_retry = service.api_request(
                "GET", target, timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if not (200 <= after_status < 300):
                return _upstream_err(UpdateUserResult, tlog, after_status, after_data, after_retry)
            after = UserData(**after_data)

            tlog.success()
            return UpdateUserResult(success=True, statusCode=200, data=UpdateUserData(before=before, after=after))
        except Exception as exc:
            return _handle_request_exc(UpdateUserResult, tlog, exc)

    @mcp.tool(
        name="delete_user",
        description=(
            "DESTRUCTIVE — REQUIRES EXPLICIT USER CONFIRMATION BEFORE CALLING. "
            "Permanently deletes a user object by id or userPrincipalName. The user's mailbox, "
            "license assignments, and group memberships move to a temporary container recoverable "
            "for 30 days; after that they are permanently deleted and the resources freed. "
            "NEVER call this tool autonomously or as part of an automated flow. You MUST stop, "
            "tell the user exactly which user will be deleted and that it becomes permanent after "
            "30 days, and wait for their explicit written confirmation before proceeding."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=True, openWorldHint=True),
    )
    def delete_user(
        id: str = Field(description="The id (GUID) or userPrincipalName of the user to delete."),
    ) -> DeleteUserResult:
        tlog = ToolLogger(logger, "delete_user")

        try:
            data, status, retry_after = service.api_request(
                "DELETE", f"/users/{id}", timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return DeleteUserResult(success=True, statusCode=status, data=DeleteUserData())
            return _upstream_err(DeleteUserResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(DeleteUserResult, tlog, exc)

    @mcp.tool(
        name="revoke_sign_in_sessions",
        description=(
            "Revokes all of the signed-in user's refresh and session tokens issued to applications, "
            "or another user's when `user_id` is given (requires higher-privileged admin "
            "permissions), forcing re-authentication everywhere. There can be a delay of a few "
            "minutes before tokens are actually revoked, and this call can't be undone. Returns "
            "whether the revocation succeeded."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=True),
    )
    def revoke_sign_in_sessions(
        user_id: str | None = Field(
            default=None,
            description=(
                "id (GUID) or userPrincipalName of another user to revoke sessions for instead of "
                "the signed-in user. Requires User.RevokeSessions.All or higher. Omit to revoke "
                "your own sessions."
            ),
        ),
    ) -> RevokeSignInSessionsResult:
        tlog = ToolLogger(logger, "revoke_sign_in_sessions")

        target = f"/users/{user_id}/revokeSignInSessions" if user_id else "/me/revokeSignInSessions"

        try:
            data, status, retry_after = service.api_request(
                "POST", target, body=None, timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return RevokeSignInSessionsResult(
                    success=True, statusCode=status, data=RevokeSignInSessionsData(**data))
            return _upstream_err(RevokeSignInSessionsResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(RevokeSignInSessionsResult, tlog, exc)

    @mcp.tool(
        name="get_users_delta",
        description=(
            "Gets incremental changes to user objects since the last delta query, returning changed "
            "users plus either `odata_next_link` (more pages in this round) or `odata_delta_link` "
            "(round complete — save it to start the next round)."
        ),
        annotations=ToolAnnotations(readOnlyHint=True, destructiveHint=False, openWorldHint=True),
    )
    def get_users_delta(
        deltatoken: str | None = Field(
            default=None,
            description=(
                "The full @odata.deltaLink URL from a previous round, to start the next round of "
                "change tracking. Omit to start a new round from scratch."
            ),
        ),
        skiptoken: str | None = Field(
            default=None,
            description=(
                "The full @odata.nextLink URL from the previous call, to continue paging through the "
                "current round. Omit if not continuing a page."
            ),
        ),
        select: str | None = Field(
            default=None,
            description=(
                "Comma-separated properties to return; id is always returned. Only takes effect on "
                "the initial call of a tracking round. Omit to use the default property set."
            ),
        ),
        filter: str | None = Field(
            default=None,
            description="Only \"id eq '{value}'\" is supported (combine with or for up to 50 ids). Omit to apply no filter.",
        ),
        return_minimal: bool = Field(
            default=False,
            description=(
                "When true, sends the Prefer: return=minimal header so a request using "
                "@odata.deltaLink returns only properties that changed since that link was issued, "
                "instead of the full default property set. Defaults to false."
            ),
        ),
    ) -> GetUsersDeltaResult:
        tlog = ToolLogger(logger, "get_users_delta")

        params: dict[str, Any] = {}
        if deltatoken is not None:
            params["$deltatoken"] = deltatoken
        if skiptoken is not None:
            params["$skiptoken"] = skiptoken
        if select is not None:
            params["$select"] = select
        if filter is not None:
            params["$filter"] = filter

        headers = {"Prefer": "return=minimal"} if return_minimal else None

        try:
            data, status, retry_after = service.api_request(
                "GET", "/users/delta", params=params or None, extra_headers=headers,
                timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return GetUsersDeltaResult(success=True, statusCode=status, data=GetUsersDeltaData(**data))
            return _upstream_err(GetUsersDeltaResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(GetUsersDeltaResult, tlog, exc)

    @mcp.tool(
        name="change_password",
        description=(
            "Updates the signed-in user's own password. This action can't be undone and the "
            "previous password isn't recoverable from the API after the call. Returns no content "
            "on success."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=True),
    )
    def change_password(
        currentPassword: str = Field(description="The user's current password."),
        newPassword: str = Field(description="The new password to set."),
    ) -> ChangePasswordResult:
        tlog = ToolLogger(logger, "change_password")

        payload = {"currentPassword": currentPassword, "newPassword": newPassword}

        try:
            data, status, retry_after = service.api_request(
                "POST", "/me/changePassword", body=payload, timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return ChangePasswordResult(success=True, statusCode=status, data=ChangePasswordData())
            return _upstream_err(ChangePasswordResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(ChangePasswordResult, tlog, exc)
