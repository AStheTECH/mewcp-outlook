from pydantic import BaseModel, ConfigDict, Field

from ._base import ToolResult


class UserData(BaseModel):
    model_config = ConfigDict(extra="allow")

    id: str | None = None
    businessPhones: list[str] = []
    displayName: str | None = None
    givenName: str | None = None
    jobTitle: str | None = None
    mail: str | None = None
    mobilePhone: str | None = None
    officeLocation: str | None = None
    preferredLanguage: str | None = None
    surname: str | None = None
    userPrincipalName: str | None = None


class UserResult(ToolResult):
    data: UserData | None = None


class UserListItem(BaseModel):
    model_config = ConfigDict(extra="allow")

    id: str | None = None
    displayName: str | None = None
    givenName: str | None = None
    surname: str | None = None
    jobTitle: str | None = None
    mail: str | None = None
    businessPhones: list[str] = []
    mobilePhone: str | None = None
    officeLocation: str | None = None
    preferredLanguage: str | None = None
    userPrincipalName: str | None = None


class ListUsersData(BaseModel):
    model_config = ConfigDict(extra="allow")

    value: list[UserListItem] = Field(default_factory=list)


class ListUsersResult(ToolResult):
    data: ListUsersData | None = None


class CreateUserData(BaseModel):
    model_config = ConfigDict(extra="allow")

    id: str = Field(description="Unique identifier of the newly created user.")
    displayName: str | None = None
    givenName: str | None = None
    jobTitle: str | None = None
    mail: str | None = None
    mobilePhone: str | None = None
    officeLocation: str | None = None
    preferredLanguage: str | None = None
    surname: str | None = None
    userPrincipalName: str | None = None
    businessPhones: list[str] = []


class CreateUserResult(ToolResult):
    data: CreateUserData | None = None


class UpdateUserData(BaseModel):
    model_config = ConfigDict(extra="allow")

    before: UserData
    after: UserData


class UpdateUserResult(ToolResult):
    data: UpdateUserData | None = None


class DeleteUserData(BaseModel):
    """Graph returns 204 No Content for this endpoint."""

    model_config = ConfigDict(extra="allow")


class DeleteUserResult(ToolResult):
    data: DeleteUserData | None = None


class GetUsersDeltaData(BaseModel):
    model_config = ConfigDict(extra="allow", populate_by_name=True)

    value: list[UserListItem] = Field(default_factory=list)
    odata_next_link: str | None = Field(
        default=None,
        alias="@odata.nextLink",
        description="Present when more pages remain in this round of change tracking.",
    )
    odata_delta_link: str | None = Field(
        default=None,
        alias="@odata.deltaLink",
        description="Present when this round is complete; save it to start the next round.",
    )


class GetUsersDeltaResult(ToolResult):
    data: GetUsersDeltaData | None = None
