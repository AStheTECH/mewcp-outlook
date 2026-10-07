from pydantic import BaseModel, ConfigDict

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
