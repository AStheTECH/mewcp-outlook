from typing_extensions import Any, TypedDict

class ToolError(TypedDict):
    error: str

OutlookToolResponse = dict[str, Any] | ToolError
ApiObjectResponse = dict[str, Any] | ToolError