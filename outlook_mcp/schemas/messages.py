from typing import Any

from pydantic import BaseModel, ConfigDict, Field

from ._base import ToolResult


class CreateDraftMessageData(BaseModel):
    model_config = ConfigDict(extra="allow")

    id: str = Field(description="Unique identifier of the new draft, consumed by send_draft_message and get_message.")
    isDraft: bool | None = None
    parentFolderId: str | None = None
    webLink: str | None = None
    subject: str | None = None
    body: dict[str, Any] | None = None
    bodyPreview: str | None = None
    importance: str | None = Field(default=None, description="Low, Normal, or High.")
    toRecipients: list[dict[str, Any]] | None = None
    ccRecipients: list[dict[str, Any]] | None = None
    bccRecipients: list[dict[str, Any]] | None = None
    replyTo: list[dict[str, Any]] | None = None
    createdDateTime: str | None = None
    lastModifiedDateTime: str | None = None
    receivedDateTime: str | None = None
    sentDateTime: str | None = Field(
        default=None,
        description="For an unsent draft this reflects creation time, not an actual send.",
    )


class CreateDraftMessageResult(ToolResult):
    data: CreateDraftMessageData | None = None


class CreateDraftToForwardMessageData(BaseModel):
    model_config = ConfigDict(extra="allow")

    id: str | None = Field(default=None, description="Unique identifier of the new forward draft, consumed by send_draft_message.")
    receivedDateTime: str | None = None
    sentDateTime: str | None = None
    hasAttachments: bool | None = None
    subject: str | None = None
    body: dict[str, Any] | None = None
    bodyPreview: str | None = None


class CreateDraftToForwardMessageResult(ToolResult):
    data: CreateDraftToForwardMessageData | None = None


class CreateDraftToReplyAllData(BaseModel):
    model_config = ConfigDict(extra="allow")

    id: str | None = Field(default=None, description="Unique identifier of the new reply-all draft, consumed by send_draft_message.")
    receivedDateTime: str | None = None
    sentDateTime: str | None = None
    hasAttachments: bool | None = None
    subject: str | None = None
    body: dict[str, Any] | None = None
    bodyPreview: str | None = None


class CreateDraftToReplyAllResult(ToolResult):
    data: CreateDraftToReplyAllData | None = None


class CreateDraftToReplyData(BaseModel):
    model_config = ConfigDict(extra="allow")

    id: str | None = Field(default=None, description="Unique identifier of the new reply draft, consumed by send_draft_message.")
    receivedDateTime: str | None = None
    sentDateTime: str | None = None
    hasAttachments: bool | None = None
    subject: str | None = None
    body: dict[str, Any] | None = None
    bodyPreview: str | None = None


class CreateDraftToReplyResult(ToolResult):
    data: CreateDraftToReplyData | None = None


class ForwardMessageData(BaseModel):
    """Graph returns 202 Accepted with no response body for this endpoint."""

    model_config = ConfigDict(extra="allow")


class ForwardMessageResult(ToolResult):
    data: ForwardMessageData | None = None


class GetMessageData(BaseModel):
    model_config = ConfigDict(extra="allow", populate_by_name=True)

    id: str = Field(description="Unique identifier of the message.")
    subject: str | None = None
    body: dict[str, Any] | None = None
    bodyPreview: str | None = None
    sender: dict[str, Any] | None = None
    from_: dict[str, Any] | None = Field(
        default=None,
        alias="from",
        description="Who the message is from; differs from `sender` only in shared-mailbox or delegate scenarios.",
    )
    toRecipients: list[dict[str, Any]] | None = None
    ccRecipients: list[dict[str, Any]] | None = None
    bccRecipients: list[dict[str, Any]] | None = None
    replyTo: list[dict[str, Any]] | None = None
    isRead: bool | None = None
    isDraft: bool | None = None
    hasAttachments: bool | None = None
    parentFolderId: str | None = None
    conversationId: str | None = None
    webLink: str | None = None
    flag: dict[str, Any] | None = Field(
        default=None,
        description="Follow-up flag state, e.g. {flagStatus: 'notFlagged' | 'flagged' | 'complete'}.",
    )


class GetMessageResult(ToolResult):
    data: GetMessageData | None = None


class MessageListItem(BaseModel):
    model_config = ConfigDict(extra="allow")

    id: str | None = Field(
        default=None,
        description=(
            "Unique identifier of the message, consumed by get_message, reply_message, "
            "reply_all_message, forward_message, and the create_draft_to_* tools."
        ),
    )
    subject: str | None = None
    sender: dict[str, Any] | None = None


class ListMessagesData(BaseModel):
    model_config = ConfigDict(extra="allow", populate_by_name=True)

    value: list[MessageListItem] = Field(default_factory=list)
    odata_next_link: str | None = Field(
        default=None,
        alias="@odata.nextLink",
        description="URL to fetch the next page when present; pass it through as-is instead of constructing `skip` manually.",
    )


class ListMessagesResult(ToolResult):
    data: ListMessagesData | None = None


class ReplyAllMessageData(BaseModel):
    """Graph returns 202 Accepted with no response body for this endpoint."""

    model_config = ConfigDict(extra="allow")


class ReplyAllMessageResult(ToolResult):
    data: ReplyAllMessageData | None = None


class ReplyMessageData(BaseModel):
    """Graph returns 202 Accepted with no response body for this endpoint."""

    model_config = ConfigDict(extra="allow")


class ReplyMessageResult(ToolResult):
    data: ReplyMessageData | None = None


class SendDraftMessageData(BaseModel):
    """Graph returns 202 Accepted with no response body for this endpoint."""

    model_config = ConfigDict(extra="allow")


class SendDraftMessageResult(ToolResult):
    data: SendDraftMessageData | None = None


class SendMailData(BaseModel):
    """Graph returns 202 Accepted with no response body for this endpoint."""

    model_config = ConfigDict(extra="allow")


class SendMailResult(ToolResult):
    data: SendMailData | None = None


class UpdateMessageData(BaseModel):
    model_config = ConfigDict(extra="allow")

    before: GetMessageData
    after: GetMessageData


class UpdateMessageResult(ToolResult):
    data: UpdateMessageData | None = None


class DeleteMessageData(BaseModel):
    """Graph returns 204 No Content for this endpoint."""

    model_config = ConfigDict(extra="allow")


class DeleteMessageResult(ToolResult):
    data: DeleteMessageData | None = None


class CopyMessageData(BaseModel):
    model_config = ConfigDict(extra="allow")

    id: str | None = Field(default=None, description="Unique identifier of the new copy, in the destination folder.")
    parentFolderId: str | None = None
    receivedDateTime: str | None = None
    sentDateTime: str | None = None
    hasAttachments: bool | None = None
    subject: str | None = None
    body: dict[str, Any] | None = None
    bodyPreview: str | None = None


class CopyMessageResult(ToolResult):
    data: CopyMessageData | None = None


class GetMessageDeltaData(BaseModel):
    model_config = ConfigDict(extra="allow", populate_by_name=True)

    value: list[dict[str, Any]] = Field(default_factory=list)
    odata_next_link: str | None = Field(
        default=None,
        alias="@odata.nextLink",
        description="Present when more changes remain to page through in this round.",
    )
    odata_delta_link: str | None = Field(
        default=None,
        alias="@odata.deltaLink",
        description="Present when this round of change tracking is complete; save it to start the next round.",
    )


class GetMessageDeltaResult(ToolResult):
    data: GetMessageDeltaData | None = None


class MoveMessageData(BaseModel):
    model_config = ConfigDict(extra="allow")

    before: GetMessageData
    after: GetMessageData


class MoveMessageResult(ToolResult):
    data: MoveMessageData | None = None


class PermanentlyDeleteMessageData(BaseModel):
    """Graph returns 204 No Content for this endpoint."""

    model_config = ConfigDict(extra="allow")


class PermanentlyDeleteMessageResult(ToolResult):
    data: PermanentlyDeleteMessageData | None = None
