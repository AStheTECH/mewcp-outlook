"""Messages group: create_draft_message, create_draft_to_forward_message,
create_draft_to_reply_all, create_draft_to_reply, forward_message, get_message,
list_messages, reply_all_message, reply_message, send_draft_message, send_mail,
update_message, delete_message, copy_message, get_message_delta, move_message,
permanently_delete_message."""

import logging
from typing import Any

from fastmcp import FastMCP
from mcp.types import ToolAnnotations
from pydantic import Field

from .. import service
from ..config import CONNECT_TIMEOUT, READ_TIMEOUT
from ..logging_utils import ToolLogger
from ..schemas.messages import (
    CopyMessageData,
    CopyMessageResult,
    CreateDraftMessageData,
    CreateDraftMessageResult,
    CreateDraftToForwardMessageData,
    CreateDraftToForwardMessageResult,
    CreateDraftToReplyAllData,
    CreateDraftToReplyAllResult,
    CreateDraftToReplyData,
    CreateDraftToReplyResult,
    DeleteMessageData,
    DeleteMessageResult,
    ForwardMessageData,
    ForwardMessageResult,
    GetMessageData,
    GetMessageDeltaData,
    GetMessageDeltaResult,
    GetMessageResult,
    ListMessagesData,
    ListMessagesResult,
    MoveMessageData,
    MoveMessageResult,
    PermanentlyDeleteMessageData,
    PermanentlyDeleteMessageResult,
    ReplyAllMessageData,
    ReplyAllMessageResult,
    ReplyMessageData,
    ReplyMessageResult,
    SendDraftMessageData,
    SendDraftMessageResult,
    SendMailData,
    SendMailResult,
    UpdateMessageData,
    UpdateMessageResult,
)
from ._helpers import _err, _handle_request_exc, _upstream_err

logger = logging.getLogger("outlook-mcp.tools.messages")


def register_messages_tools(mcp: FastMCP) -> None:

    @mcp.tool(
        name="create_draft_message",
        description=(
            "Creates a draft of a new message and saves it to the Drafts folder without sending it. "
            "Returns the created draft, including the `id` needed to send it later with send_draft_message."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=True),
    )
    def create_draft_message(
        subject: str | None = Field(
            default=None,
            description="Subject line of the draft message. Omit to create the draft with no subject.",
        ),
        importance: str | None = Field(
            default=None,
            description="Importance of the message: Low, Normal, or High. Omit to use the default (Normal).",
        ),
        body: dict[str, Any] | None = Field(
            default=None,
            description=(
                "Message body as {contentType: 'Text' or 'HTML', content: <string>}. "
                "Omit to create the draft with an empty body."
            ),
        ),
        toRecipients: list[dict[str, Any]] | None = Field(
            default=None,
            description=(
                "Recipients of the draft, each as {emailAddress: {address: <string>, name: <string>}}. "
                "Omit to create the draft with no recipients, to be added later."
            ),
        ),
        internetMessageHeaders: list[dict[str, Any]] | None = Field(
            default=None,
            description=(
                "Custom Internet message headers to attach, each as {name: <string>, value: <string>}. "
                "Omit to send no custom headers."
            ),
        ),
    ) -> CreateDraftMessageResult:
        tlog = ToolLogger(logger, "create_draft_message")

        if importance is not None and importance not in ("Low", "Normal", "High"):
            return _err(CreateDraftMessageResult, tlog, "VALIDATION_ERROR",
                        "importance must be one of Low, Normal, High", 400)
        if body is not None and body.get("contentType") is not None and body["contentType"] not in ("Text", "HTML"):
            return _err(CreateDraftMessageResult, tlog, "VALIDATION_ERROR",
                        "body.contentType must be Text or HTML", 400)

        payload: dict[str, Any] = {}
        if subject is not None:
            payload["subject"] = subject
        if importance is not None:
            payload["importance"] = importance
        if body is not None:
            payload["body"] = body
        if toRecipients is not None:
            payload["toRecipients"] = toRecipients
        if internetMessageHeaders is not None:
            payload["internetMessageHeaders"] = internetMessageHeaders

        try:
            data, status, retry_after = service.api_request(
                "POST", "/me/messages", body=payload,
                timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return CreateDraftMessageResult(success=True, statusCode=status, data=CreateDraftMessageData(**data))
            return _upstream_err(CreateDraftMessageResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(CreateDraftMessageResult, tlog, exc)

    @mcp.tool(
        name="create_draft_to_forward_message",
        description=(
            "Creates a draft Forward message for an existing message, which can be edited and sent later "
            "with send_draft_message. Specify recipients via either the top-level `toRecipients` or "
            "`message.toRecipients` (exactly one, not both and not neither), and a lead-in note via either "
            "`comment` or `message.body` (not both). Returns the created draft."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=True),
    )
    def create_draft_to_forward_message(
        id: str = Field(description="ID of the message to forward."),
        toRecipients: list[dict[str, Any]] | None = Field(
            default=None,
            description=(
                "Recipients for the forward draft, each as {emailAddress: {address, name}}. "
                "Specify either this or message.toRecipients, not both."
            ),
        ),
        message: dict[str, Any] | None = Field(
            default=None,
            description=(
                "Message properties to pre-set on the draft (e.g. toRecipients, body). Specify either "
                "message.toRecipients or the top-level toRecipients, not both; either message.body or "
                "comment, not both."
            ),
        ),
        comment: str | None = Field(
            default=None,
            description="Comment prepended to the forwarded message; can be an empty string. Mutually exclusive with message.body.",
        ),
    ) -> CreateDraftToForwardMessageResult:
        tlog = ToolLogger(logger, "create_draft_to_forward_message")

        has_top_to = toRecipients is not None
        has_msg_to = bool(message and message.get("toRecipients") is not None)
        if has_top_to and has_msg_to:
            return _err(CreateDraftToForwardMessageResult, tlog, "VALIDATION_ERROR",
                        "specify either toRecipients or message.toRecipients, not both", 400)
        if not has_top_to and not has_msg_to:
            return _err(CreateDraftToForwardMessageResult, tlog, "VALIDATION_ERROR",
                        "specify either toRecipients or message.toRecipients", 400)

        has_comment = comment is not None
        has_msg_body = bool(message and message.get("body") is not None)
        if has_comment and has_msg_body:
            return _err(CreateDraftToForwardMessageResult, tlog, "VALIDATION_ERROR",
                        "specify either comment or message.body, not both", 400)

        payload: dict[str, Any] = {}
        if toRecipients is not None:
            payload["toRecipients"] = toRecipients
        if message is not None:
            payload["message"] = message
        if comment is not None:
            payload["comment"] = comment

        try:
            data, status, retry_after = service.api_request(
                "POST", f"/me/messages/{id}/createForward", body=payload,
                timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return CreateDraftToForwardMessageResult(
                    success=True, statusCode=status, data=CreateDraftToForwardMessageData(**data))
            return _upstream_err(CreateDraftToForwardMessageResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(CreateDraftToForwardMessageResult, tlog, exc)

    @mcp.tool(
        name="create_draft_to_reply_all",
        description=(
            "Creates a draft Reply All message for an existing message, addressed to the sender and all "
            "original recipients, which can be edited and sent later with send_draft_message. Specify a "
            "lead-in note via either `comment` or `message.body`, not both. Returns the created draft."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=True),
    )
    def create_draft_to_reply_all(
        id: str = Field(description="ID of the message being replied to."),
        comment: str | None = Field(
            default=None,
            description="Plain-text comment prepended to the reply; can be an empty string. Mutually exclusive with message.body.",
        ),
        message: dict[str, Any] | None = Field(
            default=None,
            description="Message properties to pre-set on the draft. Specify either this (message.body) or comment, not both.",
        ),
    ) -> CreateDraftToReplyAllResult:
        tlog = ToolLogger(logger, "create_draft_to_reply_all")

        has_comment = comment is not None
        has_msg_body = bool(message and message.get("body") is not None)
        if has_comment and has_msg_body:
            return _err(CreateDraftToReplyAllResult, tlog, "VALIDATION_ERROR",
                        "specify either comment or message.body, not both", 400)

        payload: dict[str, Any] = {}
        if comment is not None:
            payload["comment"] = comment
        if message is not None:
            payload["message"] = message

        try:
            data, status, retry_after = service.api_request(
                "POST", f"/me/messages/{id}/createReplyAll", body=payload,
                timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return CreateDraftToReplyAllResult(
                    success=True, statusCode=status, data=CreateDraftToReplyAllData(**data))
            return _upstream_err(CreateDraftToReplyAllResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(CreateDraftToReplyAllResult, tlog, exc)

    @mcp.tool(
        name="create_draft_to_reply",
        description=(
            "Creates a draft Reply message for an existing message, addressed to the sender, which can be "
            "edited and sent later with send_draft_message. Specify a lead-in note via either `comment` or "
            "`message.body`, not both; use `message.toRecipients` to add recipients beyond the original "
            "sender. Returns the created draft."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=True),
    )
    def create_draft_to_reply(
        id: str = Field(description="ID of the message being replied to."),
        comment: str | None = Field(
            default=None,
            description="Plain-text comment prepended to the reply; can be an empty string. Mutually exclusive with message.body.",
        ),
        message: dict[str, Any] | None = Field(
            default=None,
            description=(
                "Message properties to pre-set on the draft, e.g. toRecipients to add recipients beyond "
                "the original sender. Specify either this (message.body) or comment, not both."
            ),
        ),
    ) -> CreateDraftToReplyResult:
        tlog = ToolLogger(logger, "create_draft_to_reply")

        has_comment = comment is not None
        has_msg_body = bool(message and message.get("body") is not None)
        if has_comment and has_msg_body:
            return _err(CreateDraftToReplyResult, tlog, "VALIDATION_ERROR",
                        "specify either comment or message.body, not both", 400)

        payload: dict[str, Any] = {}
        if comment is not None:
            payload["comment"] = comment
        if message is not None:
            payload["message"] = message

        try:
            data, status, retry_after = service.api_request(
                "POST", f"/me/messages/{id}/createReply", body=payload,
                timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return CreateDraftToReplyResult(
                    success=True, statusCode=status, data=CreateDraftToReplyData(**data))
            return _upstream_err(CreateDraftToReplyResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(CreateDraftToReplyResult, tlog, exc)

    @mcp.tool(
        name="forward_message",
        description=(
            "Forwards an existing message to the given recipients in a single call and saves the forward "
            "to Sent Items. Requires at least one recipient in `toRecipients`; use "
            "create_draft_to_forward_message instead if the forward needs editing before it's sent. "
            "Returns no content on success (202 Accepted)."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=True),
    )
    def forward_message(
        id: str = Field(description="ID of the message to forward."),
        toRecipients: list[dict[str, Any]] = Field(
            description="Recipients to forward to, each as {emailAddress: {name, address}}.",
        ),
        comment: str | None = Field(
            default=None,
            description="Comment prepended to the forwarded message; can be an empty string. Omit for no comment.",
        ),
    ) -> ForwardMessageResult:
        tlog = ToolLogger(logger, "forward_message")

        if not toRecipients:
            return _err(ForwardMessageResult, tlog, "VALIDATION_ERROR",
                        "toRecipients must contain at least one recipient", 400)

        payload: dict[str, Any] = {"toRecipients": toRecipients}
        if comment is not None:
            payload["comment"] = comment

        try:
            data, status, retry_after = service.api_request(
                "POST", f"/me/messages/{id}/forward", body=payload,
                timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return ForwardMessageResult(success=True, statusCode=status, data=ForwardMessageData())
            return _upstream_err(ForwardMessageResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(ForwardMessageResult, tlog, exc)

    @mcp.tool(
        name="get_message",
        description=(
            "Retrieves the full properties of a single message by its id, including sender, recipients, "
            "body, and read/draft/attachment flags. Use `select` to request only specific properties."
        ),
        annotations=ToolAnnotations(readOnlyHint=True, destructiveHint=False, openWorldHint=True),
    )
    def get_message(
        id: str = Field(description="Unique identifier of the message."),
        select: str | None = Field(
            default=None,
            description=(
                "Comma-separated list of properties to return instead of the full message object, e.g. "
                "'subject,sender'. Omit to return the full message object with all properties."
            ),
        ),
    ) -> GetMessageResult:
        tlog = ToolLogger(logger, "get_message")

        params: dict[str, Any] = {}
        if select is not None:
            params["$select"] = select

        try:
            data, status, retry_after = service.api_request(
                "GET", f"/me/messages/{id}", params=params or None,
                timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return GetMessageResult(success=True, statusCode=status, data=GetMessageData(**data))
            return _upstream_err(GetMessageResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(GetMessageResult, tlog, exc)

    @mcp.tool(
        name="list_messages",
        description=(
            "Lists messages in the signed-in user's mailbox, including the Deleted Items and Clutter "
            "folders, with optional OData filtering, sorting, searching, and paging. Returns a page of up "
            "to `top` messages; follow the response's `odata_next_link` to page through further results "
            "instead of setting `skip` manually."
        ),
        annotations=ToolAnnotations(readOnlyHint=True, destructiveHint=False, openWorldHint=True),
    )
    def list_messages(
        select: str | None = Field(
            default=None,
            description=(
                "Comma-separated list of properties to return, e.g. 'subject,sender'. Omit to return the "
                "default message properties."
            ),
        ),
        filter: str | None = Field(
            default=None,
            description=(
                "OData filter expression to restrict the messages returned, e.g. \"isRead eq false\". "
                "Omit to apply no OData filter, returning all messages."
            ),
        ),
        orderby: str | None = Field(
            default=None,
            description=(
                "Property to sort the results by, e.g. 'receivedDateTime desc'. Omit to use the API's "
                "default ordering."
            ),
        ),
        search: str | None = Field(
            default=None,
            description=(
                "Free-text search expression to match against message content, e.g. 'subject:invoice'. "
                "Omit to apply no search filtering."
            ),
        ),
        top: int = Field(default=10, description="Page size, 1-1000. Defaults to 10."),
        skip: int | None = Field(
            default=None,
            description=(
                "Number of results to skip. Omit to start from the first result with no results skipped. "
                "Prefer following the returned `@odata.nextLink` for paging instead of setting this manually."
            ),
        ),
    ) -> ListMessagesResult:
        tlog = ToolLogger(logger, "list_messages")

        if top < 1 or top > 1000:
            return _err(ListMessagesResult, tlog, "VALIDATION_ERROR", "top must be between 1 and 1000", 400)

        params: dict[str, Any] = {"$top": top}
        if select is not None:
            params["$select"] = select
        if filter is not None:
            params["$filter"] = filter
        if orderby is not None:
            params["$orderby"] = orderby
        if search is not None:
            params["$search"] = search
        if skip is not None:
            params["$skip"] = skip

        try:
            data, status, retry_after = service.api_request(
                "GET", "/me/messages", params=params,
                timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return ListMessagesResult(success=True, statusCode=status, data=ListMessagesData(**data))
            return _upstream_err(ListMessagesResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(ListMessagesResult, tlog, exc)

    @mcp.tool(
        name="reply_all_message",
        description=(
            "Replies to the sender and all recipients of an existing message in a single call and saves "
            "the reply to Sent Items. Use create_draft_to_reply_all instead if the reply needs editing "
            "before it's sent. Returns no content on success (202 Accepted)."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=True),
    )
    def reply_all_message(
        id: str = Field(description="ID of the message being replied to."),
        comment: str | None = Field(
            default=None,
            description="Plain-text comment prepended to the reply; can be an empty string. Omit for no comment.",
        ),
    ) -> ReplyAllMessageResult:
        tlog = ToolLogger(logger, "reply_all_message")

        payload: dict[str, Any] = {}
        if comment is not None:
            payload["comment"] = comment

        try:
            data, status, retry_after = service.api_request(
                "POST", f"/me/messages/{id}/replyAll", body=payload,
                timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return ReplyAllMessageResult(success=True, statusCode=status, data=ReplyAllMessageData())
            return _upstream_err(ReplyAllMessageResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(ReplyAllMessageResult, tlog, exc)

    @mcp.tool(
        name="reply_message",
        description=(
            "Replies to the sender of an existing message in a single call and saves the reply to Sent "
            "Items. Use `message` to override properties such as adding extra `toRecipients`, and use "
            "create_draft_to_reply instead if the reply needs editing before it's sent. Returns no "
            "content on success (202 Accepted)."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=True),
    )
    def reply_message(
        id: str = Field(description="ID of the message being replied to."),
        comment: str | None = Field(
            default=None,
            description="Plain-text comment prepended to the reply; can be an empty string. Omit for no comment.",
        ),
        message: dict[str, Any] | None = Field(
            default=None,
            description=(
                "Message properties to override, e.g. adding extra toRecipients beyond the original "
                "sender. Omit to reply using only the `comment` text with no additional property overrides."
            ),
        ),
    ) -> ReplyMessageResult:
        tlog = ToolLogger(logger, "reply_message")

        payload: dict[str, Any] = {}
        if comment is not None:
            payload["comment"] = comment
        if message is not None:
            payload["message"] = message

        try:
            data, status, retry_after = service.api_request(
                "POST", f"/me/messages/{id}/reply", body=payload,
                timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return ReplyMessageResult(success=True, statusCode=status, data=ReplyMessageData())
            return _upstream_err(ReplyMessageResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(ReplyMessageResult, tlog, exc)

    @mcp.tool(
        name="send_draft_message",
        description=(
            "Sends a previously created draft message by its id and saves it to Sent Items. The draft "
            "must come from create_draft_message, create_draft_to_reply, create_draft_to_reply_all, or "
            "create_draft_to_forward_message. Returns no content on success (202 Accepted)."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=True),
    )
    def send_draft_message(
        id: str = Field(description="ID of the draft to send."),
    ) -> SendDraftMessageResult:
        tlog = ToolLogger(logger, "send_draft_message")

        try:
            data, status, retry_after = service.api_request(
                "POST", f"/me/messages/{id}/send", body=None,
                timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return SendDraftMessageResult(success=True, statusCode=status, data=SendDraftMessageData())
            return _upstream_err(SendDraftMessageResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(SendDraftMessageResult, tlog, exc)

    @mcp.tool(
        name="send_mail",
        description=(
            "Composes and sends a new message in a single call without creating a draft first, saving it "
            "to Sent Items unless `saveToSentItems` is set to false. `message` must be a full message "
            "resource, e.g. subject, body, toRecipients, ccRecipients, attachments. Returns no content on "
            "success (202 Accepted)."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=True),
    )
    def send_mail(
        message: dict[str, Any] = Field(
            description="Full message resource to send, e.g. subject, body, toRecipients, ccRecipients, attachments.",
        ),
        saveToSentItems: bool | None = Field(
            default=None,
            description="Whether to save the sent message to Sent Items. Defaults to true; set false to skip saving.",
        ),
    ) -> SendMailResult:
        tlog = ToolLogger(logger, "send_mail")

        if not message:
            return _err(SendMailResult, tlog, "VALIDATION_ERROR", "message must contain the message to send", 400)

        payload: dict[str, Any] = {"message": message}
        if saveToSentItems is not None:
            payload["saveToSentItems"] = saveToSentItems

        try:
            data, status, retry_after = service.api_request(
                "POST", "/me/sendMail", body=payload,
                timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return SendMailResult(success=True, statusCode=status, data=SendMailData())
            return _upstream_err(SendMailResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(SendMailResult, tlog, exc)

    @mcp.tool(
        name="update_message",
        description=(
            "Updates fields on an existing message — most body/recipient/subject fields are "
            "updatable only while the message is still a draft. Supply only the fields to "
            "change; others keep their current value. NOTE: this overwrites the current field "
            "values and the prior state isn't stored by the API after the call; the response "
            "includes both the before and after state so you have a full record of what changed."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=True),
    )
    def update_message(
        id: str = Field(description="ID of the message to update."),
        bccRecipients: list[dict[str, Any]] | None = Field(default=None, description="The Bcc recipients for the message. Updatable only if isDraft is true. Omit to leave unchanged."),
        body: dict[str, Any] | None = Field(default=None, description="The body of the message ({contentType, content}). Updatable only if isDraft is true. Omit to leave unchanged."),
        categories: list[str] | None = Field(default=None, description="The categories associated with the message. Omit to leave unchanged."),
        ccRecipients: list[dict[str, Any]] | None = Field(default=None, description="The Cc recipients for the message. Updatable only if isDraft is true. Omit to leave unchanged."),
        flag: dict[str, Any] | None = Field(default=None, description="Follow-up flag state, e.g. {flagStatus: 'notFlagged' | 'flagged' | 'complete'}. Omit to leave unchanged."),
        from_: dict[str, Any] | None = Field(default=None, description="The mailbox owner and sender, sent to the API as `from`; must correspond to the actual mailbox used. Omit to leave unchanged."),
        importance: str | None = Field(default=None, description="Low, Normal, or High. Omit to leave unchanged."),
        inferenceClassification: str | None = Field(default=None, description="focused or other. Omit to leave unchanged."),
        internetMessageId: str | None = Field(default=None, description="The message ID per RFC2822. Updatable only if isDraft is true. Omit to leave unchanged."),
        isDeliveryReceiptRequested: bool | None = Field(default=None, description="Whether a delivery receipt is requested for the message. Omit to leave unchanged."),
        isRead: bool | None = Field(default=None, description="Whether the message has been read. Omit to leave unchanged."),
        isReadReceiptRequested: bool | None = Field(default=None, description="Whether a read receipt is requested for the message. Omit to leave unchanged."),
        multiValueExtendedProperties: list[dict[str, Any]] | None = Field(default=None, description="Multi-value extended properties. Updatable only if isDraft is true. Omit to leave unchanged."),
        replyTo: list[dict[str, Any]] | None = Field(default=None, description="Email addresses to use when replying. Updatable only if isDraft is true. Omit to leave unchanged."),
        sender: dict[str, Any] | None = Field(default=None, description="The account actually used to generate the message; must correspond to the actual mailbox used. Omit to leave unchanged."),
        singleValueExtendedProperties: list[dict[str, Any]] | None = Field(default=None, description="Single-value extended properties. Updatable only if isDraft is true. Omit to leave unchanged."),
        subject: str | None = Field(default=None, description="The subject of the message. Updatable only if isDraft is true. Omit to leave unchanged."),
        toRecipients: list[dict[str, Any]] | None = Field(default=None, description="The To recipients for the message. Updatable only if isDraft is true. Omit to leave unchanged."),
    ) -> UpdateMessageResult:
        tlog = ToolLogger(logger, "update_message")

        fields = {
            "bccRecipients": bccRecipients, "body": body, "categories": categories,
            "ccRecipients": ccRecipients, "flag": flag, "from": from_, "importance": importance,
            "inferenceClassification": inferenceClassification, "internetMessageId": internetMessageId,
            "isDeliveryReceiptRequested": isDeliveryReceiptRequested, "isRead": isRead,
            "isReadReceiptRequested": isReadReceiptRequested,
            "multiValueExtendedProperties": multiValueExtendedProperties, "replyTo": replyTo,
            "sender": sender, "singleValueExtendedProperties": singleValueExtendedProperties,
            "subject": subject, "toRecipients": toRecipients,
        }
        payload = {k: v for k, v in fields.items() if v is not None}
        if not payload:
            return _err(UpdateMessageResult, tlog, "VALIDATION_ERROR",
                        "at least one field to change must be provided", 400)

        try:
            before_data, before_status, before_retry = service.api_request(
                "GET", f"/me/messages/{id}", timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if not (200 <= before_status < 300):
                return _upstream_err(UpdateMessageResult, tlog, before_status, before_data, before_retry)
            before = GetMessageData(**before_data)

            after_data, after_status, after_retry = service.api_request(
                "PATCH", f"/me/messages/{id}", body=payload, timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if not (200 <= after_status < 300):
                return _upstream_err(UpdateMessageResult, tlog, after_status, after_data, after_retry)
            after = GetMessageData(**after_data)

            tlog.success()
            return UpdateMessageResult(success=True, statusCode=200, data=UpdateMessageData(before=before, after=after))
        except Exception as exc:
            return _handle_request_exc(UpdateMessageResult, tlog, exc)

    @mcp.tool(
        name="delete_message",
        description=(
            "DESTRUCTIVE — REQUIRES EXPLICIT USER CONFIRMATION BEFORE CALLING. "
            "Deletes a message by moving it to Deleted Items. This is a soft delete — the "
            "message is recoverable from Deleted Items, though items already in the recoverable "
            "items deletions folder may not be deletable this way. For an unrecoverable delete "
            "use permanently_delete_message; to relocate instead of deleting, use move_message. "
            "NEVER call this tool autonomously or as part of an automated flow. You MUST stop, "
            "tell the user exactly which message will be deleted, and wait for their explicit "
            "written confirmation before proceeding."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=True, openWorldHint=True),
    )
    def delete_message(
        id: str = Field(description="ID of the message to delete."),
    ) -> DeleteMessageResult:
        tlog = ToolLogger(logger, "delete_message")

        try:
            data, status, retry_after = service.api_request(
                "DELETE", f"/me/messages/{id}", timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return DeleteMessageResult(success=True, statusCode=status, data=DeleteMessageData())
            return _upstream_err(DeleteMessageResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(DeleteMessageResult, tlog, exc)

    @mcp.tool(
        name="copy_message",
        description=(
            "Copies a message to a destination folder, leaving the original message untouched "
            "at its original id. Returns the new copy, which gets its own id in the destination "
            "folder."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=True),
    )
    def copy_message(
        id: str = Field(description="ID of the message to copy."),
        destinationId: str = Field(description="The destination folder's ID, or a well-known folder name."),
    ) -> CopyMessageResult:
        tlog = ToolLogger(logger, "copy_message")

        payload = {"destinationId": destinationId}

        try:
            data, status, retry_after = service.api_request(
                "POST", f"/me/messages/{id}/copy", body=payload, timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return CopyMessageResult(success=True, statusCode=status, data=CopyMessageData(**data))
            return _upstream_err(CopyMessageResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(CopyMessageResult, tlog, exc)

    @mcp.tool(
        name="get_message_delta",
        description=(
            "Gets messages added, updated, or deleted in a mail folder since the last sync "
            "round, returning a page of changes plus either `odata_next_link` (more pages in "
            "this round) or `odata_delta_link` (round complete — save it to start the next "
            "round). A deleted entry appears with `@removed`/`reason: deleted`; folder-level "
            "sync can also emit entries that don't reflect an actual change to the message itself."
        ),
        annotations=ToolAnnotations(readOnlyHint=True, destructiveHint=False, openWorldHint=True),
    )
    def get_message_delta(
        id: str = Field(description="The mail folder to track changes in (a folder ID, or a well-known folder name such as inbox)."),
        deltatoken: str | None = Field(
            default=None,
            description="State token from the previous delta call's @odata.deltaLink, to start the next round of change tracking. Omit to start a new round.",
        ),
        skiptoken: str | None = Field(
            default=None,
            description="State token from the previous delta call's @odata.nextLink, to continue the current round. Omit if not continuing a page.",
        ),
        changeType: str | None = Field(
            default=None,
            description="Filter by change type: created, updated, or deleted. Must be set on the initial call of a round. Omit to receive all change types.",
        ),
        select: str | None = Field(
            default=None,
            description="Comma-separated fields to return; id is always returned. Must be set on the initial call of a round. Omit to use the default property set.",
        ),
        top: int | None = Field(
            default=None,
            description="Maximum items to return per page. Must be set on the initial call of a round. Omit to use the API default.",
        ),
        expand: str | None = Field(
            default=None,
            description="Related resources to expand inline. Must be set on the initial call of a round. Omit to expand nothing.",
        ),
        filter: str | None = Field(
            default=None,
            description="Only \"receivedDateTime ge {value}\" or \"receivedDateTime gt {value}\" is supported. Must be set on the initial call of a round. Omit to apply no filter.",
        ),
        orderby: str | None = Field(
            default=None,
            description="Only \"receivedDateTime desc\" is supported; otherwise return order isn't guaranteed. Must be set on the initial call of a round. Omit to use the API's default ordering.",
        ),
        max_page_size: int | None = Field(
            default=None,
            description="Sent as the Prefer: odata.maxpagesize={value} header, capping how many messages are returned per page. Omit to use the API default.",
        ),
    ) -> GetMessageDeltaResult:
        tlog = ToolLogger(logger, "get_message_delta")

        params: dict[str, Any] = {}
        if deltatoken is not None:
            params["$deltatoken"] = deltatoken
        if skiptoken is not None:
            params["$skiptoken"] = skiptoken
        if changeType is not None:
            params["changeType"] = changeType
        if select is not None:
            params["$select"] = select
        if top is not None:
            params["$top"] = top
        if expand is not None:
            params["$expand"] = expand
        if filter is not None:
            params["$filter"] = filter
        if orderby is not None:
            params["$orderby"] = orderby

        headers = {"Prefer": f"odata.maxpagesize={max_page_size}"} if max_page_size is not None else None

        try:
            data, status, retry_after = service.api_request(
                "GET", f"/me/mailFolders/{id}/messages/delta", params=params or None, extra_headers=headers,
                timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return GetMessageDeltaResult(success=True, statusCode=status, data=GetMessageDeltaData(**data))
            return _upstream_err(GetMessageDeltaResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(GetMessageDeltaResult, tlog, exc)

    @mcp.tool(
        name="move_message",
        description=(
            "Moves a message to a destination folder. This creates a new copy of the message "
            "there and removes the original — the id the message had before the move stops "
            "being valid. NOTE: the original is not preserved after the call; the response "
            "includes both the before and after state so you have a full record of what changed."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=True),
    )
    def move_message(
        id: str = Field(description="ID of the message to move."),
        destinationId: str = Field(description="The destination folder's ID, or a well-known folder name (e.g. deleteditems)."),
    ) -> MoveMessageResult:
        tlog = ToolLogger(logger, "move_message")

        try:
            before_data, before_status, before_retry = service.api_request(
                "GET", f"/me/messages/{id}", timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if not (200 <= before_status < 300):
                return _upstream_err(MoveMessageResult, tlog, before_status, before_data, before_retry)
            before = GetMessageData(**before_data)

            after_data, after_status, after_retry = service.api_request(
                "POST", f"/me/messages/{id}/move", body={"destinationId": destinationId},
                timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if not (200 <= after_status < 300):
                return _upstream_err(MoveMessageResult, tlog, after_status, after_data, after_retry)
            after = GetMessageData(**after_data)

            tlog.success()
            return MoveMessageResult(success=True, statusCode=200, data=MoveMessageData(before=before, after=after))
        except Exception as exc:
            return _handle_request_exc(MoveMessageResult, tlog, exc)

    @mcp.tool(
        name="permanently_delete_message",
        description=(
            "DESTRUCTIVE — REQUIRES EXPLICIT USER CONFIRMATION BEFORE CALLING. "
            "Permanently deletes a message into the mailbox's purges folder, bypassing Deleted "
            "Items entirely. This is harder than delete_message, and the message becomes "
            "unrecoverable once any retention hold expires. NEVER call this tool autonomously or "
            "as part of an automated flow. You MUST stop, tell the user exactly which message "
            "will be permanently deleted, and wait for their explicit written confirmation "
            "before proceeding."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=True, openWorldHint=True),
    )
    def permanently_delete_message(
        user_id: str = Field(description="The mailbox owner's user ID or user principal name."),
        id: str = Field(description="ID of the message to permanently delete."),
    ) -> PermanentlyDeleteMessageResult:
        tlog = ToolLogger(logger, "permanently_delete_message")

        try:
            data, status, retry_after = service.api_request(
                "POST", f"/users/{user_id}/messages/{id}/permanentDelete",
                timeout=(CONNECT_TIMEOUT, READ_TIMEOUT),
            )
            if 200 <= status < 300:
                tlog.success()
                return PermanentlyDeleteMessageResult(
                    success=True, statusCode=status, data=PermanentlyDeleteMessageData())
            return _upstream_err(PermanentlyDeleteMessageResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(PermanentlyDeleteMessageResult, tlog, exc)
