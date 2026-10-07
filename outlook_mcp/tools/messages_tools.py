"""Messages group: create_draft_message, create_draft_to_forward_message,
create_draft_to_reply_all, create_draft_to_reply, forward_message, get_message,
list_messages, reply_all_to_a_message, reply_to_a_message, send_draft_message, send_mail."""

import logging
from typing import Any

from fastmcp import FastMCP
from mcp.types import ToolAnnotations
from pydantic import Field

from .. import service
from ..config import CONNECT_TIMEOUT, READ_TIMEOUT
from ..logging_utils import ToolLogger
from ..schemas.messages import (
    CreateDraftMessageData,
    CreateDraftMessageResult,
    CreateDraftToForwardMessageData,
    CreateDraftToForwardMessageResult,
    CreateDraftToReplyAllData,
    CreateDraftToReplyAllResult,
    CreateDraftToReplyData,
    CreateDraftToReplyResult,
    ForwardMessageData,
    ForwardMessageResult,
    GetMessageData,
    GetMessageResult,
    ListMessagesData,
    ListMessagesResult,
    ReplyAllToAMessageData,
    ReplyAllToAMessageResult,
    ReplyToAMessageData,
    ReplyToAMessageResult,
    SendDraftMessageData,
    SendDraftMessageResult,
    SendMailData,
    SendMailResult,
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
        subject: str | None = Field(default=None, description="Subject line of the draft message."),
        importance: str | None = Field(
            default=None,
            description="Importance of the message: Low, Normal, or High. Omit to use the default (Normal).",
        ),
        body: dict[str, Any] | None = Field(
            default=None,
            description="Message body as {contentType: 'Text' or 'HTML', content: <string>}.",
        ),
        toRecipients: list[dict[str, Any]] | None = Field(
            default=None,
            description="Recipients of the draft, each as {emailAddress: {address: <string>, name: <string>}}.",
        ),
        internetMessageHeaders: list[dict[str, Any]] | None = Field(
            default=None,
            description="Custom Internet message headers to attach, each as {name: <string>, value: <string>}.",
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
            "`comment` or `message.body` (not both)."
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
            "lead-in note via either `comment` or `message.body`, not both."
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
            "`message.body`, not both; use `message.toRecipients` to add recipients beyond the original sender."
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
            "create_draft_to_forward_message instead if the forward needs editing before it's sent."
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
            description="Comma-separated list of properties to return instead of the full message object, e.g. 'subject,sender'.",
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
            description="Comma-separated list of properties to return, e.g. 'subject,sender'.",
        ),
        filter: str | None = Field(
            default=None,
            description="OData filter expression to restrict the messages returned, e.g. \"isRead eq false\".",
        ),
        orderby: str | None = Field(
            default=None,
            description="Property to sort the results by, e.g. 'receivedDateTime desc'.",
        ),
        search: str | None = Field(
            default=None,
            description="Search expression to match against message content.",
        ),
        top: int = Field(default=10, description="Page size, 1-1000. Defaults to 10."),
        skip: int | None = Field(
            default=None,
            description="Number of results to skip. Prefer following the returned next-page link for paging instead of setting this manually.",
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
        name="reply_all_to_a_message",
        description=(
            "Replies to the sender and all recipients of an existing message in a single call and saves "
            "the reply to Sent Items. Use create_draft_to_reply_all instead if the reply needs editing "
            "before it's sent."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=True),
    )
    def reply_all_to_a_message(
        id: str = Field(description="ID of the message being replied to."),
        comment: str | None = Field(
            default=None,
            description="Plain-text comment prepended to the reply; can be an empty string. Omit for no comment.",
        ),
    ) -> ReplyAllToAMessageResult:
        tlog = ToolLogger(logger, "reply_all_to_a_message")

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
                return ReplyAllToAMessageResult(success=True, statusCode=status, data=ReplyAllToAMessageData())
            return _upstream_err(ReplyAllToAMessageResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(ReplyAllToAMessageResult, tlog, exc)

    @mcp.tool(
        name="reply_to_a_message",
        description=(
            "Replies to the sender of an existing message in a single call and saves the reply to Sent "
            "Items. Use `message` to override properties such as adding extra `toRecipients`, and use "
            "create_draft_to_reply instead if the reply needs editing before it's sent."
        ),
        annotations=ToolAnnotations(readOnlyHint=False, destructiveHint=False, openWorldHint=True),
    )
    def reply_to_a_message(
        id: str = Field(description="ID of the message being replied to."),
        comment: str | None = Field(
            default=None,
            description="Plain-text comment prepended to the reply; can be an empty string. Omit for no comment.",
        ),
        message: dict[str, Any] | None = Field(
            default=None,
            description="Message properties to override, e.g. adding extra toRecipients beyond the original sender.",
        ),
    ) -> ReplyToAMessageResult:
        tlog = ToolLogger(logger, "reply_to_a_message")

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
                return ReplyToAMessageResult(success=True, statusCode=status, data=ReplyToAMessageData())
            return _upstream_err(ReplyToAMessageResult, tlog, status, data, retry_after)
        except Exception as exc:
            return _handle_request_exc(ReplyToAMessageResult, tlog, exc)

    @mcp.tool(
        name="send_draft_message",
        description=(
            "Sends a previously created draft message by its id and saves it to Sent Items. The draft "
            "must come from create_draft_message, create_draft_to_reply, create_draft_to_reply_all, or "
            "create_draft_to_forward_message."
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
            "resource, e.g. subject, body, toRecipients, ccRecipients, attachments."
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
