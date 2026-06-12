"""
MCP Server for Microsoft Outlook.

Provides tools for automating Microsoft Outlook via COM:
- Email management (list, search, send, reply, forward, delete, move, mark read,
  flag, categorize, save attachments, empty deleted, open in window)
- Draft management (create, update, send)
- Calendar management (list, get, create, update, delete appointments,
  respond to invitations, get free/busy)
- Contacts management (list, create, update, delete, export to CSV/vCard)
- Tasks management (list, create, update, delete, mark complete)
- Rules management (list, create inbox rules)
- Mailbox information
"""

from __future__ import annotations

import asyncio
import concurrent.futures
import json
import logging
from typing import Any

import pythoncom

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import (
    Tool,
    TextContent,
)

try:
    from .outlook_client import OutlookClient
except ImportError:
    from outlook_mcp.outlook_client import OutlookClient

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)

# ── MCP Server Setup ───────────────────────────────────────────────────

server = Server("outlook-mcp")

# Dedicated STA thread executor for COM operations
_sta_executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
_client: OutlookClient | None = None
_sta_initialized = False


def _init_sta() -> None:
    """Initialize COM on the STA worker thread and create the client."""
    global _client, _sta_initialized
    if not _sta_initialized:
        pythoncom.CoInitialize()
        _sta_initialized = True
    if _client is None:
        _client = OutlookClient()


def get_client() -> OutlookClient:
    """Get or create the Outlook client singleton."""
    if _client is None:
        future = _sta_executor.submit(_init_sta)
        future.result()
    assert _client is not None
    return _client


# ── Tool Definitions ───────────────────────────────────────────────────

TOOLS = [
    # ── Mailbox ──
    Tool(
        name="outlook_get_mailbox_info",
        description="Get information about the Outlook mailbox (accounts, folder counts, version).",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    # ── Email ──
    Tool(
        name="outlook_list_emails",
        description="List recent emails from a specified folder. Default folder is Inbox.",
        inputSchema={
            "type": "object",
            "properties": {
                "folder": {
                    "type": "string",
                    "description": "Folder name: 'inbox', 'sent', 'drafts', 'deleted'. Default: 'inbox'.",
                    "enum": ["inbox", "sent", "drafts", "deleted"],
                    "default": "inbox",
                },
                "count": {
                    "type": "integer",
                    "description": "Maximum number of emails to return. Default: 50.",
                    "default": 50,
                },
                "account_name": {
                    "type": "string",
                    "description": "Optional display name of a specific email account.",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="outlook_search_emails",
        description="Search emails with filters for subject, sender, date range, read status, etc.",
        inputSchema={
            "type": "object",
            "properties": {
                "subject": {
                    "type": "string",
                    "description": "Filter emails by subject containing this text.",
                },
                "sender": {
                    "type": "string",
                    "description": "Filter emails by sender name/email containing this text.",
                },
                "received_after": {
                    "type": "string",
                    "description": "ISO date string (e.g., '2026-01-01') for emails received after this date.",
                },
                "received_before": {
                    "type": "string",
                    "description": "ISO date string for emails received before this date.",
                },
                "unread_only": {
                    "type": "boolean",
                    "description": "Only return unread emails. Default: false.",
                    "default": False,
                },
                "folder": {
                    "type": "string",
                    "description": "Folder to search: 'inbox', 'sent', 'drafts', 'deleted'. Default: 'inbox'.",
                    "enum": ["inbox", "sent", "drafts", "deleted"],
                    "default": "inbox",
                },
                "count": {
                    "type": "integer",
                    "description": "Maximum number of results. Default: 50.",
                    "default": 50,
                },
                "account_name": {
                    "type": "string",
                    "description": "Optional display name of a specific email account.",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="outlook_get_email",
        description="Get full details of a specific email by its EntryID, including body text.",
        inputSchema={
            "type": "object",
            "properties": {
                "entry_id": {
                    "type": "string",
                    "description": "The Outlook EntryID of the email (obtained from list or search).",
                },
            },
            "required": ["entry_id"],
        },
    ),
    Tool(
        name="outlook_send_email",
        description="Send a new email from Outlook.",
        inputSchema={
            "type": "object",
            "properties": {
                "to": {
                    "type": "string",
                    "description": "Recipient email address(es), separated by semicolons.",
                },
                "subject": {
                    "type": "string",
                    "description": "Email subject line.",
                },
                "body": {
                    "type": "string",
                    "description": "Email body text.",
                },
                "cc": {
                    "type": "string",
                    "description": "CC recipients, semicolon-separated.",
                },
                "bcc": {
                    "type": "string",
                    "description": "BCC recipients, semicolon-separated.",
                },
                "html_body": {
                    "type": "boolean",
                    "description": "If True, body is treated as HTML. Default: false.",
                    "default": False,
                },
                "attachments": {
                    "type": "string",
                    "description": "JSON array of file paths to attach, e.g. ['C:\\file.pdf'].",
                },
                "importance": {
                    "type": "string",
                    "description": "Importance: 'low', 'normal', 'high'. Default: 'normal'.",
                    "enum": ["low", "normal", "high"],
                    "default": "normal",
                },
            },
            "required": ["to", "subject", "body"],
        },
    ),
    Tool(
        name="outlook_create_draft",
        description="Create a draft email (does NOT send). Opens a compose window for you to review and send manually.",
        inputSchema={
            "type": "object",
            "properties": {
                "subject": {
                    "type": "string",
                    "description": "Email subject line.",
                },
                "body": {
                    "type": "string",
                    "description": "Email body text.",
                },
                "to": {
                    "type": "string",
                    "description": "Recipient email address(es), separated by semicolons.",
                },
                "cc": {
                    "type": "string",
                    "description": "CC recipients, semicolon-separated.",
                },
                "bcc": {
                    "type": "string",
                    "description": "BCC recipients, semicolon-separated.",
                },
                "html_body": {
                    "type": "boolean",
                    "description": "If True, body is treated as HTML. Default: false.",
                    "default": False,
                },
                "attachments": {
                    "type": "string",
                    "description": "JSON array of file paths to attach, e.g. ['C:\\file.pdf'].",
                },
                "importance": {
                    "type": "string",
                    "description": "Importance: 'low', 'normal', 'high'. Default: 'normal'.",
                    "enum": ["low", "normal", "high"],
                    "default": "normal",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="outlook_reply_email",
        description="Reply to an email identified by its EntryID.",
        inputSchema={
            "type": "object",
            "properties": {
                "entry_id": {
                    "type": "string",
                    "description": "The Outlook EntryID of the email to reply to.",
                },
                "body": {
                    "type": "string",
                    "description": "Reply body text (added above the original message).",
                },
                "reply_all": {
                    "type": "boolean",
                    "description": "If True, reply to all recipients. Default: false.",
                    "default": False,
                },
                "html_body": {
                    "type": "boolean",
                    "description": "If True, body is treated as HTML. Default: false.",
                    "default": False,
                },
                "attachments": {
                    "type": "string",
                    "description": "JSON array of file paths to attach.",
                },
            },
            "required": ["entry_id", "body"],
        },
    ),
    Tool(
        name="outlook_forward_email",
        description="Forward an email to new recipients.",
        inputSchema={
            "type": "object",
            "properties": {
                "entry_id": {
                    "type": "string",
                    "description": "The Outlook EntryID of the email to forward.",
                },
                "to": {
                    "type": "string",
                    "description": "Recipient email address(es), semicolon-separated.",
                },
                "body": {
                    "type": "string",
                    "description": "Additional message body text.",
                },
                "html_body": {
                    "type": "boolean",
                    "description": "If True, body is treated as HTML. Default: false.",
                    "default": False,
                },
                "attachments": {
                    "type": "string",
                    "description": "JSON array of file paths to attach.",
                },
            },
            "required": ["entry_id", "to"],
        },
    ),
    Tool(
        name="outlook_delete_email",
        description="Delete an email by its EntryID.",
        inputSchema={
            "type": "object",
            "properties": {
                "entry_id": {
                    "type": "string",
                    "description": "The Outlook EntryID of the email to delete.",
                },
            },
            "required": ["entry_id"],
        },
    ),
    Tool(
        name="outlook_mark_read",
        description="Mark an email as read or unread.",
        inputSchema={
            "type": "object",
            "properties": {
                "entry_id": {
                    "type": "string",
                    "description": "The Outlook EntryID of the email.",
                },
                "read": {
                    "type": "boolean",
                    "description": "True to mark as read, False to mark as unread. Default: true.",
                    "default": True,
                },
            },
            "required": ["entry_id"],
        },
    ),
    Tool(
        name="outlook_move_email",
        description="Move an email to a different folder.",
        inputSchema={
            "type": "object",
            "properties": {
                "entry_id": {
                    "type": "string",
                    "description": "The Outlook EntryID of the email.",
                },
                "dest_folder": {
                    "type": "string",
                    "description": "Destination folder: 'inbox', 'sent', 'drafts', 'deleted'.",
                    "enum": ["inbox", "sent", "drafts", "deleted"],
                },
            },
            "required": ["entry_id", "dest_folder"],
        },
    ),
    Tool(
        name="outlook_save_attachment",
        description="Save an attachment from an email to disk.",
        inputSchema={
            "type": "object",
            "properties": {
                "entry_id": {
                    "type": "string",
                    "description": "The Outlook EntryID of the email.",
                },
                "attachment_index": {
                    "type": "integer",
                    "description": "1-based index of the attachment to save. Default: 1.",
                    "default": 1,
                },
                "save_path": {
                    "type": "string",
                    "description": "Directory or full file path to save to. Defaults to current directory.",
                },
            },
            "required": ["entry_id"],
        },
    ),
    # ── Calendar ──
    Tool(
        name="outlook_list_calendar",
        description="List calendar events for a date range (default: today to +30 days).",
        inputSchema={
            "type": "object",
            "properties": {
                "start_date": {
                    "type": "string",
                    "description": "ISO date string for range start (e.g., '2026-01-01'). Default: today.",
                },
                "end_date": {
                    "type": "string",
                    "description": "ISO date string for range end. Default: +30 days.",
                },
                "count": {
                    "type": "integer",
                    "description": "Maximum events to return. Default: 50.",
                    "default": 50,
                },
                "account_name": {
                    "type": "string",
                    "description": "Optional display name of a specific email account.",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="outlook_create_appointment",
        description="Create a new calendar appointment. Set recipients to send a meeting invitation.",
        inputSchema={
            "type": "object",
            "properties": {
                "subject": {
                    "type": "string",
                    "description": "Appointment subject/title.",
                },
                "start_time": {
                    "type": "string",
                    "description": "Start time as ISO datetime string (e.g., '2026-06-11T14:00:00').",
                },
                "end_time": {
                    "type": "string",
                    "description": "End time as ISO datetime string (e.g., '2026-06-11T15:00:00').",
                },
                "body": {
                    "type": "string",
                    "description": "Appointment description/notes.",
                },
                "location": {
                    "type": "string",
                    "description": "Appointment location.",
                },
                "all_day": {
                    "type": "boolean",
                    "description": "If True, create an all-day event. Default: false.",
                    "default": False,
                },
                "reminder_minutes": {
                    "type": "integer",
                    "description": "Minutes before to remind. 0 for no reminder. Default: 15.",
                    "default": 15,
                },
                "recipients": {
                    "type": "string",
                    "description": "Attendees' email addresses, semicolon-separated. Sends invites if provided.",
                },
            },
            "required": ["subject", "start_time", "end_time"],
        },
    ),
    Tool(
        name="outlook_delete_appointment",
        description="Delete a calendar appointment by its EntryID.",
        inputSchema={
            "type": "object",
            "properties": {
                "entry_id": {
                    "type": "string",
                    "description": "The Outlook EntryID of the appointment.",
                },
            },
            "required": ["entry_id"],
        },
    ),
    # ── Contacts ──
    Tool(
        name="outlook_list_contacts",
        description="List contacts from the Outlook address book, optionally filtered by search text.",
        inputSchema={
            "type": "object",
            "properties": {
                "search": {
                    "type": "string",
                    "description": "Optional text to filter contacts by name or email.",
                },
                "count": {
                    "type": "integer",
                    "description": "Maximum contacts to return. Default: 100.",
                    "default": 100,
                },
                "account_name": {
                    "type": "string",
                    "description": "Optional display name of a specific email account.",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="outlook_create_contact",
        description="Create a new contact in Outlook.",
        inputSchema={
            "type": "object",
            "properties": {
                "full_name": {
                    "type": "string",
                    "description": "Contact's full name.",
                },
                "email": {
                    "type": "string",
                    "description": "Email address.",
                },
                "phone": {
                    "type": "string",
                    "description": "Business phone number.",
                },
                "mobile": {
                    "type": "string",
                    "description": "Mobile phone number.",
                },
                "company": {
                    "type": "string",
                    "description": "Company name.",
                },
                "job_title": {
                    "type": "string",
                    "description": "Job title.",
                },
            },
            "required": ["full_name"],
        },
    ),
    Tool(
        name="outlook_delete_contact",
        description="Delete a contact by its EntryID.",
        inputSchema={
            "type": "object",
            "properties": {
                "entry_id": {
                    "type": "string",
                    "description": "The Outlook EntryID of the contact.",
                },
            },
            "required": ["entry_id"],
        },
    ),
    # ── Calendar: Detail / Update / Respond ──
    Tool(
        name="outlook_get_appointment",
        description="Get full details of a calendar appointment by its EntryID.",
        inputSchema={
            "type": "object",
            "properties": {
                "entry_id": {
                    "type": "string",
                    "description": "The Outlook EntryID of the appointment.",
                },
            },
            "required": ["entry_id"],
        },
    ),
    Tool(
        name="outlook_update_appointment",
        description="Update an existing calendar appointment. Only provide fields you want to change.",
        inputSchema={
            "type": "object",
            "properties": {
                "entry_id": {
                    "type": "string",
                    "description": "The Outlook EntryID of the appointment.",
                },
                "subject": {
                    "type": "string",
                    "description": "New subject.",
                },
                "start_time": {
                    "type": "string",
                    "description": "New start time as ISO datetime string.",
                },
                "end_time": {
                    "type": "string",
                    "description": "New end time as ISO datetime string.",
                },
                "body": {
                    "type": "string",
                    "description": "New body/notes text.",
                },
                "location": {
                    "type": "string",
                    "description": "New location.",
                },
                "all_day": {
                    "type": "boolean",
                    "description": "Set all-day event.",
                },
                "reminder_minutes": {
                    "type": "integer",
                    "description": "Minutes before to remind (0 for none).",
                },
                "recipients": {
                    "type": "string",
                    "description": "Updated attendees (semicolon-separated emails).",
                },
                "send_update": {
                    "type": "boolean",
                    "description": "If True, send update to attendees. Default: false.",
                    "default": False,
                },
            },
            "required": ["entry_id"],
        },
    ),
    Tool(
        name="outlook_respond_to_invitation",
        description="Accept, decline, or tentatively accept a meeting invitation.",
        inputSchema={
            "type": "object",
            "properties": {
                "entry_id": {
                    "type": "string",
                    "description": "The Outlook EntryID of the meeting request or appointment.",
                },
                "response": {
                    "type": "string",
                    "description": "Response: 'accept', 'decline', or 'tentative'.",
                    "enum": ["accept", "decline", "tentative"],
                },
                "comment": {
                    "type": "string",
                    "description": "Optional message to include with the response.",
                },
            },
            "required": ["entry_id", "response"],
        },
    ),
    Tool(
        name="outlook_get_free_busy",
        description="Get free/busy information for the current user over a date range.",
        inputSchema={
            "type": "object",
            "properties": {
                "start_date": {
                    "type": "string",
                    "description": "ISO date string for range start. Default: today.",
                },
                "months": {
                    "type": "integer",
                    "description": "Number of months to query. Default: 1.",
                    "default": 1,
                },
                "account_name": {
                    "type": "string",
                    "description": "Optional display name of a specific email account.",
                },
            },
            "required": [],
        },
    ),
    # ── Contacts: Update / Export ──
    Tool(
        name="outlook_update_contact",
        description="Update an existing contact. Only provide fields you want to change.",
        inputSchema={
            "type": "object",
            "properties": {
                "entry_id": {
                    "type": "string",
                    "description": "The Outlook EntryID of the contact.",
                },
                "full_name": {
                    "type": "string",
                    "description": "New full name.",
                },
                "email": {
                    "type": "string",
                    "description": "New email address.",
                },
                "phone": {
                    "type": "string",
                    "description": "New business phone number.",
                },
                "mobile": {
                    "type": "string",
                    "description": "New mobile phone number.",
                },
                "home_phone": {
                    "type": "string",
                    "description": "New home phone number.",
                },
                "company": {
                    "type": "string",
                    "description": "New company name.",
                },
                "job_title": {
                    "type": "string",
                    "description": "New job title.",
                },
            },
            "required": ["entry_id"],
        },
    ),
    Tool(
        name="outlook_export_contacts",
        description="Export contacts to a file (CSV or vCard format).",
        inputSchema={
            "type": "object",
            "properties": {
                "format": {
                    "type": "string",
                    "description": "Export format: 'csv' or 'vcard'. Default: 'csv'.",
                    "enum": ["csv", "vcard"],
                    "default": "csv",
                },
                "save_path": {
                    "type": "string",
                    "description": "Directory to save the file. Default: current directory.",
                },
                "account_name": {
                    "type": "string",
                    "description": "Optional display name of a specific email account.",
                },
            },
            "required": [],
        },
    ),
    # ── Email: Flag / Categorize / Empty / Open ──
    Tool(
        name="outlook_flag_email",
        description="Flag an email for follow-up, optionally with a due date and reminder.",
        inputSchema={
            "type": "object",
            "properties": {
                "entry_id": {
                    "type": "string",
                    "description": "The Outlook EntryID of the email.",
                },
                "flag": {
                    "type": "boolean",
                    "description": "True to flag, False to remove flag. Default: true.",
                    "default": True,
                },
                "due_date": {
                    "type": "string",
                    "description": "ISO date string for the flag due date.",
                },
                "reminder_date": {
                    "type": "string",
                    "description": "ISO datetime string for a reminder.",
                },
            },
            "required": ["entry_id"],
        },
    ),
    Tool(
        name="outlook_categorize_email",
        description="Add, remove, set, or clear categories on an email.",
        inputSchema={
            "type": "object",
            "properties": {
                "entry_id": {
                    "type": "string",
                    "description": "The Outlook EntryID of the email.",
                },
                "categories": {
                    "type": "string",
                    "description": "Category name(s), semicolon-separated.",
                },
                "action": {
                    "type": "string",
                    "description": "Action: 'set' (replace), 'add' (append), 'remove' (remove specific), 'clear' (remove all). Default: 'set'.",
                    "enum": ["set", "add", "remove", "clear"],
                    "default": "set",
                },
            },
            "required": ["entry_id", "categories"],
        },
    ),
    Tool(
        name="outlook_empty_deleted",
        description="Empty the Deleted Items folder.",
        inputSchema={
            "type": "object",
            "properties": {
                "account_name": {
                    "type": "string",
                    "description": "Optional display name of a specific email account.",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="outlook_open_email",
        description="Open an email in a separate Outlook window for review.",
        inputSchema={
            "type": "object",
            "properties": {
                "entry_id": {
                    "type": "string",
                    "description": "The Outlook EntryID of the email to open.",
                },
            },
            "required": ["entry_id"],
        },
    ),
    # ── Draft Management ──
    Tool(
        name="outlook_update_draft",
        description="Update an existing draft email. Only provide fields you want to change.",
        inputSchema={
            "type": "object",
            "properties": {
                "entry_id": {
                    "type": "string",
                    "description": "The Outlook EntryID of the draft.",
                },
                "subject": {
                    "type": "string",
                    "description": "New subject.",
                },
                "body": {
                    "type": "string",
                    "description": "New body text.",
                },
                "to": {
                    "type": "string",
                    "description": "New recipients (semicolon-separated).",
                },
                "cc": {
                    "type": "string",
                    "description": "New CC recipients.",
                },
                "bcc": {
                    "type": "string",
                    "description": "New BCC recipients.",
                },
                "html_body": {
                    "type": "boolean",
                    "description": "If True, treat body as HTML. Default: false.",
                    "default": False,
                },
                "attachments": {
                    "type": "string",
                    "description": "JSON array of additional file paths to attach.",
                },
                "importance": {
                    "type": "string",
                    "description": "Importance: 'low', 'normal', 'high'.",
                    "enum": ["low", "normal", "high"],
                },
            },
            "required": ["entry_id"],
        },
    ),
    Tool(
        name="outlook_send_draft",
        description="Send an existing draft email.",
        inputSchema={
            "type": "object",
            "properties": {
                "entry_id": {
                    "type": "string",
                    "description": "The Outlook EntryID of the draft to send.",
                },
            },
            "required": ["entry_id"],
        },
    ),
    # ── Tasks ──
    Tool(
        name="outlook_list_tasks",
        description="List tasks from the Outlook Tasks folder.",
        inputSchema={
            "type": "object",
            "properties": {
                "count": {
                    "type": "integer",
                    "description": "Maximum tasks to return. Default: 50.",
                    "default": 50,
                },
                "include_completed": {
                    "type": "boolean",
                    "description": "Include completed tasks. Default: false.",
                    "default": False,
                },
                "account_name": {
                    "type": "string",
                    "description": "Optional display name of a specific email account.",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="outlook_create_task",
        description="Create a new task.",
        inputSchema={
            "type": "object",
            "properties": {
                "subject": {
                    "type": "string",
                    "description": "Task subject/title.",
                },
                "body": {
                    "type": "string",
                    "description": "Task body/notes.",
                },
                "due_date": {
                    "type": "string",
                    "description": "ISO date string for due date.",
                },
                "start_date": {
                    "type": "string",
                    "description": "ISO date string for start date.",
                },
                "importance": {
                    "type": "string",
                    "description": "Importance: 'low', 'normal', 'high'. Default: 'normal'.",
                    "enum": ["low", "normal", "high"],
                    "default": "normal",
                },
                "reminder_minutes": {
                    "type": "integer",
                    "description": "Minutes before due to remind (0 for none). Default: 0.",
                    "default": 0,
                },
            },
            "required": ["subject"],
        },
    ),
    Tool(
        name="outlook_update_task",
        description="Update an existing task. Only provide fields you want to change.",
        inputSchema={
            "type": "object",
            "properties": {
                "entry_id": {
                    "type": "string",
                    "description": "The Outlook EntryID of the task.",
                },
                "subject": {
                    "type": "string",
                    "description": "New subject.",
                },
                "body": {
                    "type": "string",
                    "description": "New body/notes.",
                },
                "due_date": {
                    "type": "string",
                    "description": "New due date (ISO date string).",
                },
                "start_date": {
                    "type": "string",
                    "description": "New start date (ISO date string).",
                },
                "status": {
                    "type": "integer",
                    "description": "0=NotStarted, 1=InProgress, 2=Complete, 3=Waiting, 4=Deferred.",
                },
                "importance": {
                    "type": "string",
                    "description": "Importance: 'low', 'normal', 'high'.",
                    "enum": ["low", "normal", "high"],
                },
                "reminder_minutes": {
                    "type": "integer",
                    "description": "Minutes before due to remind (0 for none).",
                },
                "percent_complete": {
                    "type": "integer",
                    "description": "Completion percentage (0-100).",
                },
            },
            "required": ["entry_id"],
        },
    ),
    Tool(
        name="outlook_delete_task",
        description="Delete a task by its EntryID.",
        inputSchema={
            "type": "object",
            "properties": {
                "entry_id": {
                    "type": "string",
                    "description": "The Outlook EntryID of the task.",
                },
            },
            "required": ["entry_id"],
        },
    ),
    Tool(
        name="outlook_mark_task_complete",
        description="Mark a task as complete or not started.",
        inputSchema={
            "type": "object",
            "properties": {
                "entry_id": {
                    "type": "string",
                    "description": "The Outlook EntryID of the task.",
                },
                "complete": {
                    "type": "boolean",
                    "description": "True to mark complete, False to mark not started. Default: true.",
                    "default": True,
                },
            },
            "required": ["entry_id"],
        },
    ),
    # ── Rules ──
    Tool(
        name="outlook_get_rules",
        description="Get all Outlook inbox rules.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    Tool(
        name="outlook_create_rule",
        description="Create a simple inbox rule for organizing incoming emails.",
        inputSchema={
            "type": "object",
            "properties": {
                "name": {
                    "type": "string",
                    "description": "Rule name.",
                },
                "condition_type": {
                    "type": "string",
                    "description": "Condition: 'sender' (by email) or 'subject' (by keyword). Default: 'sender'.",
                    "enum": ["sender", "subject"],
                    "default": "sender",
                },
                "condition_value": {
                    "type": "string",
                    "description": "Value for the condition (email address or subject keyword).",
                },
                "action_type": {
                    "type": "string",
                    "description": "Action: 'move', 'mark_read', 'delete', or 'categorize'.",
                    "enum": ["move", "mark_read", "delete", "categorize"],
                },
                "action_value": {
                    "type": "string",
                    "description": "Target folder name (for 'move') or category name (for 'categorize').",
                },
                "enabled": {
                    "type": "boolean",
                    "description": "Whether the rule is enabled. Default: true.",
                    "default": True,
                },
                "account_name": {
                    "type": "string",
                    "description": "Optional display name of a specific email account.",
                },
            },
            "required": ["name", "condition_value", "action_type"],
        },
    ),
]

# ── Folder Mappings ────────────────────────────────────────────────────

FOLDER_MAP = {
    "inbox": OutlookClient.OL_FOLDER_INBOX,
    "sent": OutlookClient.OL_FOLDER_SENT,
    "drafts": OutlookClient.OL_FOLDER_DRAFTS,
    "deleted": OutlookClient.OL_FOLDER_DELETED,
}

IMPORTANCE_MAP = {
    "low": OutlookClient.IMPORTANCE_LOW,
    "normal": OutlookClient.IMPORTANCE_NORMAL,
    "high": OutlookClient.IMPORTANCE_HIGH,
}


# ── Tool Handler ────────────────────────────────────────────────────────

@server.list_tools()
async def list_tools() -> list[Tool]:
    """Return the list of available tools."""
    return TOOLS


@server.call_tool()
async def call_tool(name: str, arguments: dict[str, Any]) -> list[TextContent]:
    """Handle tool calls."""
    try:
        client = get_client()

        loop = asyncio.get_running_loop()
        result = await loop.run_in_executor(
            _sta_executor,
            _execute_tool,
            name,
            arguments,
            client,
        )
        return [TextContent(type="text", text=result)]

    except Exception as e:
        logger.exception(f"Error executing tool '{name}'")
        return [TextContent(type="text", text=json.dumps({
            "error": str(e),
            "tool": name,
        }, ensure_ascii=False))]


def _execute_tool(name: str, args: dict[str, Any], client: OutlookClient) -> str:
    """Execute a tool synchronously. Called from the STA executor thread."""
    pythoncom.CoInitialize()

    result: Any = None

    # ── Mailbox ──
    if name == "outlook_get_mailbox_info":
        result = client.get_mailbox_info()

    # ── Email ──
    elif name == "outlook_list_emails":
        folder = FOLDER_MAP.get(args.get("folder", "inbox"), OutlookClient.OL_FOLDER_INBOX)
        count = args.get("count", 50)
        account = args.get("account_name")
        result = {"emails": client.list_emails(folder, count, account)}

    elif name == "outlook_search_emails":
        folder = FOLDER_MAP.get(args.get("folder", "inbox"), OutlookClient.OL_FOLDER_INBOX)
        count = args.get("count", 50)
        account = args.get("account_name")
        result = {
            "emails": client.search_emails(
                folder_type=folder,
                count=count,
                account_name=account,
                subject=args.get("subject"),
                sender=args.get("sender"),
                received_after=args.get("received_after"),
                received_before=args.get("received_before"),
                unread_only=args.get("unread_only", False),
            )
        }

    elif name == "outlook_get_email":
        result = client.get_email_by_id(args["entry_id"])

    elif name == "outlook_send_email":
        attachments = None
        if "attachments" in args and args["attachments"]:
            attachments = json.loads(args["attachments"])
        importance = IMPORTANCE_MAP.get(args.get("importance", "normal"), 1)
        result = client.send_email(
            to=args["to"],
            subject=args["subject"],
            body=args["body"],
            cc=args.get("cc", ""),
            bcc=args.get("bcc", ""),
            attachments=attachments,
            html_body=args.get("html_body", False),
            importance=importance,
        )

    elif name == "outlook_create_draft":
        attachments = None
        if "attachments" in args and args["attachments"]:
            attachments = json.loads(args["attachments"])
        importance = IMPORTANCE_MAP.get(args.get("importance", "normal"), 1)
        result = client.create_draft(
            to=args.get("to", ""),
            subject=args.get("subject", ""),
            body=args.get("body", ""),
            cc=args.get("cc", ""),
            bcc=args.get("bcc", ""),
            attachments=attachments,
            html_body=args.get("html_body", False),
            importance=importance,
        )

    elif name == "outlook_reply_email":
        attachments = None
        if "attachments" in args and args["attachments"]:
            attachments = json.loads(args["attachments"])
        result = client.reply_email(
            entry_id=args["entry_id"],
            body=args["body"],
            reply_all=args.get("reply_all", False),
            html_body=args.get("html_body", False),
            attachments=attachments,
        )

    elif name == "outlook_forward_email":
        attachments = None
        if "attachments" in args and args["attachments"]:
            attachments = json.loads(args["attachments"])
        result = client.forward_email(
            entry_id=args["entry_id"],
            to=args["to"],
            body=args.get("body", ""),
            html_body=args.get("html_body", False),
            attachments=attachments,
        )

    elif name == "outlook_delete_email":
        result = client.delete_email(args["entry_id"])

    elif name == "outlook_mark_read":
        result = client.mark_as_read(args["entry_id"], args.get("read", True))

    elif name == "outlook_move_email":
        dest = FOLDER_MAP.get(args["dest_folder"], OutlookClient.OL_FOLDER_INBOX)
        result = client.move_email(args["entry_id"], dest)

    elif name == "outlook_save_attachment":
        result = client.save_attachment(
            entry_id=args["entry_id"],
            attachment_index=args.get("attachment_index", 1),
            save_path=args.get("save_path", ""),
        )

    # ── Calendar ──
    elif name == "outlook_list_calendar":
        count = args.get("count", 50)
        account = args.get("account_name")
        result = {
            "events": client.list_calendar_events(
                start_date=args.get("start_date"),
                end_date=args.get("end_date"),
                count=count,
                account_name=account,
            )
        }

    elif name == "outlook_create_appointment":
        result = client.create_appointment(
            subject=args["subject"],
            start_time=args["start_time"],
            end_time=args["end_time"],
            body=args.get("body", ""),
            location=args.get("location", ""),
            all_day=args.get("all_day", False),
            reminder_minutes=args.get("reminder_minutes", 15),
            recipients=args.get("recipients", ""),
        )

    elif name == "outlook_delete_appointment":
        result = client.delete_appointment(args["entry_id"])

    # ── Contacts ──
    elif name == "outlook_list_contacts":
        count = args.get("count", 100)
        account = args.get("account_name")
        result = {
            "contacts": client.list_contacts(
                count=count,
                search=args.get("search", ""),
                account_name=account,
            )
        }

    elif name == "outlook_create_contact":
        result = client.create_contact(
            full_name=args["full_name"],
            email=args.get("email", ""),
            phone=args.get("phone", ""),
            mobile=args.get("mobile", ""),
            company=args.get("company", ""),
            job_title=args.get("job_title", ""),
        )

    elif name == "outlook_delete_contact":
        result = client.delete_contact(args["entry_id"])

    # ── Calendar: Detail / Update / Respond ──
    elif name == "outlook_get_appointment":
        result = client.get_appointment_by_id(args["entry_id"])

    elif name == "outlook_update_appointment":
        result = client.update_appointment(
            entry_id=args["entry_id"],
            subject=args.get("subject"),
            start_time=args.get("start_time"),
            end_time=args.get("end_time"),
            body=args.get("body"),
            location=args.get("location"),
            all_day=args.get("all_day"),
            reminder_minutes=args.get("reminder_minutes"),
            recipients=args.get("recipients"),
            send_update=args.get("send_update", False),
        )

    elif name == "outlook_respond_to_invitation":
        result = client.respond_to_invitation(
            entry_id=args["entry_id"],
            response=args["response"],
            comment=args.get("comment", ""),
        )

    elif name == "outlook_get_free_busy":
        result = client.get_free_busy(
            start_date=args.get("start_date"),
            months=args.get("months", 1),
            account_name=args.get("account_name"),
        )

    # ── Contacts: Update / Export ──
    elif name == "outlook_update_contact":
        result = client.update_contact(
            entry_id=args["entry_id"],
            full_name=args.get("full_name"),
            email=args.get("email"),
            phone=args.get("phone"),
            mobile=args.get("mobile"),
            home_phone=args.get("home_phone"),
            company=args.get("company"),
            job_title=args.get("job_title"),
        )

    elif name == "outlook_export_contacts":
        result = client.export_contacts(
            format=args.get("format", "csv"),
            save_path=args.get("save_path", ""),
            account_name=args.get("account_name"),
        )

    # ── Email: Flag / Categorize / Empty / Open ──
    elif name == "outlook_flag_email":
        result = client.flag_email(
            entry_id=args["entry_id"],
            flag=args.get("flag", True),
            due_date=args.get("due_date"),
            reminder_date=args.get("reminder_date"),
        )

    elif name == "outlook_categorize_email":
        result = client.categorize_email(
            entry_id=args["entry_id"],
            categories=args["categories"],
            action=args.get("action", "set"),
        )

    elif name == "outlook_empty_deleted":
        result = client.empty_deleted_folder(
            account_name=args.get("account_name"),
        )

    elif name == "outlook_open_email":
        result = client.open_email(args["entry_id"])

    # ── Draft Management ──
    elif name == "outlook_update_draft":
        attachments = None
        if "attachments" in args and args["attachments"]:
            attachments = json.loads(args["attachments"])
        importance = IMPORTANCE_MAP.get(args.get("importance")) if args.get("importance") else None
        result = client.update_draft(
            entry_id=args["entry_id"],
            subject=args.get("subject"),
            body=args.get("body"),
            to=args.get("to"),
            cc=args.get("cc"),
            bcc=args.get("bcc"),
            html_body=args.get("html_body", False),
            attachments=attachments,
            importance=importance,
        )

    elif name == "outlook_send_draft":
        result = client.send_draft(args["entry_id"])

    # ── Tasks ──
    elif name == "outlook_list_tasks":
        count = args.get("count", 50)
        account = args.get("account_name")
        result = {
            "tasks": client.list_tasks(
                count=count,
                include_completed=args.get("include_completed", False),
                account_name=account,
            )
        }

    elif name == "outlook_create_task":
        importance = IMPORTANCE_MAP.get(args.get("importance", "normal"), 1)
        result = client.create_task(
            subject=args["subject"],
            body=args.get("body", ""),
            due_date=args.get("due_date"),
            start_date=args.get("start_date"),
            importance=importance,
            reminder_minutes=args.get("reminder_minutes", 0),
        )

    elif name == "outlook_update_task":
        importance = IMPORTANCE_MAP.get(args.get("importance")) if args.get("importance") else None
        result = client.update_task(
            entry_id=args["entry_id"],
            subject=args.get("subject"),
            body=args.get("body"),
            due_date=args.get("due_date"),
            start_date=args.get("start_date"),
            status=args.get("status"),
            importance=importance,
            reminder_minutes=args.get("reminder_minutes"),
            percent_complete=args.get("percent_complete"),
        )

    elif name == "outlook_delete_task":
        result = client.delete_task(args["entry_id"])

    elif name == "outlook_mark_task_complete":
        result = client.mark_task_complete(
            entry_id=args["entry_id"],
            complete=args.get("complete", True),
        )

    # ── Rules ──
    elif name == "outlook_get_rules":
        result = client.get_rules()

    elif name == "outlook_create_rule":
        result = client.create_rule(
            name=args["name"],
            condition_type=args.get("condition_type", "sender"),
            condition_value=args["condition_value"],
            action_type=args["action_type"],
            action_value=args.get("action_value", ""),
            enabled=args.get("enabled", True),
            account_name=args.get("account_name"),
        )

    else:
        raise ValueError(f"Unknown tool: {name}")

    return json.dumps(result, ensure_ascii=False, default=str, indent=2)


# ── Entry Point ─────────────────────────────────────────────────────────

async def run_server() -> None:
    """Run the MCP server via stdio transport."""
    async with stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            server.create_initialization_options(),
        )


def main() -> None:
    """Entry point for the outlook-mcp command."""
    asyncio.run(run_server())


if __name__ == "__main__":
    main()
