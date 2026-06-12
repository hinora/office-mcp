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
            "required": []
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
                    "description": "Folder: inbox/sent/drafts/deleted"
                },
                "count": {
                    "type": "integer",
                    "description": "Max results"
                },
                "account_name": {
                    "type": "string",
                    "description": "Account name"
                }
            },
            "required": []
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
                    "description": "Subject filter"
                },
                "sender": {
                    "type": "string",
                    "description": "Sender filter"
                },
                "received_after": {
                    "type": "string",
                    "description": "Received after (ISO date)"
                },
                "received_before": {
                    "type": "string",
                    "description": "Received before (ISO date)"
                },
                "unread_only": {
                    "type": "boolean",
                    "description": "Unread only"
                },
                "folder": {
                    "type": "string",
                    "description": "Folder: inbox/sent/drafts/deleted"
                },
                "count": {
                    "type": "integer",
                    "description": "Max results"
                },
                "account_name": {
                    "type": "string",
                    "description": "Account name"
                }
            },
            "required": []
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
                    "description": "Email EntryID"
                }
            },
            "required": ["entry_id"]
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
                    "description": "Recipient(s), semicolon-separated"
                },
                "subject": {
                    "type": "string",
                    "description": "Subject"
                },
                "body": {
                    "type": "string",
                    "description": "Body text"
                },
                "cc": {
                    "type": "string",
                    "description": "CC, semicolon-separated"
                },
                "bcc": {
                    "type": "string",
                    "description": "BCC, semicolon-separated"
                },
                "html_body": {
                    "type": "boolean",
                    "description": "Body is HTML"
                },
                "attachments": {
                    "type": "string",
                    "description": "JSON array of file paths to attach, e.g. ['C:\\file.pdf']."
                },
                "importance": {
                    "type": "string",
                    "description": "low/normal/high"
                }
            },
            "required": ["to", "subject", "body"]
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
                    "description": "Subject"
                },
                "body": {
                    "type": "string",
                    "description": "Body text"
                },
                "to": {
                    "type": "string",
                    "description": "Recipient(s), semicolon-separated"
                },
                "cc": {
                    "type": "string",
                    "description": "CC, semicolon-separated"
                },
                "bcc": {
                    "type": "string",
                    "description": "BCC, semicolon-separated"
                },
                "html_body": {
                    "type": "boolean",
                    "description": "Body is HTML"
                },
                "attachments": {
                    "type": "string",
                    "description": "JSON array of file paths to attach, e.g. ['C:\\file.pdf']."
                },
                "importance": {
                    "type": "string",
                    "description": "low/normal/high"
                }
            },
            "required": []
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
                    "description": "Email EntryID"
                },
                "body": {
                    "type": "string",
                    "description": "Reply body text"
                },
                "reply_all": {
                    "type": "boolean",
                    "description": "Reply to all"
                },
                "html_body": {
                    "type": "boolean",
                    "description": "Body is HTML"
                },
                "attachments": {
                    "type": "string",
                    "description": "JSON array of file paths"
                }
            },
            "required": ["entry_id", "body"]
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
                    "description": "Email EntryID"
                },
                "to": {
                    "type": "string",
                    "description": "Recipient email address(es), semicolon-separated."
                },
                "body": {
                    "type": "string",
                    "description": "Additional message text"
                },
                "html_body": {
                    "type": "boolean",
                    "description": "Body is HTML"
                },
                "attachments": {
                    "type": "string",
                    "description": "JSON array of file paths"
                }
            },
            "required": ["entry_id", "to"]
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
                    "description": "Email EntryID"
                }
            },
            "required": ["entry_id"]
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
                    "description": "Email EntryID"
                },
                "read": {
                    "type": "boolean",
                    "description": "Mark read (true) or unread (false)"
                }
            },
            "required": ["entry_id"]
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
                    "description": "Email EntryID"
                },
                "dest_folder": {
                    "type": "string",
                    "description": "Dest folder: inbox/sent/drafts/deleted"
                }
            },
            "required": ["entry_id", "dest_folder"]
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
                    "description": "Email EntryID"
                },
                "attachment_index": {
                    "type": "integer",
                    "description": "Attachment index (1-based)"
                },
                "save_path": {
                    "type": "string",
                    "description": "Save path"
                }
            },
            "required": ["entry_id"]
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
                    "description": "Start date (ISO)"
                },
                "end_date": {
                    "type": "string",
                    "description": "End date (ISO)"
                },
                "count": {
                    "type": "integer",
                    "description": "Max results"
                },
                "account_name": {
                    "type": "string",
                    "description": "Account name"
                }
            },
            "required": []
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
                    "description": "Subject"
                },
                "start_time": {
                    "type": "string",
                    "description": "Start (ISO datetime)"
                },
                "end_time": {
                    "type": "string",
                    "description": "End (ISO datetime)"
                },
                "body": {
                    "type": "string",
                    "description": "Notes"
                },
                "location": {
                    "type": "string",
                    "description": "Location"
                },
                "all_day": {
                    "type": "boolean",
                    "description": "All-day event"
                },
                "reminder_minutes": {
                    "type": "integer",
                    "description": "Reminder minutes (0=none)"
                },
                "recipients": {
                    "type": "string",
                    "description": "Attendees, semicolon-separated"
                }
            },
            "required": ["subject", "start_time", "end_time"]
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
                    "description": "Appointment EntryID"
                }
            },
            "required": ["entry_id"]
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
                    "description": "Filter by name/email"
                },
                "count": {
                    "type": "integer",
                    "description": "Max results"
                },
                "account_name": {
                    "type": "string",
                    "description": "Account name"
                }
            },
            "required": []
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
                    "description": "Full name"
                },
                "email": {
                    "type": "string",
                    "description": "Email"
                },
                "phone": {
                    "type": "string",
                    "description": "Phone"
                },
                "mobile": {
                    "type": "string",
                    "description": "Mobile"
                },
                "company": {
                    "type": "string",
                    "description": "Company"
                },
                "job_title": {
                    "type": "string",
                    "description": "Job title"
                }
            },
            "required": ["full_name"]
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
                    "description": "Contact EntryID"
                }
            },
            "required": ["entry_id"]
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
                    "description": "Appointment EntryID"
                }
            },
            "required": ["entry_id"]
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
                    "description": "Appointment EntryID"
                },
                "subject": {
                    "type": "string",
                    "description": "Subject"
                },
                "start_time": {
                    "type": "string",
                    "description": "Start (ISO datetime)"
                },
                "end_time": {
                    "type": "string",
                    "description": "End (ISO datetime)"
                },
                "body": {
                    "type": "string",
                    "description": "Notes"
                },
                "location": {
                    "type": "string",
                    "description": "Location"
                },
                "all_day": {
                    "type": "boolean",
                    "description": "All-day event"
                },
                "reminder_minutes": {
                    "type": "integer",
                    "description": "Reminder minutes (0=none)"
                },
                "recipients": {
                    "type": "string",
                    "description": "Attendees, semicolon-separated"
                },
                "send_update": {
                    "type": "boolean",
                    "description": "Send update to attendees"
                }
            },
            "required": ["entry_id"]
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
                    "description": "Meeting EntryID"
                },
                "response": {
                    "type": "string",
                    "description": "accept/decline/tentative"
                },
                "comment": {
                    "type": "string",
                    "description": "Response comment"
                }
            },
            "required": ["entry_id", "response"]
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
                    "description": "Start date (ISO)"
                },
                "months": {
                    "type": "integer",
                    "description": "Months"
                },
                "account_name": {
                    "type": "string",
                    "description": "Account name"
                }
            },
            "required": []
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
                    "description": "Contact EntryID"
                },
                "full_name": {
                    "type": "string",
                    "description": "Full name"
                },
                "email": {
                    "type": "string",
                    "description": "Email"
                },
                "phone": {
                    "type": "string",
                    "description": "Phone"
                },
                "mobile": {
                    "type": "string",
                    "description": "Mobile"
                },
                "home_phone": {
                    "type": "string",
                    "description": "Home phone"
                },
                "company": {
                    "type": "string",
                    "description": "Company"
                },
                "job_title": {
                    "type": "string",
                    "description": "Job title"
                }
            },
            "required": ["entry_id"]
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
                    "description": "csv/vcard"
                },
                "save_path": {
                    "type": "string",
                    "description": "Save path"
                },
                "account_name": {
                    "type": "string",
                    "description": "Account name"
                }
            },
            "required": []
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
                    "description": "Email EntryID"
                },
                "flag": {
                    "type": "boolean",
                    "description": "Flag on/off"
                },
                "due_date": {
                    "type": "string",
                    "description": "Due date (ISO)"
                },
                "reminder_date": {
                    "type": "string",
                    "description": "Reminder (ISO datetime)"
                }
            },
            "required": ["entry_id"]
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
                    "description": "Email EntryID"
                },
                "categories": {
                    "type": "string",
                    "description": "Categories, semicolon-separated"
                },
                "action": {
                    "type": "string",
                    "description": "set/add/remove/clear"
                }
            },
            "required": ["entry_id", "categories"]
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
                    "description": "Account name"
                }
            },
            "required": []
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
                    "description": "Email EntryID"
                }
            },
            "required": ["entry_id"]
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
                    "description": "Draft EntryID"
                },
                "subject": {
                    "type": "string",
                    "description": "Subject"
                },
                "body": {
                    "type": "string",
                    "description": "Body text"
                },
                "to": {
                    "type": "string",
                    "description": "Recipients, semicolon-separated"
                },
                "cc": {
                    "type": "string",
                    "description": "CC, semicolon-separated"
                },
                "bcc": {
                    "type": "string",
                    "description": "BCC, semicolon-separated"
                },
                "html_body": {
                    "type": "boolean",
                    "description": "If True, treat body as HTML. Default: false."
                },
                "attachments": {
                    "type": "string",
                    "description": "JSON array of file paths"
                },
                "importance": {
                    "type": "string",
                    "description": "low/normal/high"
                }
            },
            "required": ["entry_id"]
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
                    "description": "Draft EntryID"
                }
            },
            "required": ["entry_id"]
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
                    "description": "Max results"
                },
                "include_completed": {
                    "type": "boolean",
                    "description": "Include completed tasks"
                },
                "account_name": {
                    "type": "string",
                    "description": "Account name"
                }
            },
            "required": []
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
                    "description": "Subject"
                },
                "body": {
                    "type": "string",
                    "description": "Notes"
                },
                "due_date": {
                    "type": "string",
                    "description": "Due date (ISO)"
                },
                "start_date": {
                    "type": "string",
                    "description": "Start date (ISO)"
                },
                "importance": {
                    "type": "string",
                    "description": "low/normal/high"
                },
                "reminder_minutes": {
                    "type": "integer",
                    "description": "Reminder minutes (0=none)"
                }
            },
            "required": ["subject"]
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
                    "description": "Task EntryID"
                },
                "subject": {
                    "type": "string",
                    "description": "Subject"
                },
                "body": {
                    "type": "string",
                    "description": "Notes"
                },
                "due_date": {
                    "type": "string",
                    "description": "Due date (ISO)"
                },
                "start_date": {
                    "type": "string",
                    "description": "Start date (ISO)"
                },
                "status": {
                    "type": "integer",
                    "description": "0=NotStarted,1=InProgress,2=Complete,3=Waiting,4=Deferred"
                },
                "importance": {
                    "type": "string",
                    "description": "low/normal/high"
                },
                "reminder_minutes": {
                    "type": "integer",
                    "description": "Minutes before due to remind (0 for none)."
                },
                "percent_complete": {
                    "type": "integer",
                    "description": "Percent complete (0-100)"
                }
            },
            "required": ["entry_id"]
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
                    "description": "Task EntryID"
                }
            },
            "required": ["entry_id"]
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
                    "description": "Task EntryID"
                },
                "complete": {
                    "type": "boolean",
                    "description": "Complete (true) or not started (false)"
                }
            },
            "required": ["entry_id"]
        },
    ),
    # ── Rules ──
    Tool(
        name="outlook_get_rules",
        description="Get all Outlook inbox rules.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": []
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
                    "description": "Rule name"
                },
                "condition_type": {
                    "type": "string",
                    "description": "sender/subject"
                },
                "condition_value": {
                    "type": "string",
                    "description": "Email or keyword"
                },
                "action_type": {
                    "type": "string",
                    "description": "move/mark_read/delete/categorize"
                },
                "action_value": {
                    "type": "string",
                    "description": "Folder or category name"
                },
                "enabled": {
                    "type": "boolean",
                    "description": "Rule enabled"
                },
                "account_name": {
                    "type": "string",
                    "description": "Account name"
                }
            },
            "required": ["name", "condition_value", "action_type"]
        },
    )
]

# ── Folder Mappings ────────────────────────────────────────────────────

FOLDER_MAP = {
    "inbox": OutlookClient.OL_FOLDER_INBOX,
    "sent": OutlookClient.OL_FOLDER_SENT,
    "drafts": OutlookClient.OL_FOLDER_DRAFTS,
    "deleted": OutlookClient.OL_FOLDER_DELETED
}

IMPORTANCE_MAP = {
    "low": OutlookClient.IMPORTANCE_LOW,
    "normal": OutlookClient.IMPORTANCE_NORMAL,
    "high": OutlookClient.IMPORTANCE_HIGH
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
        result = await asyncio.wait_for(
            loop.run_in_executor(
                _sta_executor,
                _execute_tool,
                name,
                arguments,
                client,
            ),
            timeout=60.0,
        )
        return [TextContent(type="text", text=result)]

    except asyncio.TimeoutError:
        logger.error(f"Tool '{name}' timed out after 60 seconds")
        return [TextContent(type="text", text=json.dumps({
            "error": f"Tool '{name}' timed out after 60 seconds",
            "tool": name
        }, ensure_ascii=False))]

    except Exception as e:
        logger.exception(f"Error executing tool '{name}'")
        return [TextContent(type="text", text=json.dumps({
            "error": str(e),
            "tool": name
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
