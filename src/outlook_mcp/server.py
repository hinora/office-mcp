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

TOOLS: list[Tool] = [
    # ── 1. Mailbox ──
    Tool(
        name="outlook_mailbox",
        description="Get Outlook mailbox info (accounts, folder counts, version).",
        inputSchema={"type": "object", "properties": {}, "required": []},
    ),
    # ── 2. Email (17 actions) ──
    Tool(
        name="outlook_email",
        description="Manage Outlook emails. Use 'action' to specify the operation.",
        inputSchema={
            "type": "object",
            "properties": {
                "action": {
                    "type": "string",
                    "description": "Action: list, search, get, create_draft, delete, move, mark_read, flag, categorize, save_attachment, empty_deleted, open, update_draft",
                    "enum": ["list", "search", "get", "create_draft", "delete", "move", "mark_read", "flag", "categorize", "save_attachment", "empty_deleted", "open", "update_draft"],
                },
                "folder": {"type": "string", "description": "Folder: inbox/sent/drafts/deleted (default: inbox)"},
                "count": {"type": "integer", "description": "Max results (default: 20)"},
                "offset": {"type": "integer", "description": "Skip first N results for pagination (default: 0)"},
                "fields": {"type": "string", "description": "Comma-separated fields to return (e.g. 'subject,sender_name,received_time'). Omit for compact summary."},
                "account_name": {"type": "string", "description": "Account name"},
                "entry_id": {"type": "string", "description": "Email/Draft EntryID for get/update_draft/move/mark_read/flag/categorize/delete/save_attachment/open"},
                "subject": {"type": "string", "description": "Subject filter (search) or email subject"},
                "sender": {"type": "string", "description": "Sender filter (search)"},
                "received_after": {"type": "string", "description": "Received after ISO date (search)"},
                "received_before": {"type": "string", "description": "Received before ISO date (search)"},
                "unread_only": {"type": "boolean", "description": "Unread only (search)"},
                "to": {"type": "string", "description": "Recipients, semicolon-separated"},
                "cc": {"type": "string", "description": "CC, semicolon-separated"},
                "bcc": {"type": "string", "description": "BCC, semicolon-separated"},
                "body": {"type": "string", "description": "Email body text"},
                "html_body": {"type": "boolean", "description": "Treat body as HTML"},
                "attachments": {"type": "string", "description": "JSON array of file paths"},
                "importance": {"type": "string", "description": "low/normal/high"},
                "dest_folder": {"type": "string", "description": "Destination folder: inbox/sent/drafts/deleted (move)"},
                "read": {"type": "boolean", "description": "Mark read (true) or unread (false)"},
                "flag": {"type": "boolean", "description": "Flag on/off (flag action)"},
                "due_date": {"type": "string", "description": "Due date ISO (flag)"},
                "reminder_date": {"type": "string", "description": "Reminder ISO datetime (flag)"},
                "categories": {"type": "string", "description": "Categories, semicolon-separated (categorize)"},
                "cat_action": {"type": "string", "description": "set/add/remove/clear (categorize)"},
                "attachment_index": {"type": "integer", "description": "1-based attachment index (save_attachment, default: 1)"},
                "save_path": {"type": "string", "description": "Save path for attachment"},
            },
            "required": ["action"],
        },
    ),
    # ── 3. Calendar (7 actions) ──
    Tool(
        name="outlook_calendar",
        description="Manage Outlook calendar. Use 'action' to specify the operation.",
        inputSchema={
            "type": "object",
            "properties": {
                "action": {
                    "type": "string",
                    "description": "Action: list, get, create, update, delete, respond, freebusy",
                    "enum": ["list", "get", "create", "update", "delete", "respond", "freebusy"],
                },
                "entry_id": {"type": "string", "description": "Appointment EntryID (get/update/delete/respond)"},
                "start_date": {"type": "string", "description": "Start date ISO (list/freebusy)"},
                "end_date": {"type": "string", "description": "End date ISO (list)"},
                "count": {"type": "integer", "description": "Max results (list, default: 20)"},
                "offset": {"type": "integer", "description": "Skip first N results for pagination (list, default: 0)"},
                "fields": {"type": "string", "description": "Comma-separated fields to return. Omit for compact summary."},
                "account_name": {"type": "string", "description": "Account name"},
                "subject": {"type": "string", "description": "Subject (create/update)"},
                "start_time": {"type": "string", "description": "Start ISO datetime (create/update)"},
                "end_time": {"type": "string", "description": "End ISO datetime (create/update)"},
                "body": {"type": "string", "description": "Notes (create/update)"},
                "location": {"type": "string", "description": "Location (create/update)"},
                "all_day": {"type": "boolean", "description": "All-day event (create/update)"},
                "reminder_minutes": {"type": "integer", "description": "Reminder minutes, 0=none (create/update)"},
                "recipients": {"type": "string", "description": "Attendees, semicolon-separated (create/update)"},
                "send_update": {"type": "boolean", "description": "Send update to attendees (update)"},
                "response": {"type": "string", "description": "accept/decline/tentative (respond)"},
                "comment": {"type": "string", "description": "Response comment (respond)"},
                "months": {"type": "integer", "description": "Number of months (freebusy, default: 1)"},
            },
            "required": ["action"],
        },
    ),
    # ── 4. Contacts (6 actions) ──
    Tool(
        name="outlook_contact",
        description="Manage Outlook contacts. Use 'action' to specify the operation.",
        inputSchema={
            "type": "object",
            "properties": {
                "action": {
                    "type": "string",
                    "description": "Action: list, get, create, update, delete, export",
                    "enum": ["list", "get", "create", "update", "delete", "export"],
                },
                "entry_id": {"type": "string", "description": "Contact EntryID (get/update/delete)"},
                "search": {"type": "string", "description": "Filter by name/email (list)"},
                "count": {"type": "integer", "description": "Max results (list, default: 50)"},
                "offset": {"type": "integer", "description": "Skip first N results for pagination (list, default: 0)"},
                "fields": {"type": "string", "description": "Comma-separated fields to return. Omit for compact summary."},
                "account_name": {"type": "string", "description": "Account name"},
                "full_name": {"type": "string", "description": "Full name (create/update)"},
                "email": {"type": "string", "description": "Email (create/update)"},
                "phone": {"type": "string", "description": "Phone (create/update)"},
                "mobile": {"type": "string", "description": "Mobile (create/update)"},
                "home_phone": {"type": "string", "description": "Home phone (update)"},
                "company": {"type": "string", "description": "Company (create/update)"},
                "job_title": {"type": "string", "description": "Job title (create/update)"},
                "format": {"type": "string", "description": "csv/vcard (export, default: csv)"},
                "save_path": {"type": "string", "description": "Save path (export)"},
            },
            "required": ["action"],
        },
    ),
    # ── 5. Tasks (6 actions) ──
    Tool(
        name="outlook_task",
        description="Manage Outlook tasks. Use 'action' to specify the operation.",
        inputSchema={
            "type": "object",
            "properties": {
                "action": {
                    "type": "string",
                    "description": "Action: list, get, create, update, delete, mark_complete",
                    "enum": ["list", "get", "create", "update", "delete", "mark_complete"],
                },
                "entry_id": {"type": "string", "description": "Task EntryID (get/update/delete/mark_complete)"},
                "count": {"type": "integer", "description": "Max results (list, default: 20)"},
                "offset": {"type": "integer", "description": "Skip first N results for pagination (list, default: 0)"},
                "fields": {"type": "string", "description": "Comma-separated fields to return. Omit for compact summary."},
                "include_completed": {"type": "boolean", "description": "Include completed tasks (list)"},
                "account_name": {"type": "string", "description": "Account name"},
                "subject": {"type": "string", "description": "Subject (create/update)"},
                "body": {"type": "string", "description": "Notes (create/update)"},
                "due_date": {"type": "string", "description": "Due date ISO (create/update)"},
                "start_date": {"type": "string", "description": "Start date ISO (create/update)"},
                "importance": {"type": "string", "description": "low/normal/high (create/update)"},
                "reminder_minutes": {"type": "integer", "description": "Reminder minutes, 0=none (create/update)"},
                "status": {"type": "integer", "description": "0=NotStarted,1=InProgress,2=Complete,3=Waiting,4=Deferred (update)"},
                "percent_complete": {"type": "integer", "description": "0-100 (update)"},
                "complete": {"type": "boolean", "description": "Mark complete (true) or not started (false) (mark_complete)"},
            },
            "required": ["action"],
        },
    ),
    # ── 6. Rules (2 actions) ──
    Tool(
        name="outlook_rule",
        description="Manage Outlook inbox rules. Use 'action' to specify the operation.",
        inputSchema={
            "type": "object",
            "properties": {
                "action": {
                    "type": "string",
                    "description": "Action: list, create",
                    "enum": ["list", "create"],
                },
                "name": {"type": "string", "description": "Rule name (create)"},
                "condition_type": {"type": "string", "description": "sender/subject (create)"},
                "condition_value": {"type": "string", "description": "Email or keyword (create)"},
                "action_type": {"type": "string", "description": "move/mark_read/delete/categorize (create)"},
                "action_value": {"type": "string", "description": "Folder or category name (create)"},
                "enabled": {"type": "boolean", "description": "Rule enabled (create)"},
                "account_name": {"type": "string", "description": "Account name"},
            },
            "required": ["action"],
        },
    ),
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
        }, ensure_ascii=False, separators=(",", ":")))]

    except Exception as e:
        logger.exception(f"Error executing tool '{name}'")
        return [TextContent(type="text", text=json.dumps({
            "error": str(e),
            "tool": name
        }, ensure_ascii=False, separators=(",", ":")))]



def _parse_attachments(args: dict[str, Any]) -> list[str] | None:
    """Parse attachments JSON from arguments."""
    if "attachments" in args and args["attachments"]:
        return json.loads(args["attachments"])
    return None


def _execute_tool(name: str, args: dict[str, Any], client: OutlookClient) -> str:
    pythoncom.CoInitialize()
    result: Any = None

    # ── 1. Mailbox ──
    if name == "outlook_mailbox":
        result = client.get_mailbox_info()

    # ── 2. Email ──
    elif name == "outlook_email":
        action = args["action"]

        if action in ("list", "search"):
            folder = FOLDER_MAP.get(args.get("folder", "inbox"), OutlookClient.OL_FOLDER_INBOX)
            count = args.get("count", 20)
            offset = args.get("offset", 0)
            fields = args.get("fields")
            account = args.get("account_name")
            if action == "list":
                result = {"emails": client.list_emails(folder, count, offset, fields, account)}
            else:
                result = {"emails": client.search_emails(
                    folder_type=folder, count=count, offset=offset, fields=fields,
                    account_name=account,
                    subject=args.get("subject"), sender=args.get("sender"),
                    received_after=args.get("received_after"), received_before=args.get("received_before"),
                    unread_only=args.get("unread_only", False),
                )}
        elif action == "get":
            result = client.get_email_by_id(args["entry_id"])
        elif action == "create_draft":
            result = client.create_draft(
                to=args.get("to", ""), subject=args.get("subject", ""), body=args.get("body", ""),
                cc=args.get("cc", ""), bcc=args.get("bcc", ""),
                attachments=_parse_attachments(args), html_body=args.get("html_body", False),
                importance=IMPORTANCE_MAP.get(args.get("importance", "normal"), 1),
            )
        elif action == "delete":
            result = client.delete_email(args["entry_id"])
        elif action == "move":
            dest = FOLDER_MAP.get(args.get("dest_folder", "inbox"), OutlookClient.OL_FOLDER_INBOX)
            result = client.move_email(args["entry_id"], dest)
        elif action == "mark_read":
            result = client.mark_as_read(args["entry_id"], args.get("read", True))
        elif action == "flag":
            result = client.flag_email(args["entry_id"], args.get("flag", True), args.get("due_date"), args.get("reminder_date"))
        elif action == "categorize":
            result = client.categorize_email(args["entry_id"], args.get("categories", ""), args.get("cat_action", "set"))
        elif action == "save_attachment":
            result = client.save_attachment(args["entry_id"], args.get("attachment_index", 1), args.get("save_path", ""))
        elif action == "empty_deleted":
            result = client.empty_deleted_folder(args.get("account_name"))
        elif action == "open":
            result = client.open_email(args["entry_id"])
        elif action == "update_draft":
            result = client.update_draft(
                entry_id=args["entry_id"], subject=args.get("subject"), body=args.get("body"),
                to=args.get("to"), cc=args.get("cc"), bcc=args.get("bcc"),
                html_body=args.get("html_body", False), attachments=_parse_attachments(args),
                importance=IMPORTANCE_MAP.get(args.get("importance"), None),
            )
    # ── 3. Calendar ──
    elif name == "outlook_calendar":
        action = args["action"]

        if action == "list":
            result = {"events": client.list_calendar_events(
                start_date=args.get("start_date"), end_date=args.get("end_date"),
                count=args.get("count", 20), offset=args.get("offset", 0),
                fields=args.get("fields"),
                account_name=args.get("account_name"),
            )}
        elif action == "get":
            result = client.get_appointment_by_id(args["entry_id"])
        elif action == "create":
            result = client.create_appointment(
                subject=args["subject"], start_time=args["start_time"], end_time=args["end_time"],
                body=args.get("body", ""), location=args.get("location", ""),
                all_day=args.get("all_day", False), reminder_minutes=args.get("reminder_minutes", 15),
                recipients=args.get("recipients", ""),
            )
        elif action == "update":
            result = client.update_appointment(
                entry_id=args["entry_id"], subject=args.get("subject"),
                start_time=args.get("start_time"), end_time=args.get("end_time"),
                body=args.get("body"), location=args.get("location"),
                all_day=args.get("all_day"), reminder_minutes=args.get("reminder_minutes"),
                recipients=args.get("recipients"), send_update=args.get("send_update", False),
            )
        elif action == "delete":
            result = client.delete_appointment(args["entry_id"])
        elif action == "respond":
            result = client.respond_to_invitation(
                entry_id=args["entry_id"], response=args["response"], comment=args.get("comment", ""),
            )
        elif action == "freebusy":
            result = client.get_free_busy(
                start_date=args.get("start_date"), months=args.get("months", 1),
                account_name=args.get("account_name"),
            )

    # ── 4. Contacts ──
    elif name == "outlook_contact":
        action = args["action"]

        if action == "list":
            result = {"contacts": client.list_contacts(
                count=args.get("count", 50), offset=args.get("offset", 0),
                fields=args.get("fields"),
                search=args.get("search", ""),
                account_name=args.get("account_name"),
            )}
        elif action == "get":
            result = client.get_contact_by_id(args["entry_id"])
        elif action == "create":
            result = client.create_contact(
                full_name=args["full_name"], email=args.get("email", ""),
                phone=args.get("phone", ""), mobile=args.get("mobile", ""),
                company=args.get("company", ""), job_title=args.get("job_title", ""),
            )
        elif action == "update":
            result = client.update_contact(
                entry_id=args["entry_id"], full_name=args.get("full_name"),
                email=args.get("email"), phone=args.get("phone"),
                mobile=args.get("mobile"), home_phone=args.get("home_phone"),
                company=args.get("company"), job_title=args.get("job_title"),
            )
        elif action == "delete":
            result = client.delete_contact(args["entry_id"])
        elif action == "export":
            result = client.export_contacts(
                format=args.get("format", "csv"), save_path=args.get("save_path", ""),
                account_name=args.get("account_name"),
            )

    # ── 5. Tasks ──
    elif name == "outlook_task":
        action = args["action"]

        if action == "list":
            result = {"tasks": client.list_tasks(
                count=args.get("count", 20), offset=args.get("offset", 0),
                fields=args.get("fields"),
                include_completed=args.get("include_completed", False),
                account_name=args.get("account_name"),
            )}
        elif action == "get":
            result = client.get_task_by_id(args["entry_id"])
        elif action == "create":
            result = client.create_task(
                subject=args["subject"], body=args.get("body", ""),
                due_date=args.get("due_date"), start_date=args.get("start_date"),
                importance=IMPORTANCE_MAP.get(args.get("importance", "normal"), 1),
                reminder_minutes=args.get("reminder_minutes", 0),
            )
        elif action == "update":
            result = client.update_task(
                entry_id=args["entry_id"], subject=args.get("subject"),
                body=args.get("body"), due_date=args.get("due_date"),
                start_date=args.get("start_date"), status=args.get("status"),
                importance=args.get("importance"),
                reminder_minutes=args.get("reminder_minutes"),
                percent_complete=args.get("percent_complete"),
            )
        elif action == "delete":
            result = client.delete_task(args["entry_id"])
        elif action == "mark_complete":
            result = client.mark_task_complete(args["entry_id"], args.get("complete", True))

    # ── 6. Rules ──
    elif name == "outlook_rule":
        action = args["action"]
        if action == "list":
            result = client.get_rules()
        elif action == "create":
            result = client.create_rule(
                name=args["name"], condition_type=args.get("condition_type", "sender"),
                condition_value=args["condition_value"], action_type=args.get("action_type", "move"),
                action_value=args.get("action_value", ""), enabled=args.get("enabled", True),
                account_name=args.get("account_name"),
            )

    else:
        return json.dumps({"error": f"Unknown tool: {name}"}, separators=(",", ":"))

    return json.dumps(result, ensure_ascii=False, default=str, separators=(",", ":"))


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
