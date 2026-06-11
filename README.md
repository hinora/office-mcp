# MCP Servers for Windows Office Automation

This project provides [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) servers for automating Microsoft Windows desktop applications on your local machine. These servers enable AI assistants (like GitHub Copilot, Claude, etc.) to interact with your local Office applications via COM automation.

## Included MCP Servers

| Server | Description |
|---|---|
| **wps-excel-mcp** | Automate WPS Office Excel (create, read, format workbooks, charts, etc.) |
| **outlook-mcp** | Automate Microsoft Outlook (email, calendar, contacts) |

---

## Prerequisites

- **Windows** (COM automation is Windows-only)
- **Python 3.10+**
- Required Python packages:
  - `mcp` — MCP Python SDK
  - `pywin32` — Windows COM automation
- **WPS Office** installed (for wps-excel-mcp)
- **Microsoft Outlook** installed (for outlook-mcp)

## Installation

```bash
# Clone or navigate to the project directory
cd wps-mcp

# Install dependencies
pip install -r requirements.txt

# Or install the package in development mode
pip install -e .
```

## Build Standalone EXEs (No Python Required)

You can build standalone `.exe` files that users can run **without installing Python or any dependencies**. The only requirement on the target machine is the respective Office application installed.

### One-Click Build

```batch
# Build all MCP servers:
build_exe.bat

# Or build individually:
build_exe.bat --wps
build_exe.bat --outlook
```

Or manually:

```bash
pip install pyinstaller
python build_exe.py            # Build both
python build_exe.py --wps      # WPS Excel MCP only
python build_exe.py --outlook  # Outlook MCP only
```

Output:
- `dist\wps-excel-mcp.exe`
- `dist\outlook-mcp.exe`

### Using the EXEs

Add them to your MCP client config:

```json
{
  "mcpServers": {
    "wps-excel-mcp": {
      "command": "C:\\path\\to\\wps-excel-mcp.exe"
    },
    "outlook-mcp": {
      "command": "C:\\path\\to\\outlook-mcp.exe"
    }
  }
}
```

## Configuration (Development Mode)

Add the servers to your MCP client configuration:

### VS Code / GitHub Copilot

Add to your `.vscode/mcp.json`:

```json
{
  "servers": {
    "wps-excel-mcp": {
      "command": "python",
      "args": ["-m", "wps_mcp.server"],
      "cwd": "d:/work/wps-mcp/src"
    },
    "outlook-mcp": {
      "command": "python",
      "args": ["-m", "outlook_mcp.server"],
      "cwd": "d:/work/wps-mcp/src"
    }
  }
}
```

Or use the installed entry points:

```json
{
  "servers": {
    "wps-excel-mcp": {
      "command": "wps-excel-mcp"
    },
    "outlook-mcp": {
      "command": "outlook-mcp"
    }
  }
}
```

---

# WPS Excel MCP (`wps-excel-mcp`)

Automate WPS Office Excel via COM automation.

## Features

| Category | Operations |
|---|---|
| **Workbook** | Create, open, save, close, list |
| **Worksheet** | Add, rename, delete, activate, list |
| **Cells** | Get/set single cell, get/set range, clear |
| **Formatting** | Bold, font size, fill color, alignment, number format, merge/unmerge |
| **Rows/Columns** | Insert, delete, resize (height/width) |
| **Charts** | Add column, line, pie, bar, area, scatter charts |
| **Search** | Find cells by text content |
| **Macros** | Run VBA macros |
| **Window** | Show/hide WPS Excel window |

## Tool Reference

### Workbook Tools

| Tool | Description |
|---|---|
| `wps_get_app_info` | Get WPS Excel version, open workbooks count |
| `wps_create_workbook` | Create a new workbook |
| `wps_open_workbook` | Open an existing `.xlsx`/`.xls`/`.et` file |
| `wps_save_workbook` | Save the active workbook |
| `wps_close_workbook` | Close the active workbook |
| `wps_list_workbooks` | List all open workbooks |
| `wps_show_window` | Show the WPS Excel window |
| `wps_hide_window` | Hide the WPS Excel window |

### Worksheet Tools

| Tool | Description |
|---|---|
| `wps_list_sheets` | List all sheets with names, types, visibility |
| `wps_add_sheet` | Add a new worksheet |
| `wps_rename_sheet` | Rename a worksheet |
| `wps_delete_sheet` | Delete a worksheet |
| `wps_activate_sheet` | Activate (focus) a sheet |

### Cell Tools

| Tool | Description |
|---|---|
| `wps_get_cell_value` | Get a single cell's value |
| `wps_set_cell_value` | Set a single cell's value |
| `wps_get_range_values` | Get a 2D array of values from a range |
| `wps_set_range_values` | Set values in a range |
| `wps_clear_cell` | Clear cell/range contents |
| `wps_find_cell` | Find a cell by text search |

### Formatting Tools

| Tool | Description |
|---|---|
| `wps_set_font_bold` | Set bold on/off |
| `wps_set_font_size` | Set font size |
| `wps_set_cell_color` | Set background/fill color |
| `wps_set_alignment` | Set horizontal alignment |
| `wps_set_number_format` | Set number/date format |
| `wps_merge_cells` | Merge a range |
| `wps_unmerge_cells` | Unmerge a range |

### Row/Column Tools

| Tool | Description |
|---|---|
| `wps_get_used_range` | Get used range address and dimensions |
| `wps_insert_row` | Insert a row |
| `wps_insert_column` | Insert a column |
| `wps_delete_row` | Delete a row |
| `wps_delete_column` | Delete a column |
| `wps_set_row_height` | Set row height |
| `wps_set_column_width` | Set column width |

### Chart & Macro Tools

| Tool | Description |
|---|---|
| `wps_add_chart` | Add a chart (column, line, pie, bar, area, scatter) |
| `wps_run_macro` | Run a VBA macro by name |

## Project Structure

```
wps-mcp/
├── pyproject.toml              # Project metadata & dependencies
├── requirements.txt            # Pip dependencies
├── README.md                   # This file
├── build_exe.py                # Build standalone .exe files
├── build_exe.bat               # Build script wrapper
└── src/
    ├── wps_mcp/                # WPS Excel MCP Server
    │   ├── __init__.py
    │   ├── server.py           # MCP server with tool definitions & handlers
    │   ├── wps_client.py       # WPS Excel COM client
    │   └── tools/
    │       └── __init__.py
    └── outlook_mcp/            # Outlook MCP Server
        ├── __init__.py
        ├── server.py           # MCP server with tool definitions & handlers
        └── outlook_client.py   # Outlook COM client
```

## How It Works

1. Each MCP server runs as a subprocess and communicates via **stdio** (standard input/output) using JSON-RPC.
2. When an AI assistant calls a tool, the server dispatches it to the appropriate handler.
3. The handler calls the COM client which uses `win32com` to automate the target application via its COM interface.
4. Results are serialized to JSON and returned to the AI assistant.

---

# Outlook MCP (`outlook-mcp`)

Automate Microsoft Outlook via COM automation.

## Features

| Category | Operations |
|---|---|
| **Email** | List, search, get details, send, reply, forward, delete, move, mark read |
| **Attachments** | Save email attachments to disk |
| **Calendar** | List events, create appointments, delete |
| **Contacts** | List, search, create, delete |
| **Mailbox** | Get mailbox info (accounts, folder counts, version) |

## Tool Reference

### Mailbox Tools

| Tool | Description |
|---|---|
| `outlook_get_mailbox_info` | Get Outlook version, accounts, folder counts |

### Email Tools

| Tool | Description |
|---|---|
| `outlook_list_emails` | List recent emails from Inbox, Sent, Drafts, or Deleted |
| `outlook_search_emails` | Search emails by subject, sender, date range, read status |
| `outlook_get_email` | Get full details of an email by EntryID (including body) |
| `outlook_send_email` | Send a new email (with optional CC, BCC, attachments, HTML) |
| `outlook_reply_email` | Reply to an email (or Reply All) |
| `outlook_forward_email` | Forward an email to new recipients |
| `outlook_delete_email` | Delete an email by EntryID |
| `outlook_mark_read` | Mark an email as read or unread |
| `outlook_move_email` | Move an email to a different folder |
| `outlook_save_attachment` | Save an attachment from an email to disk |

### Calendar Tools

| Tool | Description |
|---|---|
| `outlook_list_calendar` | List calendar events for a date range |
| `outlook_create_appointment` | Create a new appointment (with optional meeting invitations) |
| `outlook_delete_appointment` | Delete an appointment by EntryID |

### Contacts Tools

| Tool | Description |
|---|---|
| `outlook_list_contacts` | List or search contacts |
| `outlook_create_contact` | Create a new contact |
| `outlook_delete_contact` | Delete a contact by EntryID |

## Usage Examples

### Check your inbox

```
> Show me my 5 most recent emails
```

The assistant calls `outlook_list_emails` with `count=5`.

### Send an email

```
> Send an email to john@example.com with subject "Q3 Report" and
  body "Hi John, please find the Q3 report attached."
```

The assistant calls `outlook_send_email`.

### Search emails

```
> Find all unread emails from "Jane" this week
```

The assistant calls `outlook_search_emails` with `sender="Jane"`, `unread_only=true`, and a date range.

### Check calendar

```
> What's on my calendar for next Monday?
```

The assistant calls `outlook_list_calendar` with appropriate start/end dates.

## COM ProgID

The client uses `Outlook.Application` as the COM ProgID. If Outlook is already running, it connects to the existing instance.

## Troubleshooting

### "Could not connect to Microsoft Outlook"

- Ensure Microsoft Outlook is installed and running
- Verify Outlook is properly registered (the COM interface should be available)
- If using a 64-bit version of Python, ensure a 64-bit version of Outlook is installed (or vice versa)

### "ModuleNotFoundError: No module named 'win32com'"

```bash
pip install pywin32
```

### COM errors / crashes

- Ensure Outlook is up to date
- Some operations may trigger Outlook security prompts (Outlook's security model restricts programmatic access)

## License

MIT
