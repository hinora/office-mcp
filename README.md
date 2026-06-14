# MCP Servers for Windows Office Automation

This project provides [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) servers for automating Microsoft Windows desktop applications on your local machine. These servers enable AI assistants (like GitHub Copilot, Claude, etc.) to interact with your local Office applications via COM automation.

## Included MCP Servers

| Server | Description |
|---|---|
| **wps-excel-mcp** | Automate WPS Office Excel (create, read, format workbooks, charts, etc.) |
| **wps-word-mcp** | Automate WPS Office Word (create, edit, format documents, tables, etc.) |
| **outlook-mcp** | Automate Microsoft Outlook (email, calendar, contacts) |

---

## Prerequisites

- **Windows** (COM automation is Windows-only)
- **Python 3.10+**
- Required Python packages:
  - `mcp` — MCP Python SDK
  - `pywin32` — Windows COM automation
- **WPS Office** installed (for wps-excel-mcp, wps-word-mcp)
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
build_exe.bat --word
build_exe.bat --outlook
```

Or manually:

```bash
pip install pyinstaller
python build_exe.py            # Build all
python build_exe.py --wps      # WPS Excel MCP only
python build_exe.py --word     # WPS Word MCP only
python build_exe.py --outlook  # Outlook MCP only
```

Output:
- `dist\wps-excel-mcp.exe`
- `dist\wps-word-mcp.exe`
- `dist\outlook-mcp.exe`

### Using the EXEs

Add them to your MCP client config:

```json
{
  "mcpServers": {
    "wps-excel-mcp": {
      "command": "C:\\path\\to\\wps-excel-mcp.exe"
    },
    "wps-word-mcp": {
      "command": "C:\\path\\to\\wps-word-mcp.exe"
    },
    "outlook-mcp": {
      "command": "C:\\path\\to\\outlook-mcp.exe"
    }
  }
}
```

### Pre-built Releases

Every GitHub Release automatically builds a **`wps-mcp.zip`** containing all three MCP servers as standalone `.exe` files, plus a setup guide. No Python or build tools required — just download, extract, and configure your MCP client.

**Download the latest release**: [Releases](https://github.com/hinora/wps-mcp/releases)

The zip includes:
- `wps-excel-mcp.exe` — WPS Excel MCP server
- `wps-word-mcp.exe` — WPS Word MCP server
- `outlook-mcp.exe` — Outlook MCP server
- `README_SETUP.txt` — Step-by-step setup guide

## Configuration (Development Mode)

Add the servers to your MCP client configuration:

### VS Code / GitHub Copilot

Add to your `.vscode/mcp.json`:

```json
{
  "servers": {
    "wps-excel-mcp": {
      "command": "python",
      "args": ["-m", "wps_excel_mcp.server"],
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
| **Workbook** | Create, open, save, close, list, activate |
| **Worksheet** | Add, rename, delete, activate, copy, move, hide/unhide |
| **Cells** | Get/set/clear values, get/set formulas |
| **Ranges** | Get/set/clear, copy/paste, sort, find/replace, used range, remove duplicates |
| **Formatting** | Bold, italic, font name/size/color, underline, fill color, alignment, vertical alignment, number format, wrap text, merge/unmerge, borders, conditional formatting, data validation |
| **Rows/Columns** | Insert, delete, resize, autofit, group/ungroup, freeze/unfreeze panes |
| **Charts** | Add column, line, pie, bar, area, scatter charts |
| **Data** | AutoFilter, text-to-columns |
| **Protection** | Protect/unprotect sheets |
| **Page Setup** | Print area, orientation, margins, headers/footers |
| **Export** | Export to PDF |
| **Misc** | Named ranges, hyperlinks, comments, pivot tables, sparklines, pictures, shapes, gridlines |
| **Macros** | Run VBA macros |
| **Window** | Show/hide WPS Excel window, get app info |

## Tool Reference

All tools use a unified `action`-based API. Each tool accepts an `action` parameter plus action-specific parameters and an optional `sheet_name`.

### `wps_app` — Application Control
| Action | Description |
|---|---|
| `info` | Get WPS Excel version and open workbooks count |
| `show` | Show the WPS Excel window |
| `hide` | Hide the WPS Excel window |

### `wps_workbook` — Workbook Management
| Action | Description |
|---|---|
| `create` | Create a new workbook |
| `open` | Open a file (`filepath`) |
| `save` | Save the active workbook (optional `filepath`) |
| `close` | Close the active workbook (`save`: bool) |
| `list` | List all open workbooks |
| `activate` | Activate a workbook by `name` |

### `wps_sheet` — Worksheet Management
| Action | Description |
|---|---|
| `list` | List all sheets with names, types, visibility |
| `add` | Add a new worksheet (optional `name`) |
| `rename` | Rename a sheet (`old_name` → `new_name`) |
| `delete` | Delete a sheet by `name` |
| `activate` | Activate (focus) a sheet by `name` |
| `copy` | Copy a sheet (`source_name`, optional `new_name`, `before`, `after`) |
| `move` | Move a sheet (`source_name`, optional `before`, `after`) |
| `hide` | Hide a sheet by `name` |
| `unhide` | Unhide a sheet by `name` |

### `wps_cell` — Single Cell Operations
| Action | Description |
|---|---|
| `get` | Get a cell's value (`cell_ref`) |
| `set` | Set a cell's value (`cell_ref`, `value`) |
| `clear` | Clear a cell's contents (`cell_ref`) |
| `set_formula` | Set a formula (`cell_ref`, `formula`, e.g. `=SUM(B2:B10)`) |
| `get_formula` | Get a cell's formula (`cell_ref`) |

### `wps_range` — Multi-Cell Range Operations
| Action | Description |
|---|---|
| `get` | Get values from a `range_ref` as a 2D array |
| `set` | Set values at `start_cell` (`values` as JSON 2D array) |
| `clear` | Clear a range |
| `copy` | Copy a `range_ref` |
| `paste` | Paste to `dest_cell` (optional `paste_special`: values/formats/formulas/all/transpose) |
| `sort` | Sort a range by `sort_key` column |
| `find` | Find first occurrence of `search_text` |
| `find_next` | Find next occurrence |
| `find_replace` | Find `find_text` and replace with `replace_text` |
| `used_range` | Get used range address, row count, column count |
| `remove_duplicates` | Remove duplicate rows (`columns` as JSON array) |

### `wps_format` — Cell Formatting
Apply multiple formats in a single call. All parameters apply to `cell_or_range`.

| Parameter | Description |
|---|---|
| `bold` | Bold on/off (bool) |
| `italic` | Italic on/off (bool) |
| `font_name` | Font family name |
| `font_size` | Font size (integer) |
| `font_color` | Font color (RGB hex, e.g. `FF0000`) |
| `underline` | Underline style: none/single/double |
| `fill_color` | Background/fill color (RGB hex) |
| `alignment` | Horizontal alignment: left/center/right |
| `vertical_alignment` | Vertical alignment: top/center/bottom |
| `number_format` | Number/date format string |
| `wrap_text` | Text wrap on/off (bool) |
| `merge` | True=merge, False=unmerge |
| `border_style` | Border style: thin/medium/thick |
| `border_color` | Border color (RGB hex) |
| `outline_only` | Apply border to outline only (bool) |
| `cond_operator` | Conditional format operator |
| `cond_formula` | Conditional format formula/value |
| `cond_font_color` | Conditional format font color |
| `cond_bg_color` | Conditional format background color |
| `cond_bold` | Conditional format bold (bool) |
| `clear_cond` | Clear conditional formats (bool) |
| `validation_type` | Data validation type |
| `formula1` / `formula2` | Data validation formulas |
| `ignore_blank` | Validation: ignore blank (bool) |
| `show_dropdown` | Validation: show dropdown (bool) |
| `error_title` / `error_message` | Validation error alert |

### `wps_rowcol` — Row/Column Operations
| Action | Description |
|---|---|
| `insert` | Insert a `type` (row/column) at `position` |
| `delete` | Delete a `type` at `position` |
| `resize` | Set row height or column width (`type`, `position`, `size`) |
| `autofit` | Auto-fit rows or columns (`type`, optional `start`/`end`) |
| `group` | Group rows/columns (`start` to `end`) |
| `ungroup` | Ungroup rows/columns (`start` to `end`) |
| `freeze` | Freeze panes at `cell_ref` (default: B2) |
| `unfreeze` | Unfreeze panes |

### `wps_chart` — Chart Creation
| Parameter | Description |
|---|---|
| `chart_type` | column/line/pie/bar/area/scatter |
| `range_ref` | Data range for the chart |
| `left`, `top`, `width`, `height` | Chart position and size |

### `wps_data` — Data Operations
| Action | Description |
|---|---|
| `filter` | Apply AutoFilter to a `range_ref` |
| `text_to_columns` | Split text in `range_ref` by `delimiter` |

### `wps_protection` — Sheet Protection
| Action | Description |
|---|---|
| `protect` | Protect the sheet (optional `password`, `allow_sort`, `allow_filter`, etc.) |
| `unprotect` | Unprotect the sheet (optional `password`) |

### `wps_page_setup` — Page Layout
| Action | Description |
|---|---|
| `print_area` | Set print area to `range_ref` |
| `clear_print_area` | Clear print area |
| `orientation` | Set page orientation |
| `margins` | Set margins (`left`, `right`, `top`, `bottom`) |
| `header_footer` | Set headers/footers (`left_header`, `center_header`, `right_header`, `left_footer`, `center_footer`, `right_footer`) |

### `wps_export` — Export
| Parameter | Description |
|---|---|
| `filepath` | Output PDF file path |
| `sheet_name` | Optional sheet name |

### `wps_macro` — Macros
| Parameter | Description |
|---|---|
| `macro_name` | Name of the VBA macro to run |

### `wps_misc` — Miscellaneous Operations
| Action | Description |
|---|---|
| `create_named_range` | Create a named range (`name`, `refers_to`) |
| `delete_named_range` | Delete a named range |
| `list_named_ranges` | List all named ranges |
| `add_hyperlink` | Add a hyperlink (`cell_ref`, `address`, `text_to_display`, `screen_tip`) |
| `remove_hyperlink` | Remove hyperlinks from `cell_or_range` |
| `add_comment` | Add a comment to `cell_ref` (`text`) |
| `delete_comment` | Delete comment from `cell_ref` |
| `create_pivot_table` | Create a pivot table |
| `add_sparkline` | Add a sparkline chart |
| `insert_picture` | Insert a picture (`filepath`, `left`, `top`, `width`, `height`) |
| `insert_shape` | Insert a shape (`shape_type`) |
| `toggle_gridlines` | Show/hide gridlines (`visible`: bool) |

## Project Structure

```
wps-mcp/
├── pyproject.toml              # Project metadata & dependencies
├── requirements.txt            # Pip dependencies
├── README.md                   # This file
├── build_exe.py                # Build standalone .exe files
├── build_exe.bat               # Build script wrapper
└── src/
    ├── wps_excel_mcp/          # WPS Excel MCP Server
    │   ├── __init__.py
    │   ├── server.py           # MCP server with tool definitions & handlers
    │   ├── wps_client.py       # WPS Excel COM client
    │   └── tools/
    │       └── __init__.py
    ├── wps_word_mcp/           # WPS Word MCP Server
    │   ├── __init__.py
    │   ├── server.py           # MCP server with tool definitions & handlers
    │   ├── word_client.py      # WPS Word COM client
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

# WPS Word MCP (`wps-word-mcp`)

Automate WPS Office Word via COM automation.

## Features

| Category | Operations |
|---|---|
| **Document** | Create, open, save, close, list, activate, get/set properties, protect/unprotect |
| **Text** | Get, set, type, append, prepend, get selected text |
| **Paragraphs** | Add, get count, get/set text, insert before, delete, alignment, spacing |
| **Font** | Bold, italic, underline, font name/size/color, highlighting |
| **Find/Replace** | Find text, find & replace (with match case, whole word, replace all) |
| **Tables** | Add, count, get data, set cell, add row/column, delete, style |
| **Page Layout** | Orientation, margins, size, columns, borders, headers, footers, page numbers |
| **Styles** | Apply style, bullet/number lists, remove list |
| **Insert** | Picture, page break, section break, hyperlink, table of contents, bookmark |
| **Export** | Export to PDF, print |
| **Other** | Watermark, track changes, zoom, comments, range text |

## Tool Reference

All tools use a unified `action`-based API.

### `word_app` — Application Control
| Action | Description |
|---|---|
| `info` | Get WPS Word version and open documents count |
| `show` | Show the WPS Word window |
| `hide` | Hide the WPS Word window |
| `quit` | Quit WPS Word |

### `word_document` — Document Management
| Action | Description |
|---|---|
| `create` | Create a new document |
| `open` | Open a file (`filepath`) |
| `save` | Save the active document (optional `filepath`) |
| `close` | Close the active document (`save`: bool) |
| `list` | List all open documents |
| `activate` | Activate a document by `name` |
| `get_properties` | Get document properties (author, title, etc.) |
| `set_properties` | Set document properties (`author`, `title`, `subject`, `keywords`) |
| `protect` | Protect the document with optional `password` |
| `unprotect` | Unprotect the document with optional `password` |

### `word_text` — Text Operations
| Action | Description |
|---|---|
| `get` | Get all document text |
| `set` | Replace all text (`text`) |
| `type` | Type text at cursor position |
| `append` | Append text at end of document |
| `prepend` | Prepend text at start of document |
| `get_selected` | Get currently selected text |

### `word_paragraph` — Paragraph Operations
| Action | Description |
|---|---|
| `add` | Add a paragraph (optional `text`) |
| `get_count` | Get total paragraph count |
| `get_text` | Get text of paragraph by `index` (1-based) |
| `set_text` | Set text of paragraph by `index` |
| `insert_before` | Insert paragraph before `index` |
| `delete` | Delete paragraph by `index` |
| `alignment` | Set alignment (`alignment`: left/center/right/justify) |
| `spacing` | Set paragraph spacing (`before`, `after`, `line_spacing`) |

### `word_font` — Font Formatting
Apply to a `range_spec`: selection (default), content, or `start=X,end=Y`.

| Parameter | Description |
|---|---|
| `bold` | Bold on/off (bool) |
| `italic` | Italic on/off (bool) |
| `underline` | Underline on/off (bool) |
| `font_name` | Font family name |
| `font_size` | Font size (number) |
| `font_color` | Font color (RGB hex) |
| `highlight_index` | Highlight color (0=None, 6=Yellow, 7=Green, 2=Blue) |

### `word_find` — Find & Replace
| Action | Description |
|---|---|
| `find` | Find `search_text` (`match_case`, `match_whole_word`) |
| `find_replace` | Find `find_text` and replace with `replace_text` (`replace_all`: bool) |

### `word_table` — Table Operations
| Action | Description |
|---|---|
| `add` | Add a table (`rows`, `cols`, optional `text`) |
| `count` | Get table count |
| `get_data` | Get table data by `index` (1-based) |
| `set_cell` | Set cell text (`table_index`, `row`, `col`, `text`) |
| `add_row` | Add a row to table |
| `add_col` | Add a column to table |
| `delete` | Delete table by `index` |
| `style` | Apply table `style_name` |

### `word_page` — Page Layout
| Action | Description |
|---|---|
| `orientation` | Set orientation (`orientation`: portrait/landscape) |
| `margins` | Set margins (`left`, `right`, `top`, `bottom`) |
| `size` | Set page size (`width`, `height`) |
| `columns` | Set columns (`num_columns`, `spacing`) |
| `borders` | Set page borders (`line_style`, `line_width`, `distance`) |
| `header` | Add header (`text`) |
| `footer` | Add footer (`text`) |
| `page_numbers` | Insert page numbers (`position`: bottom/top) |

### `word_insert` — Insert Elements
| Action | Description |
|---|---|
| `picture` | Insert a picture (`filepath`, `left`, `top`, `width`, `height`) |
| `page_break` | Insert a page break |
| `section_break` | Add a section break |
| `hyperlink` | Add a hyperlink (`address`, `text_to_display`) |
| `toc` | Insert table of contents |
| `bookmark` | Add bookmark (`name`) |
| `goto_bookmark` | Navigate to bookmark (`name`) |

### `word_export` — Export & Print
| Action | Description |
|---|---|
| `pdf` | Export to PDF (`filepath`) |
| `print` | Print document (optional `copies`) |

### Other Word Tools
| Tool | Description |
|---|---|
| `word_style` | Apply styles: `apply_style` (`style_name`), `list` (bullet/number), `remove_list` |
| `word_watermark` | Add text watermark (`text`, `font_size`, `color`, `layout`: diagonal/horizontal) |
| `word_track_changes` | Toggle track changes (`enable`: bool) |
| `word_zoom` | Set zoom level (`percentage`) |
| `word_comment` | Add a comment (`text`, optional `range_spec`) |
| `word_range` | Get text from a character range (`start`, `end`, 0-based) |

---

# Outlook MCP (`outlook-mcp`)

Automate Microsoft Outlook via COM automation.

## Features

| Category | Operations |
|---|---|
| **Mailbox** | Get Outlook version, accounts, folder counts |
| **Email** | List, search, get details, create/update drafts, delete, move, mark read, flag, categorize, save attachments, empty deleted, open in window |
| **Calendar** | List, get, create, update, delete appointments, respond to invitations, get free/busy |
| **Contacts** | List, get, create, update, delete, export to CSV/vCard |
| **Tasks** | List, get, create, update, delete, mark complete |
| **Rules** | List inbox rules, create rules |

## Tool Reference

All tools use a unified `action`-based API. Each tool accepts an `action` parameter plus action-specific parameters.

### `outlook_mailbox` — Mailbox Info
Get Outlook version, accounts, and folder counts. No parameters required.

### `outlook_email` — Email Management

| Action | Description |
|---|---|
| `list` | List emails from a folder (`folder`: inbox/sent/drafts/deleted; `count`, `offset`, `fields`, `account_name`) |
| `search` | Search emails (`subject`, `sender`, `received_after`, `received_before`, `unread_only`, etc.) |
| `get` | Get full email details by `entry_id` (includes body) |
| `create_draft` | Create a draft email (`to`, `subject`, `body`, `cc`, `bcc`, `attachments`, `html_body`, `importance`) |
| `update_draft` | Update an existing draft (`entry_id`, and any of `subject`, `body`, `to`, `cc`, `bcc`, `html_body`, `attachments`, `importance`) |
| `delete` | Delete an email by `entry_id` |
| `move` | Move an email to `dest_folder` (inbox/sent/drafts/deleted) |
| `mark_read` | Mark as read or unread (`read`: bool) |
| `flag` | Flag/unflag an email (`flag`: bool, optional `due_date`, `reminder_date`) |
| `categorize` | Set categories (`categories`, `cat_action`: set/add/remove/clear) |
| `save_attachment` | Save an attachment to disk (`attachment_index`, `save_path`) |
| `empty_deleted` | Empty the Deleted Items folder (optional `account_name`) |
| `open` | Open an email in its own window by `entry_id` |

> **Note:** To send an email, use `create_draft` to create it, then `update_draft` to modify it, and the draft can be sent manually from Outlook. The MCP uses a draft-based workflow instead of direct sending for security.

### `outlook_calendar` — Calendar Management

| Action | Description |
|---|---|
| `list` | List calendar events (`start_date`, `end_date`, `count`, `offset`, `fields`, `account_name`) |
| `get` | Get appointment details by `entry_id` |
| `create` | Create an appointment (`subject`, `start_time`, `end_time`, `body`, `location`, `all_day`, `reminder_minutes`, `recipients`) |
| `update` | Update an appointment (`entry_id`, plus any fields; `send_update`: bool) |
| `delete` | Delete an appointment by `entry_id` |
| `respond` | Respond to an invitation (`response`: accept/decline/tentative, optional `comment`) |
| `freebusy` | Get free/busy information (`start_date`, `months`, `account_name`) |

### `outlook_contact` — Contacts Management

| Action | Description |
|---|---|
| `list` | List contacts (`count`, `offset`, `fields`, `search`, `account_name`) |
| `get` | Get contact details by `entry_id` |
| `create` | Create a contact (`full_name`, `email`, `phone`, `mobile`, `company`, `job_title`) |
| `update` | Update a contact (`entry_id`, plus any fields; `home_phone` also supported) |
| `delete` | Delete a contact by `entry_id` |
| `export` | Export contacts to file (`format`: csv/vcard, `save_path`, `account_name`) |

### `outlook_task` — Tasks Management

| Action | Description |
|---|---|
| `list` | List tasks (`count`, `offset`, `fields`, `include_completed`, `account_name`) |
| `get` | Get task details by `entry_id` |
| `create` | Create a task (`subject`, `body`, `due_date`, `start_date`, `importance`, `reminder_minutes`) |
| `update` | Update a task (`entry_id`, plus any fields; `status`, `percent_complete` also supported) |
| `delete` | Delete a task by `entry_id` |
| `mark_complete` | Mark task complete or not started (`complete`: bool) |

### `outlook_rule` — Rules Management

| Action | Description |
|---|---|
| `list` | List all inbox rules |
| `create` | Create a rule (`name`, `condition_type`: sender/subject, `condition_value`, `action_type`: move/mark_read/delete/categorize, `action_value`, `enabled`, `account_name`) |

## Usage Examples

### Check your inbox

```
> Show me my 5 most recent emails
```

The assistant calls `outlook_email` with `action=list`, `count=5`.

### Create a draft email

```
> Create a draft email to john@example.com with subject "Q3 Report" and
  body "Hi John, please find the Q3 report attached."
```

The assistant calls `outlook_email` with `action=create_draft`.

### Search emails

```
> Find all unread emails from "Jane" this week
```

The assistant calls `outlook_email` with `action=search`, `sender="Jane"`, `unread_only=true`, and a date range.

### Check calendar

```
> What's on my calendar for next Monday?
```

The assistant calls `outlook_calendar` with `action=list` and appropriate start/end dates.

### Manage tasks

```
> Show me my incomplete tasks
```

The assistant calls `outlook_task` with `action=list`, `include_completed=false`.

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
