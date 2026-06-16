# MCP Servers for Windows Office Automation

This project provides [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) servers for automating Microsoft Windows desktop applications on your local machine. These servers enable AI assistants (like GitHub Copilot, Claude, etc.) to interact with your local Office applications via COM automation.

## Included MCP Servers

| Server | Description |
|---|---|
| **wps-excel-mcp** | Automate WPS Office Excel (create, read, format workbooks, charts, etc.) |
| **wps-word-mcp** | Automate WPS Office Word (create, edit, format documents, tables, etc.) |
| **wps-slide-mcp** | Automate WPS Office Slide (create, edit presentations, shapes, slide shows, etc.) |
| **outlook-mcp** | Automate Microsoft Outlook (email, calendar, contacts) |
| **whatsapp-mcp** | Read WhatsApp Web data (chats, messages, contacts, media) via Chrome DevTools Protocol (CDP) |
| **mcp-meta** | Web tools: search the web (Brave API) and fetch web page content |

---

## Prerequisites

- **Windows** (COM automation is Windows-only)
- **Python 3.10+**
- Required Python packages:
  - `mcp` — MCP Python SDK
  - `pywin32` — Windows COM automation
  - `websockets` — WebSocket client (for WhatsApp CDP)
- **WPS Office** installed (for wps-excel-mcp, wps-word-mcp, wps-slide-mcp)
- **Microsoft Outlook** installed (for outlook-mcp)
- **Google Chrome** with remote debugging and **WhatsApp Web** logged in (for whatsapp-mcp)

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
build_exe.bat --slide
build_exe.bat --outlook
build_exe.bat --meta
build_exe.bat --whatsapp
```

Or manually:

```bash
pip install pyinstaller
python build_exe.py            # Build all
python build_exe.py --wps      # WPS Excel MCP only
python build_exe.py --word     # WPS Word MCP only
python build_exe.py --slide    # WPS Slide MCP only
python build_exe.py --outlook  # Outlook MCP only
python build_exe.py --meta     # MCP Meta only
python build_exe.py --whatsapp # WhatsApp MCP only
```

Output:
- `dist\wps-excel-mcp.exe`
- `dist\wps-word-mcp.exe`
- `dist\wps-slide-mcp.exe`
- `dist\outlook-mcp.exe`
- `dist\mcp-meta.exe`
- `dist\whatsapp-mcp.exe`

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
    "wps-slide-mcp": {
      "command": "C:\\path\\to\\wps-slide-mcp.exe"
    },
    "outlook-mcp": {
      "command": "C:\\path\\to\\outlook-mcp.exe"
    },
    "mcp-meta": {
      "command": "C:\\path\\to\\mcp-meta.exe",
      "env": {
        "BRAVE_API_KEY": "your-brave-api-key-here"
      }
    },
    "whatsapp-mcp": {
      "command": "C:\\path\\to\\whatsapp-mcp.exe"
    }
  }
}
```

### Pre-built Releases

Every GitHub Release automatically builds a **`wps-mcp.zip`** containing all six MCP servers as standalone `.exe` files, plus a setup guide. No Python or build tools required — just download, extract, and configure your MCP client.

**Download the latest release**: [Releases](https://github.com/hinora/office-mcp/releases)

The zip includes:
- `wps-excel-mcp.exe` — WPS Excel MCP server
- `wps-word-mcp.exe` — WPS Word MCP server
- `wps-slide-mcp.exe` — WPS Slide MCP server
- `outlook-mcp.exe` — Outlook MCP server
- `mcp-meta.exe` — MCP Meta server (web search, web fetch)
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
    "wps-word-mcp": {
      "command": "python",
      "args": ["-m", "wps_word_mcp.server"],
      "cwd": "d:/work/wps-mcp/src"
    },
    "wps-slide-mcp": {
      "command": "python",
      "args": ["-m", "wps_slide_mcp.server"],
      "cwd": "d:/work/wps-mcp/src"
    },
    "outlook-mcp": {
      "command": "python",
      "args": ["-m", "outlook_mcp.server"],
      "cwd": "d:/work/wps-mcp/src"
    },
    "mcp-meta": {
      "command": "python",
      "args": ["-m", "mcp_meta.server"],
      "cwd": "d:/work/wps-mcp/src",
      "env": {
        "BRAVE_API_KEY": "your-brave-api-key-here"
      }
    },
    "whatsapp-mcp": {
      "command": "python",
      "args": ["-m", "whatsapp_mcp.server"],
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
    "wps-word-mcp": {
      "command": "wps-word-mcp"
    },
    "wps-slide-mcp": {
      "command": "wps-slide-mcp"
    },
    "outlook-mcp": {
      "command": "outlook-mcp"
    },
    "mcp-meta": {
      "command": "mcp-meta",
      "env": {
        "BRAVE_API_KEY": "your-brave-api-key-here"
      }
    },
    "whatsapp-mcp": {
      "command": "whatsapp-mcp"
    }
  }
}
```

---

# MCP Meta (`mcp-meta`)

A lightweight MCP server providing **web search** and **web fetch** capabilities for AI assistants.

## Features

| Tool | Description |
|---|---|
| **web-search** | Search the web via configurable search providers (Brave Search API supported) |
| **web-fetch** | Fetch and extract readable text content from any URL |

## Configuration

### Environment Variables

| Variable | Required | Default | Description |
|---|---|---|---|
| `SEARCH_PROVIDER` | No | `brave` | Search provider to use (only `brave` currently) |
| `BRAVE_API_KEY` | Yes (for web-search) | — | Brave Search API key (get one free at https://brave.com/search/api/) |

### Prerequisites

- **Python 3.10+**
- Required packages: `mcp`
- **Brave Search API key** for web-search functionality

## Tool Reference

### `web-search` — Web Search

Search the web using Brave Search API.

| Parameter | Type | Required | Description |
|---|---|---|---|
| `query` | string | Yes | Search query string |
| `count` | integer | No | Number of results (default: 10, max: 20) |
| `offset` | integer | No | Pagination offset (default: 0) |
| `country` | string | No | Two-letter country code (e.g. US, DE, FR) |
| `search_lang` | string | No | Search language code (e.g. en, de, fr) |
| `freshness` | string | No | Filter by recency: `pd` (past day), `pw` (past week), `pm` (past month), `py` (past year) |

**Response fields:** `provider`, `query`, `total_results`, `results[]` (each with `title`, `url`, `description`, `age`, `language`).

### `web-fetch` — Web Fetch

Fetch and extract readable text from a URL. Strips HTML tags, scripts, styles, and navigation elements.

| Parameter | Type | Required | Description |
|---|---|---|---|
| `url` | string | Yes | The URL to fetch content from |
| `max_length` | integer | No | Max characters of content returned (default: 10000) |
| `raw` | boolean | No | Return raw HTML instead of extracted text (default: false) |

**Response fields:** `url`, `content_type`, `content`, `length`.

### Usage Examples

```
> Search the web for "latest TypeScript 5.5 features"
```
The assistant calls `web-search` with `query="latest TypeScript 5.5 features"`.

```
> Fetch and summarize the content from https://example.com/article
```
The assistant calls `web-fetch` with `url="https://example.com/article"`.

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
    ├── wps_slide_mcp/          # WPS Slide MCP Server
    │   ├── __init__.py
    │   ├── server.py           # MCP server with tool definitions & handlers
    │   ├── slide_client.py     # WPS Slide COM client
    │   └── tools/
    │       └── __init__.py
    ├── outlook_mcp/            # Outlook MCP Server
    │   ├── __init__.py
    │   ├── server.py           # MCP server with tool definitions & handlers
    │   └── outlook_client.py   # Outlook COM client
    ├── mcp_meta/               # MCP Meta Server (web tools)
    │   ├── __init__.py
    │   ├── server.py           # MCP server: web-search, web-fetch
    │   └── tools/
    │       └── __init__.py
    └── whatsapp_mcp/           # WhatsApp MCP Server
        ├── __init__.py
        ├── server.py           # MCP server: WhatsApp Web bridge
        ├── cdp_client.py       # Chrome DevTools Protocol client
        └── tools/
            └── __init__.py
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

# WPS Slide MCP (`wps-slide-mcp`)

Automate WPS Office Slide (Presentation) via COM automation.

## Features

| Category | Operations |
|---|---|
| **Presentation** | Create, open, save, close, list, activate, get properties, set slide size |
| **Slides** | Add, delete, duplicate, move, go to, count, list, set background, set transition |
| **Shapes** | Text boxes, pictures, rectangles, ovals, arrows, lines, delete, list, position, fill, line, rotation, z-order, group/ungroup, copy, paste, duplicate |
| **Text** | Get/set shape text |
| **Font** | Bold, italic, underline, font name/size/color, alignment |
| **Tables** | Add tables, set cell text, get table data |
| **Notes** | Get/set speaker notes per slide |
| **Export** | Export to PDF, export single slide as image |
| **Slide Show** | Start, start from slide, stop |
| **Animations** | Add/clear shape animations (20+ effect types) |
| **Find/Replace** | Find/replace text across all slides |
| **Master/Layout** | Slide master info, apply layouts, set master background |
| **Advanced** | Hyperlinks, charts, video/audio media, headers/footers, slide numbers |

## Tool Reference

All 16 tools use a unified `action`-based API. `slide_index` is 1-based; use 0 for the currently active slide.

### `app` — Application Control
| Action | Description |
|---|---|
| `info` | Get WPS Slide version and open presentations count |
| `show` | Show the WPS Slide window |
| `hide` | Hide the WPS Slide window |
| `quit` | Quit WPS Slide |

### `pres` — Presentation Management
| Action | Description |
|---|---|
| `create` | Create a new presentation |
| `open` | Open a file (`filepath`) |
| `save` | Save the active presentation (optional `filepath`) |
| `close` | Close the active presentation (`save`: bool) |
| `list` | List all open presentations |
| `activate` | Activate a presentation by `name` |
| `get_properties` | Get presentation properties (name, slides count, slide size) |
| `set_slide_size` | Set slide dimensions (`width`, `height` in points) |

### `slide` — Slide Operations
| Action | Description |
|---|---|
| `add` | Add a slide (optional `layout_index`: 1=Title, 2=Content, etc.) |
| `delete` | Delete a slide by `index` (1-based) |
| `duplicate` | Duplicate a slide by `index` |
| `move` | Move a slide from `index` to `to_index` |
| `go_to` | Navigate to slide by `index` |
| `count` | Get total slide count |
| `list` | List all slides with index, name, and shape count |
| `set_background` | Set solid background color (`color_rgb` as hex) |
| `set_transition` | Set transition effect (`transition_type`: 1=Fade, 2=Push, 8=Wipe, 25=Zoom, `duration` in seconds) |

### `shape_add` — Add Shapes
| Action | Description |
|---|---|
| `text_box` | Add a text box (`text`, `left`, `top`, `width`, `height`, `slide_index`) |
| `picture` | Add a picture (`filepath`, `left`, `top`, `width`, `height`, `slide_index`) |
| `rectangle` | Add a rectangle (`left`, `top`, `width`, `height`, `slide_index`) |
| `oval` | Add an oval/ellipse (`left`, `top`, `width`, `height`, `slide_index`) |
| `arrow` | Add a right arrow (`left`, `top`, `width`, `height`, `slide_index`) |
| `line` | Add a line (`begin_x`, `begin_y`, `end_x`, `end_y`, `slide_index`) |

### `shape_format` — Format Shapes
| Action | Description |
|---|---|
| `set_position` | Set position/size (`name_or_index`, `left`, `top`, `width`, `height`) |
| `set_fill` | Set fill color (`name_or_index`, `color_rgb` as hex) |
| `set_line` | Set outline (`name_or_index`, `color_rgb`, `weight`) |
| `set_rotation` | Set rotation in degrees (`name_or_index`, `rotation`) |
| `set_zorder` | Set z-order: front/back/forward/backward (`name_or_index`, `zorder`) |

### `shape_organize` — Organize Shapes
| Action | Description |
|---|---|
| `group` | Group shapes (`names` as JSON array of shape names/indices) |
| `ungroup` | Ungroup a grouped shape (`name_or_index`) |
| `copy` | Copy a shape to another slide (`name_or_index`, `dest_slide_index`) |
| `paste` | Paste copied shape (`dest_slide_index`) |
| `duplicate` | Duplicate a shape (`name_or_index`) |
| `delete` | Delete a shape by `name_or_index` |
| `list` | List all shapes on a slide (`slide_index`) |
| `count` | Get shape count on a slide (`slide_index`) |

> **Note:** `slide_index=0` targets the currently active/visible slide. Use `slide` with `action=go_to` to navigate first.

### `text` — Text Operations
| Action | Description |
|---|---|
| `get` | Get text from a shape (`name_or_index`) |
| `set` | Set shape text (`name_or_index`, `text`) |

### `font` — Font Formatting
Apply to a shape (`name_or_index`). Use `start`/`length` for partial text formatting.

| Parameter | Description |
|---|---|
| `bold` | Bold on/off (bool) |
| `italic` | Italic on/off (bool) |
| `underline` | Underline on/off (bool) |
| `font_name` | Font family name |
| `font_size` | Font size (number) |
| `font_color` | Font color (hex RGB, e.g. FF0000) |
| `alignment` | Text alignment: left/center/right/justify |

### `table` — Table Operations
| Action | Description |
|---|---|
| `add` | Add a table (`rows`, `cols`, `left`, `top`, `width`, `height`, `slide_index`) |
| `set_cell` | Set cell text (`shape_name`, `row`, `col`, `text`) |
| `get_data` | Get all table data as 2D array (`shape_name`) |

### `notes` — Speaker Notes
| Action | Description |
|---|---|
| `get` | Get speaker notes for a slide (`slide_index`) |
| `set` | Set speaker notes (`slide_index`, `text`) |

### `export` — Export
| Action | Description |
|---|---|
| `pdf` | Export presentation to PDF (`filepath`) |
| `slide_image` | Export a single slide as image (`filepath`, `slide_index`, `img_width`, `img_height`) |

### `slideshow` — Slide Show Control
| Action | Description |
|---|---|
| `start` | Start slide show from beginning |
| `start_from` | Start slide show from `slide_index` |
| `stop` | Stop the running slide show |

### `animate` — Shape Animations
| Action | Description |
|---|---|
| `add` | Add animation (`name_or_index`, `effect_type`: appear/fly/blinds/box/dissolve/fade/zoom/spin/grow_shrink/float etc., `trigger`: on_click/with_previous/after_previous, `duration`, `delay`) |
| `clear` | Clear animations from shape (`name_or_index`) |

### `find` — Find & Replace
| Action | Description |
|---|---|
| `find` | Find `search_text` across all slides (`match_case`, `match_whole_word`) |
| `find_replace` | Find `find_text` and replace with `replace_text` (`replace_all`: bool) |

### `master` — Slide Master & Layout
| Action | Description |
|---|---|
| `get_info` | Get slide master information |
| `apply_layout` | Apply a layout to slide (`slide_index`, `layout_index`, `master_index`) |
| `set_background` | Set master slide background (`color_rgb`, `master_index`) |

### `advanced` — Advanced Features
Use `action=help` for full details. Supported actions:

| Action | Description |
|---|---|
| `hyperlink` | Add hyperlink to shape (`address`, `text_to_display`, `name_or_index`) |
| `chart_add` | Add a chart (`chart_type`: column/line/pie/bar/area/scatter, `left`, `top`, `width`, `height`) |
| `chart_set_data` | Set chart data (`shape_name`, `categories`, `series_data`) |
| `media_video` | Insert video (`filepath`, `left`, `top`, `width`, `height`) |
| `media_audio` | Insert audio (`filepath`) |
| `slide_number` | Toggle slide numbers (`show_slide_number`) |
| `date_time` | Toggle date/time display (`show_date_time`) |
| `header_footer` | Set header/footer text (`header_text`, `footer_text`) |

## Usage Examples

### Create a presentation with a title

```
> Create a new presentation with a title slide that says "Q3 Business Review"
```

The assistant calls `pres` with `action=create`, then `shape_add` with `action=text_box`.

### Format slide content

```
> Make the text in shape 1 bold, size 36, and centered
```

The assistant calls `font` with `name_or_index=1`, `bold=true`, `font_size=36`, `alignment=center`.

### Add an image and export

```
> Insert the company logo from C:\logos\logo.png on slide 2 at position (600, 20)
  with width 200, then export as PDF to C:\output\deck.pdf
```

The assistant calls `shape_add` with `action=picture`, then `export` with `action=pdf`.

## COM ProgID

The client tries these ProgIDs in order, preferring running instances before creating new ones:
- `WPP.Application` (WPS Slide / Presentation)
- `KWPP.Application` (older WPS Slide)
- `PowerPoint.Application` (Microsoft PowerPoint fallback)

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

---

# WhatsApp MCP (`whatsapp-mcp`)

Read WhatsApp Desktop data via Chrome DevTools Protocol (CDP). **Read-only** — no sending, replying, deleting, or mutating any data.

## Features

| Category | Operations |
|---|---|
| **Chats** | List all chats, search chats by keyword |
| **Messages** | Read messages from any chat (with sender names) |
| **Contacts** | Get contact/group info (name, status, phone) |
| **Media** | Get media files, download images, view downloaded media |
| **Screenshot** | Take screenshots of the current WhatsApp view |
| **Status** | Check connection status |

## Prerequisites

- **Windows 10 / 11**
- **Chrome** with remote debugging enabled:
  ```bash
  chrome.exe --remote-debugging-port=9222
  ```
- **WhatsApp Web** (`https://web.whatsapp.com`) open and logged in within the debug Chrome

## Tool Reference

All 9 tools are purely read-only with safety protections that prevent accidental interaction.

### `whatsapp_list_chats`
List all visible chats from the WhatsApp sidebar. Returns chat name, last message preview, unread badge, and timestamp. No parameters required.

### `whatsapp_read_chat_messages`
Read recent messages from a specific chat. Opens the chat, scrolls to load history, and extracts messages with sender names.

| Parameter | Type | Required | Description |
|---|---|---|---|
| `chat_name` | string | Yes | Chat/contact name (partial match, case-insensitive) |
| `count` | integer | No | Max messages to return (default: 20) |

### `whatsapp_get_contact_info`
Get information about a contact or group. Opens the chat and reads the header — returns display name, status/bio, and phone.

| Parameter | Type | Required | Description |
|---|---|---|---|
| `contact_name` | string | Yes | Contact/group name (partial match, case-insensitive) |

### `whatsapp_search_chats`
Search chats by keyword using the WhatsApp search box.

| Parameter | Type | Required | Description |
|---|---|---|---|
| `query` | string | Yes | Search keyword |

### `whatsapp_get_chat_media`
Get recent media files (images, videos, documents) visible in a chat. Blob images are downloaded to local temp files.

| Parameter | Type | Required | Description |
|---|---|---|---|
| `chat_name` | string | Yes | Chat/contact name |
| `count` | integer | No | Max media items (default: 10) |

### `whatsapp_download_image`
Download a WhatsApp blob URL to a local file. Use after `whatsapp_get_chat_media` returns blob URLs.

| Parameter | Type | Required | Description |
|---|---|---|---|
| `url` | string | Yes | The blob: URL from a media item |
| `save_path` | string | No | Optional local file path (default: temp file) |

### `whatsapp_screenshot`
Take a screenshot of the current WhatsApp Web view. Saves to a temp PNG file. No parameters required.

### `whatsapp_view_media`
View an image/media file that was previously downloaded. Takes a file path and returns the actual image for display.

| Parameter | Type | Required | Description |
|---|---|---|---|
| `file_path` | string | Yes | Absolute path to the media file |

### `whatsapp_status`
Check whether Chrome is reachable and WhatsApp tab is found. No parameters required.

## Usage Examples

### Check your chats

```
> Show me my recent WhatsApp chats
```

The assistant calls `whatsapp_list_chats`.

### Read messages

```
> Read the last 10 messages from "Alice"
```

The assistant calls `whatsapp_read_chat_messages` with `chat_name="Alice"`, `count=10`.

### View shared media

```
> Show me the recent images from the "Project Team" group
```

The assistant calls `whatsapp_get_chat_media` with `chat_name="Project Team"`, then `whatsapp_view_media` to display each image.

## Safety Features

- **Danger-zone detection** — Blocks clicks on compose area, send button, and context menus
- **Post-click verification** — Confirms correct chat is opened after each navigation
- **Read-only guarantee** — No send, reply, delete, archive, or mutation operations

## Troubleshooting

### "Could not connect to Chrome"

- Ensure Chrome is running with `--remote-debugging-port=9222`
- Check that no other program is using port 9222
- Close all Chrome windows and restart with the debugging flag

### "WhatsApp tab not found"

- Open `https://web.whatsapp.com` in the debug Chrome instance
- Log in by scanning the QR code with your phone
- Wait for chats to load completely

## License

MIT
