# WPS MCP Server

A [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) server for automating **WPS Office Excel** on Windows. This server enables AI assistants (like GitHub Copilot, Claude, etc.) to programmatically create, read, update, and format Excel workbooks in WPS Office via COM automation.

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

## Prerequisites

- **Windows** (WPS Office COM automation is Windows-only)
- **Python 3.10+**
- **WPS Office** installed (with Excel / Spreadsheets component)
- Required Python packages:
  - `mcp` — MCP Python SDK
  - `pywin32` — Windows COM automation

## Installation

```bash
# Clone or navigate to the project directory
cd wps-mcp

# Install dependencies
pip install -r requirements.txt

# Or install the package in development mode
pip install -e .
```

## Build Standalone EXE (No Python Required)

You can build `wps-mcp.exe` — a single standalone executable that users can run **without installing Python or any dependencies**. The only requirement on the target machine is **WPS Office installed**.

### One-Click Build

```batch
# Double-click or run in terminal:
build_exe.bat
```

Or manually:

```bash
pip install pyinstaller
python build_exe.py
```

The output is at `dist\wps-mcp.exe` (~10-20 MB, self-contained).

### Using the EXE

After building, distribute `dist\wps-mcp.exe`. Users add it to their MCP client config:

```json
{
  "mcpServers": {
    "wps-mcp": {
      "command": "C:\\path\\to\\wps-mcp.exe"
    }
  }
}
```

No Python, no pip, no `requirements.txt` — just the `.exe` and WPS Office.

## Configuration

Add the server to your MCP client configuration (e.g., Claude Desktop, VS Code Copilot, etc.):

### Claude Desktop (`claude_desktop_config.json`)

```json
{
  "mcpServers": {
    "wps-mcp": {
      "command": "python",
      "args": ["-m", "wps_mcp.server"],
      "cwd": "d:/work/wps-mcp/src"
    }
  }
}
```

### VS Code / GitHub Copilot

Add to your `.vscode/mcp.json`:

```json
{
  "servers": {
    "wps-mcp": {
      "command": "python",
      "args": ["-m", "wps_mcp.server"],
      "cwd": "d:/work/wps-mcp/src"
    }
  }
}
```

Or use the installed entry point:

```json
{
  "servers": {
    "wps-mcp": {
      "command": "wps-mcp"
    }
  }
}
```

## Usage Examples

Once configured, an AI assistant can use the following tools:

### Create a workbook and populate data

```
> Create a new Excel workbook, add headers "Name", "Age", "City" in row 1,
  then add 3 rows of sample data. Make the header row bold with a blue background.
```

The assistant will call:
1. `wps_create_workbook` — creates a new workbook
2. `wps_set_range_values` — sets headers and data
3. `wps_set_font_bold` — bolds the header row
4. `wps_set_cell_color` — colors the header background

### Read data from an existing file

```
> Open the file "C:\Users\me\Documents\sales.xlsx" and tell me the total
  sales from column C
```

The assistant will call:
1. `wps_open_workbook` — opens the file
2. `wps_get_range_values` — reads the sales column
3. Computes the total (or asks you to use Excel formulas)

### Create a chart

```
> Create a bar chart from the data in range A1:B10
```

The assistant will call:
- `wps_add_chart` with `chart_type="bar"` and `range_ref="A1:B10"`

### Search for data

```
> Find all cells containing "Pending" in the sheet
```

The assistant will call:
- `wps_find_cell` with `search_text="Pending"`

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
└── src/
    └── wps_mcp/
        ├── __init__.py         # Package init
        ├── server.py           # MCP server with tool definitions & handlers
        ├── wps_client.py       # WPS Office COM client (low-level API)
        └── tools/
            └── __init__.py     # Tools package (extensible)
```

## How It Works

1. The MCP server runs as a subprocess and communicates via **stdio** (standard input/output) using JSON-RPC.
2. When an AI assistant calls a tool, the server dispatches it to the appropriate handler in `server.py`.
3. The handler calls the **WPS Excel COM client** (`wps_client.py`), which uses `win32com` to automate WPS Office via its COM interface (`ET.Application`).
4. Results are serialized to JSON and returned to the AI assistant.

### COM ProgIDs

The client tries these ProgIDs in order:
- `ET.Application` — WPS Excel (Spreadsheets)
- `KET.Application` — Older WPS Excel

If WPS Office is already running, it connects to the existing instance.

## Troubleshooting

### "Could not connect to WPS Office Excel"

- Ensure WPS Office is installed (the Excel/Spreadsheets component specifically)
- Try launching WPS Excel manually once to ensure it's properly registered
- Check that the COM ProgID `ET.Application` is registered (run `powershell Get-ChildItem HKLM:\Software\Classes\ | Where-Object { $_.PSChildName -like "*ET.Application*" }`)

### "ModuleNotFoundError: No module named 'win32com'"

```bash
pip install pywin32
```

### COM errors / crashes

- Ensure WPS Office is up to date
- Try running with the window visible (`wps_show_window`) to see if there are UI dialogs
- Some operations may not be supported in older WPS Office versions

## License

MIT
