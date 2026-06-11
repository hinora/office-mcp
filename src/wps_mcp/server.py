"""
MCP Server for WPS Office Excel.

Provides tools for automating WPS Office Excel via COM:
- Workbook management (create, open, save, close, list)
- Worksheet management (add, rename, delete, activate, list)
- Cell operations (read, write, range read/write, clear)
- Formatting (font, color, alignment, number format, merge)
- Row/column operations (insert, delete, resize)
- Chart creation
- Search/find
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
    CallToolResult,
)

try:
    from .wps_client import WPSExcelClient
except ImportError:
    # PyInstaller standalone: wps_client is bundled as wps_mcp.wps_client
    from wps_mcp.wps_client import WPSExcelClient

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)

# ── MCP Server Setup ───────────────────────────────────────────────────

server = Server("wps-mcp")

# Dedicated STA thread executor for COM operations.
# COM objects must be created and accessed from the same STA thread.
_sta_executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
_client: WPSExcelClient | None = None
_sta_initialized = False


def _init_sta() -> None:
    """Initialize COM on the STA worker thread and create the client."""
    global _client, _sta_initialized
    if not _sta_initialized:
        pythoncom.CoInitialize()
        _sta_initialized = True
    if _client is None:
        _client = WPSExcelClient(visible=False)


def get_client() -> WPSExcelClient:
    """Get or create the WPS Excel client singleton."""
    if _client is None:
        # Initialize on the STA thread
        future = _sta_executor.submit(_init_sta)
        future.result()
    assert _client is not None
    return _client


# ── Tool Definitions ───────────────────────────────────────────────────

TOOLS = [
    Tool(
        name="wps_get_app_info",
        description="Get information about the WPS Excel application (version, open workbooks count, etc.).",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    Tool(
        name="wps_create_workbook",
        description="Create a new Excel workbook in WPS Office.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    Tool(
        name="wps_open_workbook",
        description="Open an existing Excel workbook file from disk.",
        inputSchema={
            "type": "object",
            "properties": {
                "filepath": {
                    "type": "string",
                    "description": "Full path to the workbook file (.xlsx, .xls, .et, etc.).",
                },
            },
            "required": ["filepath"],
        },
    ),
    Tool(
        name="wps_save_workbook",
        description="Save the active workbook. Optionally save to a new file path.",
        inputSchema={
            "type": "object",
            "properties": {
                "filepath": {
                    "type": "string",
                    "description": "Optional path to save the workbook to. If omitted, saves to current location.",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="wps_close_workbook",
        description="Close the active workbook.",
        inputSchema={
            "type": "object",
            "properties": {
                "save": {
                    "type": "boolean",
                    "description": "Whether to save changes before closing. Default: true.",
                    "default": True,
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="wps_list_workbooks",
        description="List all currently open workbooks.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    Tool(
        name="wps_list_sheets",
        description="List all sheets in the active workbook with their names, types, and visibility.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    Tool(
        name="wps_add_sheet",
        description="Add a new worksheet to the active workbook.",
        inputSchema={
            "type": "object",
            "properties": {
                "name": {
                    "type": "string",
                    "description": "Optional name for the new sheet.",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="wps_rename_sheet",
        description="Rename a worksheet in the active workbook.",
        inputSchema={
            "type": "object",
            "properties": {
                "old_name": {
                    "type": "string",
                    "description": "Current name of the sheet.",
                },
                "new_name": {
                    "type": "string",
                    "description": "New name for the sheet.",
                },
            },
            "required": ["old_name", "new_name"],
        },
    ),
    Tool(
        name="wps_delete_sheet",
        description="Delete a worksheet from the active workbook.",
        inputSchema={
            "type": "object",
            "properties": {
                "name": {
                    "type": "string",
                    "description": "Name of the sheet to delete.",
                },
            },
            "required": ["name"],
        },
    ),
    Tool(
        name="wps_activate_sheet",
        description="Activate a specific sheet by name (bring it into focus).",
        inputSchema={
            "type": "object",
            "properties": {
                "name": {
                    "type": "string",
                    "description": "Name of the sheet to activate.",
                },
            },
            "required": ["name"],
        },
    ),
    Tool(
        name="wps_get_cell_value",
        description="Get the value of a specific cell.",
        inputSchema={
            "type": "object",
            "properties": {
                "cell_ref": {
                    "type": "string",
                    "description": "Cell reference like 'A1', 'B2', 'C10'.",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name. Uses active sheet if not specified.",
                },
            },
            "required": ["cell_ref"],
        },
    ),
    Tool(
        name="wps_set_cell_value",
        description="Set the value of a specific cell.",
        inputSchema={
            "type": "object",
            "properties": {
                "cell_ref": {
                    "type": "string",
                    "description": "Cell reference like 'A1', 'B2', 'C10'.",
                },
                "value": {
                    "type": "string",
                    "description": "The value to set. Numbers will be converted automatically.",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name. Uses active sheet if not specified.",
                },
            },
            "required": ["cell_ref", "value"],
        },
    ),
    Tool(
        name="wps_get_range_values",
        description="Get values from a range of cells. Returns a 2D array (list of rows).",
        inputSchema={
            "type": "object",
            "properties": {
                "range_ref": {
                    "type": "string",
                    "description": "Range reference like 'A1:B10', 'C5:F20'.",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name. Uses active sheet if not specified.",
                },
            },
            "required": ["range_ref"],
        },
    ),
    Tool(
        name="wps_set_range_values",
        description="Set values for a range of cells. Provide values as a 2D JSON array.",
        inputSchema={
            "type": "object",
            "properties": {
                "start_cell": {
                    "type": "string",
                    "description": "Top-left cell reference like 'A1', 'B2'.",
                },
                "values": {
                    "type": "string",
                    "description": "JSON-encoded 2D array of values. E.g., '[['Name','Age'],['John',30],['Jane',25]]'.",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name. Uses active sheet if not specified.",
                },
            },
            "required": ["start_cell", "values"],
        },
    ),
    Tool(
        name="wps_clear_cell",
        description="Clear the contents of a cell or range.",
        inputSchema={
            "type": "object",
            "properties": {
                "cell_or_range": {
                    "type": "string",
                    "description": "Cell reference (e.g., 'A1') or range (e.g., 'A1:B10').",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name. Uses active sheet if not specified.",
                },
            },
            "required": ["cell_or_range"],
        },
    ),
    Tool(
        name="wps_set_font_bold",
        description="Set font bold for a cell or range.",
        inputSchema={
            "type": "object",
            "properties": {
                "cell_or_range": {
                    "type": "string",
                    "description": "Cell (e.g., 'A1') or range (e.g., 'A1:C10').",
                },
                "bold": {
                    "type": "boolean",
                    "description": "True to make bold, False to remove bold.",
                    "default": True,
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["cell_or_range"],
        },
    ),
    Tool(
        name="wps_set_font_size",
        description="Set font size for a cell or range.",
        inputSchema={
            "type": "object",
            "properties": {
                "cell_or_range": {
                    "type": "string",
                    "description": "Cell (e.g., 'A1') or range (e.g., 'A1:C10').",
                },
                "size": {
                    "type": "integer",
                    "description": "Font size in points (e.g., 12, 14, 18).",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["cell_or_range", "size"],
        },
    ),
    Tool(
        name="wps_set_cell_color",
        description="Set the background (fill) color of a cell or range.",
        inputSchema={
            "type": "object",
            "properties": {
                "cell_or_range": {
                    "type": "string",
                    "description": "Cell (e.g., 'A1') or range (e.g., 'A1:C10').",
                },
                "color": {
                    "type": "string",
                    "description": "Color as RGB hex string (e.g., 'FF0000' for red, '00FF00' for green, '0000FF' for blue).",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["cell_or_range", "color"],
        },
    ),
    Tool(
        name="wps_set_alignment",
        description="Set horizontal text alignment for a cell or range.",
        inputSchema={
            "type": "object",
            "properties": {
                "cell_or_range": {
                    "type": "string",
                    "description": "Cell (e.g., 'A1') or range (e.g., 'A1:C10').",
                },
                "alignment": {
                    "type": "string",
                    "description": "Alignment: 'left', 'center', 'right', 'general', or 'justify'.",
                    "enum": ["left", "center", "right", "general", "justify"],
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["cell_or_range", "alignment"],
        },
    ),
    Tool(
        name="wps_set_number_format",
        description="Set the number format for a cell or range (e.g., '0.00', '#,##0', 'yyyy-mm-dd').",
        inputSchema={
            "type": "object",
            "properties": {
                "cell_or_range": {
                    "type": "string",
                    "description": "Cell (e.g., 'A1') or range (e.g., 'A1:C10').",
                },
                "format": {
                    "type": "string",
                    "description": "Excel number format string. Examples: '0.00' for 2 decimals, '#,##0' for thousands, '0%' for percent, 'yyyy-mm-dd' for date.",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["cell_or_range", "format"],
        },
    ),
    Tool(
        name="wps_merge_cells",
        description="Merge a range of cells into one.",
        inputSchema={
            "type": "object",
            "properties": {
                "range_ref": {
                    "type": "string",
                    "description": "Range to merge (e.g., 'A1:C1').",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["range_ref"],
        },
    ),
    Tool(
        name="wps_unmerge_cells",
        description="Unmerge a previously merged range of cells.",
        inputSchema={
            "type": "object",
            "properties": {
                "range_ref": {
                    "type": "string",
                    "description": "Range to unmerge (e.g., 'A1:C1').",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["range_ref"],
        },
    ),
    Tool(
        name="wps_get_used_range",
        description="Get the address (e.g., 'A1:D20') and dimensions of the used data range in a sheet.",
        inputSchema={
            "type": "object",
            "properties": {
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="wps_insert_row",
        description="Insert a new row at the specified position.",
        inputSchema={
            "type": "object",
            "properties": {
                "row": {
                    "type": "integer",
                    "description": "Row number at which to insert (e.g., 5 inserts before row 5).",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["row"],
        },
    ),
    Tool(
        name="wps_insert_column",
        description="Insert a new column at the specified position.",
        inputSchema={
            "type": "object",
            "properties": {
                "col": {
                    "type": "integer",
                    "description": "Column number at which to insert (1=A, 2=B, etc).",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["col"],
        },
    ),
    Tool(
        name="wps_delete_row",
        description="Delete a row at the specified position.",
        inputSchema={
            "type": "object",
            "properties": {
                "row": {
                    "type": "integer",
                    "description": "Row number to delete.",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["row"],
        },
    ),
    Tool(
        name="wps_delete_column",
        description="Delete a column at the specified position.",
        inputSchema={
            "type": "object",
            "properties": {
                "col": {
                    "type": "integer",
                    "description": "Column number to delete (1=A, 2=B, etc).",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["col"],
        },
    ),
    Tool(
        name="wps_set_row_height",
        description="Set the height of a row in points.",
        inputSchema={
            "type": "object",
            "properties": {
                "row": {
                    "type": "integer",
                    "description": "Row number.",
                },
                "height": {
                    "type": "number",
                    "description": "Row height in points.",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["row", "height"],
        },
    ),
    Tool(
        name="wps_set_column_width",
        description="Set the width of a column in characters.",
        inputSchema={
            "type": "object",
            "properties": {
                "col": {
                    "type": "integer",
                    "description": "Column number (1=A, 2=B, etc).",
                },
                "width": {
                    "type": "number",
                    "description": "Column width in characters.",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["col", "width"],
        },
    ),
    Tool(
        name="wps_add_chart",
        description="Add a chart to a worksheet.",
        inputSchema={
            "type": "object",
            "properties": {
                "chart_type": {
                    "type": "string",
                    "description": "Chart type: 'column', 'line', 'pie', 'bar', 'area', or 'scatter'.",
                    "enum": ["column", "line", "pie", "bar", "area", "scatter"],
                    "default": "column",
                },
                "range_ref": {
                    "type": "string",
                    "description": "Data range for the chart (e.g., 'A1:B10'). Include headers for best results.",
                },
                "left": {
                    "type": "number",
                    "description": "Left position in points.",
                    "default": 100,
                },
                "top": {
                    "type": "number",
                    "description": "Top position in points.",
                    "default": 100,
                },
                "width": {
                    "type": "number",
                    "description": "Chart width in points.",
                    "default": 400,
                },
                "height": {
                    "type": "number",
                    "description": "Chart height in points.",
                    "default": 300,
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["range_ref"],
        },
    ),
    Tool(
        name="wps_find_cell",
        description="Search for text in a worksheet and return the first matching cell.",
        inputSchema={
            "type": "object",
            "properties": {
                "search_text": {
                    "type": "string",
                    "description": "Text to search for.",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name to search in.",
                },
                "match_whole": {
                    "type": "boolean",
                    "description": "If True, match whole cell content. If False (default), match partial.",
                    "default": False,
                },
            },
            "required": ["search_text"],
        },
    ),
    Tool(
        name="wps_run_macro",
        description="Run a VBA macro by name.",
        inputSchema={
            "type": "object",
            "properties": {
                "macro_name": {
                    "type": "string",
                    "description": "The name of the macro/subroutine to run.",
                },
            },
            "required": ["macro_name"],
        },
    ),
    Tool(
        name="wps_show_window",
        description="Make the WPS Excel application window visible.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    Tool(
        name="wps_hide_window",
        description="Hide the WPS Excel application window (run in background).",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
]


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

        # Run on the dedicated STA thread (COM requires STA apartment)
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


def _execute_tool(name: str, args: dict[str, Any], client: WPSExcelClient) -> str:
    """Execute a tool synchronously. Called from the STA executor thread."""
    # Ensure COM is initialized on this thread (STA apartment)
    pythoncom.CoInitialize()

    result: Any = None

    # ── Application ──
    if name == "wps_get_app_info":
        result = client.get_app_info()

    elif name == "wps_show_window":
        client.show()
        result = {"message": "WPS Excel window is now visible."}

    elif name == "wps_hide_window":
        client.hide()
        result = {"message": "WPS Excel window is now hidden."}

    # ── Workbook ──
    elif name == "wps_create_workbook":
        wb_name = client.create_workbook()
        result = {"message": f"Created new workbook: {wb_name}", "workbook_name": wb_name}

    elif name == "wps_open_workbook":
        wb_name = client.open_workbook(args["filepath"])
        result = {"message": f"Opened workbook: {wb_name}", "workbook_name": wb_name}

    elif name == "wps_save_workbook":
        filepath = args.get("filepath")
        saved_path = client.save_workbook(filepath)
        result = {"message": f"Workbook saved to: {saved_path}", "path": saved_path}

    elif name == "wps_close_workbook":
        save = args.get("save", True)
        client.close_workbook(save)
        result = {"message": "Workbook closed." + (" (saved)" if save else " (not saved)")}

    elif name == "wps_list_workbooks":
        result = {"workbooks": client.list_workbooks()}

    # ── Worksheet ──
    elif name == "wps_list_sheets":
        result = {"sheets": client.list_sheets()}

    elif name == "wps_add_sheet":
        sheet_name = args.get("name")
        actual_name = client.add_sheet(sheet_name)
        result = {"message": f"Added sheet: {actual_name}", "sheet_name": actual_name}

    elif name == "wps_rename_sheet":
        new_name = client.rename_sheet(args["old_name"], args["new_name"])
        result = {"message": f"Renamed sheet '{args['old_name']}' to '{new_name}'"}

    elif name == "wps_delete_sheet":
        client.delete_sheet(args["name"])
        result = {"message": f"Deleted sheet: {args['name']}"}

    elif name == "wps_activate_sheet":
        sheet_name = client.activate_sheet(args["name"])
        result = {"message": f"Activated sheet: {sheet_name}"}

    # ── Cell Operations ──
    elif name == "wps_get_cell_value":
        value = client.get_cell_value(
            args["cell_ref"],
            args.get("sheet_name"),
        )
        result = {"cell": args["cell_ref"], "value": value}

    elif name == "wps_set_cell_value":
        val = _parse_value(args["value"])
        client.set_cell_value(args["cell_ref"], val, args.get("sheet_name"))
        result = {"message": f"Set {args['cell_ref']} = {args['value']}"}

    elif name == "wps_get_range_values":
        values = client.get_range_values(
            args["range_ref"],
            args.get("sheet_name"),
        )
        result = {"range": args["range_ref"], "values": values}

    elif name == "wps_set_range_values":
        values = json.loads(args["values"])
        if not isinstance(values, list):
            raise ValueError("values must be a JSON array")
        client.set_range_values(
            args["start_cell"],
            values,
            args.get("sheet_name"),
        )
        result = {"message": f"Set range starting at {args['start_cell']} with {len(values)} rows"}

    elif name == "wps_clear_cell":
        client.clear_range(args["cell_or_range"], args.get("sheet_name"))
        result = {"message": f"Cleared: {args['cell_or_range']}"}

    # ── Formatting ──
    elif name == "wps_set_font_bold":
        bold = args.get("bold", True)
        client.set_font_bold(args["cell_or_range"], bold, args.get("sheet_name"))
        result = {"message": f"Set {args['cell_or_range']} font bold = {bold}"}

    elif name == "wps_set_font_size":
        client.set_font_size(args["cell_or_range"], args["size"], args.get("sheet_name"))
        result = {"message": f"Set {args['cell_or_range']} font size = {args['size']}"}

    elif name == "wps_set_cell_color":
        color_hex = args["color"].lstrip("#")
        color_int = int(color_hex, 16)
        client.set_cell_color(args["cell_or_range"], color_int, args.get("sheet_name"))
        result = {"message": f"Set {args['cell_or_range']} fill color = #{color_hex}"}

    elif name == "wps_set_alignment":
        client.set_horizontal_alignment(
            args["cell_or_range"],
            args["alignment"],
            args.get("sheet_name"),
        )
        result = {"message": f"Set {args['cell_or_range']} alignment = {args['alignment']}"}

    elif name == "wps_set_number_format":
        client.set_number_format(
            args["cell_or_range"],
            args["format"],
            args.get("sheet_name"),
        )
        result = {"message": f"Set {args['cell_or_range']} number format = '{args['format']}'"}

    elif name == "wps_merge_cells":
        client.merge_cells(args["range_ref"], args.get("sheet_name"))
        result = {"message": f"Merged cells: {args['range_ref']}"}

    elif name == "wps_unmerge_cells":
        client.unmerge_cells(args["range_ref"], args.get("sheet_name"))
        result = {"message": f"Unmerged cells: {args['range_ref']}"}

    # ── Row / Column ──
    elif name == "wps_get_used_range":
        addr = client.get_used_range_address(args.get("sheet_name"))
        rows = client.get_row_count(args.get("sheet_name"))
        cols = client.get_column_count(args.get("sheet_name"))
        result = {
            "address": addr,
            "row_count": rows,
            "column_count": cols,
        }

    elif name == "wps_insert_row":
        client.insert_row(args["row"], args.get("sheet_name"))
        result = {"message": f"Inserted row at position {args['row']}"}

    elif name == "wps_insert_column":
        client.insert_column(args["col"], args.get("sheet_name"))
        result = {"message": f"Inserted column at position {args['col']}"}

    elif name == "wps_delete_row":
        client.delete_row(args["row"], args.get("sheet_name"))
        result = {"message": f"Deleted row {args['row']}"}

    elif name == "wps_delete_column":
        client.delete_column(args["col"], args.get("sheet_name"))
        result = {"message": f"Deleted column {args['col']}"}

    elif name == "wps_set_row_height":
        client.set_row_height(args["row"], args["height"], args.get("sheet_name"))
        result = {"message": f"Set row {args['row']} height = {args['height']}"}

    elif name == "wps_set_column_width":
        client.set_column_width(args["col"], args["width"], args.get("sheet_name"))
        result = {"message": f"Set column {args['col']} width = {args['width']}"}

    # ── Chart ──
    elif name == "wps_add_chart":
        chart_name = client.add_chart(
            chart_type=args.get("chart_type", "column"),
            range_ref=args["range_ref"],
            left=args.get("left", 100),
            top=args.get("top", 100),
            width=args.get("width", 400),
            height=args.get("height", 300),
            sheet_name=args.get("sheet_name"),
        )
        result = {"message": f"Added {args.get('chart_type', 'column')} chart: {chart_name}"}

    # ── Search ──
    elif name == "wps_find_cell":
        look_at = "whole" if args.get("match_whole", False) else "part"
        found = client.find_cell(
            args["search_text"],
            args.get("sheet_name"),
            look_at=look_at,
        )
        if found is None:
            result = {"found": False, "message": f"'{args['search_text']}' not found"}
        else:
            result = {"found": True, **found}

    # ── Macro ──
    elif name == "wps_run_macro":
        ret = client.run_macro(args["macro_name"])
        result = {"message": f"Ran macro '{args['macro_name']}'", "return_value": str(ret)}

    else:
        raise ValueError(f"Unknown tool: {name}")

    return json.dumps(result, ensure_ascii=False, default=str, indent=2)


def _parse_value(val: str) -> Any:
    """Try to parse a string value into an int or float; otherwise return as-is."""
    if not isinstance(val, str):
        return val
    # Try int
    try:
        return int(val)
    except ValueError:
        pass
    # Try float
    try:
        return float(val)
    except ValueError:
        pass
    return val


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
    """Entry point for the wps-mcp command."""
    asyncio.run(run_server())


if __name__ == "__main__":
    main()
