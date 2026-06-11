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

server = Server("wps-excel-mcp")

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
        _client = WPSExcelClient(visible=True)


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
    # ── Font & Text Formatting ──
    Tool(
        name="wps_set_font_name",
        description="Set the font name/typeface for a cell or range (e.g., 'Arial', 'Calibri', 'Times New Roman').",
        inputSchema={
            "type": "object",
            "properties": {
                "cell_or_range": {
                    "type": "string",
                    "description": "Cell (e.g., 'A1') or range (e.g., 'A1:C10').",
                },
                "font_name": {
                    "type": "string",
                    "description": "Font name (e.g., 'Arial', 'Calibri', 'Times New Roman', 'Microsoft YaHei').",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["cell_or_range", "font_name"],
        },
    ),
    Tool(
        name="wps_set_font_italic",
        description="Set or remove italic formatting for a cell or range.",
        inputSchema={
            "type": "object",
            "properties": {
                "cell_or_range": {
                    "type": "string",
                    "description": "Cell (e.g., 'A1') or range (e.g., 'A1:C10').",
                },
                "italic": {
                    "type": "boolean",
                    "description": "True to make italic, False to remove italic.",
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
        name="wps_set_font_color",
        description="Set the font (text) color for a cell or range using an RGB hex string.",
        inputSchema={
            "type": "object",
            "properties": {
                "cell_or_range": {
                    "type": "string",
                    "description": "Cell (e.g., 'A1') or range (e.g., 'A1:C10').",
                },
                "color": {
                    "type": "string",
                    "description": "Font color as RGB hex string (e.g., 'FF0000' for red, '0000FF' for blue, '000000' for black).",
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
        name="wps_set_wrap_text",
        description="Enable or disable text wrapping for a cell or range.",
        inputSchema={
            "type": "object",
            "properties": {
                "cell_or_range": {
                    "type": "string",
                    "description": "Cell (e.g., 'A1') or range (e.g., 'A1:C10').",
                },
                "wrap": {
                    "type": "boolean",
                    "description": "True to wrap text, False to unwrap.",
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
        name="wps_set_borders",
        description="Set borders for a cell or range. Supports all border edges and outline-only mode.",
        inputSchema={
            "type": "object",
            "properties": {
                "cell_or_range": {
                    "type": "string",
                    "description": "Cell (e.g., 'A1') or range (e.g., 'A1:C10').",
                },
                "border_style": {
                    "type": "string",
                    "description": "Border line style: 'thin' (default), 'medium', 'thick', 'dotted', 'dashed', 'double', 'hairline', 'none'.",
                    "enum": ["thin", "medium", "thick", "dotted", "dashed", "double", "hairline", "none"],
                    "default": "thin",
                },
                "border_color": {
                    "type": "string",
                    "description": "Optional border color as RGB hex (e.g., '000000' for black, 'FF0000' for red).",
                },
                "outline_only": {
                    "type": "boolean",
                    "description": "If True, set only outer borders. If False (default), set all borders including inner grid.",
                    "default": False,
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["cell_or_range"],
        },
    ),
    # ── AutoFit / Freeze / Filter ──
    Tool(
        name="wps_autofit_columns",
        description="Auto-fit column widths to their content. Can apply to all used columns or a specific range.",
        inputSchema={
            "type": "object",
            "properties": {
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
                "start_col": {
                    "type": "integer",
                    "description": "Optional first column number (1=A). If omitted, fits all used columns.",
                },
                "end_col": {
                    "type": "integer",
                    "description": "Optional last column number. If omitted, fits only start_col or all columns.",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="wps_autofit_rows",
        description="Auto-fit row heights to their content. Can apply to all used rows or a specific range.",
        inputSchema={
            "type": "object",
            "properties": {
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
                "start_row": {
                    "type": "integer",
                    "description": "Optional first row number. If omitted, fits all used rows.",
                },
                "end_row": {
                    "type": "integer",
                    "description": "Optional last row number. If omitted, fits only start_row or all rows.",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="wps_freeze_panes",
        description="Freeze panes at a specific cell. Rows above and columns to the left of the cell are frozen.",
        inputSchema={
            "type": "object",
            "properties": {
                "cell_ref": {
                    "type": "string",
                    "description": "Cell at which to freeze. Default 'B2' freezes first row + first column. 'A2' freezes first row only. 'B1' freezes first column only.",
                    "default": "B2",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="wps_unfreeze_panes",
        description="Remove all frozen panes from a sheet.",
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
        name="wps_auto_filter",
        description="Add or toggle AutoFilter dropdowns for a range. If no range is specified, applies to the used range.",
        inputSchema={
            "type": "object",
            "properties": {
                "range_ref": {
                    "type": "string",
                    "description": "Optional range to apply filter (e.g., 'A1:D100'). If omitted, uses the entire used range.",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": [],
        },
    ),
    # ── Sort / Copy-Paste ──
    Tool(
        name="wps_sort_range",
        description="Sort a range of cells by a specified column.",
        inputSchema={
            "type": "object",
            "properties": {
                "range_ref": {
                    "type": "string",
                    "description": "Range to sort (e.g., 'A2:D100'). Include header row if present.",
                },
                "sort_key": {
                    "type": "string",
                    "description": "Cell within the range to sort by (e.g., 'A2' to sort by column A). If omitted, sorts by first column.",
                },
                "sort_order": {
                    "type": "string",
                    "description": "Sort order: 'ascending' (default) or 'descending'.",
                    "enum": ["ascending", "descending"],
                    "default": "ascending",
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
        name="wps_copy_range",
        description="Copy a range to the clipboard.",
        inputSchema={
            "type": "object",
            "properties": {
                "range_ref": {
                    "type": "string",
                    "description": "Range to copy (e.g., 'A1:D10').",
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
        name="wps_paste_range",
        description="Paste clipboard contents to a destination cell. Supports paste special options.",
        inputSchema={
            "type": "object",
            "properties": {
                "dest_cell": {
                    "type": "string",
                    "description": "Top-left cell to paste to (e.g., 'A1').",
                },
                "paste_special": {
                    "type": "string",
                    "description": "Optional paste mode: 'values', 'formats', 'formulas', 'all' (default), 'transpose'.",
                    "enum": ["values", "formats", "formulas", "all", "transpose"],
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["dest_cell"],
        },
    ),
    # ── Find / Comment / Clear ──
    Tool(
        name="wps_find_next",
        description="Find the next occurrence after a previous wps_find_cell call.",
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
        name="wps_add_comment",
        description="Add a comment/note to a cell.",
        inputSchema={
            "type": "object",
            "properties": {
                "cell_ref": {
                    "type": "string",
                    "description": "Cell reference (e.g., 'A1').",
                },
                "text": {
                    "type": "string",
                    "description": "Comment text.",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["cell_ref", "text"],
        },
    ),
    Tool(
        name="wps_delete_comment",
        description="Remove a comment from a cell.",
        inputSchema={
            "type": "object",
            "properties": {
                "cell_ref": {
                    "type": "string",
                    "description": "Cell reference (e.g., 'A1').",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["cell_ref"],
        },
    ),
    Tool(
        name="wps_clear_formats",
        description="Clear only formatting from a cell or range (retains content).",
        inputSchema={
            "type": "object",
            "properties": {
                "cell_or_range": {
                    "type": "string",
                    "description": "Cell (e.g., 'A1') or range (e.g., 'A1:C10').",
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
        name="wps_clear_all",
        description="Clear everything (contents, formats, comments) from a cell or range.",
        inputSchema={
            "type": "object",
            "properties": {
                "cell_or_range": {
                    "type": "string",
                    "description": "Cell (e.g., 'A1') or range (e.g., 'A1:C10').",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["cell_or_range"],
        },
    ),
    # ── Conditional Formatting ──
    Tool(
        name="wps_add_conditional_format",
        description="Add a conditional formatting rule to a range (e.g., highlight cells greater than a value).",
        inputSchema={
            "type": "object",
            "properties": {
                "range_ref": {
                    "type": "string",
                    "description": "Range to apply to (e.g., 'B2:B100').",
                },
                "operator": {
                    "type": "string",
                    "description": "Comparison operator.",
                    "enum": ["greaterThan", "lessThan", "equal", "between", "greaterThanOrEqual", "lessThanOrEqual", "notEqual"],
                    "default": "greaterThan",
                },
                "formula": {
                    "type": "string",
                    "description": "Threshold value or formula (e.g., '100', '0', '=$B$1').",
                    "default": "0",
                },
                "font_color": {
                    "type": "string",
                    "description": "Optional font color as RGB hex (e.g., 'FF0000' for red).",
                },
                "bg_color": {
                    "type": "string",
                    "description": "Optional background fill color as RGB hex (e.g., 'FFFF00' for yellow).",
                },
                "bold": {
                    "type": "boolean",
                    "description": "Whether to make the font bold.",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["range_ref"],
        },
    ),
    # ── Data Validation ──
    Tool(
        name="wps_add_data_validation",
        description="Add data validation (dropdown list or input restriction) to a cell or range.",
        inputSchema={
            "type": "object",
            "properties": {
                "range_ref": {
                    "type": "string",
                    "description": "Range to apply validation to (e.g., 'C2:C100').",
                },
                "validation_type": {
                    "type": "string",
                    "description": "Validation type: 'list' for dropdown, or 'whole', 'decimal', 'date', 'time', 'textLength', 'custom'.",
                    "enum": ["list", "whole", "decimal", "date", "time", "textLength", "custom"],
                    "default": "list",
                },
                "formula1": {
                    "type": "string",
                    "description": "Validation formula. For list: 'Option1,Option2,Option3' or '=$A$1:$A$10' for range-based list. For other types: the min/allowed value.",
                    "default": "",
                },
                "formula2": {
                    "type": "string",
                    "description": "Second formula for 'between' / 'notBetween' operators (max value).",
                    "default": "",
                },
                "ignore_blank": {
                    "type": "boolean",
                    "description": "Allow blank cells. Default: true.",
                    "default": True,
                },
                "show_dropdown": {
                    "type": "boolean",
                    "description": "Show dropdown arrow for list validation. Default: true.",
                    "default": True,
                },
                "error_title": {
                    "type": "string",
                    "description": "Title for the error dialog when invalid data is entered.",
                },
                "error_message": {
                    "type": "string",
                    "description": "Error message to show when invalid data is entered.",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["range_ref"],
        },
    ),
    # ── Sheet Protection ──
    Tool(
        name="wps_protect_sheet",
        description="Protect a worksheet with optional password and permission flags.",
        inputSchema={
            "type": "object",
            "properties": {
                "password": {
                    "type": "string",
                    "description": "Optional password to protect the sheet.",
                    "default": "",
                },
                "allow_sort": {
                    "type": "boolean",
                    "description": "Allow sorting of locked cells. Default: false.",
                    "default": False,
                },
                "allow_filter": {
                    "type": "boolean",
                    "description": "Allow using AutoFilter. Default: false.",
                    "default": False,
                },
                "allow_format_cells": {
                    "type": "boolean",
                    "description": "Allow formatting cells. Default: false.",
                    "default": False,
                },
                "allow_insert_rows": {
                    "type": "boolean",
                    "description": "Allow inserting rows. Default: false.",
                    "default": False,
                },
                "allow_delete_rows": {
                    "type": "boolean",
                    "description": "Allow deleting rows. Default: false.",
                    "default": False,
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="wps_unprotect_sheet",
        description="Remove protection from a worksheet.",
        inputSchema={
            "type": "object",
            "properties": {
                "password": {
                    "type": "string",
                    "description": "Password if the sheet was protected with one.",
                    "default": "",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": [],
        },
    ),
    # ── Page Setup ──
    Tool(
        name="wps_set_print_area",
        description="Set the print area for a sheet.",
        inputSchema={
            "type": "object",
            "properties": {
                "range_ref": {
                    "type": "string",
                    "description": "Range to set as print area (e.g., 'A1:F50').",
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
        name="wps_clear_print_area",
        description="Clear the print area for a sheet.",
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
        name="wps_set_page_orientation",
        description="Set the page orientation for printing: portrait or landscape.",
        inputSchema={
            "type": "object",
            "properties": {
                "orientation": {
                    "type": "string",
                    "description": "Page orientation: 'portrait' or 'landscape'.",
                    "enum": ["portrait", "landscape"],
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["orientation"],
        },
    ),
    # ── Formula ──
    Tool(
        name="wps_set_formula",
        description="Set a formula in a cell (e.g., '=SUM(B2:B10)', '=A1*2').",
        inputSchema={
            "type": "object",
            "properties": {
                "cell_ref": {
                    "type": "string",
                    "description": "Cell reference (e.g., 'A1').",
                },
                "formula": {
                    "type": "string",
                    "description": "The Excel formula, including leading '=' (e.g., '=SUM(B2:B10)', '=VLOOKUP(A1,D:E,2,FALSE)').",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["cell_ref", "formula"],
        },
    ),
    Tool(
        name="wps_get_formula",
        description="Get the formula of a cell (not its computed value). Returns the formula string, or the literal value if no formula is set.",
        inputSchema={
            "type": "object",
            "properties": {
                "cell_ref": {
                    "type": "string",
                    "description": "Cell reference (e.g., 'A1').",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["cell_ref"],
        },
    ),
    # ── Export ──
    Tool(
        name="wps_export_to_pdf",
        description="Export the active workbook or a specific sheet to a PDF file.",
        inputSchema={
            "type": "object",
            "properties": {
                "filepath": {
                    "type": "string",
                    "description": "Full path for the output PDF file (e.g., 'C:\\report.pdf').",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name. If omitted, exports the entire workbook.",
                },
            },
            "required": ["filepath"],
        },
    ),
    # ── Find / Replace ──
    Tool(
        name="wps_find_replace",
        description="Find and replace text across a sheet. Returns the number of replacements made.",
        inputSchema={
            "type": "object",
            "properties": {
                "find_text": {
                    "type": "string",
                    "description": "Text to search for.",
                },
                "replace_text": {
                    "type": "string",
                    "description": "Replacement text.",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
                "match_case": {
                    "type": "boolean",
                    "description": "If True, match case. Default: false.",
                    "default": False,
                },
                "match_whole": {
                    "type": "boolean",
                    "description": "If True, match whole cell content only. Default: false.",
                    "default": False,
                },
            },
            "required": ["find_text", "replace_text"],
        },
    ),
    # ── Workbook Activation ──
    Tool(
        name="wps_activate_workbook",
        description="Activate a specific workbook by name when multiple workbooks are open.",
        inputSchema={
            "type": "object",
            "properties": {
                "name": {
                    "type": "string",
                    "description": "Name of the workbook to activate.",
                },
            },
            "required": ["name"],
        },
    ),
    # ── Remove Duplicates ──
    Tool(
        name="wps_remove_duplicates",
        description="Remove duplicate rows from a range. Optionally specify which columns to check for duplicates.",
        inputSchema={
            "type": "object",
            "properties": {
                "range_ref": {
                    "type": "string",
                    "description": "Range to remove duplicates from (e.g., 'A1:D100').",
                },
                "columns": {
                    "type": "string",
                    "description": "Optional JSON array of 1-based column indices within the range to check (e.g., '[1,2]'). If omitted, all columns are used.",
                },
                "has_header": {
                    "type": "boolean",
                    "description": "Whether the range includes a header row. Default: true.",
                    "default": True,
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["range_ref"],
        },
    ),
    # ── Vertical Alignment ──
    Tool(
        name="wps_set_vertical_alignment",
        description="Set vertical text alignment for a cell or range.",
        inputSchema={
            "type": "object",
            "properties": {
                "cell_or_range": {
                    "type": "string",
                    "description": "Cell (e.g., 'A1') or range (e.g., 'A1:C10').",
                },
                "alignment": {
                    "type": "string",
                    "description": "Vertical alignment: 'top', 'center', 'bottom', 'justify', 'distributed'.",
                    "enum": ["top", "center", "bottom", "justify", "distributed"],
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["cell_or_range", "alignment"],
        },
    ),
    # ── Sheet Copy / Move ──
    Tool(
        name="wps_copy_sheet",
        description="Create a copy of a worksheet within the active workbook.",
        inputSchema={
            "type": "object",
            "properties": {
                "source_name": {
                    "type": "string",
                    "description": "Name of the sheet to copy.",
                },
                "new_name": {
                    "type": "string",
                    "description": "Optional new name for the copied sheet.",
                },
                "before": {
                    "type": "string",
                    "description": "Optional sheet name to insert the copy before.",
                },
                "after": {
                    "type": "string",
                    "description": "Optional sheet name to insert the copy after.",
                },
            },
            "required": ["source_name"],
        },
    ),
    Tool(
        name="wps_move_sheet",
        description="Move (reorder) a worksheet to a new position within the active workbook.",
        inputSchema={
            "type": "object",
            "properties": {
                "source_name": {
                    "type": "string",
                    "description": "Name of the sheet to move.",
                },
                "before": {
                    "type": "string",
                    "description": "Optional sheet name to move before.",
                },
                "after": {
                    "type": "string",
                    "description": "Optional sheet name to move after. Default: moves to end.",
                },
            },
            "required": ["source_name"],
        },
    ),
    # ── Show / Hide Sheet ──
    Tool(
        name="wps_hide_sheet",
        description="Hide a worksheet (make it not visible in the tab bar).",
        inputSchema={
            "type": "object",
            "properties": {
                "name": {
                    "type": "string",
                    "description": "Name of the sheet to hide.",
                },
            },
            "required": ["name"],
        },
    ),
    Tool(
        name="wps_unhide_sheet",
        description="Unhide a previously hidden worksheet.",
        inputSchema={
            "type": "object",
            "properties": {
                "name": {
                    "type": "string",
                    "description": "Name of the sheet to unhide.",
                },
            },
            "required": ["name"],
        },
    ),
    # ── Hyperlinks ──
    Tool(
        name="wps_add_hyperlink",
        description="Add a hyperlink to a cell. Can link to URLs, files, or cell references.",
        inputSchema={
            "type": "object",
            "properties": {
                "cell_ref": {
                    "type": "string",
                    "description": "Cell reference (e.g., 'A1').",
                },
                "address": {
                    "type": "string",
                    "description": "URL, file path, or cell reference the link points to.",
                },
                "text_to_display": {
                    "type": "string",
                    "description": "Optional display text for the hyperlink.",
                },
                "screen_tip": {
                    "type": "string",
                    "description": "Optional tooltip text shown on hover.",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["cell_ref", "address"],
        },
    ),
    Tool(
        name="wps_remove_hyperlink",
        description="Remove hyperlinks from a cell or range.",
        inputSchema={
            "type": "object",
            "properties": {
                "cell_or_range": {
                    "type": "string",
                    "description": "Cell (e.g., 'A1') or range (e.g., 'A1:D10') to remove hyperlinks from.",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["cell_or_range"],
        },
    ),
    # ── Conditional Formatting Delete ──
    Tool(
        name="wps_delete_conditional_format",
        description="Remove all conditional formatting rules from a range.",
        inputSchema={
            "type": "object",
            "properties": {
                "range_ref": {
                    "type": "string",
                    "description": "Range to remove conditional formatting from (e.g., 'A1:A100').",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["range_ref"],
        },
    ),
    # ── Font Underline ──
    Tool(
        name="wps_set_font_underline",
        description="Set font underline style for a cell or range.",
        inputSchema={
            "type": "object",
            "properties": {
                "cell_or_range": {
                    "type": "string",
                    "description": "Cell (e.g., 'A1') or range (e.g., 'A1:C10').",
                },
                "underline_style": {
                    "type": "string",
                    "description": "Underline style: 'none', 'single', 'double', 'singleAccounting', 'doubleAccounting'.",
                    "enum": ["none", "single", "double", "singleAccounting", "doubleAccounting"],
                    "default": "single",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["cell_or_range"],
        },
    ),
    # ── Row / Column Grouping ──
    Tool(
        name="wps_group_rows",
        description="Group rows together for outlining (collapse/expand).",
        inputSchema={
            "type": "object",
            "properties": {
                "start_row": {
                    "type": "integer",
                    "description": "First row to group.",
                },
                "end_row": {
                    "type": "integer",
                    "description": "Last row to group.",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["start_row", "end_row"],
        },
    ),
    Tool(
        name="wps_ungroup_rows",
        description="Ungroup previously grouped rows.",
        inputSchema={
            "type": "object",
            "properties": {
                "start_row": {
                    "type": "integer",
                    "description": "First row to ungroup.",
                },
                "end_row": {
                    "type": "integer",
                    "description": "Last row to ungroup.",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["start_row", "end_row"],
        },
    ),
    Tool(
        name="wps_group_columns",
        description="Group columns together for outlining (collapse/expand).",
        inputSchema={
            "type": "object",
            "properties": {
                "start_col": {
                    "type": "integer",
                    "description": "First column to group (1=A, 2=B, etc).",
                },
                "end_col": {
                    "type": "integer",
                    "description": "Last column to group (1=A, 2=B, etc).",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["start_col", "end_col"],
        },
    ),
    Tool(
        name="wps_ungroup_columns",
        description="Ungroup previously grouped columns.",
        inputSchema={
            "type": "object",
            "properties": {
                "start_col": {
                    "type": "integer",
                    "description": "First column to ungroup (1=A, 2=B, etc).",
                },
                "end_col": {
                    "type": "integer",
                    "description": "Last column to ungroup (1=A, 2=B, etc).",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["start_col", "end_col"],
        },
    ),
    # ── Page Margins / Headers ──
    Tool(
        name="wps_set_page_margins",
        description="Set page margins (in points) for a sheet.",
        inputSchema={
            "type": "object",
            "properties": {
                "left": {
                    "type": "number",
                    "description": "Left margin in points.",
                },
                "right": {
                    "type": "number",
                    "description": "Right margin in points.",
                },
                "top": {
                    "type": "number",
                    "description": "Top margin in points.",
                },
                "bottom": {
                    "type": "number",
                    "description": "Bottom margin in points.",
                },
                "header": {
                    "type": "number",
                    "description": "Header margin in points.",
                },
                "footer": {
                    "type": "number",
                    "description": "Footer margin in points.",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="wps_set_header_footer",
        description="Set custom header and footer text for a sheet. Format codes: &P (page #), &N (total pages), &D (date), &T (time), &F (filename), &A (sheet name), &B (bold), &I (italic).",
        inputSchema={
            "type": "object",
            "properties": {
                "left_header": {
                    "type": "string",
                    "description": "Left header text.",
                    "default": "",
                },
                "center_header": {
                    "type": "string",
                    "description": "Center header text.",
                    "default": "",
                },
                "right_header": {
                    "type": "string",
                    "description": "Right header text.",
                    "default": "",
                },
                "left_footer": {
                    "type": "string",
                    "description": "Left footer text.",
                    "default": "",
                },
                "center_footer": {
                    "type": "string",
                    "description": "Center footer text (e.g., 'Page &P of &N').",
                    "default": "",
                },
                "right_footer": {
                    "type": "string",
                    "description": "Right footer text.",
                    "default": "",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": [],
        },
    ),
    # ── Text to Columns ──
    Tool(
        name="wps_text_to_columns",
        description="Split text in a column into multiple columns using a delimiter (like comma, space, tab).",
        inputSchema={
            "type": "object",
            "properties": {
                "range_ref": {
                    "type": "string",
                    "description": "Single-column range to split (e.g., 'A1:A100').",
                },
                "delimiter": {
                    "type": "string",
                    "description": "Delimiter character: ',' (comma), ';' (semicolon), '\\t' (tab), ' ' (space), '|' (pipe). Default: ','.",
                    "default": ",",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["range_ref"],
        },
    ),
    # ── Named Ranges ──
    Tool(
        name="wps_create_named_range",
        description="Create a named range in the active workbook.",
        inputSchema={
            "type": "object",
            "properties": {
                "name": {
                    "type": "string",
                    "description": "Name for the named range.",
                },
                "refers_to": {
                    "type": "string",
                    "description": "The range reference formula (e.g., '=Sheet1!$A$1:$D$10', '=Sheet1!$A:$A').",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name for scope.",
                },
            },
            "required": ["name", "refers_to"],
        },
    ),
    Tool(
        name="wps_delete_named_range",
        description="Delete a named range from the active workbook.",
        inputSchema={
            "type": "object",
            "properties": {
                "name": {
                    "type": "string",
                    "description": "Name of the named range to delete.",
                },
            },
            "required": ["name"],
        },
    ),
    Tool(
        name="wps_list_named_ranges",
        description="List all named ranges in the active workbook.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    # ── Pivot Table ──
    Tool(
        name="wps_create_pivot_table",
        description="Create a pivot table from a source data range.",
        inputSchema={
            "type": "object",
            "properties": {
                "source_range": {
                    "type": "string",
                    "description": "Source data range (e.g., 'A1:F100').",
                },
                "dest_cell": {
                    "type": "string",
                    "description": "Top-left cell where the pivot table will be placed (e.g., 'H1').",
                },
                "pivot_name": {
                    "type": "string",
                    "description": "Optional name for the pivot table.",
                    "default": "PivotTable1",
                },
                "row_fields": {
                    "type": "string",
                    "description": "Optional JSON array of field names for row labels (e.g., '["Category","SubCategory"]').",
                },
                "column_fields": {
                    "type": "string",
                    "description": "Optional JSON array of field names for column labels (e.g., '["Year"]').",
                },
                "data_fields": {
                    "type": "string",
                    "description": "Optional JSON array of field names for values (e.g., '["Amount"]').",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name for both source and destination.",
                },
            },
            "required": ["source_range", "dest_cell"],
        },
    ),
    # ── Sparklines ──
    Tool(
        name="wps_add_sparkline",
        description="Add a sparkline (mini in-cell chart) to a cell.",
        inputSchema={
            "type": "object",
            "properties": {
                "source_range": {
                    "type": "string",
                    "description": "Data range for the sparkline (e.g., 'A1:A10').",
                },
                "dest_cell": {
                    "type": "string",
                    "description": "Cell where the sparkline will be placed.",
                },
                "spark_type": {
                    "type": "string",
                    "description": "Sparkline type: 'line', 'column', or 'winloss'.",
                    "enum": ["line", "column", "winloss"],
                    "default": "line",
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["source_range", "dest_cell"],
        },
    ),
    # ── Insert Picture / Shape ──
    Tool(
        name="wps_insert_picture",
        description="Insert an image (.png, .jpg, etc.) into a sheet.",
        inputSchema={
            "type": "object",
            "properties": {
                "filepath": {
                    "type": "string",
                    "description": "Full path to the image file.",
                },
                "left": {
                    "type": "number",
                    "description": "Left position in points. Default: 100.",
                    "default": 100,
                },
                "top": {
                    "type": "number",
                    "description": "Top position in points. Default: 100.",
                    "default": 100,
                },
                "width": {
                    "type": "number",
                    "description": "Width in points. Default: 200.",
                    "default": 200,
                },
                "height": {
                    "type": "number",
                    "description": "Height in points. Default: 150.",
                    "default": 150,
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": ["filepath"],
        },
    ),
    Tool(
        name="wps_insert_shape",
        description="Insert a drawing shape (rectangle, oval, line, arrow, textbox) into a sheet.",
        inputSchema={
            "type": "object",
            "properties": {
                "shape_type": {
                    "type": "string",
                    "description": "Shape type: 'rectangle', 'oval', 'line', 'arrow', 'textbox'.",
                    "enum": ["rectangle", "oval", "line", "arrow", "textbox"],
                    "default": "rectangle",
                },
                "left": {
                    "type": "number",
                    "description": "Left position in points. Default: 100.",
                    "default": 100,
                },
                "top": {
                    "type": "number",
                    "description": "Top position in points. Default: 100.",
                    "default": 100,
                },
                "width": {
                    "type": "number",
                    "description": "Width in points. Default: 200.",
                    "default": 200,
                },
                "height": {
                    "type": "number",
                    "description": "Height in points. Default: 100.",
                    "default": 100,
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
            "required": [],
        },
    ),
    # ── Gridlines ──
    Tool(
        name="wps_toggle_gridlines",
        description="Show or hide gridlines on the active sheet.",
        inputSchema={
            "type": "object",
            "properties": {
                "visible": {
                    "type": "boolean",
                    "description": "True to show gridlines, False to hide them. Default: true.",
                    "default": True,
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name.",
                },
            },
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

    # ── Font & Text Formatting ──
    elif name == "wps_set_font_name":
        client.set_font_name(args["cell_or_range"], args["font_name"], args.get("sheet_name"))
        result = {"message": f"Set {args['cell_or_range']} font name = '{args['font_name']}'"}

    elif name == "wps_set_font_italic":
        italic = args.get("italic", True)
        client.set_font_italic(args["cell_or_range"], italic, args.get("sheet_name"))
        result = {"message": f"Set {args['cell_or_range']} font italic = {italic}"}

    elif name == "wps_set_font_color":
        color_hex = args["color"].lstrip("#")
        color_int = int(color_hex, 16)
        client.set_font_color(args["cell_or_range"], color_int, args.get("sheet_name"))
        result = {"message": f"Set {args['cell_or_range']} font color = #{color_hex}"}

    elif name == "wps_set_wrap_text":
        wrap = args.get("wrap", True)
        client.set_wrap_text(args["cell_or_range"], wrap, args.get("sheet_name"))
        result = {"message": f"Set {args['cell_or_range']} wrap text = {wrap}"}

    elif name == "wps_set_borders":
        border_style = args.get("border_style", "thin")
        border_color = None
        if args.get("border_color"):
            border_color = int(args["border_color"].lstrip("#"), 16)
        outline_only = args.get("outline_only", False)
        client.set_borders(
            args["cell_or_range"],
            border_style,
            border_color,
            outline_only,
            args.get("sheet_name"),
        )
        result = {"message": f"Set borders on {args['cell_or_range']} (style={border_style})"}

    # ── AutoFit / Freeze / Filter ──
    elif name == "wps_autofit_columns":
        client.autofit_columns(
            args.get("sheet_name"),
            args.get("start_col"),
            args.get("end_col"),
        )
        result = {"message": "Auto-fitted column widths."}

    elif name == "wps_autofit_rows":
        client.autofit_rows(
            args.get("sheet_name"),
            args.get("start_row"),
            args.get("end_row"),
        )
        result = {"message": "Auto-fitted row heights."}

    elif name == "wps_freeze_panes":
        cell_ref = args.get("cell_ref", "B2")
        client.freeze_panes(cell_ref, args.get("sheet_name"))
        result = {"message": f"Froze panes at {cell_ref}"}

    elif name == "wps_unfreeze_panes":
        client.unfreeze_panes(args.get("sheet_name"))
        result = {"message": "Unfroze all panes."}

    elif name == "wps_auto_filter":
        client.auto_filter(args.get("range_ref"), args.get("sheet_name"))
        rng = args.get("range_ref", "used range")
        result = {"message": f"Applied AutoFilter to {rng}"}

    # ── Sort / Copy-Paste ──
    elif name == "wps_sort_range":
        client.sort_range(
            args["range_ref"],
            args.get("sort_key"),
            args.get("sort_order", "ascending"),
            args.get("sheet_name"),
        )
        order = args.get("sort_order", "ascending")
        result = {"message": f"Sorted {args['range_ref']} ({order})."}

    elif name == "wps_copy_range":
        client.copy_range(args["range_ref"], args.get("sheet_name"))
        result = {"message": f"Copied {args['range_ref']} to clipboard."}

    elif name == "wps_paste_range":
        client.paste_range(
            args["dest_cell"],
            args.get("sheet_name"),
            args.get("paste_special"),
        )
        mode = args.get("paste_special", "all")
        result = {"message": f"Pasted to {args['dest_cell']} (mode={mode})."}

    # ── Find / Comment / Clear ──
    elif name == "wps_find_next":
        found = client.find_next_cell(args.get("sheet_name"))
        if found is None:
            result = {"found": False, "message": "No more matches."}
        else:
            result = {"found": True, **found}

    elif name == "wps_add_comment":
        client.add_comment(args["cell_ref"], args["text"], args.get("sheet_name"))
        result = {"message": f"Added comment to {args['cell_ref']}"}

    elif name == "wps_delete_comment":
        client.delete_comment(args["cell_ref"], args.get("sheet_name"))
        result = {"message": f"Deleted comment from {args['cell_ref']}"}

    elif name == "wps_clear_formats":
        client.clear_formats(args["cell_or_range"], args.get("sheet_name"))
        result = {"message": f"Cleared formats from {args['cell_or_range']}"}

    elif name == "wps_clear_all":
        client.clear_all(args["cell_or_range"], args.get("sheet_name"))
        result = {"message": f"Cleared all (contents + formats) from {args['cell_or_range']}"}

    # ── Conditional Formatting ──
    elif name == "wps_add_conditional_format":
        font_color = None
        bg_color = None
        if args.get("font_color"):
            font_color = int(args["font_color"].lstrip("#"), 16)
        if args.get("bg_color"):
            bg_color = int(args["bg_color"].lstrip("#"), 16)
        client.add_conditional_format(
            args["range_ref"],
            rule_type="cellValue",
            operator=args.get("operator", "greaterThan"),
            formula=args.get("formula", "0"),
            font_color=font_color,
            bg_color=bg_color,
            bold=args.get("bold"),
            sheet_name=args.get("sheet_name"),
        )
        result = {"message": f"Added conditional formatting to {args['range_ref']}"}

    # ── Data Validation ──
    elif name == "wps_add_data_validation":
        client.add_data_validation(
            args["range_ref"],
            validation_type=args.get("validation_type", "list"),
            formula1=args.get("formula1", ""),
            formula2=args.get("formula2", ""),
            ignore_blank=args.get("ignore_blank", True),
            show_dropdown=args.get("show_dropdown", True),
            error_title=args.get("error_title", ""),
            error_message=args.get("error_message", ""),
            sheet_name=args.get("sheet_name"),
        )
        result = {"message": f"Added data validation to {args['range_ref']} (type={args.get('validation_type', 'list')})"}

    # ── Sheet Protection ──
    elif name == "wps_protect_sheet":
        client.protect_sheet(
            password=args.get("password", ""),
            allow_sort=args.get("allow_sort", False),
            allow_filter=args.get("allow_filter", False),
            allow_format_cells=args.get("allow_format_cells", False),
            allow_insert_rows=args.get("allow_insert_rows", False),
            allow_delete_rows=args.get("allow_delete_rows", False),
            sheet_name=args.get("sheet_name"),
        )
        result = {"message": "Sheet protected."}

    elif name == "wps_unprotect_sheet":
        client.unprotect_sheet(
            password=args.get("password", ""),
            sheet_name=args.get("sheet_name"),
        )
        result = {"message": "Sheet unprotected."}

    # ── Page Setup ──
    elif name == "wps_set_print_area":
        client.set_print_area(args["range_ref"], args.get("sheet_name"))
        result = {"message": f"Set print area to {args['range_ref']}"}

    elif name == "wps_clear_print_area":
        client.clear_print_area(args.get("sheet_name"))
        result = {"message": "Cleared print area."}

    elif name == "wps_set_page_orientation":
        client.set_page_orientation(args["orientation"], args.get("sheet_name"))
        result = {"message": f"Set page orientation to {args['orientation']}"}

    # ── Formula ──
    elif name == "wps_set_formula":
        client.set_formula(args["cell_ref"], args["formula"], args.get("sheet_name"))
        result = {"message": f"Set {args['cell_ref']} formula = {args['formula']}"}

    elif name == "wps_get_formula":
        formula = client.get_formula(args["cell_ref"], args.get("sheet_name"))
        result = {"cell": args["cell_ref"], "formula": formula}

    # ── Export ──
    elif name == "wps_export_to_pdf":
        saved_path = client.export_to_pdf(args["filepath"], args.get("sheet_name"))
        result = {"message": f"Exported to PDF: {saved_path}", "path": saved_path}

    # ── Find / Replace ──
    elif name == "wps_find_replace":
        count = client.find_replace(
            args["find_text"],
            args["replace_text"],
            args.get("sheet_name"),
            args.get("match_case", False),
            args.get("match_whole", False),
        )
        result = {"message": f"Replaced {count} occurrences of '{args['find_text']}' with '{args['replace_text']}'", "replacements": count}

    # ── Workbook Activation ──
    elif name == "wps_activate_workbook":
        wb_name = client.activate_workbook(args["name"])
        result = {"message": f"Activated workbook: {wb_name}", "workbook_name": wb_name}

    # ── Remove Duplicates ──
    elif name == "wps_remove_duplicates":
        columns = None
        if args.get("columns"):
            columns = json.loads(args["columns"])
        client.remove_duplicates(
            args["range_ref"],
            columns,
            args.get("has_header", True),
            args.get("sheet_name"),
        )
        result = {"message": f"Removed duplicates from {args['range_ref']}"}

    # ── Vertical Alignment ──
    elif name == "wps_set_vertical_alignment":
        client.set_vertical_alignment(
            args["cell_or_range"],
            args["alignment"],
            args.get("sheet_name"),
        )
        result = {"message": f"Set {args['cell_or_range']} vertical alignment = {args['alignment']}"}

    # ── Sheet Copy / Move ──
    elif name == "wps_copy_sheet":
        new_name = client.copy_sheet(
            args["source_name"],
            args.get("new_name"),
            args.get("before"),
            args.get("after"),
        )
        result = {"message": f"Copied sheet '{args['source_name']}' to '{new_name}'", "sheet_name": new_name}

    elif name == "wps_move_sheet":
        moved_name = client.move_sheet(
            args["source_name"],
            args.get("before"),
            args.get("after"),
        )
        result = {"message": f"Moved sheet '{args['source_name']}'", "sheet_name": moved_name}

    # ── Show / Hide Sheet ──
    elif name == "wps_hide_sheet":
        client.hide_sheet(args["name"])
        result = {"message": f"Hid sheet: {args['name']}"}

    elif name == "wps_unhide_sheet":
        client.unhide_sheet(args["name"])
        result = {"message": f"Unhid sheet: {args['name']}"}

    # ── Hyperlinks ──
    elif name == "wps_add_hyperlink":
        client.add_hyperlink(
            args["cell_ref"],
            args["address"],
            args.get("text_to_display"),
            args.get("screen_tip"),
            args.get("sheet_name"),
        )
        result = {"message": f"Added hyperlink to {args['cell_ref']} -> {args['address']}"}

    elif name == "wps_remove_hyperlink":
        client.remove_hyperlink(args["cell_or_range"], args.get("sheet_name"))
        result = {"message": f"Removed hyperlinks from {args['cell_or_range']}"}

    # ── Conditional Formatting Delete ──
    elif name == "wps_delete_conditional_format":
        client.delete_conditional_format(args["range_ref"], args.get("sheet_name"))
        result = {"message": f"Removed conditional formatting from {args['range_ref']}"}

    # ── Font Underline ──
    elif name == "wps_set_font_underline":
        style = args.get("underline_style", "single")
        client.set_font_underline(args["cell_or_range"], style, args.get("sheet_name"))
        result = {"message": f"Set {args['cell_or_range']} underline = {style}"}

    # ── Row / Column Grouping ──
    elif name == "wps_group_rows":
        client.group_rows(args["start_row"], args["end_row"], args.get("sheet_name"))
        result = {"message": f"Grouped rows {args['start_row']} to {args['end_row']}"}

    elif name == "wps_ungroup_rows":
        client.ungroup_rows(args["start_row"], args["end_row"], args.get("sheet_name"))
        result = {"message": f"Ungrouped rows {args['start_row']} to {args['end_row']}"}

    elif name == "wps_group_columns":
        client.group_columns(args["start_col"], args["end_col"], args.get("sheet_name"))
        result = {"message": f"Grouped columns {args['start_col']} to {args['end_col']}"}

    elif name == "wps_ungroup_columns":
        client.ungroup_columns(args["start_col"], args["end_col"], args.get("sheet_name"))
        result = {"message": f"Ungrouped columns {args['start_col']} to {args['end_col']}"}

    # ── Page Margins / Headers ──
    elif name == "wps_set_page_margins":
        client.set_page_margins(
            left=args.get("left"),
            right=args.get("right"),
            top=args.get("top"),
            bottom=args.get("bottom"),
            header=args.get("header"),
            footer=args.get("footer"),
            sheet_name=args.get("sheet_name"),
        )
        result = {"message": "Set page margins."}

    elif name == "wps_set_header_footer":
        client.set_header_footer(
            left_header=args.get("left_header", ""),
            center_header=args.get("center_header", ""),
            right_header=args.get("right_header", ""),
            left_footer=args.get("left_footer", ""),
            center_footer=args.get("center_footer", ""),
            right_footer=args.get("right_footer", ""),
            sheet_name=args.get("sheet_name"),
        )
        result = {"message": "Set header and footer."}

    # ── Text to Columns ──
    elif name == "wps_text_to_columns":
        delim = args.get("delimiter", ",")
        client.text_to_columns(args["range_ref"], delim, args.get("sheet_name"))
        result = {"message": f"Split {args['range_ref']} by '{delim}'"}

    # ── Named Ranges ──
    elif name == "wps_create_named_range":
        name = client.create_named_range(
            args["name"],
            args["refers_to"],
            args.get("sheet_name"),
        )
        result = {"message": f"Created named range: {name}", "name": name}

    elif name == "wps_delete_named_range":
        client.delete_named_range(args["name"])
        result = {"message": f"Deleted named range: {args['name']}"}

    elif name == "wps_list_named_ranges":
        result = {"named_ranges": client.list_named_ranges()}

    # ── Pivot Table ──
    elif name == "wps_create_pivot_table":
        row_fields = json.loads(args["row_fields"]) if args.get("row_fields") else None
        col_fields = json.loads(args["column_fields"]) if args.get("column_fields") else None
        data_fields = json.loads(args["data_fields"]) if args.get("data_fields") else None
        pt_name = client.create_pivot_table(
            args["source_range"],
            args["dest_cell"],
            args.get("pivot_name", "PivotTable1"),
            row_fields,
            col_fields,
            data_fields,
            args.get("sheet_name"),
        )
        result = {"message": f"Created pivot table: {pt_name}", "pivot_name": pt_name}

    # ── Sparklines ──
    elif name == "wps_add_sparkline":
        client.add_sparkline(
            args["source_range"],
            args["dest_cell"],
            args.get("spark_type", "line"),
            args.get("sheet_name"),
        )
        result = {"message": f"Added {args.get('spark_type', 'line')} sparkline in {args['dest_cell']}"}

    # ── Insert Picture / Shape ──
    elif name == "wps_insert_picture":
        pic_name = client.insert_picture(
            args["filepath"],
            args.get("left", 100),
            args.get("top", 100),
            args.get("width", 200),
            args.get("height", 150),
            args.get("sheet_name"),
        )
        result = {"message": f"Inserted picture: {pic_name}", "shape_name": pic_name}

    elif name == "wps_insert_shape":
        shape_name = client.insert_shape(
            args.get("shape_type", "rectangle"),
            args.get("left", 100),
            args.get("top", 100),
            args.get("width", 200),
            args.get("height", 100),
            args.get("sheet_name"),
        )
        result = {"message": f"Inserted shape: {shape_name}", "shape_name": shape_name}

    # ── Gridlines ──
    elif name == "wps_toggle_gridlines":
        visible = args.get("visible", True)
        client.toggle_gridlines(visible, args.get("sheet_name"))
        result = {"message": f"Gridlines {'shown' if visible else 'hidden'}."}

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
    """Entry point for the wps-excel-mcp command."""
    asyncio.run(run_server())


if __name__ == "__main__":
    main()
