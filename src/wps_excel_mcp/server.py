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
    # PyInstaller standalone: wps_client is bundled as wps_excel_mcp.wps_client
    from wps_excel_mcp.wps_client import WPSExcelClient

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
            "required": []
        },
    ),
    Tool(
        name="wps_create_workbook",
        description="Create a new Excel workbook in WPS Office.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": []
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
                    "description": "Path to .xlsx/.xls/.et file"
                }
            },
            "required": ["filepath"]
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
                    "description": "Save path (default: current location)"
                }
            },
            "required": []
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
                    "description": "Save before closing"
                }
            },
            "required": []
        },
    ),
    Tool(
        name="wps_list_workbooks",
        description="List all currently open workbooks.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": []
        },
    ),
    Tool(
        name="wps_list_sheets",
        description="List all sheets in the active workbook with their names, types, and visibility.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": []
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
                    "description": "Sheet name"
                }
            },
            "required": []
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
                    "description": "Current sheet name"
                },
                "new_name": {
                    "type": "string",
                    "description": "New sheet name"
                }
            },
            "required": ["old_name", "new_name"]
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
                    "description": "Sheet name"
                }
            },
            "required": ["name"]
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
                    "description": "Sheet name"
                }
            },
            "required": ["name"]
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
                    "description": "Cell ref, e.g. 'A1'"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name (default: active)"
                }
            },
            "required": ["cell_ref"]
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
                    "description": "Cell ref, e.g. 'A1'"
                },
                "value": {
                    "type": "string",
                    "description": "Value (numbers auto-converted)"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name (default: active)"
                }
            },
            "required": ["cell_ref", "value"]
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
                    "description": "Range, e.g. 'A1:B10'"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name (default: active)"
                }
            },
            "required": ["range_ref"]
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
                    "description": "Top-left cell, e.g. 'A1'"
                },
                "values": {
                    "type": "string",
                    "description": "JSON 2D array, e.g. [['Name','Age'],['John',30]]"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name (default: active)"
                }
            },
            "required": ["start_cell", "values"]
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
                    "description": "Cell reference (e.g., 'A1') or range (e.g., 'A1:B10')."
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name (default: active)"
                }
            },
            "required": ["cell_or_range"]
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
                    "description": "Range, e.g. 'A1:C1'"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["range_ref"]
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
                    "description": "Range, e.g. 'A1:C1'"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["range_ref"]
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
                    "description": "Sheet name"
                }
            },
            "required": []
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
                    "description": "column/line/pie/bar/area/scatter"
                },
                "range_ref": {
                    "type": "string",
                    "description": "Data range, e.g. 'A1:B10'"
                },
                "left": {
                    "type": "number",
                    "description": "Left position in points."
                },
                "top": {
                    "type": "number",
                    "description": "Top position in points."
                },
                "width": {
                    "type": "number",
                    "description": "Width (points)"
                },
                "height": {
                    "type": "number",
                    "description": "Height (points)"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["range_ref"]
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
                    "description": "Search text"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                },
                "match_whole": {
                    "type": "boolean",
                    "description": "If True, match whole cell content. If False (default), match partial."
                }
            },
            "required": ["search_text"]
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
                    "description": "Macro name"
                }
            },
            "required": ["macro_name"]
        },
    ),
    Tool(
        name="wps_show_window",
        description="Make the WPS Excel application window visible.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": []
        },
    ),
    Tool(
        name="wps_hide_window",
        description="Hide the WPS Excel application window (run in background).",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": []
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
                    "description": "Cell/range, e.g. 'A1' or 'A1:C10'"
                },
                "border_style": {
                    "type": "string",
                    "description": "thin/medium/thick/dotted/dashed/double/hairline/none"
                },
                "border_color": {
                    "type": "string",
                    "description": "Border RGB hex"
                },
                "outline_only": {
                    "type": "boolean",
                    "description": "Outline only (not inner grid)"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["cell_or_range"]
        },
    ),
    # ── Freeze / Filter ──
            Tool(
        name="wps_freeze_panes",
        description="Freeze panes at a specific cell. Rows above and columns to the left of the cell are frozen.",
        inputSchema={
            "type": "object",
            "properties": {
                "cell_ref": {
                    "type": "string",
                    "description": "Freeze cell (B2=both, A2=row, B1=col)"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": []
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
                    "description": "Sheet name"
                }
            },
            "required": []
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
                    "description": "Range (default: used range)"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": []
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
                    "description": "Range, e.g. 'A2:D100'"
                },
                "sort_key": {
                    "type": "string",
                    "description": "Sort key cell, e.g. 'A2'"
                },
                "sort_order": {
                    "type": "string",
                    "description": "ascending/descending"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["range_ref"]
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
                    "description": "Range, e.g. 'A1:D10'"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["range_ref"]
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
                    "description": "Dest cell, e.g. 'A1'"
                },
                "paste_special": {
                    "type": "string",
                    "description": "values/formats/formulas/all/transpose"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["dest_cell"]
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
                    "description": "Sheet name"
                }
            },
            "required": []
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
                    "description": "Cell ref, e.g. 'A1'"
                },
                "text": {
                    "type": "string",
                    "description": "Comment text"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["cell_ref", "text"]
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
                    "description": "Cell ref, e.g. 'A1'"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["cell_ref"]
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
                    "description": "Cell/range, e.g. 'A1' or 'A1:C10'"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["cell_or_range"]
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
                    "description": "Cell/range, e.g. 'A1' or 'A1:C10'"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["cell_or_range"]
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
                    "description": "Range, e.g. 'B2:B100'"
                },
                "operator": {
                    "type": "string",
                    "description": "Operator"
                },
                "formula": {
                    "type": "string",
                    "description": "Threshold/formula"
                },
                "font_color": {
                    "type": "string",
                    "description": "Font RGB hex"
                },
                "bg_color": {
                    "type": "string",
                    "description": "Background RGB hex"
                },
                "bold": {
                    "type": "boolean",
                    "description": "Bold"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["range_ref"]
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
                    "description": "Range, e.g. 'C2:C100'"
                },
                "validation_type": {
                    "type": "string",
                    "description": "list/whole/decimal/date/time/textLength/custom"
                },
                "formula1": {
                    "type": "string",
                    "description": "Formula (list: 'A,B,C' or '=$A$1:$A$10')"
                },
                "formula2": {
                    "type": "string",
                    "description": "Second formula (max)"
                },
                "ignore_blank": {
                    "type": "boolean",
                    "description": "Allow blanks"
                },
                "show_dropdown": {
                    "type": "boolean",
                    "description": "Show dropdown arrow"
                },
                "error_title": {
                    "type": "string",
                    "description": "Error dialog title"
                },
                "error_message": {
                    "type": "string",
                    "description": "Error message"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["range_ref"]
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
                    "description": "Password"
                },
                "allow_sort": {
                    "type": "boolean",
                    "description": "Allow sort"
                },
                "allow_filter": {
                    "type": "boolean",
                    "description": "Allow filter"
                },
                "allow_format_cells": {
                    "type": "boolean",
                    "description": "Allow format cells"
                },
                "allow_insert_rows": {
                    "type": "boolean",
                    "description": "Allow insert rows"
                },
                "allow_delete_rows": {
                    "type": "boolean",
                    "description": "Allow delete rows"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": []
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
                    "description": "Password"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": []
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
                    "description": "Range, e.g. 'A1:F50'"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["range_ref"]
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
                    "description": "Sheet name"
                }
            },
            "required": []
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
                    "description": "portrait/landscape"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["orientation"]
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
                    "description": "Cell ref, e.g. 'A1'"
                },
                "formula": {
                    "type": "string",
                    "description": "Formula, e.g. '=SUM(B2:B10)'"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["cell_ref", "formula"]
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
                    "description": "Cell ref, e.g. 'A1'"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["cell_ref"]
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
                    "description": "Full path for the output PDF file (e.g., 'C:\\report.pdf')."
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Optional sheet name. If omitted, exports the entire workbook."
                }
            },
            "required": ["filepath"]
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
                    "description": "Search text"
                },
                "replace_text": {
                    "type": "string",
                    "description": "Replacement text"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                },
                "match_case": {
                    "type": "boolean",
                    "description": "Match case"
                },
                "match_whole": {
                    "type": "boolean",
                    "description": "Match whole cell only"
                }
            },
            "required": ["find_text", "replace_text"]
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
                    "description": "Workbook name"
                }
            },
            "required": ["name"]
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
                    "description": "Range, e.g. 'A1:D100'"
                },
                "columns": {
                    "type": "string",
                    "description": "Column indices (JSON array, 1-based)"
                },
                "has_header": {
                    "type": "boolean",
                    "description": "Has header row"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["range_ref"]
        },
    ),
    # ── Vertical Alignment ──
        # ── Sheet Copy / Move ──
    Tool(
        name="wps_copy_sheet",
        description="Create a copy of a worksheet within the active workbook.",
        inputSchema={
            "type": "object",
            "properties": {
                "source_name": {
                    "type": "string",
                    "description": "Sheet name"
                },
                "new_name": {
                    "type": "string",
                    "description": "New sheet name"
                },
                "before": {
                    "type": "string",
                    "description": "Insert before sheet"
                },
                "after": {
                    "type": "string",
                    "description": "Insert after sheet"
                }
            },
            "required": ["source_name"]
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
                    "description": "Sheet name"
                },
                "before": {
                    "type": "string",
                    "description": "Move before sheet"
                },
                "after": {
                    "type": "string",
                    "description": "Move after sheet"
                }
            },
            "required": ["source_name"]
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
                    "description": "Sheet name"
                }
            },
            "required": ["name"]
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
                    "description": "Sheet name"
                }
            },
            "required": ["name"]
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
                    "description": "Cell ref, e.g. 'A1'"
                },
                "address": {
                    "type": "string",
                    "description": "URL, file path, or cell ref"
                },
                "text_to_display": {
                    "type": "string",
                    "description": "Display text"
                },
                "screen_tip": {
                    "type": "string",
                    "description": "Tooltip"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["cell_ref", "address"]
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
                    "description": "Cell/range, e.g. 'A1' or 'A1:D10'"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["cell_or_range"]
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
                    "description": "Range, e.g. 'A1:A100'"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["range_ref"]
        },
    ),
        # ── Row / Column Grouping ──
                    # ── Page Margins / Headers ──
    Tool(
        name="wps_set_page_margins",
        description="Set page margins (in points) for a sheet.",
        inputSchema={
            "type": "object",
            "properties": {
                "left": {
                    "type": "number",
                    "description": "Left margin (points)"
                },
                "right": {
                    "type": "number",
                    "description": "Right margin (points)"
                },
                "top": {
                    "type": "number",
                    "description": "Top margin (points)"
                },
                "bottom": {
                    "type": "number",
                    "description": "Bottom margin (points)"
                },
                "header": {
                    "type": "number",
                    "description": "Header margin (points)"
                },
                "footer": {
                    "type": "number",
                    "description": "Footer margin (points)"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": []
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
                    "description": "Left header"
                },
                "center_header": {
                    "type": "string",
                    "description": "Center header"
                },
                "right_header": {
                    "type": "string",
                    "description": "Right header"
                },
                "left_footer": {
                    "type": "string",
                    "description": "Left footer"
                },
                "center_footer": {
                    "type": "string",
                    "description": "Center footer"
                },
                "right_footer": {
                    "type": "string",
                    "description": "Right footer"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": []
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
                    "description": "Range, e.g. 'A1:A100'"
                },
                "delimiter": {
                    "type": "string",
                    "description": "Delimiter character: ',' (comma), ';' (semicolon), '\\t' (tab), ' ' (space), '|' (pipe). Default: ','."
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["range_ref"]
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
                    "description": "Name"
                },
                "refers_to": {
                    "type": "string",
                    "description": "Range formula, e.g. '=Sheet1!$A$1:$D$10'"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name (scope)"
                }
            },
            "required": ["name", "refers_to"]
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
                    "description": "Name"
                }
            },
            "required": ["name"]
        },
    ),
    Tool(
        name="wps_list_named_ranges",
        description="List all named ranges in the active workbook.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": []
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
                    "description": "Source range, e.g. 'A1:F100'"
                },
                "dest_cell": {
                    "type": "string",
                    "description": "Dest cell, e.g. 'H1'"
                },
                "pivot_name": {
                    "type": "string",
                    "description": "Pivot table name"
                },
                "row_fields": {
                    "type": "string",
                    "description": "Row field names (JSON array)"
                },
                "column_fields": {
                    "type": "string",
                    "description": "Column field names (JSON array)"
                },
                "data_fields": {
                    "type": "string",
                    "description": "Value field names (JSON array)"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name (source & dest)"
                }
            },
            "required": ["source_range", "dest_cell"]
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
                    "description": "Data range, e.g. 'A1:A10'"
                },
                "dest_cell": {
                    "type": "string",
                    "description": "Dest cell"
                },
                "spark_type": {
                    "type": "string",
                    "description": "line/column/winloss"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["source_range", "dest_cell"]
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
                    "description": "Image file path"
                },
                "left": {
                    "type": "number",
                    "description": "Left (points)"
                },
                "top": {
                    "type": "number",
                    "description": "Top (points)"
                },
                "width": {
                    "type": "number",
                    "description": "Width (points)"
                },
                "height": {
                    "type": "number",
                    "description": "Height (points)"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["filepath"]
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
                    "description": "rectangle/oval/line/arrow/textbox"
                },
                "left": {
                    "type": "number",
                    "description": "Left (points)"
                },
                "top": {
                    "type": "number",
                    "description": "Top (points)"
                },
                "width": {
                    "type": "number",
                    "description": "Width (points)"
                },
                "height": {
                    "type": "number",
                    "description": "Height (points)"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": []
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
                    "description": "Show gridlines"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": []
        },
    ),
    # ── Consolidated Formatting ──
    Tool(
        name="wps_format_font",
        description="Set font properties (bold, italic, name, size, color, underline) for a cell or range.",
        inputSchema={
            "type": "object",
            "properties": {
                "cell_or_range": {
                    "type": "string",
                    "description": "Cell/range, e.g. 'A1' or 'A1:C10'"
                },
                "bold": {
                    "type": "boolean",
                    "description": "Bold on/off"
                },
                "italic": {
                    "type": "boolean",
                    "description": "Italic on/off"
                },
                "font_name": {
                    "type": "string",
                    "description": "Font name, e.g. 'Arial'"
                },
                "size": {
                    "type": "integer",
                    "description": "Font size (points)"
                },
                "color": {
                    "type": "string",
                    "description": "RGB hex, e.g. 'FF0000'=red"
                },
                "underline": {
                    "type": "string",
                    "description": "none/single/double/singleAccounting/doubleAccounting"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["cell_or_range"]
        },
    ),
    Tool(
        name="wps_format_cell",
        description="Set cell appearance (fill color, alignment, number format, wrap text) for a cell or range.",
        inputSchema={
            "type": "object",
            "properties": {
                "cell_or_range": {
                    "type": "string",
                    "description": "Cell/range, e.g. 'A1' or 'A1:C10'"
                },
                "fill_color": {
                    "type": "string",
                    "description": "RGB hex, e.g. 'FF0000'=red"
                },
                "alignment": {
                    "type": "string",
                    "description": "left/center/right/general/justify"
                },
                "vertical_alignment": {
                    "type": "string",
                    "description": "top/center/bottom/justify/distributed"
                },
                "number_format": {
                    "type": "string",
                    "description": "Format string, e.g. '0.00', '#,##0', 'yyyy-mm-dd'"
                },
                "wrap_text": {
                    "type": "boolean",
                    "description": "Wrap text on/off"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["cell_or_range"]
        },
    ),
    # ── Consolidated Row/Col ──
    Tool(
        name="wps_insert",
        description="Insert a row or column at the specified position.",
        inputSchema={
            "type": "object",
            "properties": {
                "position": {
                    "type": "integer",
                    "description": "Row or column number (1=A for columns)"
                },
                "insert_type": {
                    "type": "string",
                    "description": "'row' or 'column'"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["position", "insert_type"]
        },
    ),
    Tool(
        name="wps_delete",
        description="Delete a row or column at the specified position.",
        inputSchema={
            "type": "object",
            "properties": {
                "position": {
                    "type": "integer",
                    "description": "Row or column number (1=A for columns)"
                },
                "delete_type": {
                    "type": "string",
                    "description": "'row' or 'column'"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["position", "delete_type"]
        },
    ),
    Tool(
        name="wps_set_dimensions",
        description="Set row height or column width.",
        inputSchema={
            "type": "object",
            "properties": {
                "position": {
                    "type": "integer",
                    "description": "Row or column number (1=A for columns)"
                },
                "size": {
                    "type": "number",
                    "description": "Height (points) for rows, width (chars) for columns"
                },
                "dimension_type": {
                    "type": "string",
                    "description": "'row' or 'column'"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["position", "size", "dimension_type"]
        },
    ),
    Tool(
        name="wps_group",
        description="Group rows or columns for outlining (collapse/expand).",
        inputSchema={
            "type": "object",
            "properties": {
                "start": {
                    "type": "integer",
                    "description": "First row or column (1=A for columns)"
                },
                "end": {
                    "type": "integer",
                    "description": "Last row or column"
                },
                "group_type": {
                    "type": "string",
                    "description": "'row' or 'column'"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["start", "end", "group_type"]
        },
    ),
    Tool(
        name="wps_ungroup",
        description="Ungroup previously grouped rows or columns.",
        inputSchema={
            "type": "object",
            "properties": {
                "start": {
                    "type": "integer",
                    "description": "First row or column (1=A for columns)"
                },
                "end": {
                    "type": "integer",
                    "description": "Last row or column"
                },
                "group_type": {
                    "type": "string",
                    "description": "'row' or 'column'"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": ["start", "end", "group_type"]
        },
    ),
    Tool(
        name="wps_autofit",
        description="Auto-fit column widths or row heights to content.",
        inputSchema={
            "type": "object",
            "properties": {
                "fit_type": {
                    "type": "string",
                    "description": "'columns' or 'rows'"
                },
                "start": {
                    "type": "integer",
                    "description": "First col/row (1=A for cols)"
                },
                "end": {
                    "type": "integer",
                    "description": "Last col/row"
                },
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet name"
                }
            },
            "required": []
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
            "tool": name
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

    # ── Consolidated Formatting ──
    elif name == "wps_format_font":
        sheet = args.get("sheet_name")
        cr = args["cell_or_range"]
        if "bold" in args:
            client.set_font_bold(cr, args["bold"], sheet)
        if "italic" in args:
            client.set_font_italic(cr, args["italic"], sheet)
        if "font_name" in args:
            client.set_font_name(cr, args["font_name"], sheet)
        if "size" in args:
            client.set_font_size(cr, args["size"], sheet)
        if "color" in args:
            color_int = int(args["color"].lstrip("#"), 16)
            client.set_font_color(cr, color_int, sheet)
        if "underline" in args:
            client.set_font_underline(cr, args["underline"], sheet)
        result = {"message": f"Formatted font on {cr}"}

    elif name == "wps_format_cell":
        sheet = args.get("sheet_name")
        cr = args["cell_or_range"]
        if "fill_color" in args:
            color_int = int(args["fill_color"].lstrip("#"), 16)
            client.set_cell_color(cr, color_int, sheet)
        if "alignment" in args:
            client.set_horizontal_alignment(cr, args["alignment"], sheet)
        if "vertical_alignment" in args:
            client.set_vertical_alignment(cr, args["vertical_alignment"], sheet)
        if "number_format" in args:
            client.set_number_format(cr, args["number_format"], sheet)
        if "wrap_text" in args:
            client.set_wrap_text(cr, args["wrap_text"], sheet)
        result = {"message": f"Formatted cell {cr}"}

    # ── Backward compat: old font/cell format names ──
    elif name == "wps_set_font_bold":
        client.set_font_bold(args["cell_or_range"], args.get("bold", True), args.get("sheet_name"))
        result = {"message": f"Set {args['cell_or_range']} font bold = {args.get('bold', True)}"}
    elif name == "wps_set_font_size":
        client.set_font_size(args["cell_or_range"], args["size"], args.get("sheet_name"))
        result = {"message": f"Set {args['cell_or_range']} font size = {args['size']}"}
    elif name == "wps_set_cell_color":
        color_hex = args["color"].lstrip("#")
        client.set_cell_color(args["cell_or_range"], int(color_hex, 16), args.get("sheet_name"))
        result = {"message": f"Set {args['cell_or_range']} fill color = #{color_hex}"}
    elif name == "wps_set_alignment":
        client.set_horizontal_alignment(args["cell_or_range"], args["alignment"], args.get("sheet_name"))
        result = {"message": f"Set {args['cell_or_range']} alignment = {args['alignment']}"}
    elif name == "wps_set_number_format":
        client.set_number_format(args["cell_or_range"], args["format"], args.get("sheet_name"))
        result = {"message": f"Set {args['cell_or_range']} number format = '{args['format']}'"}
    elif name == "wps_set_font_name":
        client.set_font_name(args["cell_or_range"], args["font_name"], args.get("sheet_name"))
        result = {"message": f"Set {args['cell_or_range']} font name = '{args['font_name']}'"}
    elif name == "wps_set_font_italic":
        client.set_font_italic(args["cell_or_range"], args.get("italic", True), args.get("sheet_name"))
        result = {"message": f"Set {args['cell_or_range']} font italic = {args.get('italic', True)}"}
    elif name == "wps_set_font_color":
        color_hex = args["color"].lstrip("#")
        client.set_font_color(args["cell_or_range"], int(color_hex, 16), args.get("sheet_name"))
        result = {"message": f"Set {args['cell_or_range']} font color = #{color_hex}"}
    elif name == "wps_set_wrap_text":
        client.set_wrap_text(args["cell_or_range"], args.get("wrap", True), args.get("sheet_name"))
        result = {"message": f"Set {args['cell_or_range']} wrap text = {args.get('wrap', True)}"}
    elif name == "wps_set_vertical_alignment":
        client.set_vertical_alignment(args["cell_or_range"], args["alignment"], args.get("sheet_name"))
        result = {"message": f"Set {args['cell_or_range']} vertical alignment = {args['alignment']}"}
    elif name == "wps_set_font_underline":
        client.set_font_underline(args["cell_or_range"], args.get("underline_style", "single"), args.get("sheet_name"))
        result = {"message": f"Set {args['cell_or_range']} underline = {args.get('underline_style', 'single')}"}

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
            "column_count": cols
        }

    elif name == "wps_insert":
        pos = args["position"]
        sheet = args.get("sheet_name")
        if args["insert_type"] == "row":
            client.insert_row(pos, sheet)
            result = {"message": f"Inserted row at position {pos}"}
        else:
            client.insert_column(pos, sheet)
            result = {"message": f"Inserted column at position {pos}"}

    elif name == "wps_delete":
        pos = args["position"]
        sheet = args.get("sheet_name")
        if args["delete_type"] == "row":
            client.delete_row(pos, sheet)
            result = {"message": f"Deleted row {pos}"}
        else:
            client.delete_column(pos, sheet)
            result = {"message": f"Deleted column {pos}"}

    elif name == "wps_set_dimensions":
        pos = args["position"]
        sheet = args.get("sheet_name")
        if args["dimension_type"] == "row":
            client.set_row_height(pos, args["size"], sheet)
            result = {"message": f"Set row {pos} height = {args['size']}"}
        else:
            client.set_column_width(pos, args["size"], sheet)
            result = {"message": f"Set column {pos} width = {args['size']}"}

    # ── Backward compat: old insert/delete/size names ──
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

    # ── Freeze / Filter ──
    elif name == "wps_autofit":
        sheet = args.get("sheet_name")
        if args.get("fit_type") == "rows":
            client.autofit_rows(sheet, args.get("start"), args.get("end"))
            result = {"message": "Auto-fitted row heights."}
        else:
            client.autofit_columns(sheet, args.get("start"), args.get("end"))
            result = {"message": "Auto-fitted column widths."}

    # ── Backward compat: old autofit names ──
    elif name == "wps_autofit_columns":
        client.autofit_columns(args.get("sheet_name"), args.get("start_col"), args.get("end_col"))
        result = {"message": "Auto-fitted column widths."}
    elif name == "wps_autofit_rows":
        client.autofit_rows(args.get("sheet_name"), args.get("start_row"), args.get("end_row"))
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

    elif name == "wps_set_font_underline":
        style = args.get("underline_style", "single")
        client.set_font_underline(args["cell_or_range"], style, args.get("sheet_name"))
        result = {"message": f"Set {args['cell_or_range']} underline = {style}"}

    # ── Row / Column Grouping ──
    elif name == "wps_group":
        sheet = args.get("sheet_name")
        if args["group_type"] == "row":
            client.group_rows(args["start"], args["end"], sheet)
            result = {"message": f"Grouped rows {args['start']} to {args['end']}"}
        else:
            client.group_columns(args["start"], args["end"], sheet)
            result = {"message": f"Grouped columns {args['start']} to {args['end']}"}

    elif name == "wps_ungroup":
        sheet = args.get("sheet_name")
        if args["group_type"] == "row":
            client.ungroup_rows(args["start"], args["end"], sheet)
            result = {"message": f"Ungrouped rows {args['start']} to {args['end']}"}
        else:
            client.ungroup_columns(args["start"], args["end"], sheet)
            result = {"message": f"Ungrouped columns {args['start']} to {args['end']}"}

    # ── Backward compat: old group/ungroup names ──
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
