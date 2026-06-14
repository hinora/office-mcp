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


TOOLS: list[Tool] = [
    Tool(name="wps_app", description="WPS Excel app control. action: info/show/hide",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["info","show","hide"]}},"required":["action"]}),
    Tool(name="wps_workbook", description="Manage workbooks. action: create/open/save/close/list/activate",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["create","open","save","close","list","activate"]},"filepath":{"type":"string"},"save":{"type":"boolean"},"name":{"type":"string"}},"required":["action"]}),
    Tool(name="wps_sheet", description="Manage sheets. action: list/add/rename/delete/activate/copy/move/hide/unhide",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["list","add","rename","delete","activate","copy","move","hide","unhide"]},"name":{"type":"string"},"new_name":{"type":"string"},"old_name":{"type":"string"},"source_name":{"type":"string"},"before":{"type":"string"},"after":{"type":"string"}},"required":["action"]}),
    Tool(name="wps_cell", description="Single cell ops. action: get/set/clear/set_formula/get_formula",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["get","set","clear","set_formula","get_formula"]},"cell_ref":{"type":"string","description":"e.g. A1"},"value":{"type":"string"},"formula":{"type":"string","description":"e.g. =SUM(B2:B10)"},"sheet_name":{"type":"string"}},"required":["action","cell_ref"]}),
    Tool(name="wps_range", description="Multi-cell ops. action: get/set/clear/copy/paste/sort/find/find_next/find_replace/used_range/remove_duplicates",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["get","set","clear","copy","paste","sort","find","find_next","find_replace","used_range","remove_duplicates"]},"range_ref":{"type":"string"},"start_cell":{"type":"string"},"values":{"type":"string","description":"JSON 2D array"},"dest_cell":{"type":"string"},"paste_special":{"type":"string","description":"values/formats/formulas/all/transpose"},"sort_key":{"type":"string"},"sort_order":{"type":"string"},"search_text":{"type":"string"},"find_text":{"type":"string"},"replace_text":{"type":"string"},"match_case":{"type":"boolean"},"match_whole":{"type":"boolean"},"columns":{"type":"string","description":"JSON array for remove_duplicates"},"has_header":{"type":"boolean"},"sheet_name":{"type":"string"}},"required":["action"]}),
    Tool(name="wps_format", description="Format cells: font, borders, fill, alignment, number format, merge, conditional format, wrap, data validation.",
        inputSchema={"type":"object","properties":{"cell_or_range":{"type":"string"},"bold":{"type":"boolean"},"italic":{"type":"boolean"},"font_name":{"type":"string"},"font_size":{"type":"integer"},"font_color":{"type":"string","description":"RGB hex"},"underline":{"type":"string","description":"none/single/double"},"fill_color":{"type":"string","description":"RGB hex"},"alignment":{"type":"string","description":"left/center/right"},"vertical_alignment":{"type":"string","description":"top/center/bottom"},"number_format":{"type":"string"},"wrap_text":{"type":"boolean"},"merge":{"type":"boolean","description":"True=merge, False=unmerge"},"border_style":{"type":"string","description":"thin/medium/thick"},"border_color":{"type":"string"},"outline_only":{"type":"boolean"},"cond_operator":{"type":"string"},"cond_formula":{"type":"string"},"cond_font_color":{"type":"string"},"cond_bg_color":{"type":"string"},"cond_bold":{"type":"boolean"},"clear_cond":{"type":"boolean"},"validation_type":{"type":"string"},"formula1":{"type":"string"},"formula2":{"type":"string"},"ignore_blank":{"type":"boolean"},"show_dropdown":{"type":"boolean"},"error_title":{"type":"string"},"error_message":{"type":"string"},"sheet_name":{"type":"string"}},"required":["cell_or_range"]}),
    Tool(name="wps_rowcol", description="Row/col ops. action: insert/delete/resize/autofit/group/ungroup/freeze/unfreeze",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["insert","delete","resize","autofit","group","ungroup","freeze","unfreeze"]},"type":{"type":"string","description":"row or column"},"position":{"type":"integer"},"size":{"type":"number"},"start":{"type":"integer"},"end":{"type":"integer"},"cell_ref":{"type":"string","description":"Freeze cell"},"sheet_name":{"type":"string"}},"required":["action"]}),
    Tool(name="wps_chart", description="Add a chart.",
        inputSchema={"type":"object","properties":{"chart_type":{"type":"string","description":"column/line/pie/bar/area/scatter"},"range_ref":{"type":"string"},"left":{"type":"number"},"top":{"type":"number"},"width":{"type":"number"},"height":{"type":"number"},"sheet_name":{"type":"string"}},"required":["range_ref"]}),
    Tool(name="wps_data", description="Data ops. action: filter/text_to_columns",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["filter","text_to_columns"]},"range_ref":{"type":"string"},"delimiter":{"type":"string"},"sheet_name":{"type":"string"}},"required":["action"]}),
    Tool(name="wps_protection", description="Sheet protection. action: protect/unprotect",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["protect","unprotect"]},"password":{"type":"string"},"allow_sort":{"type":"boolean"},"allow_filter":{"type":"boolean"},"allow_format_cells":{"type":"boolean"},"allow_insert_rows":{"type":"boolean"},"allow_delete_rows":{"type":"boolean"},"sheet_name":{"type":"string"}},"required":["action"]}),
    Tool(name="wps_page_setup", description="Page layout. action: print_area/clear_print_area/orientation/margins/header_footer",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["print_area","clear_print_area","orientation","margins","header_footer"]},"range_ref":{"type":"string"},"orientation":{"type":"string"},"left":{"type":"number"},"right":{"type":"number"},"top":{"type":"number"},"bottom":{"type":"number"},"header":{"type":"number"},"footer":{"type":"number"},"left_header":{"type":"string"},"center_header":{"type":"string"},"right_header":{"type":"string"},"left_footer":{"type":"string"},"center_footer":{"type":"string"},"right_footer":{"type":"string"},"sheet_name":{"type":"string"}},"required":["action"]}),
    Tool(name="wps_macro", description="Run a VBA macro.", inputSchema={"type":"object","properties":{"macro_name":{"type":"string"}},"required":["macro_name"]}),
    Tool(name="wps_export", description="Export to PDF.", inputSchema={"type":"object","properties":{"filepath":{"type":"string"},"sheet_name":{"type":"string"}},"required":["filepath"]}),
    Tool(name="wps_misc", description="Misc ops: named ranges, hyperlinks, comments, pivot, sparkline, picture, shape, gridlines. action: create_named_range/delete_named_range/list_named_ranges/add_hyperlink/remove_hyperlink/add_comment/delete_comment/create_pivot_table/add_sparkline/insert_picture/insert_shape/toggle_gridlines",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["create_named_range","delete_named_range","list_named_ranges","add_hyperlink","remove_hyperlink","add_comment","delete_comment","create_pivot_table","add_sparkline","insert_picture","insert_shape","toggle_gridlines"]},"name":{"type":"string"},"refers_to":{"type":"string"},"cell_ref":{"type":"string"},"cell_or_range":{"type":"string"},"address":{"type":"string"},"text_to_display":{"type":"string"},"screen_tip":{"type":"string"},"text":{"type":"string"},"source_range":{"type":"string"},"dest_cell":{"type":"string"},"pivot_name":{"type":"string"},"row_fields":{"type":"string"},"column_fields":{"type":"string"},"data_fields":{"type":"string"},"spark_type":{"type":"string"},"filepath":{"type":"string"},"left":{"type":"number"},"top":{"type":"number"},"width":{"type":"number"},"height":{"type":"number"},"shape_type":{"type":"string"},"visible":{"type":"boolean"},"sheet_name":{"type":"string"}},"required":["action"]}),
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



def _parse_value(val: str):
    """Parse a string value to int/float if possible."""
    try:
        if "." in val: return float(val)
        return int(val)
    except (ValueError, TypeError): return val

def _execute_tool(name: str, args: dict[str, Any], client: WPSExcelClient) -> str:
    pythoncom.CoInitialize()
    result: Any = None
    sheet = args.get("sheet_name")

    if name == "wps_app":
        a = args["action"]
        if a == "info": result = client.get_app_info()
        elif a == "show": client.show(); result = {"message": "WPS Excel window visible."}
        elif a == "hide": client.hide(); result = {"message": "WPS Excel window hidden."}

    elif name == "wps_workbook":
        a = args["action"]
        if a == "create": n = client.create_workbook(); result = {"message": f"Created: {n}", "workbook_name": n}
        elif a == "open": n = client.open_workbook(args["filepath"]); result = {"message": f"Opened: {n}", "workbook_name": n}
        elif a == "save": p = client.save_workbook(args.get("filepath")); result = {"message": f"Saved to: {p}", "path": p}
        elif a == "close": s = args.get("save", True); client.close_workbook(s); result = {"message": "Closed." + (" (saved)" if s else "")}
        elif a == "list": result = {"workbooks": client.list_workbooks()}
        elif a == "activate": n = client.activate_workbook(args["name"]); result = {"message": f"Activated: {n}"}

    elif name == "wps_sheet":
        a = args["action"]
        if a == "list": result = {"sheets": client.list_sheets()}
        elif a == "add": n = client.add_sheet(args.get("name")); result = {"message": f"Added: {n}", "sheet_name": n}
        elif a == "rename": n = client.rename_sheet(args["old_name"], args["new_name"]); result = {"message": f"Renamed to {n}"}
        elif a == "delete": client.delete_sheet(args["name"]); result = {"message": f"Deleted: {args['name']}"}
        elif a == "activate": n = client.activate_sheet(args["name"]); result = {"message": f"Activated: {n}"}
        elif a == "copy": n = client.copy_sheet(args["source_name"], args.get("new_name"), args.get("before"), args.get("after")); result = {"message": f"Copied to: {n}"}
        elif a == "move": n = client.move_sheet(args["source_name"], args.get("before"), args.get("after")); result = {"message": f"Moved: {n}"}
        elif a == "hide": client.hide_sheet(args["name"]); result = {"message": f"Hidden: {args['name']}"}
        elif a == "unhide": client.unhide_sheet(args["name"]); result = {"message": f"Unhidden: {args['name']}"}

    elif name == "wps_cell":
        a = args["action"]; cr = args["cell_ref"]
        if a == "get": result = {"cell": cr, "value": client.get_cell_value(cr, sheet)}
        elif a == "set": client.set_cell_value(cr, _parse_value(args["value"]), sheet); result = {"message": f"Set {cr} = {args['value']}"}
        elif a == "clear": client.clear_cell(cr, sheet); result = {"message": f"Cleared {cr}"}
        elif a == "set_formula": client.set_formula(cr, args["formula"], sheet); result = {"message": f"Set formula on {cr}"}
        elif a == "get_formula": result = {"cell": cr, "formula": client.get_formula(cr, sheet)}

    elif name == "wps_range":
        a = args["action"]
        if a == "get": result = {"range": args["range_ref"], "values": client.get_range_values(args["range_ref"], sheet)}
        elif a == "set":
            vals = json.loads(args["values"])
            if not isinstance(vals, list): raise ValueError("values must be a JSON array")
            client.set_range_values(args["start_cell"], vals, sheet)
            result = {"message": f"Set range at {args['start_cell']} with {len(vals)} rows"}
        elif a == "clear": client.clear_range(args.get("range_ref", args.get("cell_or_range","")), sheet); result = {"message": "Range cleared."}
        elif a == "copy": client.copy_range(args["range_ref"], sheet); result = {"message": f"Copied {args['range_ref']}"}
        elif a == "paste": client.paste_range(args["dest_cell"], sheet, args.get("paste_special")); result = {"message": f"Pasted to {args['dest_cell']}"}
        elif a == "sort": client.sort_range(args["range_ref"], args.get("sort_key"), args.get("sort_order","ascending"), sheet); result = {"message": f"Sorted {args['range_ref']}"}
        elif a == "find":
            look_at = "whole" if args.get("match_whole", False) else "part"
            found = client.find_cell(args["search_text"], sheet, look_at)
            result = found if found else {"message": f"'{args['search_text']}' not found", "found": False}
        elif a == "find_next":
            found = client.find_next_cell(sheet)
            result = found if found else {"message": "No more occurrences", "found": False}
        elif a == "find_replace":
            cnt = client.find_replace(args["find_text"], args["replace_text"], sheet, args.get("match_case",False), args.get("match_whole",False))
            result = {"message": f"Replaced {cnt} occurrence(s)", "replacements": cnt}
        elif a == "used_range":
            addr = client.get_used_range_address(sheet)
            result = {"address": addr, "row_count": client.get_row_count(sheet), "column_count": client.get_column_count(sheet)}
        elif a == "remove_duplicates":
            cols = json.loads(args["columns"]) if args.get("columns") else None
            cnt = client.remove_duplicates(args["range_ref"], cols, args.get("has_header",True), sheet)
            result = {"message": f"Removed {cnt} duplicate row(s)"}

    elif name == "wps_format":
        cr = args["cell_or_range"]
        if "bold" in args: client.set_font_bold(cr, args["bold"], sheet)
        if "italic" in args: client.set_font_italic(cr, args["italic"], sheet)
        if "font_name" in args: client.set_font_name(cr, args["font_name"], sheet)
        if "font_size" in args: client.set_font_size(cr, args["font_size"], sheet)
        if "font_color" in args: client.set_font_color(cr, int(args["font_color"].lstrip("#"), 16), sheet)
        if "underline" in args: client.set_font_underline(cr, args["underline"], sheet)
        if "fill_color" in args: client.set_cell_color(cr, int(args["fill_color"].lstrip("#"), 16), sheet)
        if "alignment" in args: client.set_horizontal_alignment(cr, args["alignment"], sheet)
        if "vertical_alignment" in args: client.set_vertical_alignment(cr, args["vertical_alignment"], sheet)
        if "number_format" in args: client.set_number_format(cr, args["number_format"], sheet)
        if "wrap_text" in args: client.set_wrap_text(cr, args["wrap_text"], sheet)
        if "merge" in args:
            if args["merge"]: client.merge_cells(cr, sheet)
            else: client.unmerge_cells(cr, sheet)
        if "border_style" in args:
            bc = int(args["border_color"].lstrip("#"), 16) if args.get("border_color") else None
            client.set_borders(cr, args["border_style"], bc, args.get("outline_only",False), sheet)
        if "cond_operator" in args:
            client.add_conditional_format(cr, "cellValue", args["cond_operator"], args.get("cond_formula","0"),
                int(args["cond_font_color"].lstrip("#"),16) if args.get("cond_font_color") else None,
                int(args["cond_bg_color"].lstrip("#"),16) if args.get("cond_bg_color") else None,
                args.get("cond_bold"), sheet)
        if args.get("clear_cond"): client.delete_conditional_format(cr, sheet)
        if "validation_type" in args:
            client.add_data_validation(cr, args["validation_type"], args.get("formula1",""), args.get("formula2",""),
                args.get("ignore_blank",True), args.get("show_dropdown",True),
                args.get("error_title",""), args.get("error_message",""), sheet)
        result = {"message": f"Formatted {cr}"}

    elif name == "wps_rowcol":
        a = args["action"]; t = args.get("type","row")
        if a == "insert":
            if t == "row": client.insert_row(args["position"], sheet)
            else: client.insert_column(args["position"], sheet)
            result = {"message": f"Inserted {t} at {args['position']}"}
        elif a == "delete":
            if t == "row": client.delete_row(args["position"], sheet)
            else: client.delete_column(args["position"], sheet)
            result = {"message": f"Deleted {t} {args['position']}"}
        elif a == "resize":
            if t == "row": client.set_row_height(args["position"], args["size"], sheet)
            else: client.set_column_width(args["position"], args["size"], sheet)
            result = {"message": f"Set {t} {args['position']} size={args['size']}"}
        elif a == "autofit":
            if t == "row": client.autofit_rows(sheet, args.get("start"), args.get("end"))
            else: client.autofit_columns(sheet, args.get("start"), args.get("end"))
            result = {"message": f"Auto-fitted {t}s"}
        elif a == "group":
            if t == "row": client.group_rows(args["start"], args["end"], sheet)
            else: client.group_columns(args["start"], args["end"], sheet)
            result = {"message": f"Grouped {t}s {args['start']}-{args['end']}"}
        elif a == "ungroup":
            if t == "row": client.ungroup_rows(args["start"], args["end"], sheet)
            else: client.ungroup_columns(args["start"], args["end"], sheet)
            result = {"message": f"Ungrouped {t}s {args['start']}-{args['end']}"}
        elif a == "freeze": client.freeze_panes(args.get("cell_ref","B2"), sheet); result = {"message": f"Frozen at {args.get('cell_ref','B2')}"}
        elif a == "unfreeze": client.unfreeze_panes(sheet); result = {"message": "Unfrozen"}

    elif name == "wps_chart":
        n = client.add_chart(args.get("chart_type","column"), args["range_ref"], args.get("left",100), args.get("top",100), args.get("width",400), args.get("height",300), sheet)
        result = {"message": f"Added chart: {n}"}

    elif name == "wps_data":
        a = args["action"]
        if a == "filter": client.auto_filter(args.get("range_ref"), sheet); result = {"message": "AutoFilter applied."}
        elif a == "text_to_columns": client.text_to_columns(args["range_ref"], args.get("delimiter",","), sheet); result = {"message": f"Split {args['range_ref']}"}

    elif name == "wps_protection":
        if args["action"] == "protect":
            client.protect_sheet(args.get("password",""), args.get("allow_sort",False), args.get("allow_filter",False), args.get("allow_format_cells",False), args.get("allow_insert_rows",False), args.get("allow_delete_rows",False), sheet)
            result = {"message": "Sheet protected."}
        else:
            client.unprotect_sheet(args.get("password",""), sheet)
            result = {"message": "Sheet unprotected."}

    elif name == "wps_page_setup":
        a = args["action"]
        if a == "print_area": client.set_print_area(args["range_ref"], sheet); result = {"message": f"Print area: {args['range_ref']}"}
        elif a == "clear_print_area": client.clear_print_area(sheet); result = {"message": "Print area cleared."}
        elif a == "orientation": client.set_page_orientation(args["orientation"], sheet); result = {"message": f"Orientation: {args['orientation']}"}
        elif a == "margins": client.set_page_margins(args.get("left"), args.get("right"), args.get("top"), args.get("bottom"), args.get("header"), args.get("footer"), sheet); result = {"message": "Margins set."}
        elif a == "header_footer": client.set_header_footer(args.get("left_header",""), args.get("center_header",""), args.get("right_header",""), args.get("left_footer",""), args.get("center_footer",""), args.get("right_footer",""), sheet); result = {"message": "Header/footer set."}

    elif name == "wps_macro":
        result = client.run_macro(args["macro_name"])

    elif name == "wps_export":
        p = client.export_to_pdf(args["filepath"], sheet)
        result = {"message": f"Exported to PDF: {p}", "path": p}

    elif name == "wps_misc":
        a = args["action"]
        if a == "create_named_range": n = client.create_named_range(args["name"], args["refers_to"], sheet); result = {"message": f"Created named range: {n}"}
        elif a == "delete_named_range": client.delete_named_range(args["name"]); result = {"message": f"Deleted named range: {args['name']}"}
        elif a == "list_named_ranges": result = {"named_ranges": client.list_named_ranges()}
        elif a == "add_hyperlink": client.add_hyperlink(args["cell_ref"], args["address"], args.get("text_to_display"), args.get("screen_tip"), sheet); result = {"message": f"Added hyperlink to {args['cell_ref']}"}
        elif a == "remove_hyperlink": client.remove_hyperlink(args["cell_or_range"], sheet); result = {"message": "Hyperlink removed."}
        elif a == "add_comment": client.add_comment(args["cell_ref"], args["text"], sheet); result = {"message": f"Added comment to {args['cell_ref']}"}
        elif a == "delete_comment": client.delete_comment(args["cell_ref"], sheet); result = {"message": f"Deleted comment from {args['cell_ref']}"}
        elif a == "create_pivot_table":
            row_fields = json.loads(args.get("row_fields","[]")) if args.get("row_fields") else None
            col_fields = json.loads(args.get("column_fields","[]")) if args.get("column_fields") else None
            data_fields = json.loads(args.get("data_fields","[]")) if args.get("data_fields") else None
            n = client.create_pivot_table(args["source_range"], args["dest_cell"], args.get("pivot_name","PivotTable1"), row_fields, col_fields, data_fields, sheet)
            result = {"message": f"Created pivot table: {n}"}
        elif a == "add_sparkline": client.add_sparkline(args["source_range"], args["dest_cell"], args.get("spark_type","line"), sheet); result = {"message": "Added sparkline."}
        elif a == "insert_picture": n = client.insert_picture(args["filepath"], args.get("left",100), args.get("top",100), args.get("width",200), args.get("height",150), sheet); result = {"message": f"Inserted picture: {n}"}
        elif a == "insert_shape": n = client.insert_shape(args.get("shape_type","rectangle"), args.get("left",100), args.get("top",100), args.get("width",200), args.get("height",100), sheet); result = {"message": f"Inserted shape: {n}"}
        elif a == "toggle_gridlines": client.toggle_gridlines(args.get("visible",True), sheet); result = {"message": f"Gridlines {'shown' if args.get('visible',True) else 'hidden'}."}

    else: return json.dumps({"error": f"Unknown tool: {name}"}, separators=(",", ":"))
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
    """Entry point for the wps-excel-mcp command."""
    asyncio.run(run_server())


if __name__ == "__main__":
    main()
