"""
MCP Server for WPS Office Word.

Provides tools for automating WPS Office Word via COM:
- Document management (create, open, save, close, list)
- Text operations (insert, replace, get content)
- Paragraph operations (add, format, delete)
- Font formatting (bold, italic, underline, size, color, name)
- Find and replace
- Table operations (create, populate, read, delete)
- Page setup (margins, orientation, size)
- Header and footer
- Export to PDF
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
    from .word_client import WPSWordClient
except ImportError:
    # PyInstaller standalone: word_client is bundled as wps_word_mcp.word_client
    from wps_word_mcp.word_client import WPSWordClient

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)

# ── MCP Server Setup ───────────────────────────────────────────────────

server = Server("wps-word-mcp")

# Dedicated STA thread executor for COM operations.
# COM objects must be created and accessed from the same STA thread.
_sta_executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
_client: WPSWordClient | None = None
_sta_initialized = False


def _init_sta() -> None:
    """Initialize COM on the STA worker thread and create the client."""
    global _client, _sta_initialized
    if not _sta_initialized:
        pythoncom.CoInitialize()
        _sta_initialized = True
    if _client is None:
        _client = WPSWordClient(visible=True)


def get_client() -> WPSWordClient:
    """Get or create the WPS Word client singleton."""
    if _client is None:
        future = _sta_executor.submit(_init_sta)
        future.result()
    assert _client is not None
    return _client


# ── Tool Definitions ───────────────────────────────────────────────────


TOOLS: list[Tool] = [
    Tool(name="word_app", description="WPS Word app control. action: info/show/hide/quit",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["info","show","hide","quit"]}},"required":["action"]}),
    Tool(name="word_document", description="Document mgmt. action: create/open/save/close/list/activate/get_properties/set_properties/protect/unprotect",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["create","open","save","close","list","activate","get_properties","set_properties","protect","unprotect"]},"filepath":{"type":"string"},"save":{"type":"boolean"},"name":{"type":"string"},"author":{"type":"string"},"title":{"type":"string"},"subject":{"type":"string"},"keywords":{"type":"string"},"password":{"type":"string"}},"required":["action"]}),
    Tool(name="word_text", description="Text ops. action: get/set/type/append/prepend/get_selected",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["get","set","type","append","prepend","get_selected"]},"text":{"type":"string"}},"required":["action"]}),
    Tool(name="word_paragraph", description="Paragraph ops. action: add/get_count/get_text/set_text/insert_before/delete/alignment/spacing",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["add","get_count","get_text","set_text","insert_before","delete","alignment","spacing"]},"index":{"type":"integer","description":"1-based"},"text":{"type":"string"},"alignment":{"type":"string","description":"left/center/right/justify"},"before":{"type":"number"},"after":{"type":"number"},"line_spacing":{"type":"number"}},"required":["action"]}),
    Tool(name="word_font", description="Font formatting on a range (selection/content/start=X,end=Y).",
        inputSchema={"type":"object","properties":{"bold":{"type":"boolean"},"italic":{"type":"boolean"},"underline":{"type":"boolean"},"font_name":{"type":"string"},"font_size":{"type":"number"},"font_color":{"type":"string","description":"RGB hex"},"highlight_index":{"type":"integer","description":"0=None,6=Yellow,7=Green,2=Blue"},"range_spec":{"type":"string","description":"selection (default), content, or start=X,end=Y"}}}),
    Tool(name="word_find", description="Find/replace. action: find/find_replace",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["find","find_replace"]},"search_text":{"type":"string"},"find_text":{"type":"string"},"replace_text":{"type":"string"},"match_case":{"type":"boolean"},"match_whole_word":{"type":"boolean"},"replace_all":{"type":"boolean"}},"required":["action"]}),
    Tool(name="word_table", description="Table ops. action: add/count/get_data/set_cell/add_row/add_col/delete/style",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["add","count","get_data","set_cell","add_row","add_col","delete","style"]},"rows":{"type":"integer"},"cols":{"type":"integer"},"table_index":{"type":"integer"},"index":{"type":"integer","description":"1-based for delete/get_data"},"row":{"type":"integer"},"col":{"type":"integer"},"text":{"type":"string"},"style_name":{"type":"string"}},"required":["action"]}),
    Tool(name="word_page", description="Page layout. action: orientation/margins/size/columns/borders/header/footer/page_numbers",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["orientation","margins","size","columns","borders","header","footer","page_numbers"]},"orientation":{"type":"string","enum":["portrait","landscape"]},"left":{"type":"number"},"right":{"type":"number"},"top":{"type":"number"},"bottom":{"type":"number"},"width":{"type":"number"},"height":{"type":"number"},"num_columns":{"type":"integer"},"spacing":{"type":"number"},"line_style":{"type":"integer"},"line_width":{"type":"integer"},"distance":{"type":"integer"},"text":{"type":"string"},"position":{"type":"string","description":"bottom/top for page_numbers"}},"required":["action"]}),
    Tool(name="word_export", description="Export/print. action: pdf/print",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["pdf","print"]},"filepath":{"type":"string"},"copies":{"type":"integer"}},"required":["action"]}),
    Tool(name="word_style", description="Style & list formatting. action: apply_style/list/remove_list",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["apply_style","list","remove_list"]},"style_name":{"type":"string"},"list_type":{"type":"string","enum":["bullet","number"]},"range_spec":{"type":"string","description":"selection/content/start=X,end=Y"}},"required":["action"]}),
    Tool(name="word_insert", description="Insert elements. action: picture/page_break/section_break/hyperlink/toc/bookmark/goto_bookmark",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["picture","page_break","section_break","hyperlink","toc","bookmark","goto_bookmark"]},"filepath":{"type":"string"},"address":{"type":"string"},"text_to_display":{"type":"string"},"name":{"type":"string"},"left":{"type":"number"},"top":{"type":"number"},"width":{"type":"number"},"height":{"type":"number"},"range_spec":{"type":"string"}},"required":["action"]}),
    Tool(name="word_watermark", description="Add text watermark.", inputSchema={"type":"object","properties":{"text":{"type":"string"},"font_size":{"type":"integer"},"color":{"type":"string","description":"RGB hex"},"layout":{"type":"string","enum":["diagonal","horizontal"]}},"required":["text"]}),
    Tool(name="word_track_changes", description="Toggle Track Changes.", inputSchema={"type":"object","properties":{"enable":{"type":"boolean"}}}),
    Tool(name="word_zoom", description="Set zoom level.", inputSchema={"type":"object","properties":{"percentage":{"type":"integer"}},"required":["percentage"]}),
    Tool(name="word_comment", description="Add comment.", inputSchema={"type":"object","properties":{"text":{"type":"string"},"range_spec":{"type":"string"}},"required":["text"]}),
    Tool(name="word_range", description="Get text from 0-based character range.", inputSchema={"type":"object","properties":{"start":{"type":"integer"},"end":{"type":"integer"}},"required":["start","end"]}),
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
            "tool": name,
        }, ensure_ascii=False))]

    except Exception as e:
        logger.exception(f"Error executing tool '{name}'")
        return [TextContent(type="text", text=json.dumps({
            "error": str(e),
            "tool": name,
        }, ensure_ascii=False))]



def _execute_tool(name: str, args: dict[str, Any], client: WPSWordClient) -> str:
    pythoncom.CoInitialize()
    result: Any = None

    if name == "word_app":
        a = args["action"]
        if a == "info": result = client.get_app_info()
        elif a == "show": client.show(); result = {"message": "WPS Word window visible."}
        elif a == "hide": client.hide(); result = {"message": "WPS Word window hidden."}
        elif a == "quit": client.quit_app(); result = {"message": "WPS Word quit."}

    elif name == "word_document":
        a = args["action"]
        if a == "create": n = client.create_document(); result = {"message": f"Created: {n}", "document_name": n}
        elif a == "open": n = client.open_document(args["filepath"]); result = {"message": f"Opened: {n}", "document_name": n}
        elif a == "save": p = client.save_document(args.get("filepath")); result = {"message": f"Saved to: {p}", "path": p}
        elif a == "close": s = args.get("save", True); client.close_document(s); result = {"message": "Closed." + (" (saved)" if s else "")}
        elif a == "list": result = {"documents": client.list_documents()}
        elif a == "activate": n = client.activate_document(args["name"]); result = {"message": f"Activated: {n}"}
        elif a == "get_properties": result = client.get_document_properties()
        elif a == "set_properties": client.set_document_properties(args.get("author"), args.get("title"), args.get("subject"), args.get("keywords")); result = {"message": "Properties set."}
        elif a == "protect": client.protect_document(args.get("password","")); result = {"message": "Document protected."}
        elif a == "unprotect": client.unprotect_document(args.get("password","")); result = {"message": "Document unprotected."}

    elif name == "word_text":
        a = args["action"]
        if a == "get": t = client.get_text(); result = {"text": t, "length": len(t)}
        elif a == "set": client.set_text(args["text"]); result = {"message": "Text replaced.", "length": len(args["text"])}
        elif a == "type": client.type_text(args["text"]); result = {"message": "Typed."}
        elif a == "append": client.insert_text_at_end(args["text"]); result = {"message": "Appended."}
        elif a == "prepend": client.insert_text_at_start(args["text"]); result = {"message": "Prepended."}
        elif a == "get_selected": t = client.get_selected_text(); result = {"text": t, "length": len(t)}

    elif name == "word_paragraph":
        a = args["action"]
        if a == "add": idx = client.add_paragraph(args.get("text","")); result = {"message": f"Added paragraph {idx}.", "paragraph_index": idx}
        elif a == "get_count": result = {"paragraph_count": client.get_paragraph_count()}
        elif a == "get_text": result = {"paragraph_index": args["index"], "text": client.get_paragraph_text(args["index"])}
        elif a == "set_text": client.set_paragraph_text(args["index"], args["text"]); result = {"message": f"Set paragraph {args['index']}."}
        elif a == "insert_before": client.insert_paragraph_before(args["index"], args.get("text","")); result = {"message": f"Inserted before {args['index']}."}
        elif a == "delete": client.delete_paragraph(args["index"]); result = {"message": f"Deleted paragraph {args['index']}."}
        elif a == "alignment": client.set_paragraph_alignment(args["index"], args["alignment"]); result = {"message": f"Alignment = {args['alignment']}."}
        elif a == "spacing": client.set_paragraph_spacing(args["index"], args.get("before"), args.get("after"), args.get("line_spacing")); result = {"message": "Spacing set."}

    elif name == "word_font":
        rs = args.get("range_spec", "selection")
        if "bold" in args: client.set_font_bold(args["bold"], rs)
        if "italic" in args: client.set_font_italic(args["italic"], rs)
        if "underline" in args: client.set_font_underline(args["underline"], rs)
        if "font_name" in args: client.set_font_name(args["font_name"], rs)
        if "font_size" in args: client.set_font_size(args["font_size"], rs)
        if "font_color" in args: client.set_font_color(int(args["font_color"].lstrip("#"), 16), rs)
        if "highlight_index" in args: client.set_highlight(args["highlight_index"], rs)
        result = {"message": f"Font formatted on '{rs}'."}

    elif name == "word_find":
        a = args["action"]
        if a == "find":
            found = client.find_text(args["search_text"], args.get("match_case",False), args.get("match_whole_word",False))
            result = found if found else {"found": False, "message": f"'{args['search_text']}' not found."}
        elif a == "find_replace":
            cnt = client.find_replace(args["find_text"], args["replace_text"], args.get("match_case",False), args.get("match_whole_word",False), args.get("replace_all",True))
            result = {"message": f"Replaced {cnt} occurrence(s).", "replacements": cnt}

    elif name == "word_table":
        a = args["action"]
        if a == "add": idx = client.add_table(args["rows"], args["cols"], args.get("text","")); result = {"message": f"Added table {idx}.", "table_index": idx}
        elif a == "count": result = {"table_count": client.get_table_count()}
        elif a == "get_data": data = client.get_table_data(args["index"]); result = {"table_index": args["index"], "rows": len(data), "columns": len(data[0]) if data else 0, "data": data}
        elif a == "set_cell": client.set_cell_text(args["table_index"], args["row"], args["col"], args["text"]); result = {"message": f"Cell ({args['row']},{args['col']}) set."}
        elif a == "add_row": client.add_table_row(args["table_index"]); result = {"message": f"Added row to table {args['table_index']}."}
        elif a == "add_col": client.add_table_column(args["table_index"]); result = {"message": f"Added column to table {args['table_index']}."}
        elif a == "delete": client.delete_table(args["index"]); result = {"message": f"Deleted table {args['index']}."}
        elif a == "style": client.set_table_style(args["table_index"], args["style_name"]); result = {"message": f"Style '{args['style_name']}' applied."}

    elif name == "word_page":
        a = args["action"]
        if a == "orientation": client.set_page_orientation(args["orientation"]); result = {"message": f"Orientation: {args['orientation']}."}
        elif a == "margins": client.set_page_margins(args.get("left"), args.get("right"), args.get("top"), args.get("bottom")); result = {"message": "Margins set."}
        elif a == "size": client.set_page_size(args.get("width"), args.get("height")); result = {"message": "Page size set."}
        elif a == "columns": client.set_columns(args.get("num_columns",1), args.get("spacing")); result = {"message": f"Columns: {args.get('num_columns',1)}."}
        elif a == "borders": client.set_page_borders(args.get("line_style",1), args.get("line_width",4), args.get("distance",24)); result = {"message": "Page borders set."}
        elif a == "header": client.add_header(args["text"]); result = {"message": "Header set."}
        elif a == "footer": client.add_footer(args["text"]); result = {"message": "Footer set."}
        elif a == "page_numbers": client.insert_page_numbers(args.get("position","bottom")); result = {"message": "Page numbers inserted."}

    elif name == "word_export":
        if args["action"] == "pdf":
            p = client.export_to_pdf(args["filepath"]); result = {"message": f"Exported to PDF: {p}", "path": p}
        else:
            client.print_document(args.get("copies",1)); result = {"message": f"Printing {args.get('copies',1)} copies."}

    elif name == "word_style":
        a = args["action"]; rs = args.get("range_spec", "selection")
        if a == "apply_style": client.apply_style(args["style_name"], rs); result = {"message": f"Style '{args['style_name']}' applied."}
        elif a == "list": client.set_list_format(args.get("list_type","bullet"), rs); result = {"message": f"List format applied."}
        elif a == "remove_list": client.remove_list_format(rs); result = {"message": "List format removed."}

    elif name == "word_insert":
        a = args["action"]; rs = args.get("range_spec", "selection")
        if a == "picture": n = client.insert_picture(args["filepath"], args.get("left"), args.get("top"), args.get("width"), args.get("height")); result = {"message": f"Inserted picture: {n}"}
        elif a == "page_break": client.insert_page_break(); result = {"message": "Page break inserted."}
        elif a == "section_break": client.add_section_break(); result = {"message": "Section break added."}
        elif a == "hyperlink": client.add_hyperlink(args["address"], args.get("text_to_display"), rs); result = {"message": f"Hyperlink added."}
        elif a == "toc": client.insert_table_of_contents(); result = {"message": "TOC inserted."}
        elif a == "bookmark": client.add_bookmark(args["name"], rs); result = {"message": f"Bookmark '{args['name']}' added."}
        elif a == "goto_bookmark": bi = client.go_to_bookmark(args["name"]); result = {"message": f"Navigated to '{args['name']}'.", **bi}

    elif name == "word_watermark":
        ci = int(args["color"].lstrip("#"), 16) if args.get("color") else None
        client.add_watermark(args["text"], args.get("font_size",72), ci, args.get("layout","diagonal"))
        result = {"message": f"Watermark '{args['text']}' added."}

    elif name == "word_track_changes":
        e = args.get("enable", True); client.toggle_track_changes(e)
        result = {"message": f"Track changes {'enabled' if e else 'disabled'}."}

    elif name == "word_zoom":
        client.set_zoom(args["percentage"]); result = {"message": f"Zoom: {args['percentage']}%."}

    elif name == "word_comment":
        client.add_comment(args["text"], args.get("range_spec","selection"))
        result = {"message": "Comment added."}

    elif name == "word_range":
        t = client.get_range_text(args["start"], args["end"])
        result = {"start": args["start"], "end": args["end"], "text": t}

    else: return json.dumps({"error": f"Unknown tool: {name}"})
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
    """Entry point for the wps-word-mcp command."""
    asyncio.run(run_server())


if __name__ == "__main__":
    main()
