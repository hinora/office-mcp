"""
MCP Server for WPS Office Slide (Presentation).

Provides tools for automating WPS Office Slide via COM:
- Presentation management (create, open, save, close, list)
- Slide operations (add, delete, duplicate, move, list)
- Shape operations (text boxes, rectangles, ovals, arrows, lines, pictures, copy/paste)
- Text and font formatting
- Table operations
- Speaker notes
- Export to PDF / images
- Slide show control
- Transitions
- Hyperlinks
- Animations (entrance/exit/emphasis effects)
- Find and replace across slides
- Slide master and layout management
- Insert headers, footers, slide numbers, date/time
- Charts
- Media (video/audio)
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
    from .slide_client import WPSSlideClient
except ImportError:
    # PyInstaller standalone: slide_client is bundled as wps_slide_mcp.slide_client
    from wps_slide_mcp.slide_client import WPSSlideClient

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)

# ── MCP Server Setup ───────────────────────────────────────────────────

server = Server("wps-slide-mcp")

# Dedicated STA thread executor for COM operations.
# COM objects must be created and accessed from the same STA thread.
_sta_executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
_client: WPSSlideClient | None = None
_sta_initialized = False


def _init_sta() -> None:
    """Initialize COM on the STA worker thread and create the client."""
    global _client, _sta_initialized
    if not _sta_initialized:
        pythoncom.CoInitialize()
        _sta_initialized = True
    if _client is None:
        _client = WPSSlideClient(visible=True)


def get_client() -> WPSSlideClient:
    """Get or create the WPS Slide client singleton."""
    if _client is None:
        future = _sta_executor.submit(_init_sta)
        future.result()
    assert _client is not None
    return _client


# ── Tool Definitions ───────────────────────────────────────────────────

TOOLS: list[Tool] = [
    Tool(name="app", description="WPS Slide app control",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["info","show","hide","quit"]}},"required":["action"]}),
    Tool(name="pres", description="Presentation mgmt: create/open/save/close/list/activate/get_properties/set_slide_size",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["create","open","save","close","list","activate","get_properties","set_slide_size"]},"filepath":{"type":"string"},"save":{"type":"boolean"},"name":{"type":"string"},"width":{"type":"number"},"height":{"type":"number"}},"required":["action"]}),
    Tool(name="slide", description="Slide ops: add/delete/duplicate/move/go_to/count/list/set_background/set_transition",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["add","delete","duplicate","move","go_to","count","list","set_background","set_transition"]},"index":{"type":"integer"},"to_index":{"type":"integer"},"layout_index":{"type":"integer","default":1},"color_rgb":{"type":"string"},"transition_type":{"type":"integer"},"duration":{"type":"number"}},"required":["action"]}),
    Tool(name="shape_add", description="Add shapes: text_box/picture/rectangle/oval/arrow/line. Conventions: slide_index 1-based (0=active), coords in points",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["text_box","picture","rectangle","oval","arrow","line"]},"text":{"type":"string"},"filepath":{"type":"string"},"left":{"type":"number"},"top":{"type":"number"},"width":{"type":"number"},"height":{"type":"number"},"begin_x":{"type":"number"},"begin_y":{"type":"number"},"end_x":{"type":"number"},"end_y":{"type":"number"},"slide_index":{"type":"integer"}},"required":["action"]}),
    Tool(name="shape_format", description="Format shapes: set_position/set_fill/set_line/set_rotation/set_zorder. name_or_index: name(str) or 1-based index(int). Colors: hex RGB like FF0000",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["set_position","set_fill","set_line","set_rotation","set_zorder"]},"name_or_index":{},"left":{"type":"number"},"top":{"type":"number"},"width":{"type":"number"},"height":{"type":"number"},"color_rgb":{"type":"string"},"weight":{"type":"number"},"rotation":{"type":"number"},"zorder":{"type":"string","enum":["front","back","forward","backward"]},"slide_index":{"type":"integer"}},"required":["action","name_or_index"]}),
    Tool(name="shape_organize", description="Organize shapes: group/ungroup/copy/paste/duplicate/delete/list/count. names: JSON array of shape names/indices to group",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["group","ungroup","copy","paste","duplicate","delete","list","count"]},"name_or_index":{},"names":{},"dest_slide_index":{"type":"integer"},"slide_index":{"type":"integer"}},"required":["action"]}),
    Tool(name="text", description="Get/set shape text",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["get","set"]},"name_or_index":{},"text":{"type":"string"},"slide_index":{"type":"integer"}},"required":["action","name_or_index"]}),
    Tool(name="font", description="Font formatting: bold/italic/underline/font_name/font_size/font_color/alignment. start/length for partial range formatting",
        inputSchema={"type":"object","properties":{"name_or_index":{},"slide_index":{"type":"integer"},"bold":{"type":"boolean"},"italic":{"type":"boolean"},"underline":{"type":"boolean"},"font_name":{"type":"string"},"font_size":{"type":"number"},"font_color":{"type":"string"},"alignment":{"type":"string","enum":["left","center","right","justify"]},"start":{"type":"integer"},"length":{"type":"integer"}},"required":["name_or_index"]}),
    Tool(name="table", description="Table ops: add/set_cell/get_data",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["add","set_cell","get_data"]},"rows":{"type":"integer"},"cols":{"type":"integer"},"left":{"type":"number"},"top":{"type":"number"},"width":{"type":"number"},"height":{"type":"number"},"shape_name":{},"row":{"type":"integer"},"col":{"type":"integer"},"text":{"type":"string"},"slide_index":{"type":"integer"}},"required":["action"]}),
    Tool(name="notes", description="Speaker notes: get/set",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["get","set"]},"slide_index":{"type":"integer"},"text":{"type":"string"}},"required":["action","slide_index"]}),
    Tool(name="export", description="Export to PDF or slide image",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["pdf","slide_image"]},"filepath":{"type":"string"},"slide_index":{"type":"integer"},"img_width":{"type":"integer","default":1920},"img_height":{"type":"integer","default":1080}},"required":["action","filepath"]}),
    Tool(name="slideshow", description="Slide show control: start/start_from/stop",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["start","start_from","stop"]},"slide_index":{"type":"integer"}},"required":["action"]}),
    Tool(name="animate", description="Add/clear shape animations",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["add","clear"]},"name_or_index":{},"effect_type":{"type":"string","enum":["appear","fly","blinds","box","checkerboard","dissolve","fade","flash_once","peek","random_bars","spiral","split","stretch","strips","swivel","wipe","zoom","random_effects","spin","grow_shrink","float"]},"trigger":{"type":"string","enum":["on_click","with_previous","after_previous"]},"duration":{"type":"number"},"delay":{"type":"number"},"slide_index":{"type":"integer"}},"required":["action"]}),
    Tool(name="find", description="Find/replace text across all slides",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["find","find_replace"]},"search_text":{"type":"string"},"find_text":{"type":"string"},"replace_text":{"type":"string"},"match_case":{"type":"boolean"},"match_whole_word":{"type":"boolean"},"replace_all":{"type":"boolean"}},"required":["action"]}),
    Tool(name="master", description="Slide master & layout: get_info/apply_layout/set_background",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["get_info","apply_layout","set_background"]},"slide_index":{"type":"integer"},"layout_index":{"type":"integer"},"color_rgb":{"type":"string"},"master_index":{"type":"integer","default":1}},"required":["action"]}),
    Tool(name="advanced", description="Hyperlink, chart, media, headers/footers, slide numbers. Use action='help' for details.",
        inputSchema={"type":"object","properties":{"action":{"type":"string","enum":["help","hyperlink","chart_add","chart_set_data","media_video","media_audio","slide_number","date_time","header_footer"]},"address":{"type":"string"},"text_to_display":{"type":"string"},"name_or_index":{},"slide_index":{"type":"integer"},"chart_type":{"type":"string","enum":["column","line","pie","bar","area","scatter"]},"left":{"type":"number"},"top":{"type":"number"},"width":{"type":"number"},"height":{"type":"number"},"shape_name":{},"categories":{},"series_data":{},"filepath":{"type":"string"},"header_text":{"type":"string"},"footer_text":{"type":"string"},"show_slide_number":{"type":"boolean","default":true},"show_date_time":{"type":"boolean","default":false}},"required":["action"]}),
]

# ── Tool Handler ───────────────────────────────────────────────────────

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


def _parse_hex_color(hex_str: str) -> int:
    """Parse a hex color string like 'FF0000' to int."""
    return int(hex_str.lstrip("#"), 16)


def _pump_com_messages(timeout_ms: int = 50) -> None:
    """Pump pending COM/Win32 messages to prevent WPS from freezing.

    COM automation in STA threads requires message pumping. WPS applications
    (especially WPP.Application) may need to dispatch cross-thread COM calls
    back to the STA thread. Without pumping, the STA thread blocks and WPS
    freezes indefinitely.
    """
    try:
        pythoncom.PumpWaitingMessages(timeout=timeout_ms)
    except Exception:
        pass


def _execute_tool(name: str, args: dict[str, Any], client: WPSSlideClient) -> str:
    pythoncom.CoInitialize()
    result: Any = None

    _pump_com_messages()

    if name == "app":
        a = args["action"]
        if a == "info": result = client.get_app_info()
        elif a == "show": client.show(); result = {"message": "WPS Slide window visible."}
        elif a == "hide": client.hide(); result = {"message": "WPS Slide window hidden."}
        elif a == "quit": client.quit_app(); result = {"message": "WPS Slide quit."}

    elif name == "pres":
        a = args["action"]
        if a == "create": n = client.create_presentation(); result = {"message": f"Created: {n}", "presentation_name": n}
        elif a == "open": n = client.open_presentation(args["filepath"]); result = {"message": f"Opened: {n}", "presentation_name": n}
        elif a == "save": p = client.save_presentation(args.get("filepath")); result = {"message": f"Saved to: {p}", "path": p}
        elif a == "close": s = args.get("save", True); client.close_presentation(s); result = {"message": "Closed." + (" (saved)" if s else "")}
        elif a == "list": result = {"presentations": client.list_presentations()}
        elif a == "activate": n = client.activate_presentation(args["name"]); result = {"message": f"Activated: {n}"}
        elif a == "get_properties": result = client.get_presentation_properties()
        elif a == "set_slide_size": client.set_slide_size(args["width"], args["height"]); result = {"message": f"Slide size: {args['width']}x{args['height']}."}

    elif name == "slide":
        a = args["action"]
        if a == "add": idx = client.add_slide(args.get("layout_index", 1)); result = {"message": f"Added slide {idx}.", "slide_index": idx}
        elif a == "delete": client.delete_slide(args["index"]); result = {"message": f"Deleted slide {args['index']}."}
        elif a == "duplicate": idx = client.duplicate_slide(args["index"]); result = {"message": f"Duplicated slide {args['index']} -> {idx}.", "new_index": idx}
        elif a == "move": client.move_slide(args["index"], args["to_index"]); result = {"message": f"Moved slide {args['index']} to {args['to_index']}."}
        elif a == "go_to": client.go_to_slide(args["index"]); result = {"message": f"Navigated to slide {args['index']}."}
        elif a == "count": result = {"slide_count": client.get_slide_count()}
        elif a == "list": result = {"slides": client.list_slides()}
        elif a == "set_background": client.slide_set_background(args["index"], _parse_hex_color(args["color_rgb"])); result = {"message": f"Background set on slide {args['index']}."}
        elif a == "set_transition": client.slide_set_transition(args["index"], args["transition_type"], args.get("duration", 1.0)); result = {"message": f"Transition set on slide {args['index']}."}

    elif name == "shape_add":
        a = args["action"]; si = args.get("slide_index", 0)
        if a == "text_box": n = client.add_text_box(args.get("text",""), args.get("left",50), args.get("top",50), args.get("width",400), args.get("height",100), si); result = {"message": f"Added text box: {n}", "shape_name": n}
        elif a == "picture": n = client.add_picture(args["filepath"], args.get("left",50), args.get("top",50), args.get("width",-1), args.get("height",-1), si); result = {"message": f"Added picture: {n}", "shape_name": n}
        elif a == "rectangle": n = client.add_rectangle(args.get("left",50), args.get("top",50), args.get("width",200), args.get("height",100), si); result = {"message": f"Added rectangle: {n}", "shape_name": n}
        elif a == "oval": n = client.add_oval(args.get("left",50), args.get("top",50), args.get("width",200), args.get("height",100), si); result = {"message": f"Added oval: {n}", "shape_name": n}
        elif a == "arrow": n = client.add_arrow(args.get("left",50), args.get("top",50), args.get("width",200), args.get("height",50), si); result = {"message": f"Added arrow: {n}", "shape_name": n}
        elif a == "line": n = client.add_line(args.get("begin_x",50), args.get("begin_y",50), args.get("end_x",300), args.get("end_y",50), si); result = {"message": f"Added line: {n}", "shape_name": n}

    elif name == "shape_format":
        a = args["action"]; si = args.get("slide_index", 0); ni = args["name_or_index"]
        if a == "set_position": client.set_shape_position(ni, args.get("left"), args.get("top"), args.get("width"), args.get("height"), si); result = {"message": f"Position set on shape '{ni}'."}
        elif a == "set_fill": client.set_shape_fill(ni, _parse_hex_color(args["color_rgb"]), si); result = {"message": f"Fill set on shape '{ni}'."}
        elif a == "set_line": client.set_shape_line(ni, _parse_hex_color(args["color_rgb"]), args.get("weight",1.0), si); result = {"message": f"Line set on shape '{ni}'."}
        elif a == "set_rotation": client.set_shape_rotation(ni, args["rotation"], si); result = {"message": f"Rotation set to {args['rotation']}°."}
        elif a == "set_zorder": client.set_shape_zorder(ni, args["zorder"], si); result = {"message": f"Z-order '{args['zorder']}' applied."}

    elif name == "shape_organize":
        a = args["action"]; si = args.get("slide_index", 0)
        if a == "group":
            names = json.loads(args["names"]) if isinstance(args.get("names"), str) else args["names"]
            n = client.group_shapes(names, si); result = {"message": f"Grouped shapes into: {n}", "group_name": n}
        elif a == "ungroup": client.ungroup_shapes(args["name_or_index"], si); result = {"message": f"Ungrouped: {args['name_or_index']}."}
        elif a == "copy": client.copy_shape(args["name_or_index"], si); result = {"message": f"Copied shape '{args['name_or_index']}'."}
        elif a == "paste":
            dsi = args.get("dest_slide_index", si)
            n = client.paste_shape(dsi); result = {"message": f"Pasted shape: {n}", "shape_name": n}
        elif a == "duplicate": n = client.duplicate_shape(args["name_or_index"], si); result = {"message": f"Duplicated: {n}", "shape_name": n}
        elif a == "delete": client.delete_shape(args["name_or_index"], si); result = {"message": f"Deleted shape: {args['name_or_index']}."}
        elif a == "list": result = {"shapes": client.list_shapes(si)}
        elif a == "count": result = {"shape_count": client.get_shape_count(si)}

    elif name == "text":
        a = args["action"]; si = args.get("slide_index", 0)
        if a == "get": t = client.get_shape_text(args["name_or_index"], si); result = {"text": t, "length": len(t)}
        elif a == "set": client.set_shape_text(args["name_or_index"], args["text"], si); result = {"message": "Text set.", "length": len(args["text"])}

    elif name == "font":
        ni = args["name_or_index"]; si = args.get("slide_index", 0); s = args.get("start", 0); l = args.get("length", 0)
        changes = []
        if "bold" in args: client.set_font_bold(ni, args["bold"], s, l, si); changes.append("bold")
        if "italic" in args: client.set_font_italic(ni, args["italic"], s, l, si); changes.append("italic")
        if "underline" in args: client.set_font_underline(ni, args["underline"], s, l, si); changes.append("underline")
        if "font_name" in args: client.set_font_name(ni, args["font_name"], s, l, si); changes.append(f"font={args['font_name']}")
        if "font_size" in args: client.set_font_size(ni, args["font_size"], s, l, si); changes.append(f"size={args['font_size']}")
        if "font_color" in args: client.set_font_color(ni, _parse_hex_color(args["font_color"]), s, l, si); changes.append("color")
        if "alignment" in args: client.set_text_alignment(ni, args["alignment"], si); changes.append(f"align={args['alignment']}")
        result = {"message": f"Font updated: {', '.join(changes)}."}

    elif name == "table":
        a = args["action"]; si = args.get("slide_index", 0)
        if a == "add": n = client.add_table(args["rows"], args["cols"], args.get("left",50), args.get("top",50), args.get("width",600), args.get("height",300), si); result = {"message": f"Added table: {n}", "shape_name": n}
        elif a == "set_cell": client.set_table_cell(args["shape_name"], args["row"], args["col"], args["text"], si); result = {"message": f"Cell ({args['row']},{args['col']}) set."}
        elif a == "get_data": data = client.get_table_data(args["shape_name"], si); result = {"rows": len(data), "columns": len(data[0]) if data else 0, "data": data}

    elif name == "notes":
        a = args["action"]
        if a == "get": t = client.get_notes(args["slide_index"]); result = {"slide_index": args["slide_index"], "notes": t}
        elif a == "set": client.set_notes(args["slide_index"], args["text"]); result = {"message": f"Notes set on slide {args['slide_index']}."}

    elif name == "export":
        if args["action"] == "pdf":
            p = client.export_to_pdf(args["filepath"]); result = {"message": f"Exported to PDF: {p}", "path": p}
        elif args["action"] == "slide_image":
            p = client.export_slide_image(args["filepath"], args["slide_index"], args.get("img_width",1920), args.get("img_height",1080))
            result = {"message": f"Exported slide image: {p}", "path": p}

    elif name == "slideshow":
        a = args["action"]
        if a == "start": client.start_slideshow(); result = {"message": "Slide show started."}
        elif a == "start_from": client.start_slideshow_from(args["slide_index"]); result = {"message": f"Slide show started from slide {args['slide_index']}."}
        elif a == "stop": client.stop_slideshow(); result = {"message": "Slide show stopped."}

    elif name == "animate":
        a = args["action"]; si = args.get("slide_index", 0)
        if a == "add":
            client.add_animation(
                args.get("name_or_index", 0),
                args.get("effect_type", "fade"),
                args.get("trigger", "on_click"),
                args.get("duration", 1.0),
                args.get("delay", 0.0),
                si,
            )
            result = {"message": f"Animation '{args.get('effect_type', 'fade')}' added."}
        elif a == "clear":
            client.clear_animations(args.get("name_or_index", 0), si)
            result = {"message": "Animations cleared."}

    elif name == "find":
        a = args["action"]
        if a == "find":
            found = client.find_text(
                args["search_text"],
                args.get("match_case", False),
                args.get("match_whole_word", False),
            )
            result = {"matches": found, "count": len(found)}
        elif a == "find_replace":
            cnt = client.find_replace(
                args["find_text"],
                args["replace_text"],
                args.get("match_case", False),
                args.get("match_whole_word", False),
                args.get("replace_all", True),
            )
            result = {"message": f"Replaced {cnt} occurrence(s).", "replacements": cnt}

    elif name == "master":
        a = args["action"]
        if a == "get_info": result = client.get_master_info()
        elif a == "apply_layout": client.apply_layout(args["slide_index"], args["layout_index"]); result = {"message": f"Layout {args['layout_index']} applied to slide {args['slide_index']}."}
        elif a == "set_background": client.set_master_background(_parse_hex_color(args["color_rgb"]), args.get("master_index", 1)); result = {"message": "Master background set."}

    elif name == "advanced":
        a = args["action"]; si = args.get("slide_index", 0)
        if a == "help":
            result = {"message": "Advanced WPS Slide operations:\n"
                "- hyperlink: Add hyperlink to shape (args: address, text_to_display, name_or_index, slide_index)\n"
                "- chart_add: Add chart (args: chart_type, left, top, width, height, slide_index)\n"
                "- chart_set_data: Set chart data (args: shape_name, categories, series_data, slide_index)\n"
                "- media_video: Insert video (args: filepath, left, top, width, height, slide_index)\n"
                "- media_audio: Insert audio (args: filepath, slide_index)\n"
                "- slide_number: Insert slide numbers on all slides\n"
                "- date_time: Insert date/time on all slides\n"
                "- header_footer: Set header/footer (args: header_text, footer_text, show_slide_number, show_date_time)"}
        elif a == "hyperlink":
            client.add_hyperlink(args["address"], args.get("text_to_display"), args.get("name_or_index", 0), si)
            result = {"message": f"Hyperlink added: {args['address']}."}
        elif a == "chart_add":
            n = client.add_chart(args.get("chart_type", "column"), args.get("left", 50), args.get("top", 50), args.get("width", 600), args.get("height", 400), si)
            result = {"message": f"Added chart: {n}", "shape_name": n}
        elif a == "chart_set_data":
            categories = json.loads(args["categories"]) if isinstance(args.get("categories"), str) else args.get("categories", [])
            series_data = json.loads(args["series_data"]) if isinstance(args.get("series_data"), str) else args.get("series_data", [])
            client.set_chart_data(args["shape_name"], categories, series_data, si)
            result = {"message": "Chart data set."}
        elif a == "media_video":
            n = client.add_video(args["filepath"], args.get("left", 50), args.get("top", 50), args.get("width", 640), args.get("height", 480), si)
            result = {"message": f"Added video: {n}", "shape_name": n}
        elif a == "media_audio":
            n = client.add_audio(args["filepath"], si)
            result = {"message": f"Added audio: {n}", "shape_name": n}
        elif a == "slide_number": client.insert_slide_number(); result = {"message": "Slide numbers inserted."}
        elif a == "date_time": client.insert_date_time(); result = {"message": "Date/time inserted."}
        elif a == "header_footer":
            client.insert_header_footer(args.get("header_text", ""), args.get("footer_text", ""), args.get("show_slide_number", True), args.get("show_date_time", False))
            result = {"message": "Headers/footers set."}

    else:
        return json.dumps({"error": f"Unknown tool: {name}"}, separators=(",", ":"))

    _pump_com_messages()
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
    """Entry point for the wps-slide-mcp command."""
    asyncio.run(run_server())


if __name__ == "__main__":
    main()
