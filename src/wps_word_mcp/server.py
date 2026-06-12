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

TOOLS = [
    # ── Application ──
    Tool(
        name="word_get_app_info",
        description="Get information about the WPS Word application (version, open documents count, active document, etc.).",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    Tool(
        name="word_show_window",
        description="Make the WPS Word application window visible.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    Tool(
        name="word_hide_window",
        description="Hide the WPS Word application window (run in background).",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    Tool(
        name="word_quit_app",
        description="Quit the WPS Word application entirely.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),

    # ── Document Management ──
    Tool(
        name="word_create_document",
        description="Create a new blank Word document.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    Tool(
        name="word_open_document",
        description="Open an existing Word document from disk.",
        inputSchema={
            "type": "object",
            "properties": {
                "filepath": {
                    "type": "string",
                    "description": "Full path to the document file (.docx, .doc, .wps, etc.).",
                },
            },
            "required": ["filepath"],
        },
    ),
    Tool(
        name="word_save_document",
        description="Save the active document. Optionally save to a new file path.",
        inputSchema={
            "type": "object",
            "properties": {
                "filepath": {
                    "type": "string",
                    "description": "Optional path to save the document to. If omitted, saves to current location.",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="word_close_document",
        description="Close the active document.",
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
        name="word_list_documents",
        description="List all currently open documents.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    Tool(
        name="word_activate_document",
        description="Activate a specific document by name when multiple documents are open.",
        inputSchema={
            "type": "object",
            "properties": {
                "name": {
                    "type": "string",
                    "description": "Name of the document to activate.",
                },
            },
            "required": ["name"],
        },
    ),

    # ── Text Operations ──
    Tool(
        name="word_get_text",
        description="Get all text content from the active document.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    Tool(
        name="word_get_selected_text",
        description="Get the currently selected text in the document.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    Tool(
        name="word_set_text",
        description="Replace all text in the document with new text.",
        inputSchema={
            "type": "object",
            "properties": {
                "text": {
                    "type": "string",
                    "description": "The full text to set as the document content.",
                },
            },
            "required": ["text"],
        },
    ),
    Tool(
        name="word_type_text",
        description="Type/insert text at the current cursor/selection position.",
        inputSchema={
            "type": "object",
            "properties": {
                "text": {
                    "type": "string",
                    "description": "The text to insert at the current position.",
                },
            },
            "required": ["text"],
        },
    ),
    Tool(
        name="word_insert_text_at_end",
        description="Append text at the end of the document.",
        inputSchema={
            "type": "object",
            "properties": {
                "text": {
                    "type": "string",
                    "description": "The text to append.",
                },
            },
            "required": ["text"],
        },
    ),
    Tool(
        name="word_insert_text_at_start",
        description="Insert text at the beginning of the document.",
        inputSchema={
            "type": "object",
            "properties": {
                "text": {
                    "type": "string",
                    "description": "The text to insert at the start.",
                },
            },
            "required": ["text"],
        },
    ),

    # ── Paragraph Operations ──
    Tool(
        name="word_add_paragraph",
        description="Add a new paragraph at the end of the document.",
        inputSchema={
            "type": "object",
            "properties": {
                "text": {
                    "type": "string",
                    "description": "Optional text for the new paragraph.",
                    "default": "",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="word_get_paragraph_count",
        description="Get the number of paragraphs in the active document.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    Tool(
        name="word_get_paragraph_text",
        description="Get the text of a specific paragraph by 1-based index.",
        inputSchema={
            "type": "object",
            "properties": {
                "index": {
                    "type": "integer",
                    "description": "1-based paragraph index.",
                },
            },
            "required": ["index"],
        },
    ),
    Tool(
        name="word_set_paragraph_text",
        description="Set the text of a specific paragraph by 1-based index.",
        inputSchema={
            "type": "object",
            "properties": {
                "index": {
                    "type": "integer",
                    "description": "1-based paragraph index.",
                },
                "text": {
                    "type": "string",
                    "description": "New text for the paragraph.",
                },
            },
            "required": ["index", "text"],
        },
    ),
    Tool(
        name="word_insert_paragraph_before",
        description="Insert a new paragraph before a specific paragraph index.",
        inputSchema={
            "type": "object",
            "properties": {
                "index": {
                    "type": "integer",
                    "description": "1-based paragraph index to insert before.",
                },
                "text": {
                    "type": "string",
                    "description": "Optional text for the new paragraph.",
                    "default": "",
                },
            },
            "required": ["index"],
        },
    ),
    Tool(
        name="word_delete_paragraph",
        description="Delete a paragraph by 1-based index.",
        inputSchema={
            "type": "object",
            "properties": {
                "index": {
                    "type": "integer",
                    "description": "1-based paragraph index to delete.",
                },
            },
            "required": ["index"],
        },
    ),

    # ── Paragraph Formatting ──
    Tool(
        name="word_set_paragraph_alignment",
        description="Set text alignment for a specific paragraph.",
        inputSchema={
            "type": "object",
            "properties": {
                "index": {
                    "type": "integer",
                    "description": "1-based paragraph index.",
                },
                "alignment": {
                    "type": "string",
                    "description": "Alignment: 'left', 'center', 'right', or 'justify'.",
                    "enum": ["left", "center", "right", "justify"],
                },
            },
            "required": ["index", "alignment"],
        },
    ),
    Tool(
        name="word_set_paragraph_spacing",
        description="Set paragraph spacing (before, after, line spacing).",
        inputSchema={
            "type": "object",
            "properties": {
                "index": {
                    "type": "integer",
                    "description": "1-based paragraph index.",
                },
                "before": {
                    "type": "number",
                    "description": "Space before in points.",
                },
                "after": {
                    "type": "number",
                    "description": "Space after in points.",
                },
                "line_spacing": {
                    "type": "number",
                    "description": "Line spacing multiplier (e.g., 1.0, 1.5, 2.0).",
                },
            },
            "required": ["index"],
        },
    ),

    # ── Font Formatting ──
    Tool(
        name="word_set_font_bold",
        description="Set or remove bold formatting on a range (selection or whole content).",
        inputSchema={
            "type": "object",
            "properties": {
                "bold": {
                    "type": "boolean",
                    "description": "True to make bold, False to remove bold.",
                    "default": True,
                },
                "range_spec": {
                    "type": "string",
                    "description": "Where to apply: 'selection' (default), 'content' (entire document), or 'start=X,end=Y' for specific character range.",
                    "default": "selection",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="word_set_font_italic",
        description="Set or remove italic formatting on a range.",
        inputSchema={
            "type": "object",
            "properties": {
                "italic": {
                    "type": "boolean",
                    "description": "True to make italic, False to remove italic.",
                    "default": True,
                },
                "range_spec": {
                    "type": "string",
                    "description": "Where to apply: 'selection', 'content', or 'start=X,end=Y'.",
                    "default": "selection",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="word_set_font_underline",
        description="Set or remove underline formatting on a range.",
        inputSchema={
            "type": "object",
            "properties": {
                "underline": {
                    "type": "boolean",
                    "description": "True to underline, False to remove.",
                    "default": True,
                },
                "range_spec": {
                    "type": "string",
                    "description": "Where to apply: 'selection', 'content', or 'start=X,end=Y'.",
                    "default": "selection",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="word_set_font_name",
        description="Set the font name on a range (e.g., 'Arial', 'Times New Roman', 'Calibri').",
        inputSchema={
            "type": "object",
            "properties": {
                "font_name": {
                    "type": "string",
                    "description": "Font name (e.g., 'Arial', 'Calibri', 'Times New Roman').",
                },
                "range_spec": {
                    "type": "string",
                    "description": "Where to apply: 'selection', 'content', or 'start=X,end=Y'.",
                    "default": "selection",
                },
            },
            "required": ["font_name"],
        },
    ),
    Tool(
        name="word_set_font_size",
        description="Set the font size in points on a range.",
        inputSchema={
            "type": "object",
            "properties": {
                "size": {
                    "type": "number",
                    "description": "Font size in points (e.g., 12, 14, 18).",
                },
                "range_spec": {
                    "type": "string",
                    "description": "Where to apply: 'selection', 'content', or 'start=X,end=Y'.",
                    "default": "selection",
                },
            },
            "required": ["size"],
        },
    ),
    Tool(
        name="word_set_font_color",
        description="Set the font (text) color on a range using an RGB hex string.",
        inputSchema={
            "type": "object",
            "properties": {
                "color": {
                    "type": "string",
                    "description": "Font color as RGB hex string (e.g., 'FF0000' for red, '0000FF' for blue).",
                },
                "range_spec": {
                    "type": "string",
                    "description": "Where to apply: 'selection', 'content', or 'start=X,end=Y'.",
                    "default": "selection",
                },
            },
            "required": ["color"],
        },
    ),

    # ── Find / Replace ──
    Tool(
        name="word_find_text",
        description="Find text in the document. Selects the first match and returns its position.",
        inputSchema={
            "type": "object",
            "properties": {
                "search_text": {
                    "type": "string",
                    "description": "Text to search for.",
                },
                "match_case": {
                    "type": "boolean",
                    "description": "If True, match case. Default: false.",
                    "default": False,
                },
                "match_whole_word": {
                    "type": "boolean",
                    "description": "If True, match whole words only. Default: false.",
                    "default": False,
                },
            },
            "required": ["search_text"],
        },
    ),
    Tool(
        name="word_find_replace",
        description="Find and replace text in the document. Returns the number of replacements.",
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
                "match_case": {
                    "type": "boolean",
                    "description": "If True, match case. Default: false.",
                    "default": False,
                },
                "match_whole_word": {
                    "type": "boolean",
                    "description": "If True, match whole words only. Default: false.",
                    "default": False,
                },
                "replace_all": {
                    "type": "boolean",
                    "description": "If True, replace all occurrences. If False, replace only the first. Default: true.",
                    "default": True,
                },
            },
            "required": ["find_text", "replace_text"],
        },
    ),

    # ── Table Operations ──
    Tool(
        name="word_add_table",
        description="Add a table at the end of the document. Optionally populate cells with text (tab-separated rows, newline-separated lines).",
        inputSchema={
            "type": "object",
            "properties": {
                "rows": {
                    "type": "integer",
                    "description": "Number of rows in the table.",
                },
                "cols": {
                    "type": "integer",
                    "description": "Number of columns in the table.",
                },
                "text": {
                    "type": "string",
                    "description": "Optional text to populate the table. Use \\t for column separators and \\n for row separators.",
                    "default": "",
                },
            },
            "required": ["rows", "cols"],
        },
    ),
    Tool(
        name="word_get_table_count",
        description="Get the number of tables in the active document.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    Tool(
        name="word_get_table_data",
        description="Get all data from a specific table as a 2D array.",
        inputSchema={
            "type": "object",
            "properties": {
                "index": {
                    "type": "integer",
                    "description": "1-based table index.",
                },
            },
            "required": ["index"],
        },
    ),
    Tool(
        name="word_set_cell_text",
        description="Set text in a specific table cell.",
        inputSchema={
            "type": "object",
            "properties": {
                "table_index": {
                    "type": "integer",
                    "description": "1-based table index.",
                },
                "row": {
                    "type": "integer",
                    "description": "1-based row index within the table.",
                },
                "col": {
                    "type": "integer",
                    "description": "1-based column index within the table.",
                },
                "text": {
                    "type": "string",
                    "description": "Text to set in the cell.",
                },
            },
            "required": ["table_index", "row", "col", "text"],
        },
    ),
    Tool(
        name="word_add_table_row",
        description="Add a new row to the end of a specified table.",
        inputSchema={
            "type": "object",
            "properties": {
                "table_index": {
                    "type": "integer",
                    "description": "1-based table index.",
                },
            },
            "required": ["table_index"],
        },
    ),
    Tool(
        name="word_add_table_column",
        description="Add a new column to the end of a specified table.",
        inputSchema={
            "type": "object",
            "properties": {
                "table_index": {
                    "type": "integer",
                    "description": "1-based table index.",
                },
            },
            "required": ["table_index"],
        },
    ),
    Tool(
        name="word_delete_table",
        description="Delete a specific table by index.",
        inputSchema={
            "type": "object",
            "properties": {
                "index": {
                    "type": "integer",
                    "description": "1-based table index to delete.",
                },
            },
            "required": ["index"],
        },
    ),

    # ── Page Setup ──
    Tool(
        name="word_set_page_orientation",
        description="Set the page orientation: portrait or landscape.",
        inputSchema={
            "type": "object",
            "properties": {
                "orientation": {
                    "type": "string",
                    "description": "Page orientation: 'portrait' or 'landscape'.",
                    "enum": ["portrait", "landscape"],
                },
            },
            "required": ["orientation"],
        },
    ),
    Tool(
        name="word_set_page_margins",
        description="Set page margins in points (72 points = 1 inch).",
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
            },
            "required": [],
        },
    ),
    Tool(
        name="word_set_page_size",
        description="Set page width and height in points (for A4: width=595, height=842).",
        inputSchema={
            "type": "object",
            "properties": {
                "width": {
                    "type": "number",
                    "description": "Page width in points.",
                },
                "height": {
                    "type": "number",
                    "description": "Page height in points.",
                },
            },
            "required": [],
        },
    ),

    # ── Header / Footer ──
    Tool(
        name="word_set_header",
        description="Set the primary header text for the first section.",
        inputSchema={
            "type": "object",
            "properties": {
                "text": {
                    "type": "string",
                    "description": "Header text.",
                },
            },
            "required": ["text"],
        },
    ),
    Tool(
        name="word_set_footer",
        description="Set the primary footer text for the first section.",
        inputSchema={
            "type": "object",
            "properties": {
                "text": {
                    "type": "string",
                    "description": "Footer text.",
                },
            },
            "required": ["text"],
        },
    ),

    # ── Export ──
    Tool(
        name="word_export_to_pdf",
        description="Export the active document to a PDF file.",
        inputSchema={
            "type": "object",
            "properties": {
                "filepath": {
                    "type": "string",
                    "description": "Full path for the output PDF file (e.g., 'C:\\report.pdf').",
                },
            },
            "required": ["filepath"],
        },
    ),

    # ── Insert Elements ──
    Tool(
        name="word_insert_picture",
        description="Insert an image into the document.",
        inputSchema={
            "type": "object",
            "properties": {
                "filepath": {
                    "type": "string",
                    "description": "Full path to the image file (.png, .jpg, etc.).",
                },
                "left": {
                    "type": "number",
                    "description": "Optional left position in points.",
                },
                "top": {
                    "type": "number",
                    "description": "Optional top position in points.",
                },
                "width": {
                    "type": "number",
                    "description": "Optional width in points.",
                },
                "height": {
                    "type": "number",
                    "description": "Optional height in points.",
                },
            },
            "required": ["filepath"],
        },
    ),
    Tool(
        name="word_insert_page_break",
        description="Insert a page break at the current cursor position.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    Tool(
        name="word_add_section_break",
        description="Add a section break (next page) at the end of the document.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),

    # ── View / Zoom ──
    Tool(
        name="word_set_zoom",
        description="Set the zoom level of the active document window.",
        inputSchema={
            "type": "object",
            "properties": {
                "percentage": {
                    "type": "integer",
                    "description": "Zoom percentage (10-500).",
                },
            },
            "required": ["percentage"],
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


def _execute_tool(name: str, args: dict[str, Any], client: WPSWordClient) -> str:
    """Execute a tool synchronously. Called from the STA executor thread."""
    pythoncom.CoInitialize()

    result: Any = None

    # ── Application ──
    if name == "word_get_app_info":
        result = client.get_app_info()

    elif name == "word_show_window":
        client.show()
        result = {"message": "WPS Word window is now visible."}

    elif name == "word_hide_window":
        client.hide()
        result = {"message": "WPS Word window is now hidden."}

    elif name == "word_quit_app":
        client.quit_app()
        result = {"message": "WPS Word application quit."}

    # ── Document Management ──
    elif name == "word_create_document":
        doc_name = client.create_document()
        result = {"message": f"Created new document: {doc_name}", "document_name": doc_name}

    elif name == "word_open_document":
        doc_name = client.open_document(args["filepath"])
        result = {"message": f"Opened document: {doc_name}", "document_name": doc_name}

    elif name == "word_save_document":
        filepath = args.get("filepath")
        saved_path = client.save_document(filepath)
        result = {"message": f"Document saved to: {saved_path}", "path": saved_path}

    elif name == "word_close_document":
        save = args.get("save", True)
        client.close_document(save)
        result = {"message": "Document closed." + (" (saved)" if save else " (not saved)")}

    elif name == "word_list_documents":
        result = {"documents": client.list_documents()}

    elif name == "word_activate_document":
        doc_name = client.activate_document(args["name"])
        result = {"message": f"Activated document: {doc_name}", "document_name": doc_name}

    # ── Text Operations ──
    elif name == "word_get_text":
        text = client.get_text()
        result = {"text": text, "length": len(text)}

    elif name == "word_get_selected_text":
        text = client.get_selected_text()
        result = {"text": text, "length": len(text)}

    elif name == "word_set_text":
        client.set_text(args["text"])
        result = {"message": "Document text replaced.", "length": len(args["text"])}

    elif name == "word_type_text":
        client.type_text(args["text"])
        result = {"message": f"Typed text at current position."}

    elif name == "word_insert_text_at_end":
        client.insert_text_at_end(args["text"])
        result = {"message": f"Appended text at end of document."}

    elif name == "word_insert_text_at_start":
        client.insert_text_at_start(args["text"])
        result = {"message": f"Inserted text at start of document."}

    # ── Paragraph Operations ──
    elif name == "word_add_paragraph":
        text = args.get("text", "")
        idx = client.add_paragraph(text)
        result = {"message": f"Added paragraph at index {idx}.", "paragraph_index": idx}

    elif name == "word_get_paragraph_count":
        count = client.get_paragraph_count()
        result = {"paragraph_count": count}

    elif name == "word_get_paragraph_text":
        text = client.get_paragraph_text(args["index"])
        result = {"paragraph_index": args["index"], "text": text}

    elif name == "word_set_paragraph_text":
        client.set_paragraph_text(args["index"], args["text"])
        result = {"message": f"Set paragraph {args['index']} text."}

    elif name == "word_insert_paragraph_before":
        client.insert_paragraph_before(args["index"], args.get("text", ""))
        result = {"message": f"Inserted paragraph before index {args['index']}."}

    elif name == "word_delete_paragraph":
        client.delete_paragraph(args["index"])
        result = {"message": f"Deleted paragraph {args['index']}."}

    # ── Paragraph Formatting ──
    elif name == "word_set_paragraph_alignment":
        client.set_paragraph_alignment(args["index"], args["alignment"])
        result = {"message": f"Set paragraph {args['index']} alignment = {args['alignment']}."}

    elif name == "word_set_paragraph_spacing":
        client.set_paragraph_spacing(
            args["index"],
            args.get("before"),
            args.get("after"),
            args.get("line_spacing"),
        )
        result = {"message": f"Set paragraph {args['index']} spacing."}

    # ── Font Formatting ──
    elif name == "word_set_font_bold":
        bold = args.get("bold", True)
        range_spec = args.get("range_spec", "selection")
        client.set_font_bold(bold, range_spec)
        result = {"message": f"Set bold = {bold} on '{range_spec}'."}

    elif name == "word_set_font_italic":
        italic = args.get("italic", True)
        range_spec = args.get("range_spec", "selection")
        client.set_font_italic(italic, range_spec)
        result = {"message": f"Set italic = {italic} on '{range_spec}'."}

    elif name == "word_set_font_underline":
        underline = args.get("underline", True)
        range_spec = args.get("range_spec", "selection")
        client.set_font_underline(underline, range_spec)
        result = {"message": f"Set underline = {underline} on '{range_spec}'."}

    elif name == "word_set_font_name":
        range_spec = args.get("range_spec", "selection")
        client.set_font_name(args["font_name"], range_spec)
        result = {"message": f"Set font name = '{args['font_name']}' on '{range_spec}'."}

    elif name == "word_set_font_size":
        range_spec = args.get("range_spec", "selection")
        client.set_font_size(args["size"], range_spec)
        result = {"message": f"Set font size = {args['size']} on '{range_spec}'."}

    elif name == "word_set_font_color":
        color_hex = args["color"].lstrip("#")
        color_int = int(color_hex, 16)
        range_spec = args.get("range_spec", "selection")
        client.set_font_color(color_int, range_spec)
        result = {"message": f"Set font color = #{color_hex} on '{range_spec}'."}

    # ── Find / Replace ──
    elif name == "word_find_text":
        found = client.find_text(
            args["search_text"],
            args.get("match_case", False),
            args.get("match_whole_word", False),
        )
        if found is None:
            result = {"found": False, "message": f"'{args['search_text']}' not found."}
        else:
            result = found

    elif name == "word_find_replace":
        count = client.find_replace(
            args["find_text"],
            args["replace_text"],
            args.get("match_case", False),
            args.get("match_whole_word", False),
            args.get("replace_all", True),
        )
        result = {
            "message": f"Replaced '{args['find_text']}' with '{args['replace_text']}'.",
            "replacements": count,
        }

    # ── Table Operations ──
    elif name == "word_add_table":
        idx = client.add_table(
            args["rows"],
            args["cols"],
            args.get("text", ""),
        )
        result = {"message": f"Added table with {args['rows']} rows x {args['cols']} cols at index {idx}.", "table_index": idx}

    elif name == "word_get_table_count":
        count = client.get_table_count()
        result = {"table_count": count}

    elif name == "word_get_table_data":
        data = client.get_table_data(args["index"])
        result = {"table_index": args["index"], "rows": len(data), "columns": len(data[0]) if data else 0, "data": data}

    elif name == "word_set_cell_text":
        client.set_cell_text(args["table_index"], args["row"], args["col"], args["text"])
        result = {"message": f"Set table {args['table_index']} cell ({args['row']},{args['col']}) text."}

    elif name == "word_add_table_row":
        client.add_table_row(args["table_index"])
        result = {"message": f"Added row to table {args['table_index']}."}

    elif name == "word_add_table_column":
        client.add_table_column(args["table_index"])
        result = {"message": f"Added column to table {args['table_index']}."}

    elif name == "word_delete_table":
        client.delete_table(args["index"])
        result = {"message": f"Deleted table {args['index']}."}

    # ── Page Setup ──
    elif name == "word_set_page_orientation":
        client.set_page_orientation(args["orientation"])
        result = {"message": f"Set page orientation to {args['orientation']}."}

    elif name == "word_set_page_margins":
        client.set_page_margins(
            args.get("left"),
            args.get("right"),
            args.get("top"),
            args.get("bottom"),
        )
        result = {"message": "Set page margins."}

    elif name == "word_set_page_size":
        client.set_page_size(
            args.get("width"),
            args.get("height"),
        )
        result = {"message": "Set page size."}

    # ── Header / Footer ──
    elif name == "word_set_header":
        client.add_header(args["text"])
        result = {"message": "Set header text."}

    elif name == "word_set_footer":
        client.add_footer(args["text"])
        result = {"message": "Set footer text."}

    # ── Export ──
    elif name == "word_export_to_pdf":
        saved_path = client.export_to_pdf(args["filepath"])
        result = {"message": f"Exported to PDF: {saved_path}", "path": saved_path}

    # ── Insert Elements ──
    elif name == "word_insert_picture":
        shape_name = client.insert_picture(
            args["filepath"],
            args.get("left"),
            args.get("top"),
            args.get("width"),
            args.get("height"),
        )
        result = {"message": f"Inserted picture: {shape_name}", "shape_name": shape_name}

    elif name == "word_insert_page_break":
        client.insert_page_break()
        result = {"message": "Inserted page break."}

    elif name == "word_add_section_break":
        client.add_section_break()
        result = {"message": "Added section break (next page)."}

    # ── View ──
    elif name == "word_set_zoom":
        client.set_zoom(args["percentage"])
        result = {"message": f"Set zoom to {args['percentage']}%."}

    else:
        raise ValueError(f"Unknown tool: {name}")

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
