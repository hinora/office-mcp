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

    # ── Style Operations ──
    Tool(
        name="word_apply_style",
        description="Apply a named Word style (e.g., 'Heading 1', 'Normal', 'Title', 'Subtitle') to the specified range.",
        inputSchema={
            "type": "object",
            "properties": {
                "style_name": {
                    "type": "string",
                    "description": "The style name (e.g., 'Heading 1', 'Normal', 'Title', 'Subtitle').",
                },
                "range_spec": {
                    "type": "string",
                    "description": "Where to apply: 'selection' (default), 'content' (entire document), or 'start=X,end=Y'.",
                    "default": "selection",
                },
            },
            "required": ["style_name"],
        },
    ),

    # ── List Formatting ──
    Tool(
        name="word_set_list_format",
        description="Apply bullet or numbered list formatting to paragraphs in the range.",
        inputSchema={
            "type": "object",
            "properties": {
                "list_type": {
                    "type": "string",
                    "description": "Type of list: 'bullet' or 'number'.",
                    "enum": ["bullet", "number"],
                },
                "range_spec": {
                    "type": "string",
                    "description": "Where to apply: 'selection' (default), 'content', or 'start=X,end=Y'.",
                    "default": "selection",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="word_remove_list_format",
        description="Remove bullet/numbered list formatting from paragraphs.",
        inputSchema={
            "type": "object",
            "properties": {
                "range_spec": {
                    "type": "string",
                    "description": "Where to apply: 'selection' (default), 'content', or 'start=X,end=Y'.",
                    "default": "selection",
                },
            },
            "required": [],
        },
    ),

    # ── Hyperlink ──
    Tool(
        name="word_add_hyperlink",
        description="Add a hyperlink to the document (URL, file link, email).",
        inputSchema={
            "type": "object",
            "properties": {
                "address": {
                    "type": "string",
                    "description": "URL, file path, or email address for the link.",
                },
                "text_to_display": {
                    "type": "string",
                    "description": "Optional display text. Defaults to the address if not provided.",
                },
                "range_spec": {
                    "type": "string",
                    "description": "Where to add: 'selection' (default), 'content' (appends at end), or 'start=X,end=Y'.",
                    "default": "selection",
                },
            },
            "required": ["address"],
        },
    ),

    # ── Table of Contents ──
    Tool(
        name="word_insert_table_of_contents",
        description="Insert a Table of Contents at the current cursor position. Uses Heading 1-3 styles for entries.",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),

    # ── Page Numbers ──
    Tool(
        name="word_insert_page_numbers",
        description="Insert page numbers in the header or footer of the first section.",
        inputSchema={
            "type": "object",
            "properties": {
                "position": {
                    "type": "string",
                    "description": "Where to place: 'bottom' (footer, default) or 'top' (header).",
                    "enum": ["bottom", "top"],
                    "default": "bottom",
                },
            },
            "required": [],
        },
    ),

    # ── Document Properties ──
    Tool(
        name="word_get_document_properties",
        description="Get document metadata (author, title, subject, keywords, comments, etc.).",
        inputSchema={
            "type": "object",
            "properties": {},
            "required": [],
        },
    ),
    Tool(
        name="word_set_document_properties",
        description="Set document metadata (author, title, subject, keywords).",
        inputSchema={
            "type": "object",
            "properties": {
                "author": {
                    "type": "string",
                    "description": "Document author.",
                },
                "title": {
                    "type": "string",
                    "description": "Document title.",
                },
                "subject": {
                    "type": "string",
                    "description": "Document subject.",
                },
                "keywords": {
                    "type": "string",
                    "description": "Document keywords (comma separated).",
                },
            },
            "required": [],
        },
    ),

    # ── Comments ──
    Tool(
        name="word_add_comment",
        description="Add a comment to the specified range in the document.",
        inputSchema={
            "type": "object",
            "properties": {
                "text": {
                    "type": "string",
                    "description": "Comment text.",
                },
                "range_spec": {
                    "type": "string",
                    "description": "Where to anchor the comment: 'selection' (default), 'content', or 'start=X,end=Y'.",
                    "default": "selection",
                },
            },
            "required": ["text"],
        },
    ),

    # ── Highlight ──
    Tool(
        name="word_set_highlight",
        description="Set text highlight color on a range. Common colors: 6=Yellow, 7=Green, 2=Blue, 13=Pink, 0=None.",
        inputSchema={
            "type": "object",
            "properties": {
                "color_index": {
                    "type": "integer",
                    "description": "Highlight color index: 0=None, 6=Yellow, 7=Bright Green, 2=Blue, 13=Pink, 15=Gray.",
                },
                "range_spec": {
                    "type": "string",
                    "description": "Where to apply: 'selection' (default), 'content', or 'start=X,end=Y'.",
                    "default": "selection",
                },
            },
            "required": ["color_index"],
        },
    ),

    # ── Table Style ──
    Tool(
        name="word_set_table_style",
        description="Apply a built-in style to a table (e.g., 'Table Grid', 'Table Style Medium 1').",
        inputSchema={
            "type": "object",
            "properties": {
                "table_index": {
                    "type": "integer",
                    "description": "1-based table index.",
                },
                "style_name": {
                    "type": "string",
                    "description": "Table style name (e.g., 'Table Grid', 'Table Style Medium 1', 'Table Style Light 1').",
                },
            },
            "required": ["table_index", "style_name"],
        },
    ),

    # ── Page Borders ──
    Tool(
        name="word_set_page_borders",
        description="Add or modify page borders for the document. Use line_style 1 for single, 2 for dotted, 3 for dashed.",
        inputSchema={
            "type": "object",
            "properties": {
                "line_style": {
                    "type": "integer",
                    "description": "Border line style: 1=single, 2=dot, 3=dash, 4=dash-dot. Default 1.",
                    "default": 1,
                },
                "line_width": {
                    "type": "integer",
                    "description": "Border width: 4=0.5pt, 6=0.75pt, 8=1pt, 12=1.5pt. Default 4.",
                    "default": 4,
                },
                "distance": {
                    "type": "integer",
                    "description": "Distance from page edge in points. Default 24.",
                    "default": 24,
                },
            },
            "required": [],
        },
    ),

    # ── Watermark ──
    Tool(
        name="word_add_watermark",
        description="Add a text watermark to the document (e.g., 'CONFIDENTIAL', 'DRAFT').",
        inputSchema={
            "type": "object",
            "properties": {
                "text": {
                    "type": "string",
                    "description": "Watermark text (e.g., 'CONFIDENTIAL', 'DRAFT').",
                },
                "font_size": {
                    "type": "integer",
                    "description": "Font size in points. Default 72.",
                    "default": 72,
                },
                "color": {
                    "type": "string",
                    "description": "RGB hex color string (e.g., 'C0C0C0' for light gray). Default is silver.",
                },
                "layout": {
                    "type": "string",
                    "description": "Layout: 'diagonal' (default) or 'horizontal'.",
                    "enum": ["diagonal", "horizontal"],
                    "default": "diagonal",
                },
            },
            "required": ["text"],
        },
    ),

    # ── Document Protection ──
    Tool(
        name="word_protect_document",
        description="Protect the document as read-only. Optionally set a password.",
        inputSchema={
            "type": "object",
            "properties": {
                "password": {
                    "type": "string",
                    "description": "Optional password to protect with.",
                    "default": "",
                },
            },
            "required": [],
        },
    ),
    Tool(
        name="word_unprotect_document",
        description="Remove protection from the document (password needed if set).",
        inputSchema={
            "type": "object",
            "properties": {
                "password": {
                    "type": "string",
                    "description": "Password if the document was protected with one.",
                    "default": "",
                },
            },
            "required": [],
        },
    ),

    # ── Track Changes ──
    Tool(
        name="word_toggle_track_changes",
        description="Enable or disable Track Changes (revision tracking).",
        inputSchema={
            "type": "object",
            "properties": {
                "enable": {
                    "type": "boolean",
                    "description": "True to enable track changes, False to disable. Default: true.",
                    "default": True,
                },
            },
            "required": [],
        },
    ),

    # ── Columns ──
    Tool(
        name="word_set_columns",
        description="Set multi-column layout for the first section (newsletter-style).",
        inputSchema={
            "type": "object",
            "properties": {
                "num_columns": {
                    "type": "integer",
                    "description": "Number of columns (1-4). 1 = single column. Default 1.",
                    "default": 1,
                },
                "spacing": {
                    "type": "number",
                    "description": "Column spacing in points.",
                },
            },
            "required": [],
        },
    ),

    # ── Bookmarks ──
    Tool(
        name="word_add_bookmark",
        description="Add a bookmark at the specified range (for navigation).",
        inputSchema={
            "type": "object",
            "properties": {
                "name": {
                    "type": "string",
                    "description": "Bookmark name (must be unique, use underscores instead of spaces).",
                },
                "range_spec": {
                    "type": "string",
                    "description": "Where to add: 'selection' (default), 'content', or 'start=X,end=Y'.",
                    "default": "selection",
                },
            },
            "required": ["name"],
        },
    ),
    Tool(
        name="word_go_to_bookmark",
        description="Navigate to a bookmark by name and return its position.",
        inputSchema={
            "type": "object",
            "properties": {
                "name": {
                    "type": "string",
                    "description": "Bookmark name to navigate to.",
                },
            },
            "required": ["name"],
        },
    ),

    # ── Print ──
    Tool(
        name="word_print_document",
        description="Print the active document to the default printer.",
        inputSchema={
            "type": "object",
            "properties": {
                "copies": {
                    "type": "integer",
                    "description": "Number of copies to print. Default 1.",
                    "default": 1,
                },
            },
            "required": [],
        },
    ),

    # ── Range Text ──
    Tool(
        name="word_get_range_text",
        description="Get text from a specific 0-based character range (e.g., start=0, end=100).",
        inputSchema={
            "type": "object",
            "properties": {
                "start": {
                    "type": "integer",
                    "description": "0-based start character position.",
                },
                "end": {
                    "type": "integer",
                    "description": "0-based end character position.",
                },
            },
            "required": ["start", "end"],
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

    # ── Style Operations ──
    elif name == "word_apply_style":
        range_spec = args.get("range_spec", "selection")
        client.apply_style(args["style_name"], range_spec)
        result = {"message": f"Applied style '{args['style_name']}' on '{range_spec}'."}

    # ── List Formatting ──
    elif name == "word_set_list_format":
        list_type = args.get("list_type", "bullet")
        range_spec = args.get("range_spec", "selection")
        client.set_list_format(list_type, range_spec)
        result = {"message": f"Applied {list_type} list format on '{range_spec}'."}

    elif name == "word_remove_list_format":
        range_spec = args.get("range_spec", "selection")
        client.remove_list_format(range_spec)
        result = {"message": f"Removed list format from '{range_spec}'."}

    # ── Hyperlink ──
    elif name == "word_add_hyperlink":
        range_spec = args.get("range_spec", "selection")
        client.add_hyperlink(
            args["address"],
            args.get("text_to_display"),
            range_spec,
        )
        result = {"message": f"Added hyperlink to '{args['address']}'."}

    # ── Table of Contents ──
    elif name == "word_insert_table_of_contents":
        client.insert_table_of_contents()
        result = {"message": "Inserted Table of Contents."}

    # ── Page Numbers ──
    elif name == "word_insert_page_numbers":
        client.insert_page_numbers(args.get("position", "bottom"))
        result = {"message": f"Inserted page numbers in {args.get('position', 'bottom')}."}

    # ── Document Properties ──
    elif name == "word_get_document_properties":
        result = client.get_document_properties()

    elif name == "word_set_document_properties":
        client.set_document_properties(
            args.get("author"),
            args.get("title"),
            args.get("subject"),
            args.get("keywords"),
        )
        result = {"message": "Set document properties."}

    # ── Comments ──
    elif name == "word_add_comment":
        range_spec = args.get("range_spec", "selection")
        client.add_comment(args["text"], range_spec)
        result = {"message": f"Added comment on '{range_spec}'."}

    # ── Highlight ──
    elif name == "word_set_highlight":
        range_spec = args.get("range_spec", "selection")
        client.set_highlight(args["color_index"], range_spec)
        result = {"message": f"Set highlight color index={args['color_index']} on '{range_spec}'."}

    # ── Table Style ──
    elif name == "word_set_table_style":
        client.set_table_style(args["table_index"], args["style_name"])
        result = {"message": f"Applied style '{args['style_name']}' to table {args['table_index']}."}

    # ── Page Borders ──
    elif name == "word_set_page_borders":
        client.set_page_borders(
            args.get("line_style", 1),
            args.get("line_width", 4),
            args.get("distance", 24),
        )
        result = {"message": "Set page borders."}

    # ── Watermark ──
    elif name == "word_add_watermark":
        color_hex = args.get("color")
        color_int = None
        if color_hex:
            color_int = int(color_hex.lstrip("#"), 16)
        client.add_watermark(
            args["text"],
            args.get("font_size", 72),
            color_int,
            args.get("layout", "diagonal"),
        )
        result = {"message": f"Added watermark: '{args['text']}'."}

    # ── Document Protection ──
    elif name == "word_protect_document":
        client.protect_document(args.get("password", ""))
        result = {"message": "Document protected (read-only)."}

    elif name == "word_unprotect_document":
        client.unprotect_document(args.get("password", ""))
        result = {"message": "Document unprotected."}

    # ── Track Changes ──
    elif name == "word_toggle_track_changes":
        enable = args.get("enable", True)
        client.toggle_track_changes(enable)
        result = {"message": f"Track changes {'enabled' if enable else 'disabled'}."}

    # ── Columns ──
    elif name == "word_set_columns":
        client.set_columns(
            args.get("num_columns", 1),
            args.get("spacing"),
        )
        result = {"message": f"Set columns to {args.get('num_columns', 1)}."}

    # ── Bookmarks ──
    elif name == "word_add_bookmark":
        range_spec = args.get("range_spec", "selection")
        client.add_bookmark(args["name"], range_spec)
        result = {"message": f"Added bookmark '{args['name']}'."}

    elif name == "word_go_to_bookmark":
        bookmark_info = client.go_to_bookmark(args["name"])
        result = {"message": f"Navigated to bookmark '{args['name']}'.", **bookmark_info}

    # ── Print ──
    elif name == "word_print_document":
        client.print_document(args.get("copies", 1))
        result = {"message": f"Printing {args.get('copies', 1)} copy/copies."}

    # ── Range Text ──
    elif name == "word_get_range_text":
        text = client.get_range_text(args["start"], args["end"])
        result = {"start": args["start"], "end": args["end"], "text": text}

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
