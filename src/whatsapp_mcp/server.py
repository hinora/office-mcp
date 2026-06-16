r"""
MCP Server for WhatsApp Web — READ-ONLY bridge via Chrome CDP.

                         ─── SETUP INSTRUCTIONS ───

  Launch Chrome with remote debugging BEFORE running this server:

    "C:\Program Files\Google\Chrome\Application\chrome.exe" --app=https://web.whatsapp.com --remote-debugging-port=9222 --user-data-dir="C:\WhatsAppMCPProfile"

  ─────────────── END SETUP INSTRUCTIONS ────────────────────────────────

Provides 9 read-only MCP tools. No send, reply, delete, archive, or mutate.
"""

from __future__ import annotations

import asyncio
import base64
import json
import logging
import tempfile
from typing import Any

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

try:
    from .cdp_client import WhatsAppClient
except ImportError:
    from whatsapp_mcp.cdp_client import WhatsAppClient

logger = logging.getLogger(__name__)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)

server = Server("whatsapp-mcp")
_client: WhatsAppClient | None = None


async def get_client() -> WhatsAppClient:
    global _client
    if _client is None:
        _client = WhatsAppClient()
    return _client


async def _eval(expression: str, await_promise: bool = False) -> Any:
    client = await get_client()
    return await client.evaluate(expression, await_promise)


async def _eval_json(expression: str, await_promise: bool = False) -> Any:
    raw = await _eval(expression, await_promise)
    if isinstance(raw, str):
        return json.loads(raw)
    return raw


# ── Click Safety ───────────────────────────────────────────────────────

# Get bounding boxes of danger zones (compose area, send button) in the
# current viewport. These are regions where a click could accidentally
# send a message, reply, or trigger other destructive actions.
DANGER_ZONE_JS = r"""
(function() {
    const zones = [];
    // Message compose area (footer with text input and send button)
    const footer = document.querySelector(
        'footer[data-testid="conversation-compose-box"], ' +
        '[data-testid="conversation-compose-box"], ' +
        'footer[role="contentinfo"]'
    );
    if (footer) {
        const r = footer.getBoundingClientRect();
        zones.push({id: 'compose_area', x: r.left, y: r.top,
                     width: r.width, height: r.height});
    }
    // Send button (standalone, sometimes outside footer in older layouts)
    const sendBtn = document.querySelector(
        'button[data-testid="compose-btn-send"], ' +
        'span[data-icon="send"]'
    );
    if (sendBtn) {
        const r = sendBtn.getBoundingClientRect();
        zones.push({id: 'send_button', x: r.left, y: r.top,
                     width: r.width, height: r.height});
    }
    // Popup menus / context menus (archive, delete, etc.)
    const menus = document.querySelectorAll(
        '[data-testid="chat-list-archive-panel"], ' +
        'div[role="menu"], div[data-animate-menu-item="true"]'
    );
    for (const m of menus) {
        const r = m.getBoundingClientRect();
        if (r.width > 0 && r.height > 0) {
            zones.push({id: 'context_menu', x: r.left, y: r.top,
                         width: r.width, height: r.height});
        }
    }
    return JSON.stringify({danger_zones: zones});
})();
"""


async def _click_at(x: float, y: float) -> None:
    """Click at coordinates with danger-zone validation.

    Before clicking, queries the page for danger zones (compose area,
    send button, context menus) and refuses to click if the coordinates
    fall within any of them.
    """
    # Safety: check for danger zones before clicking
    try:
        danger = await _eval_json(DANGER_ZONE_JS)
        for dz in danger.get("danger_zones", []):
            dx, dy, dw, dh = dz["x"], dz["y"], dz["width"], dz["height"]
            if dx <= x <= dx + dw and dy <= y <= dy + dh:
                raise RuntimeError(
                    f"SAFETY BLOCK: Click at ({x:.0f}, {y:.0f}) would land in "
                    f"danger zone '{dz['id']}' "
                    f"(x={dx:.0f}, y={dy:.0f}, w={dw:.0f}, h={dh:.0f}). "
                    f"Refusing to click to prevent accidental message send, "
                    f"reply, archive, or delete."
                )
    except RuntimeError:
        raise
    except Exception:
        # If danger-zone check itself fails, still proceed but log it
        logger.warning("Danger-zone check failed, proceeding with click", exc_info=True)

    client = await get_client()
    await client.click_at(x, y)


# ── JS Snippets ────────────────────────────────────────────────────────

# Convert blob URLs to base64 data URLs. Pass a JSON array of URLs.
BLOB_TO_DATA_URL_JS = r"""
(async function() {
    const URLS = %s;
    const results = [];
    for (const url of URLS) {
        try {
            const resp = await fetch(url);
            const blob = await resp.blob();
            const mime = blob.type || 'image/png';
            const dataUrl = await new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onloadend = () => resolve(reader.result);
                reader.onerror = reject;
                reader.readAsDataURL(blob);
            });
            results.push({url: url, data_url: dataUrl, mime: mime, size: blob.size});
        } catch(e) {
            results.push({url: url, error: e.message});
        }
    }
    return JSON.stringify(results);
})();
"""

# Get the bounding box of the first sidebar row matching a chat name.
FIND_CHAT_BOX_JS = r"""
(function() {
    const NAME = %s;
    const pane = document.querySelector('#pane-side');
    if (!pane) return JSON.stringify({error: 'Chat sidebar not found'});
    const rows = pane.querySelectorAll('[role="row"]');
    for (const r of rows) {
        const titleEl = r.querySelector('span[title]');
        const name = (titleEl && titleEl.getAttribute('title')) || r.textContent || '';
        if (name.toLowerCase().includes(NAME.toLowerCase())) {
            const rect = r.getBoundingClientRect();
            return JSON.stringify({
                found: true,
                name: name,
                x: rect.left + rect.width / 2,
                y: rect.top + rect.height / 2,
            });
        }
    }
    return JSON.stringify({error: 'Chat not found: ' + NAME});
})();
"""

# Extract messages from the currently open chat.
# Targets only leaf elements with data-pre-plain-text (unique per message).
# The new WhatsApp format includes a date: "[HH:MM, DD/MM/YYYY] Name: "
EXTRACT_MESSAGES_JS = r"""
(function() {
    const COUNT = %d;
    const container = document.querySelector(
        '[data-testid="conversation-panel-messages"]'
    );
    if (!container) return JSON.stringify({error: 'Message pane not found'});

    // Scroll up to load older messages
    container.scrollTop = 0;
    setTimeout(() => { container.scrollTop = 0; }, 300);

    // Target only leaf elements with data-pre-plain-text (unique per message)
    const rows = container.querySelectorAll('[data-pre-plain-text]');
    const results = [];
    for (const row of rows) {
        const preText = row.getAttribute('data-pre-plain-text') || '';
        // Format: "[HH:MM, DD/MM/YYYY] Name: " or "[HH:MM, Name] "
        let sender = '';
        const preMatch = preText.match(/\]\s*(.+?)\s*:?\s*$/);
        if (preMatch) sender = preMatch[1].trim();

        // Get text from the selectable-text or copyable-text child
        const selectable = row.querySelector(
            '.selectable-text, [data-testid="msg-text"]'
        );
        let text = selectable
            ? selectable.textContent.trim()
            : row.textContent.trim();
        // Strip the [HH:MM, ...] prefix if it leaked into textcontent
        text = text.replace(/^\[?\d{1,2}:\d{2}[,\]]\s*[^\n]*?\n?/, '').trim();

        if (!text && !sender) continue;
        results.push({sender: sender || 'Unknown', text: text || ''});
        if (results.length >= COUNT) break;
    }
    return JSON.stringify({
        count: results.length,
        messages: results.reverse(),
        total_in_view: rows.length
    });
})();
"""

# Get contact info from currently open chat.
CONTACT_INFO_JS = r"""
(function() {
    const header = document.querySelector(
        '[data-testid="conversation-header"], ' +
        '[data-testid="conversation-info-header"]'
    );
    if (!header) return JSON.stringify({error: 'Header not found'});

    const titleEl = header.querySelector(
        '[data-testid="conversation-info-header-chat-title"], ' +
        'span[title]'
    );
    const displayName = titleEl
        ? (titleEl.getAttribute('title') || titleEl.textContent || '').trim()
        : '';

    const subtitleEl = header.querySelector(
        '[data-testid="chat-subtitle"]'
    );
    const status = subtitleEl ? subtitleEl.textContent.trim() : '';

    return JSON.stringify({
        name: displayName,
        status: status,
        phone: '',
    });
})();
"""

# Type text into the search box and wait for results.
SEARCH_AND_GET_RESULTS_JS = r"""
(async function() {
    const QUERY = %s;
    const searchBox = document.querySelector(
        '#side input[type="text"], ' +
        '[data-testid="chat-list-search-container"] input[type="text"]'
    );
    if (!searchBox) return JSON.stringify({error: 'Search box not found'});

    // Focus and clear
    searchBox.focus();
    searchBox.select();
    // Clear by setting value and dispatching input event
    const nativeInputValueSetter = Object.getOwnPropertyDescriptor(
        window.HTMLInputElement.prototype, 'value'
    ).set;
    nativeInputValueSetter.call(searchBox, '');
    searchBox.dispatchEvent(new Event('input', {bubbles: true}));
    await new Promise(r => setTimeout(r, 300));

    // Type query
    nativeInputValueSetter.call(searchBox, QUERY);
    searchBox.dispatchEvent(new Event('input', {bubbles: true}));
    await new Promise(r => setTimeout(r, 2000));

    // Read results
    const pane = document.querySelector('#pane-side');
    if (!pane) return JSON.stringify({error: 'Sidebar not found'});
    const rows = pane.querySelectorAll('[role="row"]');
    const results = [];
    for (const r of rows) {
        const titleEl = r.querySelector('span[title]');
        const name = titleEl ? titleEl.getAttribute('title') : '';
        if (!name) continue;
        const previewEls = r.querySelectorAll('span[dir="auto"]');
        let preview = '';
        for (const s of previewEls) {
            const t = s.textContent.trim();
            if (t && t !== name && !t.match(/^\d{1,2}:\d{2}/)) {
                preview = t; break;
            }
        }
        results.push({name, preview});
    }
    return JSON.stringify({query: QUERY, results: results.slice(0, 20)});
})();
"""

# List visible chats from sidebar.
LIST_CHATS_JS = r"""
(function() {
    const pane = document.querySelector('#pane-side');
    if (!pane) return JSON.stringify({error: 'Chat sidebar not found'});
    const rows = pane.querySelectorAll('[role="row"]');
    const results = [];
    for (const r of rows) {
        const titleEl = r.querySelector('span[title]');
        const name = titleEl ? titleEl.getAttribute('title') : '';
        if (!name) continue;
        const previewEls = r.querySelectorAll('span[dir="auto"]');
        let preview = '';
        for (const s of previewEls) {
            const t = s.textContent.trim();
            if (t && t !== name && !t.match(/^\d{1,2}:\d{2}/)) {
                preview = t; break;
            }
        }
        const badge = r.querySelector(
            '[data-testid="icon-unread-count"], ' +
            'span[aria-label*="unread"]'
        );
        const unread = badge
            ? (badge.textContent || badge.getAttribute('aria-label') || '').trim()
            : '';
        let time = '';
        const allSpans = r.querySelectorAll('span');
        for (const s of allSpans) {
            const t = s.textContent.trim();
            if (t.match(/^\d{1,2}:\d{2}/) ||
                t.match(/^(yesterday|today|sun|mon|tue|wed|thu|fri|sat)/i)) {
                time = t; break;
            }
        }
        results.push({name, preview, unread, time});
    }
    return JSON.stringify({chats: results.slice(0, 50), total: results.length});
})();
"""

# Get bounding box of search box.
FIND_SEARCH_BOX_JS = r"""
(function() {
    const el = document.querySelector(
        '#side input[type="text"], ' +
        '[data-testid="chat-list-search-container"] input[type="text"]'
    );
    if (!el) return JSON.stringify({error: 'Search box not found'});
    const rect = el.getBoundingClientRect();
    return JSON.stringify({
        found: true,
        x: rect.left + rect.width / 2,
        y: rect.top + rect.height / 2,
    });
})();
"""

# ── Tool Definitions ───────────────────────────────────────────────────

TOOLS: list[Tool] = [
    Tool(
        name="whatsapp_list_chats",
        description=(
            "List all visible chats from the WhatsApp sidebar. "
            "Returns each chat's name, last message preview, unread badge, "
            "and timestamp. Purely read-only."
        ),
        inputSchema={"type": "object", "properties": {}, "required": []},
    ),
    Tool(
        name="whatsapp_read_chat_messages",
        description=(
            "Read recent messages from a specific WhatsApp chat. "
            "Opens the chat, scrolls to load history, and extracts messages. "
            "Each message includes sender name and text. Purely read-only."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "chat_name": {
                    "type": "string",
                    "description": "Name of the chat/contact (partial match, case-insensitive)",
                },
                "count": {
                    "type": "integer",
                    "description": "Max messages to return (default: 20)",
                },
            },
            "required": ["chat_name"],
        },
    ),
    Tool(
        name="whatsapp_get_contact_info",
        description=(
            "Get information about a WhatsApp contact or group. "
            "Opens the chat and reads the header. Returns display name, "
            "status/bio, and phone. Purely read-only."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "contact_name": {
                    "type": "string",
                    "description": "Name of the contact or group (partial match, case-insensitive)",
                },
            },
            "required": ["contact_name"],
        },
    ),
    Tool(
        name="whatsapp_search_chats",
        description=(
            "Search WhatsApp chats by keyword. Types into the search box "
            "and returns matching chats. Purely read-only."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "query": {
                    "type": "string",
                    "description": "Search keyword to find matching chats",
                },
            },
            "required": ["query"],
        },
    ),
    Tool(
        name="whatsapp_get_chat_media",
        description=(
            "Get recent media files (images, videos, documents) visible "
            "in a chat. Blob images are downloaded to local temp files "
            "and returned as readable file paths. Purely read-only."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "chat_name": {
                    "type": "string",
                    "description": "Name of the chat/contact (partial match, case-insensitive)",
                },
                "count": {
                    "type": "integer",
                    "description": "Max media items to return (default: 10)",
                },
            },
            "required": ["chat_name"],
        },
    ),
    Tool(
        name="whatsapp_download_image",
        description=(
            "Download a WhatsApp image blob URL to a local file. "
            "Use this after whatsapp_get_chat_media returns blob:// URLs. "
            "Saves to a temp PNG/JPG file and returns the path so the "
            "agent can view it. Purely read-only."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "url": {
                    "type": "string",
                    "description": "The blob: URL from a media item",
                },
                "save_path": {
                    "type": "string",
                    "description": "Optional local file path to save to (default: temp file)",
                },
            },
            "required": ["url"],
        },
    ),
    Tool(
        name="whatsapp_screenshot",
        description=(
            "Take a screenshot of the current WhatsApp Web view. "
            "Saves to a temporary PNG file and returns the file path. "
            "Purely read-only."
        ),
        inputSchema={"type": "object", "properties": {}, "required": []},
    ),
    Tool(
        name="whatsapp_view_media",
        description=(
            "View an image or media file that was previously downloaded. "
            "Takes a file path (from whatsapp_get_chat_media, "
            "whatsapp_download_image, or whatsapp_screenshot) and returns "
            "the actual image so you can see it. Purely read-only."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "file_path": {
                    "type": "string",
                    "description": "Absolute path to the media file to view",
                },
            },
            "required": ["file_path"],
        },
    ),
    Tool(
        name="whatsapp_status",
        description=(
            "Get the current connection status for the WhatsApp MCP bridge. "
            "Reports whether Chrome is reachable and the WhatsApp tab is found."
        ),
        inputSchema={"type": "object", "properties": {}, "required": []},
    ),
]


# ── Tool Handler ───────────────────────────────────────────────────────

@server.list_tools()
async def list_tools() -> list[Tool]:
    return TOOLS


@server.call_tool()
async def call_tool(name: str, arguments: dict[str, Any]) -> list[TextContent]:
    try:
        return await _dispatch(name, arguments)
    except Exception as e:
        logger.exception(f"Error executing tool '{name}'")
        return [TextContent(type="text", text=json.dumps({
            "error": str(e), "tool": name,
        }, ensure_ascii=False, separators=(",", ":")))]


async def _dispatch(name: str, args: dict[str, Any]) -> list[TextContent]:
    if name == "whatsapp_list_chats":
        return _to_text(await _eval_json(LIST_CHATS_JS))

    elif name == "whatsapp_read_chat_messages":
        chat_name = args["chat_name"]
        count = args.get("count", 20)
        return _to_text(await _read_messages(chat_name, count))

    elif name == "whatsapp_get_contact_info":
        return _to_text(await _get_contact(args["contact_name"]))

    elif name == "whatsapp_search_chats":
        return _to_text(await _search_chats(args["query"]))

    elif name == "whatsapp_get_chat_media":
        return _to_text(await _get_media(args["chat_name"], args.get("count", 10)))

    elif name == "whatsapp_download_image":
        dl = await _download_image(args["url"], args.get("save_path"))
        if "error" in dl:
            return _to_text(dl)
        return await _read_image_to_text(dl["saved_path"], extra=dl)

    elif name == "whatsapp_screenshot":
        client = await get_client()
        img_bytes = await client.screenshot("png")
        tmp = tempfile.NamedTemporaryFile(
            suffix=".png", delete=False, prefix="whatsapp_"
        )
        tmp.write(img_bytes)
        tmp.close()
        b64 = base64.b64encode(img_bytes).decode("ascii")
        data_url = f"data:image/png;base64,{b64}"
        return _image_text(
            data_url=data_url,
            file_path=tmp.name,
            size_bytes=len(img_bytes),
            mime_type="image/png",
        )

    elif name == "whatsapp_view_media":
        file_path = args["file_path"]
        return await _read_image_to_text(file_path)

    elif name == "whatsapp_status":
        client = await get_client()
        try:
            status = await client.connect()
            connected = (status == "ok")
        except Exception as e:
            connected = False
            status = str(e)
        return _to_text({"connected": connected, "detail": status})

    else:
        return _to_text({"error": f"Unknown tool: {name}"})


def _to_text(data: dict[str, Any]) -> list[TextContent]:
    """Wrap a dict as a TextContent list."""
    return [TextContent(type="text", text=json.dumps(
        data, ensure_ascii=False, default=str, separators=(",", ":")
    ))]


async def _read_image_to_text(file_path: str, extra: dict | None = None) -> list[TextContent]:
    """Read a local image file and return text with an embedded Markdown data: URI.

    Chatbox and many chat UIs auto-render Markdown images with data: URIs,
    so the agent/LLM can actually see the image.
    """
    from pathlib import Path
    p = Path(file_path)
    if not p.exists():
        return _to_text({"error": f"File not found: {file_path}"})

    raw_bytes = p.read_bytes()
    b64 = base64.b64encode(raw_bytes).decode("ascii")

    ext = p.suffix.lower()
    mime_map = {
        ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
        ".webp": "image/webp", ".gif": "image/gif", ".bmp": "image/bmp",
    }
    mime = mime_map.get(ext, "image/png")
    data_url = f"data:{mime};base64,{b64}"

    return _image_text(
        data_url=data_url,
        file_path=str(p),
        size_bytes=len(raw_bytes),
        mime_type=mime,
        extra=extra,
    )


def _image_text(
    data_url: str,
    file_path: str,
    size_bytes: int,
    mime_type: str,
    extra: dict | None = None,
) -> list[TextContent]:
    """Build a text response with an embedded Markdown data: URI image.

    Strategy: return TWO TextContent blocks:
      1. A Markdown image tag that Chatbox renders visually.
      2. Metadata as JSON so the agent also gets text context.
    """
    # Short mime label for display
    label = mime_type.split("/")[-1].upper()
    info = extra.copy() if extra else {}
    info.update({
        "file_path": file_path,
        "size_bytes": size_bytes,
        "mime_type": mime_type,
    })

    return [TextContent(type="text", text=(
        f"📷 **{label} image** ({size_bytes:,} bytes)\n\n"
        f"![image]({data_url})\n\n"
        f"```json\n{json.dumps(info, ensure_ascii=False, default=str, indent=2)}\n```"
    ))]

    info["size_bytes"] = len(raw_bytes)
    info["mime_type"] = mime
    result.append(TextContent(type="text", text=json.dumps(info)))
    return result


# ── Post-Click Verification ────────────────────────────────────────────

# Verify that the currently open chat header matches the expected name.
VERIFY_CHAT_OPEN_JS = r"""
(function() {
    const EXPECTED = %s;
    const header = document.querySelector(
        '[data-testid="conversation-header"], ' +
        '[data-testid="conversation-info-header"]'
    );
    if (!header) return JSON.stringify({verified: false, reason: 'Header not found'});
    const titleEl = header.querySelector(
        '[data-testid="conversation-info-header-chat-title"], ' +
        'span[title]'
    );
    const actual = titleEl
        ? (titleEl.getAttribute('title') || titleEl.textContent || '').trim()
        : '';
    const match = actual.toLowerCase().includes(EXPECTED.toLowerCase());
    return JSON.stringify({
        verified: match,
        actual_name: actual,
        expected: EXPECTED
    });
})();
"""


# ── Tool Implementations (multi-step CDP) ──────────────────────────────

async def _open_chat(chat_name: str) -> dict:
    """Click a chat in the sidebar to open it. Returns the box info or error.

    Includes post-click verification: after clicking, checks that the
    conversation header actually shows the expected chat name. If the
    wrong chat opened (e.g., mis-click on a different element), returns
    an error instead of silently operating on the wrong conversation.
    """
    await _clear_search()  # ensure full chat list is visible
    await asyncio.sleep(0.3)
    js = FIND_CHAT_BOX_JS % json.dumps(chat_name)
    info = await _eval_json(js)
    if "error" in info:
        return info
    await _click_at(info["x"], info["y"])
    await asyncio.sleep(1.5)
    # Post-click verification: did the right chat actually open?
    verify_js = VERIFY_CHAT_OPEN_JS % json.dumps(chat_name)
    verify = await _eval_json(verify_js)
    if not verify.get("verified"):
        return {
            "error": (
                f"Chat verification failed: expected '{chat_name}' but "
                f"header shows '{verify.get('actual_name', 'unknown')}'. "
                f"Reason: {verify.get('reason', 'name mismatch')}. "
                f"Target chat may have scrolled out of view or DOM changed."
            )
        }
    return info


async def _read_messages(chat_name: str, count: int) -> dict:
    """Open a chat and extract messages."""
    info = await _open_chat(chat_name)
    if "error" in info:
        return info
    # Wait a bit more for messages to render, then scroll
    await asyncio.sleep(0.5)
    js = EXTRACT_MESSAGES_JS % count
    return await _eval_json(js)


async def _get_contact(contact_name: str) -> dict:
    """Open a chat and read contact header info."""
    info = await _open_chat(contact_name)
    if "error" in info:
        return info
    await asyncio.sleep(0.5)
    result = await _eval_json(CONTACT_INFO_JS)
    result["chat_name"] = info.get("name", contact_name)
    return result


async def _search_chats(query: str) -> dict:
    """Focus search box, type query, and return results."""
    # First, click the search box
    box = await _eval_json(FIND_SEARCH_BOX_JS)
    if "error" in box:
        return box
    await _click_at(box["x"], box["y"])
    await asyncio.sleep(0.5)
    # Now type and wait for results
    js = SEARCH_AND_GET_RESULTS_JS % json.dumps(query)
    return await _eval_json(js, await_promise=True)


async def _clear_search() -> None:
    """Clear the search box and return to the full chat list."""
    await _eval("""
    (function(){
        // Press Escape to exit search mode
        const event = new KeyboardEvent('keydown', {
            key: 'Escape', code: 'Escape', keyCode: 27,
            which: 27, bubbles: true
        });
        document.dispatchEvent(event);
        // Also clear the search box directly
        const input = document.querySelector(
            '#side input[type="text"], ' +
            '[data-testid="chat-list-search-container"] input[type="text"]'
        );
        if (input) {
            const setter = Object.getOwnPropertyDescriptor(
                window.HTMLInputElement.prototype, 'value'
            ).set;
            setter.call(input, '');
            input.dispatchEvent(new Event('input', {bubbles: true}));
        }
    })()
    """)
    await asyncio.sleep(0.5)


async def _download_blobs(blob_urls: list[str]) -> list[dict]:
    """Convert blob:// URLs to local temp files by reading them from the page via CDP.

    Uses fetch() inside the WhatsApp page to read each blob, converts to
    base64 data URL, returns it to Python, and saves to a temp file.
    Returns a list of {url, saved_path, mime, size_bytes} dicts.
    """
    if not blob_urls:
        return []
    js = BLOB_TO_DATA_URL_JS % json.dumps(blob_urls)
    raw = await _eval(js, await_promise=True)
    entries = json.loads(raw) if isinstance(raw, str) else raw
    results = []
    for e in entries:
        if "error" in e:
            results.append({"url": e["url"], "error": e["error"]})
            continue
        data_url = e.get("data_url", "")
        if not data_url or "," not in data_url:
            results.append({"url": e["url"], "error": "empty data_url"})
            continue
        # data_url format: "data:image/png;base64,ABC123..."
        header, b64 = data_url.split(",", 1)
        ext = ".png"
        if "jpeg" in header or "jpg" in header:
            ext = ".jpg"
        elif "webp" in header:
            ext = ".webp"
        elif "gif" in header:
            ext = ".gif"
        elif "mp4" in header:
            ext = ".mp4"
        try:
            raw_bytes = base64.b64decode(b64)
        except Exception:
            results.append({"url": e["url"], "error": "base64 decode failed"})
            continue
        tmp = tempfile.NamedTemporaryFile(
            suffix=ext, delete=False, prefix="whatsapp_media_"
        )
        tmp.write(raw_bytes)
        tmp.close()
        results.append({
            "url": e["url"],
            "saved_path": tmp.name,
            "mime": e.get("mime", ""),
            "size_bytes": len(raw_bytes),
        })
    return results


async def _get_media(chat_name: str, count: int) -> dict:
    """Open a chat, extract media info, and download blob images to local files."""
    info = await _open_chat(chat_name)
    if "error" in info:
        return info
    await asyncio.sleep(0.5)
    js = r"""
    (function() {
        const COUNT = %d;
        const container = document.querySelector(
            '[data-testid="conversation-panel-messages"]'
        );
        if (!container) return JSON.stringify({error: 'Message pane not found'});
        container.scrollTop = 0;
        const media = [];
        const imgs = container.querySelectorAll(
            'img[src*="blob"], img[src*="whatsapp"], img[src*="cdn"]'
        );
        for (const img of imgs) {
            if (media.length >= COUNT) break;
            const src = img.getAttribute('src') || '';
            const alt = img.getAttribute('alt') || '';
            if (src) media.push({type: 'image', src: src, caption: alt});
        }
        const videos = container.querySelectorAll('video');
        for (const v of videos) {
            if (media.length >= COUNT) break;
            const src = v.getAttribute('src') || '';
            if (src) media.push({type: 'video', src: src});
        }
        const docs = container.querySelectorAll(
            '[data-testid="document-thumb"], [data-testid="audio-player"]'
        );
        for (const d of docs) {
            if (media.length >= COUNT) break;
            media.push({type: 'document', info: d.textContent.trim().slice(0, 100)});
        }
        return JSON.stringify({chat: %s, media: media, count: media.length});
    })();
    """ % (count, json.dumps(info.get("name", chat_name)))
    data = await _eval_json(js)
    if "error" in data:
        return data

    # Download blob images/videos to local temp files
    blob_urls = [
        m["src"] for m in data.get("media", [])
        if m.get("src", "").startswith("blob:")
    ]
    if blob_urls:
        downloads = await _download_blobs(blob_urls)
        # Map blob URL → local path
        by_url = {d["url"]: d for d in downloads}
        for m in data["media"]:
            url = m.get("src", "")
            if url in by_url:
                dl = by_url[url]
                if "saved_path" in dl:
                    m["saved_path"] = dl["saved_path"]
                    m["size_bytes"] = dl.get("size_bytes", 0)
                    m["mime"] = dl.get("mime", "")
                else:
                    m["download_error"] = dl.get("error", "unknown")
    return data


async def _download_image(url: str, save_path: str | None = None) -> dict:
    """Download a single WhatsApp blob/source image to a local file."""
    downloads = await _download_blobs([url])
    if not downloads:
        return {"error": "No data returned from browser"}
    dl = downloads[0]
    if "error" in dl:
        return {"error": dl["error"], "url": url}
    if save_path:
        import shutil
        shutil.move(dl["saved_path"], save_path)
        dl["saved_path"] = save_path
    return dl




# ── Entry Point ─────────────────────────────────────────────────────────

async def run_server() -> None:
    async with stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream, write_stream,
            server.create_initialization_options(),
        )


def main() -> None:
    asyncio.run(run_server())


if __name__ == "__main__":
    main()
