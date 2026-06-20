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
# Scrolls the row into view first so the coordinates are valid for clicking.
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
            // Scroll the row into view so coordinates are in viewport
            r.scrollIntoView({block: 'center', behavior: 'instant'});
            // Small delay to let layout settle
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
# Queries [data-pre-plain-text] globally (not scoped to a specific container
# since WhatsApp Web no longer uses a dedicated conversation-panel-messages
# data-testid). Each element has a data-pre-plain-text attribute like
# "[18:30, 3/29/2023] Hinora: " from which we parse the sender name.
EXTRACT_MESSAGES_JS = r"""
(function() {
    const COUNT = %d;

    // Scroll the message area (any scrollable ancestor) to load older messages
    var scrollable = document.querySelector(
        '[data-testid="msg-container"], ' +
        '#app div[style*="overflow"]'
    );
    if (scrollable && scrollable.scrollTop !== undefined) {
        scrollable.scrollTop = 0;
    }

    // Query ALL elements with data-pre-plain-text (globally — no specific
    // container needed since sidebar rows don't carry this attribute)
    const rows = document.querySelectorAll('[data-pre-plain-text]');
    const results = [];
    for (const row of rows) {
        const preText = row.getAttribute('data-pre-plain-text') || '';
        // Format: "[HH:MM, DD/MM/YYYY] Name: " or "[HH:MM, Name] "
        let sender = '';
        const preMatch = preText.match(/\]\s*(.+?)\s*:?\s*$/);
        if (preMatch) sender = preMatch[1].trim();

        // Get text from selectable-text child (newer layout) or msg-text (older)
        const selectable = row.querySelector(
            '[data-testid="selectable-text"], ' +
            '.selectable-text, ' +
            '.copyable-text, ' +
            '[data-testid="msg-text"]'
        );
        let text = selectable
            ? selectable.textContent.trim()
            : row.textContent.trim();
        // Strip the [HH:MM, ...] prefix if it leaked into textContent
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
# Uses span[title] (most reliable across WhatsApp versions) for the display
# name, and scans the conversation area for status/subtitle text.
CONTACT_INFO_JS = r"""
(function() {
    var displayName = '';
    var status = '';

    // ── Best source: span[title] anywhere in #app ───────────────
    // Only picks up span[title] that are NOT in the sidebar
    // (sidebar titles are chat list rows, not the conversation header)
    var titleSpans = document.querySelectorAll('#app span[title]');
    for (var i = 0; i < titleSpans.length; i++) {
        var s = titleSpans[i];
        var t = s.getAttribute('title') || '';
        // Skip sidebar rows and wordmark/logo elements
        if (s.closest('#side') || s.closest('#pane-side')) continue;
        if (t === 'wa-wordmark-refreshed' || t.match(/^w[ads]{1,2}-/i)) continue;
        if (t.length > 0) {
            displayName = t;
            break;
        }
    }

    // ── Fallback: any heading-like element in the conversation area ──
    if (!displayName) {
        var headings = document.querySelectorAll(
            '#app h1, #app h2, #app [role="heading"]'
        );
        for (var j = 0; j < headings.length; j++) {
            var h = headings[j];
            if (!h.closest('#side') && !h.closest('#pane-side')) {
                var txt = h.textContent.trim();
                if (txt && txt.length < 80) {
                    displayName = txt;
                    break;
                }
            }
        }
    }

    // ── Status/subtitle ─────────────────────────────────────────
    var subtitleEl = document.querySelector('[data-testid="chat-subtitle"]');
    if (subtitleEl) {
        status = subtitleEl.textContent.trim();
    }

    // ── Phone (try to find it in the contact panel) ─────────────
    var phone = '';
    var phoneEl = document.querySelector(
        '[data-testid="chat-subtitle"] span[dir="auto"], ' +
        'span[title*="+"], ' +
        'span[title*="84"], ' +
        'span[title*="1-"]'
    );
    if (phoneEl) {
        var p = phoneEl.getAttribute('title') || phoneEl.textContent || '';
        var pm = p.match(/\+[\d\s\-\(\)]{6,20}/);
        if (pm) phone = pm[0].trim();
    }

    return JSON.stringify({
        name: displayName || 'Unknown',
        status: status || '',
        phone: phone
    });
})();
"""

# Type text into the search box and wait for results.
# Uses a polling loop (up to 6 s) to handle slow WhatsApp search rendering.
SEARCH_AND_GET_RESULTS_JS = r"""
(async function() {
    const QUERY = %s;
    const searchBox = document.querySelector(
        '#side input[type="text"], ' +
        '[data-testid="chat-list-search-container"] input[type="text"]'
    );
    if (!searchBox) return JSON.stringify({error: 'Search box not found'});

    // Helper: extract names from the sidebar rows right now
    const collectRows = () => {
        const pane = document.querySelector('#pane-side');
        if (!pane) return [];
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
        return results;
    };

    // Focus and clear
    searchBox.focus();
    searchBox.select();
    const nativeInputValueSetter = Object.getOwnPropertyDescriptor(
        window.HTMLInputElement.prototype, 'value'
    ).set;
    nativeInputValueSetter.call(searchBox, '');
    searchBox.dispatchEvent(new Event('input', {bubbles: true}));
    await new Promise(r => setTimeout(r, 300));

    // Snapshot rows BEFORE typing to compare later
    const preRows = collectRows().map(c => c.name);

    // Type query
    nativeInputValueSetter.call(searchBox, QUERY);
    searchBox.dispatchEvent(new Event('input', {bubbles: true}));

    // Poll for results — up to 6 s, checking every 500 ms.
    // Stop early when the row set has actually changed (search rendered).
    const MAX_WAIT = 6000;
    const POLL_INTERVAL = 500;
    let waited = 0;
    let finalResults = [];
    while (waited < MAX_WAIT) {
        await new Promise(r => setTimeout(r, POLL_INTERVAL));
        waited += POLL_INTERVAL;
        const current = collectRows();
        const names = current.map(c => c.name);
        // If the list changed vs pre-search state, results have loaded
        const changed = names.length !== preRows.length ||
            !names.every((n, i) => n === preRows[i]);
        if (changed && current.length > 0) {
            finalResults = current;
            break;
        }
        // Keep the last batch even if no clear change detected
        if (current.length > 0) finalResults = current;
    }
    // Fallback: one final read
    if (finalResults.length === 0) {
        finalResults = collectRows();
    }
    return JSON.stringify({query: QUERY, results: finalResults.slice(0, 20)});
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


# Serialize all tool calls — only ONE at a time.
# Concurrent calls would race on the same WhatsApp Web page (clicks,
# typing, DOM reads) and produce corrupted/interleaved results.
_tool_lock = asyncio.Lock()


@server.call_tool()
async def call_tool(name: str, arguments: dict[str, Any]) -> list[TextContent]:
    async with _tool_lock:
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

# Verify that the currently open chat matches the expected name.
# Uses a layered approach:
#   1. Check if a message container exists — proves ANY chat is open
#   2. Search the full DOM for the expected name in text content
#   3. Fall back to span[title] and document.title
VERIFY_CHAT_OPEN_JS = r"""
(function() {
    const EXPECTED = %s;
    const lowered = EXPECTED.toLowerCase();

    // ── Layer 1: Is any chat actually open? ─────────────────────
    const msgPane = document.querySelector(
        '[data-testid="conversation-panel-messages"], ' +
        '#main [data-testid="msg-container"], ' +
        '[data-testid="conversation-panel-messages"] > div'
    );
    const hasOpenChat = !!msgPane;

    // ── Layer 2: Does the expected name appear in the page? ────
    // Scan all span[title] attributes (sidebar rows + conversation header)
    const titleSpans = document.querySelectorAll('span[title]');
    for (const s of titleSpans) {
        const t = s.getAttribute('title') || '';
        if (t.toLowerCase().includes(lowered)) {
            return JSON.stringify({verified: true, actual_name: t, source: 'span[title]'});
        }
    }

    // ── Layer 3: document.title (works in tab mode, not PWA) ───
    const docTitle = (document.title || '').replace(/^\(\d+\)\s*/, '').trim();
    if (docTitle && docTitle.toLowerCase().includes(lowered)) {
        return JSON.stringify({verified: true, actual_name: docTitle, source: 'title'});
    }

    // ── Layer 4: URL hash (works in tab mode, not PWA) ─────────
    const hash = location.hash || '';
    if (hash && (hash.toLowerCase().includes(encodeURIComponent(EXPECTED).toLowerCase()) ||
        hash.toLowerCase().includes(lowered))) {
        return JSON.stringify({verified: true, actual_name: EXPECTED, source: 'url'});
    }

    // ── Layer 5: body innerText substring scan ─────────────────
    const bodyText = (document.body.innerText || '').substring(0, 500).toLowerCase();
    const foundInBody = bodyText.includes(lowered);

    // ── Failure diagnostics ────────────────────────────────────
    return JSON.stringify({
        verified: false,
        has_open_chat: hasOpenChat,
        found_in_body: foundInBody,
        reason: hasOpenChat && foundInBody
            ? 'Chat seems open but name not confirmed via title spans'
            : (!hasOpenChat ? 'No message pane visible' : 'Name not found in page'),
        actual_name: 'unknown',
        doc_title: docTitle.substring(0, 80),
        hash_sample: hash.substring(0, 80)
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
    Retries verification with progressive waits to handle slow DOM rendering.
    """
    await _clear_search()  # ensure full chat list is visible
    await asyncio.sleep(0.3)
    js = FIND_CHAT_BOX_JS % json.dumps(chat_name)
    info = await _eval_json(js)
    if "error" in info:
        return info
    await _click_at(info["x"], info["y"])

    # Post-click verification with retries for slow DOM rendering
    verify_js = VERIFY_CHAT_OPEN_JS % json.dumps(chat_name)
    for attempt, wait in enumerate((1.5, 2.0, 2.5), 1):
        await asyncio.sleep(wait)
        verify = await _eval_json(verify_js)
        if verify.get("verified"):
            logger.info(f"Chat '{chat_name}' verified via {verify.get('source')} (attempt {attempt})")
            return info
        reason = verify.get("reason", "")
        has_open = verify.get("has_open_chat", False)
        # Retry if no message pane yet (DOM still loading) or
        # name not in page yet (slow rendering).
        # Break early if a message pane exists but wrong name — different chat opened.
        if not has_open:
            logger.info(f"No message pane yet on attempt {attempt}, retrying...")
        elif verify.get("found_in_body"):
            # Chat seems open, just verification didn't confirm — one more retry
            logger.info(f"Chat open but name not in title spans on attempt {attempt}, retrying...")
        else:
            logger.warning(f"Wrong chat opened (has pane but name absent), breaking")
            break

    return {
        "error": (
            f"Chat verification failed: expected '{chat_name}'. "
            f"Has open chat: {verify.get('has_open_chat', False)}. "
            f"Name in body: {verify.get('found_in_body', False)}. "
            f"Reason: {verify.get('reason', 'unknown')}. "
            f"Title: {verify.get('doc_title', 'N/A')}. "
            f"Target chat may have scrolled out of view or DOM changed."
        )
    }


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
    known_name = info.get("name", contact_name)
    result["chat_name"] = known_name
    # Prefer the verified name from _open_chat over unreliable DOM probes
    if not result.get("name") or result["name"] in ("Unknown", "wa-wordmark-refreshed"):
        result["name"] = known_name
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
    """Clear the search box and return to the full chat list.

    Dispatches Escape on the search input element (not document) to
    properly trigger React's synthetic event handlers.
    Falls back to clearing the input value directly.
    """
    await _eval("""
    (function(){
        var input = document.querySelector(
            '#side input[type="text"], ' +
            '[data-testid="chat-list-search-container"] input[type="text"]'
        );
        if (!input) return;
        // Focus the input first so it receives the key event
        input.focus();
        // Dispatch Escape keydown on the input itself (not document)
        input.dispatchEvent(new KeyboardEvent('keydown', {
            key: 'Escape', code: 'Escape', keyCode: 27,
            which: 27, bubbles: true, cancelable: true
        }));
        // Also try input event to clear
        var setter = Object.getOwnPropertyDescriptor(
            window.HTMLInputElement.prototype, 'value'
        );
        if (setter && setter.set) {
            setter.set.call(input, '');
            input.dispatchEvent(new Event('input', {bubbles: true}));
            input.dispatchEvent(new Event('change', {bubbles: true}));
        }
        // Blur to trigger any blur handlers
        input.blur();
    })()
    """)
    await asyncio.sleep(1.0)


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
