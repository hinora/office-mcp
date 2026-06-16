"""
Chrome DevTools Protocol (CDP) client for WhatsApp Web.

Connects to an ALREADY RUNNING Chrome instance via the remote debugging port
(http://localhost:9222). Does NOT launch a new browser. Scans open tabs to find
web.whatsapp.com and provides a JavaScript evaluation interface for read-only
tool queries.

Architecture:
    GET http://localhost:9222/json  →  list of open page targets
    ws://localhost:9222/devtools/page/<id>  →  CDP WebSocket to WhatsApp page
    Runtime.evaluate  →  run JS in the page and return results
    Page.captureScreenshot  →  take a screenshot
"""

from __future__ import annotations

import asyncio
import base64
import json
import logging
import urllib.request
import urllib.error
from typing import Any

import websockets

logger = logging.getLogger(__name__)

# ── Constants ──────────────────────────────────────────────────────────

CHROME_DEBUG_URL = "http://localhost:9222"
WHATSAPP_URL_PATTERN = "web.whatsapp.com"
WS_CONNECT_TIMEOUT = 10.0


# ── CDP Client ─────────────────────────────────────────────────────────


class CDPClient:
    """Manages a CDP WebSocket connection to a single Chrome page target."""

    def __init__(self, ws_url: str, page_title: str, page_url: str):
        self.ws_url = ws_url
        self.page_title = page_title
        self.page_url = page_url
        self._ws: websockets.WebSocketClientProtocol | None = None
        self._msg_id = 0
        self._connected = False

    @property
    def connected(self) -> bool:
        return self._connected

    async def connect(self) -> None:
        """Open the CDP WebSocket and enable required domains."""
        try:
            self._ws = await websockets.connect(
                self.ws_url,
                max_size=2**23,
                ping_interval=30,
                ping_timeout=10,
                open_timeout=WS_CONNECT_TIMEOUT,
                close_timeout=5,
            )
            self._connected = True
            logger.info(f"Connected to CDP: {self.page_title}")
            await self._send_cmd("Runtime.enable")
            await self._send_cmd("Page.enable")
        except Exception:
            self._connected = False
            raise

    async def disconnect(self) -> None:
        """Close the CDP WebSocket connection."""
        self._connected = False
        if self._ws:
            try:
                await self._ws.close()
            except Exception:
                pass
            self._ws = None
        logger.info(f"Disconnected from CDP: {self.page_title}")

    async def evaluate(self, expression: str, await_promise: bool = False) -> Any:
        """Run a JavaScript expression in the page and return the result value.

        ⚠️ SECURITY: This can execute ARBITRARY JavaScript in the WhatsApp page,
        including destructive actions (send message, delete, archive, etc.).
        Callers MUST only pass READ-ONLY expressions that do not mutate state.

        Args:
            expression: JavaScript code to evaluate (MUST be read-only).
            await_promise: If True, wait for the returned Promise to resolve.

        Returns:
            The result value (converted to Python types via returnByValue).
        """
        result = await self._send_cmd("Runtime.evaluate", {
            "expression": expression,
            "awaitPromise": await_promise,
            "returnByValue": True,
        })
        res = result.get("result", {})
        if res.get("type") == "object" and res.get("subtype") == "error":
            desc = res.get("description", "unknown error")
            raise RuntimeError(f"JS evaluation error: {desc}")
        return res.get("value")

    async def capture_screenshot(self, format: str = "png", quality: int = 80) -> bytes:
        """Take a screenshot of the page.

        Returns raw bytes (PNG or JPEG).
        """
        params: dict = {"format": format}
        if format == "jpeg":
            params["quality"] = quality
        result = await self._send_cmd("Page.captureScreenshot", params)
        data_b64 = result.get("data", "")
        return base64.b64decode(data_b64)

    async def click_at(self, x: float, y: float) -> None:
        """Perform a left-click at screen coordinates via CDP Input domain.

        Uses real OS-level mouse events that React event handlers respond to.
        """
        await self._send_cmd("Input.dispatchMouseEvent", {
            "type": "mousePressed",
            "x": x, "y": y,
            "button": "left", "clickCount": 1,
        })
        await asyncio.sleep(0.05)
        await self._send_cmd("Input.dispatchMouseEvent", {
            "type": "mouseReleased",
            "x": x, "y": y,
            "button": "left", "clickCount": 1,
        })

    async def _send_cmd(self, method: str, params: dict | None = None) -> dict:
        """Send a CDP command and wait for the response."""
        if not self._ws:
            raise RuntimeError("Not connected")
        self._msg_id += 1
        cmd = {
            "id": self._msg_id,
            "method": method,
            "params": params or {},
        }
        await self._ws.send(json.dumps(cmd))

        while True:
            raw = await self._ws.recv()
            resp = json.loads(raw)
            if resp.get("id") == self._msg_id:
                if "error" in resp:
                    err = resp["error"]
                    raise RuntimeError(
                        f"CDP error ({err.get('code')}): "
                        f"{err.get('message', 'unknown')}"
                    )
                return resp.get("result", {})


# ── Browser Discovery ──────────────────────────────────────────────────


def fetch_page_targets() -> list[dict]:
    """Fetch the list of open page targets from Chrome's /json endpoint."""
    try:
        req = urllib.request.Request(
            f"{CHROME_DEBUG_URL}/json",
            headers={"User-Agent": "WhatsApp-MCP/1.0"},
        )
        with urllib.request.urlopen(req, timeout=5) as resp:
            data = json.loads(resp.read().decode())
        return [t for t in data if t.get("type") == "page"]
    except (urllib.error.URLError, OSError) as e:
        logger.debug(f"Cannot reach Chrome debug port: {e}")
        return []
    except Exception as e:
        logger.debug(f"Error fetching page targets: {e}")
        return []


def find_whatsapp_target(targets: list[dict]) -> dict | None:
    """Find the WhatsApp Web page target among open pages."""
    for t in targets:
        if WHATSAPP_URL_PATTERN in (t.get("url", "")):
            return t
    return None


# ── WhatsApp Client (convenience wrapper) ──────────────────────────────


class WhatsAppClient:
    """High-level client that connects to Chrome's CDP, finds the WhatsApp tab,
    and provides on-demand JavaScript evaluation for tool queries.

    Reuses a single CDP WebSocket connection across calls. Auto-reconnects
    if the connection drops.
    """

    def __init__(self):
        self._cdp: CDPClient | None = None
        self._lock = asyncio.Lock()

    @property
    def connected(self) -> bool:
        return self._cdp is not None and self._cdp.connected

    async def connect(self) -> str:
        """Connect to WhatsApp Web via CDP.

        Returns "ok" on success, or an error message string.
        """
        async with self._lock:
            if self._cdp and self._cdp.connected:
                return "ok"

            if self._cdp:
                try:
                    await self._cdp.disconnect()
                except Exception:
                    pass
                self._cdp = None

            targets = await asyncio.to_thread(fetch_page_targets)
            if not targets:
                return (
                    f"Chrome debug port ({CHROME_DEBUG_URL}) is not reachable. "
                    "Make sure Chrome is running with --remote-debugging-port=9222"
                )
            wa = find_whatsapp_target(targets)
            if not wa:
                return (
                    f"WhatsApp Web tab not found among "
                    f"{len(targets)} open tabs"
                )
            ws_url = wa.get("webSocketDebuggerUrl")
            if not ws_url:
                return "WhatsApp tab found but missing webSocketDebuggerUrl"

            self._cdp = CDPClient(
                ws_url=ws_url,
                page_title=wa.get("title", "WhatsApp"),
                page_url=wa.get("url", ""),
            )
            try:
                await self._cdp.connect()
                logger.info("WhatsApp client connected.")
                return "ok"
            except Exception as e:
                self._cdp = None
                return f"CDP connection failed: {e}"

    async def disconnect(self) -> None:
        """Close the CDP connection."""
        async with self._lock:
            if self._cdp:
                try:
                    await self._cdp.disconnect()
                except Exception:
                    pass
                self._cdp = None

    async def evaluate(self, expression: str, await_promise: bool = False) -> Any:
        """Run JavaScript in the WhatsApp page. Auto-connects if needed.

        ⚠️ SECURITY: This can execute ARBITRARY JavaScript in the WhatsApp page,
        including destructive actions (send message, delete, archive, etc.).
        Callers MUST only pass READ-ONLY expressions that do not mutate state.

        Returns the JS result value, or raises RuntimeError on failure.
        """
        status = await self.connect()
        if status != "ok":
            raise RuntimeError(status)
        assert self._cdp is not None
        return await self._cdp.evaluate(expression, await_promise)

    async def click_at(self, x: float, y: float) -> None:
        """Click at coordinates on the WhatsApp page. Auto-connects if needed."""
        status = await self.connect()
        if status != "ok":
            raise RuntimeError(status)
        assert self._cdp is not None
        await self._cdp.click_at(x, y)

    async def screenshot(self, format: str = "png", quality: int = 80) -> bytes:
        """Take a screenshot of the WhatsApp page. Auto-connects if needed."""
        status = await self.connect()
        if status != "ok":
            raise RuntimeError(status)
        assert self._cdp is not None
        return await self._cdp.capture_screenshot(format, quality)
