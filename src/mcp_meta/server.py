"""
MCP Server for meta tools: web search and web fetch.

Provides:
- web-search: Search the web via configurable providers (Brave, etc.)
- web-fetch:  Fetch and extract text content from a URL

Environment variables:
    SEARCH_PROVIDER  - "brave" (default)
    BRAVE_API_KEY    - Brave Search API key (required)
"""

from __future__ import annotations

import asyncio
import json
import logging
import os
import re
import urllib.error
import urllib.request
from typing import Any
from urllib.parse import urlencode

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)

# ── MCP Server Setup ───────────────────────────────────────────────────

server = Server("mcp-meta")

# ── Provider Configuration ─────────────────────────────────────────────

# Which search provider to use (default: brave)
SEARCH_PROVIDER = os.environ.get("SEARCH_PROVIDER", "brave").strip().lower()

# Brave Search API config
BRAVE_SEARCH_URL = "https://api.search.brave.com/res/v1/web/search"


# ── Tool Definitions ───────────────────────────────────────────────────

TOOLS: list[Tool] = [
    Tool(
        name="web-search",
        description=(
            "Search the web. Currently supports Brave Search API. "
            "Set BRAVE_API_KEY env var (get a key at https://brave.com/search/api/). "
            "Control provider via SEARCH_PROVIDER env var."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "query": {
                    "type": "string",
                    "description": "Search query string",
                },
                "count": {
                    "type": "integer",
                    "description": "Number of results to return (default: 10, max: 20)",
                },
                "offset": {
                    "type": "integer",
                    "description": "Pagination offset for results (default: 0)",
                },
                "country": {
                    "type": "string",
                    "description": "Two-letter country code (e.g. US, DE, FR)",
                },
                "search_lang": {
                    "type": "string",
                    "description": "Search language code (e.g. en, de, fr)",
                },
                "freshness": {
                    "type": "string",
                    "description": "Filter by freshness: pd (past day), pw (past week), pm (past month), py (past year)",
                },
            },
            "required": ["query"],
        },
    ),
    Tool(
        name="web-fetch",
        description="Fetch and extract readable text content from a URL.",
        inputSchema={
            "type": "object",
            "properties": {
                "url": {
                    "type": "string",
                    "description": "The URL to fetch content from",
                },
                "max_length": {
                    "type": "integer",
                    "description": "Maximum characters of content to return (default: 10000)",
                },
                "raw": {
                    "type": "boolean",
                    "description": "Return raw HTML content instead of extracted text (default: false)",
                },
            },
            "required": ["url"],
        },
    ),
]


# ═══════════════════════════════════════════════════════════════════════
#  Search Providers
# ═══════════════════════════════════════════════════════════════════════


def _search_brave(
    query: str,
    count: int = 10,
    offset: int = 0,
    country: str = "",
    search_lang: str = "",
    freshness: str = "",
) -> dict:
    """Call the Brave Search API."""
    api_key = os.environ.get("BRAVE_API_KEY", "")
    if not api_key:
        return {
            "error": (
                "BRAVE_API_KEY environment variable is not set. "
                "Get a free API key at https://brave.com/search/api/"
            ),
            "provider": "brave",
        }

    params: dict[str, Any] = {
        "q": query,
        "count": max(1, min(count, 20)),
        "offset": max(0, offset),
    }
    if country:
        params["country"] = country
    if search_lang:
        params["search_lang"] = search_lang
    if freshness:
        params["freshness"] = freshness

    url = f"{BRAVE_SEARCH_URL}?{urlencode(params)}"

    req = urllib.request.Request(url)
    req.add_header("X-Subscription-Token", api_key)
    req.add_header("Accept", "application/json")

    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        return {"error": f"Brave API HTTP {e.code}: {body[:500]}", "provider": "brave"}
    except Exception as e:
        return {"error": f"Brave API error: {e}", "provider": "brave"}

    # Extract web results
    results: list[dict] = []
    web_results = data.get("web", {}).get("results", [])
    for r in web_results:
        results.append({
            "title": r.get("title", ""),
            "url": r.get("url", ""),
            "description": r.get("description", ""),
            "age": r.get("age", ""),
            "language": r.get("language", ""),
        })

    return {
        "provider": "brave",
        "query": query,
        "total_results": len(results),
        "results": results,
    }


# ── Provider Router ────────────────────────────────────────────────────

def web_search(
    query: str,
    count: int = 10,
    offset: int = 0,
    country: str = "",
    search_lang: str = "",
    freshness: str = "",
) -> dict:
    """Route web search to the configured provider."""
    if SEARCH_PROVIDER == "brave":
        return _search_brave(
            query=query,
            count=count,
            offset=offset,
            country=country,
            search_lang=search_lang,
            freshness=freshness,
        )
    else:
        supported = ["brave"]
        return {
            "error": (
                f"Unsupported search provider: '{SEARCH_PROVIDER}'. "
                f"Supported providers: {', '.join(supported)}. "
                f"Set SEARCH_PROVIDER env var to one of these values."
            ),
            "provider": SEARCH_PROVIDER,
        }


# ═══════════════════════════════════════════════════════════════════════
#  Web Fetch
# ═══════════════════════════════════════════════════════════════════════

_HTML_TAG_RE = re.compile(r"<[^>]+>", re.DOTALL)
_SCRIPT_STYLE_RE = re.compile(
    r"<(script|style|noscript|iframe|svg|canvas|nav|footer|header|aside)[^>]*>.*?</\1>",
    re.DOTALL | re.IGNORECASE,
)
_WHITESPACE_RE = re.compile(r"\s+")


def _extract_text(html: str) -> str:
    """Extract readable text from HTML by stripping tags and normalizing whitespace."""
    # Remove script, style, noscript, iframe, svg, canvas, nav, footer, header, aside
    text = _SCRIPT_STYLE_RE.sub(" ", html)
    # Remove remaining HTML tags
    text = _HTML_TAG_RE.sub(" ", text)
    # Decode common HTML entities
    text = text.replace("&amp;", "&")
    text = text.replace("&lt;", "<")
    text = text.replace("&gt;", ">")
    text = text.replace("&quot;", '"')
    text = text.replace("&#39;", "'")
    text = text.replace("&nbsp;", " ")
    # Normalize whitespace
    text = _WHITESPACE_RE.sub(" ", text).strip()
    return text


def web_fetch(url: str, max_length: int = 10000, raw: bool = False) -> dict:
    """Fetch a URL and return its content."""
    req = urllib.request.Request(url)
    req.add_header(
        "User-Agent",
        "Mozilla/5.0 (compatible; MCP-Meta/1.0; +https://github.com)",
    )
    req.add_header(
        "Accept",
        "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    )
    req.add_header("Accept-Language", "en-US,en;q=0.5")

    try:
        with urllib.request.urlopen(req, timeout=20) as resp:
            content_type = resp.headers.get("Content-Type", "")
            raw_bytes = resp.read()

            # Try to decode with the declared charset
            charset = "utf-8"
            if "charset=" in content_type:
                charset = content_type.split("charset=")[-1].split(";")[0].strip()

            try:
                html = raw_bytes.decode(charset, errors="replace")
            except (LookupError, UnicodeDecodeError):
                html = raw_bytes.decode("utf-8", errors="replace")

    except urllib.error.HTTPError as e:
        return {"error": f"HTTP {e.code}: {e.reason}", "url": url}
    except urllib.error.URLError as e:
        return {"error": f"URL Error: {e.reason}", "url": url}
    except Exception as e:
        return {"error": str(e), "url": url}

    if raw:
        content = html[:max_length]
    else:
        content = _extract_text(html)[:max_length]

    return {
        "url": url,
        "content_type": content_type,
        "content": content,
        "length": len(content),
    }


# ═══════════════════════════════════════════════════════════════════════
#  MCP Tool Handlers
# ═══════════════════════════════════════════════════════════════════════


@server.list_tools()
async def list_tools() -> list[Tool]:
    """Return the list of available tools."""
    return TOOLS


@server.call_tool()
async def call_tool(name: str, arguments: dict[str, Any]) -> list[TextContent]:
    """Handle tool calls."""
    try:
        loop = asyncio.get_running_loop()

        if name == "web-search":
            result = await loop.run_in_executor(
                None,
                lambda: web_search(
                    query=arguments["query"],
                    count=arguments.get("count", 10),
                    offset=arguments.get("offset", 0),
                    country=arguments.get("country", ""),
                    search_lang=arguments.get("search_lang", ""),
                    freshness=arguments.get("freshness", ""),
                ),
            )

        elif name == "web-fetch":
            result = await loop.run_in_executor(
                None,
                lambda: web_fetch(
                    url=arguments["url"],
                    max_length=arguments.get("max_length", 10000),
                    raw=arguments.get("raw", False),
                ),
            )

        else:
            result = {"error": f"Unknown tool: {name}"}

        return [TextContent(type="text", text=json.dumps(result, ensure_ascii=False))]

    except Exception as e:
        logger.exception(f"Error executing tool '{name}'")
        return [
            TextContent(
                type="text",
                text=json.dumps(
                    {"error": str(e), "tool": name},
                    ensure_ascii=False,
                ),
            )
        ]


# ── Server Entry Point ─────────────────────────────────────────────────


async def run_server() -> None:
    """Run the MCP server via stdio transport."""
    async with stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            server.create_initialization_options(),
        )


def main() -> None:
    """Entry point for the mcp-meta command."""
    asyncio.run(run_server())


if __name__ == "__main__":
    main()
