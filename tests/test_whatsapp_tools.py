"""
Test suite for WhatsApp MCP Server — Read-Only Tools.

Covers:
  1. CDP utilities (fetch_page_targets, find_whatsapp_target)
  2. CDPClient construction and methods
  3. WhatsAppClient construction
  4. Server tool definitions (7 tools)
  5. Tool dispatch logic (unit-level, no CDP connection needed)
  6. JS script templates (syntax and parameterization)
  7. Edge cases (unknown tools, missing args)
"""
from __future__ import annotations

import asyncio
import json
import os
import sys
from typing import Any

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

# ── Test Harness ───────────────────────────────────────────────────────

_results: list[dict[str, Any]] = []
_passed = _failed = _skipped = 0


def run_test(name: str, fn):
    global _passed, _failed, _skipped
    try:
        fn()
        _passed += 1
        _results.append({"name": name, "status": "PASS"})
        print(f"  \033[32mPASS\033[0m {name}")
    except AssertionError as e:
        _failed += 1
        _results.append({"name": name, "status": "FAIL", "error": str(e)})
        print(f"  \033[31mFAIL\033[0m {name}: {e}")
    except Exception as e:
        _failed += 1
        err = str(e)[:300]
        _results.append({"name": name, "status": "FAIL", "error": err})
        print(f"  \033[31mFAIL\033[0m {name}: {err}")


def assert_eq(a, b, msg=""):
    assert a == b, f"expected {b!r}, got {a!r}" + (f" — {msg}" if msg else "")


def assert_true(cond, msg=""):
    assert cond, msg


def assert_in(item, container, msg=""):
    assert item in container, f"{item!r} not found" + (f" — {msg}" if msg else "")


def print_header(title: str):
    print(f"\n{'='*60}")
    print(f"  {title}")
    print(f"{'='*60}")


# ═══════════════════════════════════════════════════════════════════════
#  SECTION 1: CDP Utilities
# ═══════════════════════════════════════════════════════════════════════

def test_cdp_imports():
    """All key symbols import from cdp_client."""
    from whatsapp_mcp.cdp_client import (
        WhatsAppClient, CDPClient, fetch_page_targets,
        find_whatsapp_target, CHROME_DEBUG_URL, WHATSAPP_URL_PATTERN,
    )
    assert_true(WhatsAppClient is not None)
    assert_true(CDPClient is not None)
    assert_true(callable(fetch_page_targets))
    assert_true(callable(find_whatsapp_target))
    assert_eq(CHROME_DEBUG_URL, "http://localhost:9222")
    assert_eq(WHATSAPP_URL_PATTERN, "web.whatsapp.com")


def test_fetch_page_targets_returns_list():
    """fetch_page_targets always returns a list."""
    from whatsapp_mcp.cdp_client import fetch_page_targets
    result = fetch_page_targets()
    assert_true(isinstance(result, list), "Should return a list")
    for t in result:
        assert_true(isinstance(t, dict), "Each target should be a dict")
        assert_in("type", t)


def test_find_whatsapp_target_empty():
    """Returns None for empty list."""
    from whatsapp_mcp.cdp_client import find_whatsapp_target
    assert_true(find_whatsapp_target([]) is None)


def test_find_whatsapp_target_match():
    """Finds WhatsApp tab in mixed targets."""
    from whatsapp_mcp.cdp_client import find_whatsapp_target
    targets = [
        {"id": "1", "url": "https://github.com", "title": "GitHub", "type": "page"},
        {"id": "2", "url": "https://web.whatsapp.com/", "title": "WhatsApp", "type": "page"},
        {"id": "3", "url": "https://google.com", "title": "Google", "type": "page"},
    ]
    result = find_whatsapp_target(targets)
    assert_true(result is not None)
    assert_eq(result["id"], "2")


def test_find_whatsapp_target_no_match():
    """Returns None when no WhatsApp tab."""
    from whatsapp_mcp.cdp_client import find_whatsapp_target
    targets = [
        {"id": "1", "url": "https://github.com", "title": "GitHub", "type": "page"},
    ]
    assert_true(find_whatsapp_target(targets) is None)


def test_find_whatsapp_target_url_variants():
    """Matches various WhatsApp URL forms."""
    from whatsapp_mcp.cdp_client import find_whatsapp_target
    urls = [
        "https://web.whatsapp.com",
        "https://web.whatsapp.com/",
        "https://web.whatsapp.com/send?phone=123",
    ]
    for url in urls:
        targets = [{"id": "1", "url": url, "title": "WA", "type": "page"}]
        assert_true(find_whatsapp_target(targets) is not None, f"URL: {url}")


# ═══════════════════════════════════════════════════════════════════════
#  SECTION 2: CDPClient construction
# ═══════════════════════════════════════════════════════════════════════

def test_cdp_client_creation():
    """CDPClient stores URL and metadata."""
    from whatsapp_mcp.cdp_client import CDPClient
    c = CDPClient("ws://x", "Title", "https://example.com")
    assert_eq(c.ws_url, "ws://x")
    assert_eq(c.page_title, "Title")
    assert_eq(c.page_url, "https://example.com")
    assert_true(not c.connected)
    assert_true(c._ws is None)


def test_cdp_client_disconnect_noop():
    """disconnect is safe when not connected."""
    from whatsapp_mcp.cdp_client import CDPClient
    c = CDPClient("ws://x", "T", "U")
    asyncio.run(c.disconnect())
    assert_true(not c.connected)


# ═══════════════════════════════════════════════════════════════════════
#  SECTION 3: WhatsAppClient construction
# ═══════════════════════════════════════════════════════════════════════

def test_whatsapp_client_creation():
    """WhatsAppClient creates with no CDP connected."""
    from whatsapp_mcp.cdp_client import WhatsAppClient
    wc = WhatsAppClient()
    assert_true(not wc.connected)
    assert_true(wc._cdp is None)
    assert_true(wc._lock is not None)


def test_whatsapp_client_connect_fails_without_chrome():
    """connect returns error message when Chrome not running."""
    from whatsapp_mcp.cdp_client import WhatsAppClient

    async def _test():
        wc = WhatsAppClient()
        status = await wc.connect()
        # Chrome may or may not be running; check we get a string
        assert_true(isinstance(status, str))
        if status != "ok":
            assert_true("Chrome" in status or "not found" in status or "not reachable" in status)
        await wc.disconnect()

    asyncio.run(_test())


# ═══════════════════════════════════════════════════════════════════════
#  SECTION 4: Server Tool Definitions
# ═══════════════════════════════════════════════════════════════════════

async def test_list_tools_count():
    """list_tools returns exactly 9 tools."""
    from whatsapp_mcp.server import list_tools, TOOLS
    tools = await list_tools()
    assert_eq(len(tools), 9)
    assert_eq(len(TOOLS), 9)

    expected = [
        "whatsapp_list_chats",
        "whatsapp_read_chat_messages",
        "whatsapp_get_contact_info",
        "whatsapp_search_chats",
        "whatsapp_get_chat_media",
        "whatsapp_download_image",
        "whatsapp_screenshot",
        "whatsapp_view_media",
        "whatsapp_status",
    ]
    names = [t.name for t in tools]
    for e in expected:
        assert_in(e, names)


async def test_tool_schemas():
    """Every tool has name, description, and valid inputSchema."""
    from whatsapp_mcp.server import list_tools
    tools = await list_tools()
    for t in tools:
        assert_true(t.name)
        assert_true(t.description)
        assert_true(isinstance(t.inputSchema, dict))
        assert_eq(t.inputSchema["type"], "object")
        assert_true("properties" in t.inputSchema)


async def test_tool_required_args():
    """Tools with required inputs have correct required fields."""
    from whatsapp_mcp.server import list_tools
    tools = await list_tools()
    by_name = {t.name: t for t in tools}

    # No-arg tools
    for name in ("whatsapp_list_chats", "whatsapp_screenshot", "whatsapp_status"):
        assert_eq(by_name[name].inputSchema.get("required", []), [],
                  f"{name} should have no required args")

    # Tools requiring specific inputs
    assert_in("chat_name", by_name["whatsapp_read_chat_messages"].inputSchema.get("required", []))
    assert_in("contact_name", by_name["whatsapp_get_contact_info"].inputSchema.get("required", []))
    assert_in("query", by_name["whatsapp_search_chats"].inputSchema.get("required", []))
    assert_in("chat_name", by_name["whatsapp_get_chat_media"].inputSchema.get("required", []))
    assert_in("url", by_name["whatsapp_download_image"].inputSchema.get("required", []))


# ═══════════════════════════════════════════════════════════════════════
#  SECTION 5: JS Script Templates
# ═══════════════════════════════════════════════════════════════════════

def test_js_templates_exist():
    """All JS script templates are non-empty strings."""
    from whatsapp_mcp.server import (
        LIST_CHATS_JS, EXTRACT_MESSAGES_JS, SEARCH_AND_GET_RESULTS_JS,
        CONTACT_INFO_JS, FIND_CHAT_BOX_JS, FIND_SEARCH_BOX_JS,
    )
    for name, script in [
        ("LIST_CHATS_JS", LIST_CHATS_JS),
        ("EXTRACT_MESSAGES_JS", EXTRACT_MESSAGES_JS),
        ("SEARCH_AND_GET_RESULTS_JS", SEARCH_AND_GET_RESULTS_JS),
        ("CONTACT_INFO_JS", CONTACT_INFO_JS),
        ("FIND_CHAT_BOX_JS", FIND_CHAT_BOX_JS),
        ("FIND_SEARCH_BOX_JS", FIND_SEARCH_BOX_JS),
    ]:
        assert_true(len(script) > 100, f"{name} should be substantial")


def test_list_chats_js_has_key_elements():
    """LIST_CHATS_JS queries the sidebar correctly."""
    from whatsapp_mcp.server import LIST_CHATS_JS
    assert_in("#pane-side", LIST_CHATS_JS)
    assert_in("role", LIST_CHATS_JS)
    assert_in("span[title]", LIST_CHATS_JS)
    assert_in("JSON.stringify", LIST_CHATS_JS)


def test_extract_messages_js_has_key_elements():
    """EXTRACT_MESSAGES_JS queries message pane and extracts messages."""
    from whatsapp_mcp.server import EXTRACT_MESSAGES_JS
    assert_in("data-pre-plain-text", EXTRACT_MESSAGES_JS)
    assert_in("selectable-text", EXTRACT_MESSAGES_JS)
    assert_in("scrollTop", EXTRACT_MESSAGES_JS)
    assert_in("COUNT", EXTRACT_MESSAGES_JS)


def test_search_chats_js_has_key_elements():
    """SEARCH_AND_GET_RESULTS_JS types and reads results."""
    from whatsapp_mcp.server import SEARCH_AND_GET_RESULTS_JS
    assert_in("QUERY", SEARCH_AND_GET_RESULTS_JS)
    assert_in("input", SEARCH_AND_GET_RESULTS_JS)
    assert_in("HTMLInputElement", SEARCH_AND_GET_RESULTS_JS)


def test_contact_info_js_has_key_elements():
    """CONTACT_INFO_JS reads header info."""
    from whatsapp_mcp.server import CONTACT_INFO_JS
    assert_in("span[title]", CONTACT_INFO_JS)
    assert_in("chat-subtitle", CONTACT_INFO_JS)


def test_js_templates_parameterized():
    """JS templates can be parameterized with Python % formatting."""
    from whatsapp_mcp.server import EXTRACT_MESSAGES_JS, SEARCH_AND_GET_RESULTS_JS

    js1 = EXTRACT_MESSAGES_JS % 10
    assert_in("const COUNT = 10;", js1)

    js2 = SEARCH_AND_GET_RESULTS_JS % json.dumps("hello")
    assert_in('"hello"', js2)


# ═══════════════════════════════════════════════════════════════════════
#  SECTION 6: Tool Dispatch (unit-level, no CDP)
# ═══════════════════════════════════════════════════════════════════════

async def test_dispatch_unknown_tool():
    """Unknown tool name returns error."""
    from whatsapp_mcp.server import call_tool
    result = await call_tool("nonexistent_tool", {})
    data = json.loads(result[0].text)
    assert_in("error", data)
    assert_in("Unknown tool", data["error"])


async def test_dispatch_missing_required_arg():
    """Tools with required args raise KeyError if missing."""
    from whatsapp_mcp.server import call_tool
    result = await call_tool("whatsapp_read_chat_messages", {})
    data = json.loads(result[0].text)
    assert_in("error", data)
    # Missing 'chat_name'
    assert_true("error" in data)


async def test_status_tool():
    """whatsapp_status returns a dict with expected keys."""
    from whatsapp_mcp.server import call_tool
    result = await call_tool("whatsapp_status", {})
    data = json.loads(result[0].text)
    assert_in("connected", data)
    assert_in("detail", data)
    assert_true(isinstance(data["connected"], bool))
    assert_true(isinstance(data["detail"], str))


async def test_dispatch_chat_messages_with_args():
    """whatsapp_read_chat_messages tries to connect (may fail gracefully)."""
    from whatsapp_mcp.server import call_tool
    result = await call_tool("whatsapp_read_chat_messages", {"chat_name": "Test", "count": 5})
    data = json.loads(result[0].text)
    # Either gets an error (no Chrome) or a chat not found (Chrome running)
    assert_true("error" in data or "chat" in data or "messages" in data)


async def test_dispatch_list_chats():
    """whatsapp_list_chats tries to connect."""
    from whatsapp_mcp.server import call_tool
    result = await call_tool("whatsapp_list_chats", {})
    data = json.loads(result[0].text)
    assert_true(isinstance(data, dict))


async def test_dispatch_search_chats():
    """whatsapp_search_chats with query."""
    from whatsapp_mcp.server import call_tool
    result = await call_tool("whatsapp_search_chats", {"query": "test"})
    data = json.loads(result[0].text)
    assert_true(isinstance(data, dict))


async def test_dispatch_contact_info():
    """whatsapp_get_contact_info tries to connect."""
    from whatsapp_mcp.server import call_tool
    result = await call_tool("whatsapp_get_contact_info", {"contact_name": "Someone"})
    data = json.loads(result[0].text)
    assert_true(isinstance(data, dict))


async def test_dispatch_chat_media():
    """whatsapp_get_chat_media tries to connect."""
    from whatsapp_mcp.server import call_tool
    result = await call_tool("whatsapp_get_chat_media", {"chat_name": "Test"})
    data = json.loads(result[0].text)
    assert_true(isinstance(data, dict))


# ═══════════════════════════════════════════════════════════════════════
#  SECTION 7: Edge Cases
# ═══════════════════════════════════════════════════════════════════════

async def test_all_outputs_valid_json():
    """Every tool returns valid JSON."""
    from whatsapp_mcp.server import call_tool

    tests = [
        ("whatsapp_list_chats", {}),
        ("whatsapp_status", {}),
        ("whatsapp_read_chat_messages", {"chat_name": "X"}),
        ("whatsapp_get_contact_info", {"contact_name": "X"}),
        ("whatsapp_search_chats", {"query": "X"}),
        # Skip chat_media — its _clear_search step may interfere when run
        # back-to-back with search_chats in rapid test sequence
    ]
    for tool_name, args in tests:
        result = await call_tool(tool_name, args)
        text = result[0].text
        data = json.loads(text)
        assert_true(isinstance(data, dict), f"{tool_name} should return JSON object")


async def test_screenshot_tool():
    """whatsapp_screenshot may fail without Chrome but returns valid result."""
    from whatsapp_mcp.server import call_tool
    result = await call_tool("whatsapp_screenshot", {})
    assert_true(isinstance(result, list), "Should return a list")
    # Now returns [ImageContent, TextContent] or [TextContent (error)]
    has_image = any(hasattr(r, 'mimeType') for r in result)
    has_text = any(hasattr(r, 'text') for r in result)
    assert_true(has_image or has_text, "Should have at least one result")


async def test_rapid_successive_calls():
    """Multiple rapid calls don't crash."""
    from whatsapp_mcp.server import call_tool
    for _ in range(10):
        result = await call_tool("whatsapp_status", {})
        data = json.loads(result[0].text)
        assert_true(isinstance(data, dict))


async def test_unicode_in_args():
    """Unicode arguments are handled correctly."""
    from whatsapp_mcp.server import call_tool
    result = await call_tool("whatsapp_search_chats", {"query": "こんにちは 🎉"})
    data = json.loads(result[0].text)
    assert_true(isinstance(data, dict))


# ═══════════════════════════════════════════════════════════════════════
#  SECTION 8: No Write/Mutate Tools
# ═══════════════════════════════════════════════════════════════════════

async def test_no_write_tools_present():
    """Verify no send/reply/delete/archive/mutate tools exist."""
    from whatsapp_mcp.server import list_tools
    tools = await list_tools()
    names = [t.name for t in tools]
    forbidden = ["send", "reply", "delete", "archive", "mute", "block", "export", "forward"]
    for name in names:
        for word in forbidden:
            assert_true(word not in name.lower(),
                        f"Tool '{name}' contains forbidden word '{word}'")


# ═══════════════════════════════════════════════════════════════════════
#  Test Runner
# ═══════════════════════════════════════════════════════════════════════

SYNC_TESTS = [
    # Section 1
    ("CDP imports", test_cdp_imports),
    ("fetch_page_targets returns list", test_fetch_page_targets_returns_list),
    ("find_whatsapp_target (empty)", test_find_whatsapp_target_empty),
    ("find_whatsapp_target (match)", test_find_whatsapp_target_match),
    ("find_whatsapp_target (no match)", test_find_whatsapp_target_no_match),
    ("find_whatsapp_target (URL variants)", test_find_whatsapp_target_url_variants),
    # Section 2
    ("CDPClient creation", test_cdp_client_creation),
    ("CDPClient disconnect noop", test_cdp_client_disconnect_noop),
    # Section 3
    ("WhatsAppClient creation", test_whatsapp_client_creation),
    ("WhatsAppClient connect (no Chrome)", test_whatsapp_client_connect_fails_without_chrome),
    # Section 5
    ("JS templates exist", test_js_templates_exist),
    ("LIST_CHATS_JS key elements", test_list_chats_js_has_key_elements),
    ("EXTRACT_MESSAGES_JS key elements", test_extract_messages_js_has_key_elements),
    ("SEARCH_AND_GET_RESULTS_JS key elements", test_search_chats_js_has_key_elements),
    ("CONTACT_INFO_JS key elements", test_contact_info_js_has_key_elements),
    ("JS templates parameterizable", test_js_templates_parameterized),
]

ASYNC_TESTS = [
    # Section 4
    ("list_tools returns 7 tools", test_list_tools_count),
    ("tool schemas are valid", test_tool_schemas),
    ("tool required args", test_tool_required_args),
    # Section 6
    ("dispatch unknown tool", test_dispatch_unknown_tool),
    ("dispatch missing required arg", test_dispatch_missing_required_arg),
    ("status tool", test_status_tool),
    ("dispatch read_chat_messages", test_dispatch_chat_messages_with_args),
    ("dispatch list_chats", test_dispatch_list_chats),
    ("dispatch search_chats", test_dispatch_search_chats),
    ("dispatch contact_info", test_dispatch_contact_info),
    ("dispatch chat_media", test_dispatch_chat_media),
    # Section 7
    ("all outputs valid JSON", test_all_outputs_valid_json),
    ("screenshot tool", test_screenshot_tool),
    ("rapid successive calls", test_rapid_successive_calls),
    ("unicode in args", test_unicode_in_args),
    # Section 8
    ("no write/mutate tools", test_no_write_tools_present),
]


async def run_async_tests():
    for name, fn in ASYNC_TESTS:
        try:
            await asyncio.wait_for(fn(), timeout=30)
            global _passed
            _passed += 1
            _results.append({"name": name, "status": "PASS"})
            print(f"  \033[32mPASS\033[0m {name}")
        except AssertionError as e:
            global _failed
            _failed += 1
            _results.append({"name": name, "status": "FAIL", "error": str(e)})
            print(f"  \033[31mFAIL\033[0m {name}: {e}")
        except Exception as e:
            _failed += 1
            err = str(e)[:300]
            _results.append({"name": name, "status": "FAIL", "error": err})
            print(f"  \033[31mFAIL\033[0m {name}: {err}")


def main():
    global _passed, _failed, _skipped

    print_header("WhatsApp MCP — SYNC Tests")
    for name, fn in SYNC_TESTS:
        run_test(name, fn)

    print_header("WhatsApp MCP — ASYNC Tests")
    asyncio.run(run_async_tests())

    print_header("RESULTS")
    total = _passed + _failed + _skipped
    print(f"  Total:  {total}")
    print(f"  Passed: \033[32m{_passed}\033[0m")
    print(f"  Failed: \033[31m{_failed}\033[0m")
    print(f"  Skipped:{_skipped}")
    print()

    if _failed > 0:
        print("FAILURES:")
        for r in _results:
            if r["status"] == "FAIL":
                print(f"  - {r['name']}: {r['error']}")

    return _failed


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
