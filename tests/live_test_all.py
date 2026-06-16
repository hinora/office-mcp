"""
Live test of all WhatsApp MCP tools against running Chrome.
"""
import asyncio, json, sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from whatsapp_mcp.server import (
    _eval_json, _eval, _click_at, _open_chat,
    _read_messages, _get_contact, _search_chats, _get_media,
    LIST_CHATS_JS, FIND_CHAT_BOX_JS, FIND_SEARCH_BOX_JS,
    CONTACT_INFO_JS, EXTRACT_MESSAGES_JS, SEARCH_AND_GET_RESULTS_JS,
    _download_image,
)
from whatsapp_mcp.cdp_client import WhatsAppClient


async def test_all():
    results = {}
    print("=" * 60)
    print("  WhatsApp MCP — ALL TOOLS LIVE TEST")
    print("=" * 60)

    # ── 1. whatsapp_status ────────────────────────────────────────────
    print("\n[1] whatsapp_status")
    try:
        wc = WhatsAppClient()
        status = await wc.connect()
        connected = (status == "ok")
        results["status"] = {"connected": connected, "detail": status}
        print(f"    connected={connected}, detail='{status[:80]}'")
        if not connected:
            print("    SKIPPING remaining tests — Chrome not reachable")
            return
    except Exception as e:
        print(f"    FAIL: {e}")
        return

    # ── 2. whatsapp_list_chats ────────────────────────────────────────
    print("\n[2] whatsapp_list_chats")
    try:
        data = await _eval_json(LIST_CHATS_JS)
        if "error" in data:
            print(f"    FAIL: {data['error']}")
            results["list_chats"] = data
        else:
            chats = data.get("chats", [])
            print(f"    OK: {data.get('total', 0)} total, {len(chats)} returned")
            for c in chats[:5]:
                print(f"        {c['name'][:35]:35s} | unread={c.get('unread',''):4s} | {c.get('time','')}")
            results["list_chats"] = {"count": len(chats), "total": data.get("total")}
    except Exception as e:
        print(f"    FAIL: {e}")
        results["list_chats"] = {"error": str(e)}

    # ── 3. whatsapp_read_chat_messages ────────────────────────────────
    print("\n[3] whatsapp_read_chat_messages")
    first_chat = None
    try:
        list_data = await _eval_json(LIST_CHATS_JS)
        if list_data.get("chats"):
            first_chat = list_data["chats"][0]["name"]
        if not first_chat:
            print("    SKIP: no chats found")
            results["read_messages"] = {"error": "no chats"}
        else:
            print(f"    Opening chat: '{first_chat}'")
            data = await _read_messages(first_chat, 10)
            if "error" in data:
                print(f"    FAIL: {data['error']}")
                results["read_messages"] = data
            else:
                msgs = data.get("messages", [])
                print(f"    OK: {len(msgs)} messages")
                for m in msgs[:5]:
                    print(f"        [{m.get('sender','?')[:20]}]: {m.get('text','')[:60]}")
                results["read_messages"] = {"count": len(msgs), "total_in_view": data.get("total_in_view")}
    except Exception as e:
        print(f"    FAIL: {e}")
        results["read_messages"] = {"error": str(e)}

    # ── 4. whatsapp_get_contact_info ──────────────────────────────────
    print("\n[4] whatsapp_get_contact_info")
    try:
        if not first_chat:
            print("    SKIP: no chat available")
            results["contact_info"] = {"error": "no chat"}
        else:
            print(f"    Getting info for: '{first_chat}'")
            data = await _get_contact(first_chat)
            if "error" in data:
                print(f"    FAIL: {data['error']}")
                results["contact_info"] = data
            else:
                print(f"    OK: name='{data.get('name','')[:40]}', status='{data.get('status','')[:40]}', phone='{data.get('phone','')}'")
                results["contact_info"] = data
    except Exception as e:
        print(f"    FAIL: {e}")
        results["contact_info"] = {"error": str(e)}

    # ── 5. whatsapp_search_chats ──────────────────────────────────────
    print("\n[5] whatsapp_search_chats")
    try:
        # First go back to chat list by clicking an empty area or pressing escape
        await _eval("""
        (function(){
            const backBtn = document.querySelector('[data-testid="back"], [aria-label="Back"]');
            if (backBtn) backBtn.click();
        })()
        """)
        await asyncio.sleep(1)

        print("    Searching for: 'a'")
        data = await _search_chats("a")
        if "error" in data:
            print(f"    FAIL: {data['error']}")
            results["search_chats"] = data
        else:
            res = data.get("results", [])
            print(f"    OK: {len(res)} results")
            for r in res[:5]:
                print(f"        {r.get('name','')[:40]}")
            results["search_chats"] = {"count": len(res)}
    except Exception as e:
        print(f"    FAIL: {e}")
        results["search_chats"] = {"error": str(e)}

    # ── 6. whatsapp_get_chat_media ────────────────────────────────────
    print("\n[6] whatsapp_get_chat_media")
    first_blob = None
    try:
        if not first_chat:
            print("    SKIP: no chat")
            results["chat_media"] = {"error": "no chat"}
        else:
            print(f"    Getting media from: '{first_chat}'")
            data = await _get_media(first_chat, 10)
            if "error" in data:
                print(f"    FAIL: {data['error']}")
                results["chat_media"] = data
            else:
                media = data.get("media", [])
                print(f"    OK: {len(media)} items")
                for m in media[:5]:
                    saved = m.get("saved_path", m.get("src", "")[:60])
                    print(f"        {m.get('type','?')}: {saved}")
                results["chat_media"] = {"count": len(media)}
                # Remember first blob URL for download test
                for m in media:
                    if m.get("src", "").startswith("blob:") and m.get("saved_path"):
                        first_blob = m["src"]
                        break
    except Exception as e:
        print(f"    FAIL: {e}")
        results["chat_media"] = {"error": str(e)}

    # ── 6b. whatsapp_download_image ──────────────────────────────────
    print("\n[6b] whatsapp_download_image")
    try:
        if not first_blob:
            print("    SKIP: no blob URL to test")
            results["download_image"] = {"error": "no blob url"}
        else:
            print(f"    Downloading: {first_blob[:60]}...")
            data = await _download_image(first_blob)
            if "error" in data:
                print(f"    FAIL: {data['error']}")
                results["download_image"] = data
            else:
                path = data.get("saved_path", "?")
                size = data.get("size_bytes", 0)
                has_url = bool(data.get("data_url"))
                local_file = path
                print(f"    OK: {size} bytes → {path}")
                results["download_image"] = {"saved_path": path, "size_bytes": size}
    except Exception as e:
        print(f"    FAIL: {e}")
        results["download_image"] = {"error": str(e)}

    # ── 7. whatsapp_screenshot ────────────────────────────────────────
    print("\n[7] whatsapp_screenshot")
    try:
        img_bytes = await wc.screenshot("png")
        import tempfile
        tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False, prefix="whatsapp_test_")
        tmp.write(img_bytes)
        tmp.close()
        print(f"    OK: {len(img_bytes)} bytes → {tmp.name}")
        results["screenshot"] = {"size_bytes": len(img_bytes), "path": tmp.name}
    except Exception as e:
        print(f"    FAIL: {e}")
        results["screenshot"] = {"error": str(e)}

    # ── Summary ───────────────────────────────────────────────────────
    print("\n" + "=" * 60)
    print("  RESULTS SUMMARY")
    print("=" * 60)
    passed = 0
    failed = 0
    for tool_name, data in results.items():
        if "error" in str(data) and not isinstance(data, dict):
            continue
        if isinstance(data, dict) and "error" in data:
            print(f"  \033[31mFAIL\033[0m {tool_name}: {data['error']}")
            failed += 1
        else:
            print(f"  \033[32mPASS\033[0m {tool_name}")
            passed += 1
    print(f"\n  {passed} passed, {failed} failed, {8 - passed - failed} skipped")
    print()

    await wc.disconnect()


if __name__ == "__main__":
    asyncio.run(test_all())
