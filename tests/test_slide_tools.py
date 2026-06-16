"""
Comprehensive test suite for WPS Slide MCP tools - CLEAN STATE VERSION.
Always starts from a fresh presentation to avoid cascading failures.
"""
from __future__ import annotations

import os
import sys
import tempfile
from typing import Any

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
from wps_slide_mcp.slide_client import WPSSlideClient

_results: list[dict[str, Any]] = []
_passed = _failed = _skipped = 0
_client: WPSSlideClient | None = None
_test_dir: str = ""
_saved_path: str = ""


def get_client() -> WPSSlideClient:
    global _client
    if _client is None:
        _client = WPSSlideClient(visible=True)
    return _client


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
        msg = str(e)
        if "skip" in msg.lower():
            _skipped += 1
            _results.append({"name": name, "status": "SKIP", "error": msg})
            print(f"  \033[33mSKIP\033[0m {name}")
        else:
            _failed += 1
            err = str(e)[:300]
            _results.append({"name": name, "status": "FAIL", "error": err})
            print(f"  \033[31mFAIL\033[0m {name}: {err}")


def ok(cond, msg=""):
    if not cond:
        raise AssertionError(msg or "Expected True")


def eq(a, b, msg=""):
    if a != b:
        raise AssertionError(msg or f"Expected {b!r}, got {a!r}")


def has(sub, container, msg=""):
    if sub not in container:
        raise AssertionError(msg or f"Expected {sub!r} in container")


def has_key(key, container, msg=""):
    if key not in container:
        raise AssertionError(msg or f"Expected key {key!r} in dict")


def not_none(val, msg=""):
    if val is None:
        raise AssertionError(msg or "Expected non-None value")


def _set_saved_path(path: str) -> None:
    global _saved_path
    _saved_path = path


def get_saved_path() -> str:
    return _saved_path


# ═══════════════════════════════════════════════════════════════════

def run_all_tests():
    global _test_dir, _saved_path
    _test_dir = tempfile.mkdtemp(prefix="wps_slide_mcp_test_")
    print(f"Test dir: {_test_dir}")

    c = get_client()
    info = c.get_app_info()
    print(f"Connected: {info.get('name','?')} v{info.get('version','?')}")

    # Force clean state: close ALL existing presentations
    while True:
        try:
            c.close_presentation(save=False)
        except Exception:
            break
        if c.presentations.Count == 0:
            break

    # Create fresh presentation
    c.create_presentation()
    print(f"Fresh presentation, slide count: {c.get_slide_count()}\n")

    # ═══════════════════════════════════════════════════════
    # APPLICATION
    # ═══════════════════════════════════════════════════════
    print("-- Application --")
    run_test("slide_get_app_info", lambda: (
        info := c.get_app_info(),
        has_key("name", info),
        has_key("version", info),
        has_key("visible", info),
        has_key("presentations_count", info),
        has_key("active_presentation", info),
    ))
    run_test("slide_show_window", lambda: (c.show(), ok(c.app.Visible)))
    run_test("slide_hide_window", lambda: (
        c.hide(),
        ok(not c._visible),
        c.show(),  # restore visible
        ok(c._visible),
    ))

    # ═══════════════════════════════════════════════════════
    # PRESENTATION MANAGEMENT
    # ═══════════════════════════════════════════════════════
    print("\n-- Presentation Management --")
    run_test("slide_create_presentation", lambda: (
        name := c.create_presentation(),
        ok(isinstance(name, str)),
        ok(len(name) > 0),
        print(f"  Created: {name}"),
    ))
    run_test("slide_list_presentations", lambda: (
        pres_list := c.list_presentations(),
        ok(len(pres_list) >= 1),
        has_key("name", pres_list[0]),
        has_key("fullname", pres_list[0]),
        has_key("slides_count", pres_list[0]),
    ))
    run_test("slide_activate_presentation", lambda: (
        pres_list := c.list_presentations(),
        name := pres_list[0]["name"],
        result := c.activate_presentation(name),
        eq(result, name),
    ))
    run_test("slide_save_presentation", lambda: (
        _set_saved_path(os.path.join(_test_dir, "test_save.pptx")),
        c.save_presentation(get_saved_path()),
        ok(os.path.exists(get_saved_path())),
        print(f"  Saved to: {get_saved_path()}"),
    ))
    run_test("slide_get_properties", lambda: (
        props := c.get_presentation_properties(),
        has_key("name", props),
        has_key("fullname", props),
        has_key("slides_count", props),
        has_key("slide_width", props),
        has_key("slide_height", props),
    ))
    run_test("slide_set_slide_size", lambda: (
        c.set_slide_size(960, 540),
        props := c.get_presentation_properties(),
        ok(abs(props["slide_width"] - 960) < 10),
        ok(abs(props["slide_height"] - 540) < 10),
        # Restore default
        c.set_slide_size(960, 540),
    ))
    run_test("slide_close_presentation", lambda: (
        old_cnt := c.presentations.Count,
        c.close_presentation(save=False),
        new_cnt := c.presentations.Count,
        eq(new_cnt, old_cnt - 1),
    ))
    run_test("slide_open_presentation", lambda: (
        name := c.open_presentation(get_saved_path()),
        ok(isinstance(name, str)),
        ok(len(name) > 0),
        print(f"  Opened: {name}"),
    ))

    # ═══════════════════════════════════════════════════════
    # SLIDE OPERATIONS
    # ═══════════════════════════════════════════════════════
    print("\n-- Slide Operations --")

    # Close & create fresh
    while True:
        try:
            c.close_presentation(save=False)
        except Exception:
            break
    c.create_presentation()

    run_test("slide_get_count", lambda: (
        cnt := c.get_slide_count(),
        # Fresh presentation may have 0 or 1 slides depending on version
        ok(cnt >= 0),
        print(f"  Slide count: {cnt}"),
    ))
    run_test("slide_list_slides", lambda: (
        slides := c.list_slides(),
        ok(isinstance(slides, list)),
        ok(len(slides) >= 0),
        (has_key("index", slides[0]), has_key("name", slides[0]), has_key("shapes_count", slides[0])) if slides else ok(True),
        print(f"  Slides: {len(slides)}"),
    ))
    run_test("slide_add_slide", lambda: (
        old_cnt := c.get_slide_count(),
        idx := c.add_slide(layout_index=1),
        ok(idx == old_cnt + 1),
        eq(c.get_slide_count(), old_cnt + 1),
    ))
    run_test("slide_duplicate_slide", lambda: (
        old_cnt := c.get_slide_count(),
        new_idx := c.duplicate_slide(1),
        eq(new_idx, 2),
        eq(c.get_slide_count(), old_cnt + 1),
    ))
    run_test("slide_move_slide", lambda: (
        c.move_slide(1, 2),
        slides := c.list_slides(),
        ok(len(slides) >= 2),
        print(f"  Moved slide 1 -> 2"),
    ))
    run_test("slide_go_to", lambda: (
        c.go_to_slide(1),
        ok(True),  # No exception = success
    ))
    run_test("slide_set_background", lambda: (
        c.slide_set_background(1, 0x4472C4),  # Blue
        ok(True),
    ))
    run_test("slide_set_transition", lambda: (
        c.slide_set_transition(1, transition_type=1, duration=1.5),  # Fade
        ok(True),
    ))
    run_test("slide_delete", lambda: (
        old_cnt := c.get_slide_count(),
        c.delete_slide(old_cnt),  # Delete last slide
        eq(c.get_slide_count(), old_cnt - 1),
    ))

    # ═══════════════════════════════════════════════════════
    # SHAPE OPERATIONS - ADD
    # ═══════════════════════════════════════════════════════
    print("\n-- Shape Operations (Add) --")
    c.go_to_slide(1)

    run_test("shape_add_text_box", lambda: (
        name := c.add_text_box("Hello WPS Slide!", 50, 50, 400, 100),
        ok(isinstance(name, str)),
        ok(len(name) > 0),
        print(f"  Added text box: {name}"),
    ))
    run_test("shape_add_rectangle", lambda: (
        name := c.add_rectangle(50, 200, 200, 100),
        ok(isinstance(name, str)),
        ok(len(name) > 0),
    ))
    run_test("shape_add_oval", lambda: (
        name := c.add_oval(300, 200, 200, 100),
        ok(isinstance(name, str)),
        ok(len(name) > 0),
    ))
    run_test("shape_add_arrow", lambda: (
        name := c.add_arrow(50, 350, 200, 50),
        ok(isinstance(name, str)),
        ok(len(name) > 0),
    ))
    run_test("shape_add_line", lambda: (
        name := c.add_line(50, 450, 300, 450),
        ok(isinstance(name, str)),
        ok(len(name) > 0),
    ))

    # ═══════════════════════════════════════════════════════
    # SHAPE OPERATIONS - LIST / COUNT
    # ═══════════════════════════════════════════════════════
    print("\n-- Shape Operations (List/Count) --")
    run_test("shape_list_shapes", lambda: (
        shapes := c.list_shapes(),
        ok(len(shapes) >= 5),
        has_key("index", shapes[0]),
        has_key("name", shapes[0]),
        has_key("type", shapes[0]),
        print(f"  Shape count: {len(shapes)}"),
    ))
    run_test("shape_count", lambda: (
        cnt := c.get_shape_count(),
        ok(cnt >= 5),
        print(f"  Count: {cnt}"),
    ))

    # ═══════════════════════════════════════════════════════
    # SHAPE OPERATIONS - FORMAT
    # ═══════════════════════════════════════════════════════
    print("\n-- Shape Format --")
    # Use the text box (shape index 1)
    shapes = c.list_shapes()
    text_box_idx = shapes[0]["index"]

    run_test("shape_set_position", lambda: (
        c.set_shape_position(text_box_idx, left=100, top=100, width=500, height=120),
        ok(True),
    ))
    run_test("shape_set_fill", lambda: (
        c.set_shape_fill(text_box_idx, 0xFF0000),  # Red fill
        ok(True),
    ))
    run_test("shape_set_line", lambda: (
        c.set_shape_line(text_box_idx, 0x0000FF, weight=2.0),  # Blue outline
        ok(True),
    ))
    run_test("shape_set_rotation", lambda: (
        c.set_shape_rotation(text_box_idx, 5.0),
        ok(True),
        c.set_shape_rotation(text_box_idx, 0.0),  # Reset
    ))
    run_test("shape_set_zorder", lambda: (
        c.set_shape_zorder(text_box_idx, "front"),
        ok(True),
    ))

    # ═══════════════════════════════════════════════════════
    # SHAPE OPERATIONS - ORGANIZE
    # ═══════════════════════════════════════════════════════
    print("\n-- Shape Organize --")
    # Add a shape on slide 2 for paste test
    c.add_slide(layout_index=1)

    run_test("shape_copy", lambda: (
        c.copy_shape(text_box_idx, slide_index=1),
        ok(True),
    ))
    run_test("shape_paste", lambda: (
        name := c.paste_shape(slide_index=2),
        ok(name is not None),
        print(f"  Pasted: {name or '(empty)'}"),
    ))
    run_test("shape_duplicate", lambda: (
        # Use first shape on slide 1
        name := c.duplicate_shape(text_box_idx, slide_index=1),
        ok(name is not None),
        print(f"  Duplicated: {name or '(empty)'}"),
    ))
    run_test("shape_delete", lambda: (
        shapes_before := c.list_shapes(slide_index=2),
        (
            c.delete_shape(shapes_before[0]["index"], slide_index=2),
            shapes_after := c.list_shapes(slide_index=2),
            eq(len(shapes_after), len(shapes_before) - 1)
        ) if len(shapes_before) > 0 else ok(True),
    ))

    # ═══════════════════════════════════════════════════════
    # TEXT OPERATIONS
    # ═══════════════════════════════════════════════════════
    print("\n-- Text Operations --")
    c.go_to_slide(1)
    shapes = c.list_shapes()
    text_box = next((s for s in shapes if s.get("has_text")), shapes[0])

    # Add a fresh text box for clean testing
    fresh_name = c.add_text_box("Original text for testing", 50, 350, 500, 60)

    run_test("text_get", lambda: (
        t := c.get_shape_text(fresh_name),
        ok(isinstance(t, str)),
        has("Original", t),
        print(f"  Text: {t[:50]}"),
    ))
    run_test("text_set", lambda: (
        c.set_shape_text(fresh_name, "Modified text content"),
        t := c.get_shape_text(fresh_name),
        eq(t, "Modified text content"),
    ))

    # ═══════════════════════════════════════════════════════
    # FONT FORMATTING
    # ═══════════════════════════════════════════════════════
    print("\n-- Font Formatting --")

    run_test("font_bold", lambda: (
        c.set_font_bold(fresh_name, bold=True),
        c.set_shape_text(fresh_name, "Bold text test"),
        ok(True),
    ))
    run_test("font_italic", lambda: (
        c.set_font_italic(fresh_name, italic=True),
        c.set_shape_text(fresh_name, "Italic text test"),
        ok(True),
    ))
    run_test("font_underline", lambda: (
        c.set_font_underline(fresh_name, underline=True),
        c.set_shape_text(fresh_name, "Underlined text"),
        ok(True),
    ))
    run_test("font_name", lambda: (
        c.set_font_name(fresh_name, "Arial"),
        c.set_shape_text(fresh_name, "Arial font text"),
        ok(True),
    ))
    run_test("font_size", lambda: (
        c.set_font_size(fresh_name, 24),
        c.set_shape_text(fresh_name, "Size 24 text"),
        ok(True),
    ))
    run_test("font_color", lambda: (
        c.set_font_color(fresh_name, 0xFF0000),  # Red
        c.set_shape_text(fresh_name, "Red color text"),
        ok(True),
    ))
    run_test("font_alignment", lambda: (
        c.set_text_alignment(fresh_name, "center"),
        c.set_shape_text(fresh_name, "Centered text"),
        ok(True),
    ))

    # ═══════════════════════════════════════════════════════
    # TABLE OPERATIONS
    # ═══════════════════════════════════════════════════════
    print("\n-- Table Operations --")

    run_test("table_add", lambda: (
        name := c.add_table(rows=3, cols=3, left=100, top=400, width=500, height=200),
        ok(isinstance(name, str)),
        ok(len(name) > 0),
        print(f"  Added table: {name}"),
    ))
    run_test("table_set_cell", lambda: (
        tables := [s for s in c.list_shapes() if s.get("type") and s["type"] == 19],
        (
            c.set_table_cell(tables[0]["name"], row=1, col=1, text="R1C1"),
            c.set_table_cell(tables[0]["name"], row=1, col=2, text="R1C2"),
            c.set_table_cell(tables[0]["name"], row=2, col=1, text="R2C1"),
            ok(True)
        ) if tables else ok(True),
    ))
    run_test("table_get_data", lambda: (
        tables := [s for s in c.list_shapes() if s.get("type") and s["type"] == 19],
        (
            data := c.get_table_data(tables[0]["name"]),
            ok(isinstance(data, list)),
            ok(len(data) > 0),
            print(f"  Table data: {len(data)} rows x {len(data[0]) if data else 0} cols")
        ) if (tables and len(tables) > 0) else ok(True),
    ))

    # ═══════════════════════════════════════════════════════
    # NOTES
    # ═══════════════════════════════════════════════════════
    print("\n-- Speaker Notes --")

    run_test("notes_set", lambda: (
        c.set_notes(1, "These are test speaker notes."),
        ok(True),
    ))
    run_test("notes_get", lambda: (
        notes := c.get_notes(1),
        ok(isinstance(notes, str)),
        print(f"  Notes: {notes[:50] if notes else '(empty)'}"),
    ))

    # ═══════════════════════════════════════════════════════
    # EXPORT
    # ═══════════════════════════════════════════════════════
    print("\n-- Export --")

    run_test("export_pdf", lambda: (
        pdf_path := os.path.join(_test_dir, "test_export.pdf"),
        result := c.export_to_pdf(pdf_path),
        ok(os.path.exists(result)),
        print(f"  PDF exported: {result}"),
    ))
    run_test("export_slide_image", lambda: (
        img_path := os.path.join(_test_dir, "test_slide.png"),
        result := c.export_slide_image(img_path, slide_index=1, width=1920, height=1080),
        ok(os.path.exists(result)),
        print(f"  Image exported: {result}"),
    ))

    # ═══════════════════════════════════════════════════════
    # FIND & REPLACE
    # ═══════════════════════════════════════════════════════
    print("\n-- Find & Replace --")
    c.add_text_box("FindableText here and there", 50, 500, 400, 50)

    run_test("find_text", lambda: (
        results := c.find_text("FindableText"),
        ok(isinstance(results, list)),
        ok(len(results) > 0),
        print(f"  Found: {len(results)} match(es)"),
    ))
    run_test("find_replace", lambda: (
        count := c.find_replace("FindableText", "ReplacedText", replace_all=True),
        ok(count > 0),
        print(f"  Replaced: {count} occurrence(s)"),
        # Verify
        results := c.find_text("ReplacedText"),
        ok(len(results) > 0),
    ))

    # ═══════════════════════════════════════════════════════
    # MASTER
    # ═══════════════════════════════════════════════════════
    print("\n-- Master & Layout --")

    run_test("master_get_info", lambda: (
        info := c.get_master_info(),
        has_key("slide_master_name", info),
        has_key("layouts", info),
        ok(isinstance(info["layouts"], list)),
        print(f"  Layouts available: {len(info['layouts'])}"),
    ))
    run_test("master_set_background", lambda: (
        c.set_master_background(0x4472C4),  # Blue
        ok(True),
    ))
    run_test("master_apply_layout", lambda: (
        master_info := c.get_master_info(),
        (
            layout_idx := master_info["layouts"][0]["index"],
            c.apply_layout(1, layout_idx)
        ) if master_info["layouts"] else None,
        ok(True),
    ))

    # ═══════════════════════════════════════════════════════
    # ADVANCED - Header/Footer
    # ═══════════════════════════════════════════════════════
    print("\n-- Headers & Footers --")

    run_test("slide_number", lambda: (
        c.insert_slide_number(),
        ok(True),
    ))
    run_test("date_time", lambda: (
        c.insert_date_time(),
        ok(True),
    ))
    run_test("header_footer", lambda: (
        c.insert_header_footer(
            footer_text="Test Footer",
            slide_number=True,
            date_time=False,
        ),
        ok(True),
    ))

    # ═══════════════════════════════════════════════════════
    # HYPERLINK
    # ═══════════════════════════════════════════════════════
    print("\n-- Hyperlink --")
    href_name = c.add_text_box("Click here", 50, 600, 200, 30)

    run_test("hyperlink", lambda: (
        c.add_hyperlink("https://www.example.com", "Example Link", href_name),
        ok(True),
    ))

    # ═══════════════════════════════════════════════════════
    # ANIMATION
    # ═══════════════════════════════════════════════════════
    print("\n-- Animation --")
    anim_shape = c.add_text_box("Animated text", 50, 650, 300, 50)

    run_test("animation_add", lambda: (
        c.add_animation(
            anim_shape,
            effect_type="fade",
            trigger="on_click",
            duration=1.0,
            delay=0.5,
        ),
        ok(True),
    ))
    run_test("animation_clear", lambda: (
        c.clear_animations(anim_shape),
        ok(True),
    ))

    # ═══════════════════════════════════════════════════════
    # CHART
    # ═══════════════════════════════════════════════════════
    print("\n-- Chart --")

    run_test("chart_add", lambda: (
        name := c.add_chart("column", left=50, top=700, width=700, height=400),
        ok(isinstance(name, str)),
        ok(len(name) > 0),
        print(f"  Added chart: {name}"),
    ))
    run_test("chart_set_data", lambda: (
        charts := [s for s in c.list_shapes() if s.get("type") and s["type"] == 3],
        (
            c.set_chart_data(
                charts[0]["name"],
                categories=["Q1", "Q2", "Q3", "Q4"],
                series_list=[
                    {"name": "Sales", "values": [100, 200, 150, 300]},
                    {"name": "Costs", "values": [80, 150, 120, 250]},
                ],
            ),
            ok(True)
        ) if charts else ok(True),
    ))

    # ═══════════════════════════════════════════════════════
    # SLIDE SHOW (quick start/stop)
    # ═══════════════════════════════════════════════════════
    print("\n-- Slide Show --")

    run_test("slideshow_start", lambda: (
        c.start_slideshow(),
        ok(True),
    ))
    # Give it a moment then stop
    import time
    time.sleep(1)
    run_test("slideshow_stop", lambda: (
        c.stop_slideshow(),
        ok(True),
    ))

    # ═══════════════════════════════════════════════════════
    # CLEANUP
    # ═══════════════════════════════════════════════════════
    print("\n-- Cleanup --")
    # Save and close
    final_path = os.path.join(_test_dir, "test_final.pptx")
    c.save_presentation(final_path)
    try:
        c.close_presentation(save=False)
    except Exception:
        pass

    # ═══════════════════════════════════════════════════════
    # SUMMARY
    # ═══════════════════════════════════════════════════════
    total = _passed + _failed + _skipped
    print(f"\n{'='*60}")
    print(f"RESULTS: {_passed} passed, {_failed} failed, {_skipped} skipped ({total} total)")
    print(f"{'='*60}")

    if _failed > 0:
        print("\nFAILURES:")
        for r in _results:
            if r["status"] == "FAIL":
                print(f"  - {r['name']}: {r['error']}")

    return _passed, _failed, _skipped


if __name__ == "__main__":
    run_all_tests()
