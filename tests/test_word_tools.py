"""
Comprehensive test suite for WPS Word MCP tools - CLEAN STATE VERSION.
Always starts from a fresh document to avoid cascading failures.
"""
from __future__ import annotations

import os
import sys
import tempfile
from typing import Any

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))
from wps_word_mcp.word_client import WPSWordClient

_results: list[dict[str, Any]] = []
_passed = _failed = _skipped = 0
_client: WPSWordClient | None = None
_test_dir: str = ""
_saved_path: str = ""


def get_client() -> WPSWordClient:
    global _client
    if _client is None:
        _client = WPSWordClient(visible=True)
    return _client


def run_test(name: str, fn):
    global _passed, _failed, _skipped
    try:
        fn()
        _passed += 1
        _results.append({"name": name, "status": "PASS"})
        print(f"  PASS {name}")
    except AssertionError as e:
        _failed += 1
        _results.append({"name": name, "status": "FAIL", "error": str(e)})
        print(f"  FAIL {name}: {e}")
    except Exception as e:
        msg = str(e)
        if "skip" in msg.lower():
            _skipped += 1
            _results.append({"name": name, "status": "SKIP", "error": msg})
            print(f"  SKIP {name}")
        else:
            _failed += 1
            err = str(e)[:300]
            _results.append({"name": name, "status": "FAIL", "error": err})
            print(f"  FAIL {name}: {err}")


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
    _test_dir = tempfile.mkdtemp(prefix="wps_word_mcp_test_")
    print(f"Test dir: {_test_dir}")

    c = get_client()
    info = c.get_app_info()
    print(f"Connected: {info.get('name','?')} v{info.get('version','?')}")

    # Force clean state: close ALL existing documents
    while True:
        try:
            c.close_document(save=False)
        except Exception:
            break
        if c.documents.Count == 0:
            break

    # Create fresh document
    c.create_document()
    print(f"Fresh document, doc count: {c.documents.Count}\n")

    # ═══════════════════════════════════════════════════════
    # APPLICATION
    # ═══════════════════════════════════════════════════════
    print("-- Application --")
    run_test("word_get_app_info", lambda: (
        info := c.get_app_info(),
        has_key("name", info),
        has_key("version", info),
        has_key("documents_count", info),
        has_key("active_document", info),
    ))
    run_test("word_show_window", lambda: (c.show(), ok(c.app.Visible)))
    run_test("word_hide_window", lambda: (c.hide(), c.show()))  # restore visible

    # ═══════════════════════════════════════════════════════
    # DOCUMENT MANAGEMENT
    # ═══════════════════════════════════════════════════════
    print("\n-- Document Management --")
    run_test("word_create_document", lambda: (
        name := c.create_document(),
        ok(isinstance(name, str)),
        ok(len(name) > 0),
        print(f"  Created: {name}"),
    ))
    run_test("word_list_documents", lambda: (
        docs := c.list_documents(),
        ok(len(docs) >= 1),
        has_key("name", docs[0]),
        has_key("fullname", docs[0]),
    ))
    run_test("word_activate_document", lambda: (
        docs := c.list_documents(),
        name := docs[0]["name"],
        result := c.activate_document(name),
        eq(result, name),
    ))
    run_test("word_save_document", lambda: (
        c.set_text("Hello from WPS Word MCP test suite!"),
        _set_saved_path(os.path.join(_test_dir, "test_save.docx")),
        c.save_document(get_saved_path()),
        ok(os.path.exists(get_saved_path())),
        print(f"  Saved to: {get_saved_path()}"),
    ))
    run_test("word_close_document", lambda: (
        # Close without saving (we already saved)
        old_cnt := c.documents.Count,
        c.close_document(save=False),
        new_cnt := c.documents.Count,
        eq(new_cnt, old_cnt - 1),
    ))
    run_test("word_open_document", lambda: (
        name := c.open_document(get_saved_path()),
        ok(isinstance(name, str)),
        ok(len(name) > 0),
        print(f"  Opened: {name}"),
    ))

    # ═══════════════════════════════════════════════════════
    # TEXT OPERATIONS
    # ═══════════════════════════════════════════════════════
    print("\n-- Text Operations --")
    # Re-open saved doc for clean state
    while True:
        try:
            c.close_document(save=False)
        except Exception:
            break
    c.open_document(get_saved_path())

    run_test("word_get_text", lambda: (
        text := c.get_text(),
        ok(isinstance(text, str)),
        has("Hello from WPS Word", text),
    ))
    run_test("word_set_text", lambda: (
        c.set_text("Line 1\rLine 2\rLine 3"),
        text := c.get_text(),
        has("Line 1", text),
        has("Line 3", text),
    ))
    run_test("word_type_text", lambda: (
        # Type at the end
        c.set_text("Start."),
        sel := c.selection,
        sel.EndKey(Unit=6),  # wdStory=6 (go to end)
        c.type_text(" More text added."),
        text := c.get_text(),
        has("More text added", text),
    ))
    run_test("word_get_selected_text", lambda: (
        c.set_text("Select me please"),
        sel := c.selection,
        sel.WholeStory(),
        sel.MoveEnd(Unit=1, Count=-6),  # wdCharacter=1, leave "please" unselected
        text := c.get_selected_text(),
        ok("Select me" in text),
    ))
    run_test("word_insert_text_at_end", lambda: (
        c.set_text("Beginning"),
        c.insert_text_at_end(" and End"),
        text := c.get_text(),
        has("Beginning", text),
        has("End", text),
    ))
    run_test("word_insert_text_at_start", lambda: (
        c.set_text("Original"),
        c.insert_text_at_start("Prefix: "),
        text := c.get_text(),
        has("Prefix:", text),
        has("Original", text),
    ))

    # ═══════════════════════════════════════════════════════
    # PARAGRAPH OPERATIONS
    # ═══════════════════════════════════════════════════════
    print("\n-- Paragraph Operations --")
    run_test("word_get_paragraph_count", lambda: (
        c.set_text("Para1\rPara2\rPara3\r"),
        count := c.get_paragraph_count(),
        ok(count >= 3),
        print(f"  Paragraph count: {count}"),
    ))
    run_test("word_get_paragraph_text", lambda: (
        c.set_text("Alpha\rBravo\rCharlie\r"),
        eq(c.get_paragraph_text(1).strip(), "Alpha"),
        eq(c.get_paragraph_text(2).strip(), "Bravo"),
    ))
    run_test("word_set_paragraph_text", lambda: (
        c.set_text("One\rTwo\rThree\r"),
        c.set_paragraph_text(2, "Modified"),
        eq(c.get_paragraph_text(2).strip(), "Modified"),
    ))
    run_test("word_add_paragraph", lambda: (
        c.set_text("First\rSecond\r"),
        old_count := c.get_paragraph_count(),
        idx := c.add_paragraph("New Para"),
        new_count := c.get_paragraph_count(),
        ok(new_count >= old_count),
        text := c.get_paragraph_text(idx),
        eq(text.strip(), "New Para"),
    ))
    run_test("word_insert_paragraph_before", lambda: (
        c.set_text("A\rC\r"),
        c.insert_paragraph_before(2, "B"),
        eq(c.get_paragraph_text(2).strip(), "B"),
    ))
    run_test("word_delete_paragraph", lambda: (
        c.set_text("Keep\rDelete\rKeep2\r"),
        old_count := c.get_paragraph_count(),
        c.delete_paragraph(2),
        new_count := c.get_paragraph_count(),
        eq(new_count, old_count - 1),
    ))

    # ═══════════════════════════════════════════════════════
    # PARAGRAPH FORMATTING
    # ═══════════════════════════════════════════════════════
    print("\n-- Paragraph Formatting --")
    run_test("word_set_paragraph_alignment", lambda: (
        c.set_text("Centered text\r"),
        c.set_paragraph_alignment(1, "center"),
        para := c.active_document.Paragraphs(1),
        eq(para.Alignment, 1),  # wdAlignParagraphCenter=1
    ))
    run_test("word_set_paragraph_alignment_right", lambda: (
        c.set_text("Right text\r"),
        c.set_paragraph_alignment(1, "right"),
        eq(c.active_document.Paragraphs(1).Alignment, 2),  # wdAlignParagraphRight=2
    ))
    run_test("word_set_paragraph_alignment_justify", lambda: (
        c.set_text("Justified text\r"),
        c.set_paragraph_alignment(1, "justify"),
        eq(c.active_document.Paragraphs(1).Alignment, 3),  # wdAlignParagraphJustify=3
    ))
    run_test("word_set_paragraph_spacing", lambda: (
        c.set_text("Spaced text\r"),
        c.set_paragraph_spacing(1, before=12, after=6, line_spacing=1.5),
        pf := c.active_document.Paragraphs(1).Format,
        eq(pf.SpaceBefore, 12),
        eq(pf.SpaceAfter, 6),
        ok(abs(pf.LineSpacing - 1.5) < 0.01),
    ))

    # ═══════════════════════════════════════════════════════
    # FONT FORMATTING
    # ═══════════════════════════════════════════════════════
    print("\n-- Font Formatting --")
    run_test("word_set_font_bold", lambda: (
        c.set_text("Bold text"),
        c.set_font_bold(True, "content"),
        ok(c.active_document.Content.Font.Bold == -1),  # Word COM: -1 = True
    ))
    run_test("word_set_font_bold_off", lambda: (
        c.set_text("Not bold"),
        c.set_font_bold(False, "content"),
        ok(c.active_document.Content.Font.Bold == 0),  # Word COM: 0 = False
    ))
    run_test("word_set_font_italic", lambda: (
        c.set_text("Italic text"),
        c.set_font_italic(True, "content"),
        ok(c.active_document.Content.Font.Italic == -1),  # Word COM: -1 = True
    ))
    run_test("word_set_font_italic_off", lambda: (
        c.set_text("Not italic"),
        c.set_font_italic(False, "content"),
        ok(c.active_document.Content.Font.Italic == 0),  # Word COM: 0 = False
    ))
    run_test("word_set_font_underline", lambda: (
        c.set_text("Underlined"),
        c.set_font_underline(True, "content"),
        ok(c.active_document.Content.Font.Underline == 1),
    ))
    run_test("word_set_font_name", lambda: (
        c.set_text("Font test"),
        c.set_font_name("Arial", "content"),
        eq(c.active_document.Content.Font.Name, "Arial"),
    ))
    run_test("word_set_font_size", lambda: (
        c.set_text("Sized text"),
        c.set_font_size(18, "content"),
        eq(c.active_document.Content.Font.Size, 18),
    ))
    run_test("word_set_font_color", lambda: (
        c.set_text("Colored text"),
        c.set_font_color(0x0000FF, "content"),  # Red in BGR
        ok(c.active_document.Content.Font.Color == 0x0000FF),
    ))

    # ═══════════════════════════════════════════════════════
    # FIND / REPLACE
    # ═══════════════════════════════════════════════════════
    print("\n-- Find / Replace --")
    run_test("word_find_text_found", lambda: (
        c.set_text("The quick brown fox jumps over the lazy dog"),
        result := c.find_text("brown"),
        not_none(result),
        eq(result["found"], True),
        has("brown", result["text"]),
    ))
    run_test("word_find_text_not_found", lambda: (
        c.set_text("Nothing to find here"),
        result := c.find_text("xyzzy_missing_text"),
        eq(result, None),
    ))
    run_test("word_find_text_match_case", lambda: (
        c.set_text("Cat cat CAT"),
        result := c.find_text("Cat", match_case=True),
        not_none(result),
        eq(result["text"], "Cat"),
    ))
    run_test("word_find_replace_all", lambda: (
        c.set_text("apple banana apple cherry apple"),
        count := c.find_replace("apple", "orange", replace_all=True),
        text := c.get_text(),
        ok("apple" not in text),
        ok("orange" in text),
    ))
    run_test("word_find_replace_single", lambda: (
        c.set_text("One Two Two Three"),
        count := c.find_replace("Two", "X", replace_all=False),
        text := c.get_text(),
        # Only first occurrence should be replaced
        ok("X" in text),
        ok("Two" in text),  # second one still there
    ))
    run_test("word_find_replace_match_case", lambda: (
        c.set_text("Word word WORD"),
        c.find_replace("Word", "X", match_case=True, replace_all=True),
        text := c.get_text(),
        ok("X" in text),
        ok("word" in text),  # lowercase preserved
    ))

    # ═══════════════════════════════════════════════════════
    # TABLE OPERATIONS
    # ═══════════════════════════════════════════════════════
    print("\n-- Table Operations --")
    run_test("word_add_table", lambda: (
        c.set_text("Before table."),
        idx := c.add_table(3, 2, "Name\tAge\nJohn\t30\nJane\t25"),
        ok(idx >= 1),
        count := c.get_table_count(),
        ok(count >= 1),
        print(f"  Table index: {idx}, count: {count}"),
    ))
    run_test("word_get_table_count", lambda: (
        count := c.get_table_count(),
        ok(count >= 1),
    ))
    run_test("word_get_table_data", lambda: (
        data := c.get_table_data(1),
        ok(len(data) == 3),   # 3 rows
        ok(len(data[0]) == 2),  # 2 columns
        eq(data[0][0], "Name"),
        eq(data[0][1], "Age"),
        eq(data[1][0], "John"),
    ))
    run_test("word_set_cell_text", lambda: (
        c.set_cell_text(1, 2, 1, "Jack"),
        data := c.get_table_data(1),
        eq(data[1][0], "Jack"),
    ))
    run_test("word_add_table_row", lambda: (
        c.add_table_row(1),
        data := c.get_table_data(1),
        ok(len(data) == 4),  # was 3, now 4
    ))
    run_test("word_add_table_column", lambda: (
        c.add_table_column(1),
        data := c.get_table_data(1),
        ok(len(data[0]) == 3),  # was 2, now 3
    ))
    run_test("word_add_table_empty", lambda: (
        c.set_text("Before empty table."),
        idx := c.add_table(2, 3),
        data := c.get_table_data(idx),
        print(f"  Empty table idx={idx}, rows={len(data)}, cols={len(data[0]) if data else 0}"),
        ok(len(data) >= 2),  # Word may add extra row
        ok(len(data[0]) >= 3),  # Word may add extra col marker
    ))
    run_test("word_delete_table", lambda: (
        old_count := c.get_table_count(),
        c.delete_table(old_count),  # delete the last one
        ok(c.get_table_count() == old_count - 1),
    ))

    # ═══════════════════════════════════════════════════════
    # PAGE SETUP
    # ═══════════════════════════════════════════════════════
    print("\n-- Page Setup --")
    run_test("word_set_page_orientation_portrait", lambda: (
        c.set_page_orientation("portrait"),
        eq(c.active_document.PageSetup.Orientation, 0),
    ))
    run_test("word_set_page_orientation_landscape", lambda: (
        c.set_page_orientation("landscape"),
        eq(c.active_document.PageSetup.Orientation, 1),
        c.set_page_orientation("portrait"),  # restore
    ))
    run_test("word_set_page_margins", lambda: (
        c.set_page_margins(left=72, right=72, top=72, bottom=72),
        ps := c.active_document.PageSetup,
        eq(ps.LeftMargin, 72),
        eq(ps.TopMargin, 72),
        eq(ps.BottomMargin, 72),
    ))
    run_test("word_set_page_size", lambda: (
        c.set_page_size(width=595, height=842),
        ps := c.active_document.PageSetup,
        ok(abs(ps.PageWidth - 595) < 2),  # small tolerance
        ok(abs(ps.PageHeight - 842) < 2),
    ))

    # ═══════════════════════════════════════════════════════
    # HEADER / FOOTER
    # ═══════════════════════════════════════════════════════
    print("\n-- Header / Footer --")
    run_test("word_set_header", lambda: (
        c.set_text("Document body text"),
        c.add_header("Test Header"),
        section := c.active_document.Sections(1),
        header := section.Headers(1),
        ok("Test Header" in header.Range.Text),
    ))
    run_test("word_set_footer", lambda: (
        c.add_footer("Test Footer"),
        section := c.active_document.Sections(1),
        footer := section.Footers(1),
        ok("Test Footer" in footer.Range.Text),
    ))

    # ═══════════════════════════════════════════════════════
    # INSERT ELEMENTS
    # ═══════════════════════════════════════════════════════
    print("\n-- Insert Elements --")
    run_test("word_insert_page_break", lambda: (
        orig_count := c.get_paragraph_count(),
        c.set_text("Page 1 text"),
        c.selection.EndKey(Unit=6),
        c.insert_page_break(),
        c.type_text("Page 2 text"),
        new_count := c.get_paragraph_count(),
        ok(new_count >= orig_count + 1),
    ))
    run_test("word_add_section_break", lambda: (
        orig_sections := c.get_section_count(),
        c.add_section_break(),
        new_sections := c.get_section_count(),
        ok(new_sections > orig_sections),
    ))

    # ═══════════════════════════════════════════════════════
    # EXPORT TO PDF
    # ═══════════════════════════════════════════════════════
    print("\n-- Export --")
    run_test("word_export_to_pdf", lambda: (
        c.set_text("PDF Export Test Document\nLine 2\nLine 3"),
        pdf_path := os.path.join(_test_dir, "test_export.pdf"),
        result := c.export_to_pdf(pdf_path),
        ok(os.path.exists(result)),
        file_size := os.path.getsize(result),
        ok(file_size > 0),
        print(f"  PDF: {result} ({file_size} bytes)"),
    ))

    # ═══════════════════════════════════════════════════════
    # VIEW / ZOOM
    # ═══════════════════════════════════════════════════════
    print("\n-- View --")
    run_test("word_set_zoom", lambda: (
        c.set_zoom(120),
        zoom := c.app.ActiveWindow.View.Zoom.Percentage,
        eq(zoom, 120),
    ))

    # ═══════════════════════════════════════════════════════
    # CLEANUP
    # ═══════════════════════════════════════════════════════
    print("\n-- Cleanup --")
    try:
        c.close_document(save=False)
    except Exception:
        pass


# ═══════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    print("=" * 60)
    print("  WPS Word MCP — Full Tools Test Suite")
    print("=" * 60 + "\n")

    run_all_tests()

    total = _passed + _failed + _skipped
    print(f"\n{'='*60}")
    print(f"  Results: {_passed} passed, {_failed} failed, {_skipped} skipped  ({total} total)")
    print(f"{'='*60}")

    if _failed > 0:
        print("\nFailed tests:")
        for r in _results:
            if r["status"] == "FAIL":
                print(f"  - {r['name']}: {r['error']}")

    if _skipped > 0:
        print("\nSkipped tests:")
        for r in _results:
            if r["status"] == "SKIP":
                print(f"  - {r['name']}: {r['error']}")

    input("\nPress Enter to exit...")
