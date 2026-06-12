"""
WPS Office Word COM client.

Provides a high-level Python interface to automate WPS Office Word
using the COM automation API (WPS.Application).
"""

from __future__ import annotations

import os
import re
from typing import Any

import pythoncom
import win32com.client


class WPSWordClient:
    """Client for interacting with WPS Office Word via COM automation."""

    # WPS Office Word ProgIDs to try (in order of preference)
    _PROG_IDS = [
        "WPS.Application",   # WPS Word
        "KWPS.Application",  # Older WPS Word
        "Word.Application",  # Fallback to MS Word if WPS not available
    ]

    def __init__(self, visible: bool = False) -> None:
        """
        Initialize the WPS Word client.

        Args:
            visible: If True, make the WPS Word window visible.
        """
        self._app: Any = None
        self._visible = visible
        self._connect()

    def _connect(self) -> None:
        """Connect to a running WPS Word instance or create a new one."""
        pythoncom.CoInitialize()

        existing_instance = False

        for prog_id in self._PROG_IDS:
            # First try getting an already running instance
            try:
                self._app = win32com.client.GetActiveObject(prog_id)
                existing_instance = True
            except Exception:
                pass

            # If no running instance, create a new one
            if self._app is None:
                try:
                    self._app = win32com.client.Dispatch(prog_id)
                except Exception:
                    continue

            # Verify the connection works by accessing Documents
            if self._app is not None:
                try:
                    _ = self._app.Documents.Count
                    break  # Connection verified
                except Exception:
                    self._app = None
                    existing_instance = False
                    continue

        if self._app is None:
            raise RuntimeError(
                "Could not connect to WPS Office Word. "
                "Please ensure WPS Office is installed."
            )

        if not existing_instance:
            self._app.Visible = self._visible

    @property
    def app(self) -> Any:
        """Get the underlying COM application object."""
        if self._app is None:
            self._connect()
        return self._app

    @property
    def documents(self) -> Any:
        """Get the Documents collection."""
        return self.app.Documents

    @property
    def active_document(self) -> Any | None:
        """Get the active document, or None if no document is open."""
        try:
            return self.app.ActiveDocument
        except Exception:
            return None

    @property
    def selection(self) -> Any | None:
        """Get the current selection, or None if unavailable."""
        try:
            return self.app.Selection
        except Exception:
            return None

    # ── Application ──────────────────────────────────────────────────

    def get_app_info(self) -> dict[str, str]:
        """Get information about the WPS Word application."""
        info: dict[str, str] = {}
        try:
            info["name"] = str(self.app.Name)
        except Exception:
            info["name"] = "WPS Word (unknown)"
        try:
            info["version"] = str(self.app.Version)
        except Exception:
            info["version"] = "unknown"
        info["visible"] = str(self._visible)
        try:
            info["documents_count"] = str(self.documents.Count)
        except Exception:
            info["documents_count"] = "unknown"
        try:
            doc = self.active_document
            info["active_document"] = doc.Name if doc else "none"
        except Exception:
            info["active_document"] = "none"
        return info

    def show(self) -> None:
        """Make the WPS Word window visible."""
        self.app.Visible = True
        self._visible = True

    def hide(self) -> None:
        """Hide the WPS Word window."""
        self.app.Visible = False
        self._visible = False

    def quit_app(self) -> None:
        """Quit the WPS Word application."""
        try:
            self.app.Quit(SaveChanges=0)  # wdDoNotSaveChanges
        except Exception:
            pass
        self._app = None

    # ── Document Operations ──────────────────────────────────────────

    def create_document(self) -> str:
        """
        Create a new blank document.

        Returns:
            The name of the new document.
        """
        doc = self.documents.Add()
        return doc.Name

    def open_document(self, filepath: str) -> str:
        """
        Open an existing document from disk.

        Args:
            filepath: Full path to the document file (.docx, .doc, .wps, etc.).

        Returns:
            The name of the opened document.
        """
        full_path = os.path.abspath(filepath)
        if not os.path.exists(full_path):
            raise FileNotFoundError(f"Document not found: {full_path}")
        doc = self.documents.Open(full_path)
        return doc.Name

    def save_document(self, filepath: str | None = None) -> str:
        """
        Save the active document.

        Args:
            filepath: Path to save to. If None, saves to current location.

        Returns:
            The full path where the document was saved.
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document to save.")

        if filepath:
            full_path = os.path.abspath(filepath)
            doc.SaveAs(full_path)
        else:
            doc.Save()
            full_path = doc.FullName

        return full_path

    def close_document(self, save: bool = True) -> None:
        """
        Close the active document.

        Args:
            save: If True, save changes before closing.
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document to close.")
        save_val = -1 if save else 0  # wdSaveChanges=-1, wdDoNotSaveChanges=0
        doc.Close(SaveChanges=save_val)

    def list_documents(self) -> list[dict[str, str]]:
        """
        List all open documents.

        Returns:
            List of dicts with 'name', 'fullname', and 'path'.
        """
        result = []
        for i in range(1, self.documents.Count + 1):
            doc = self.documents.Item(i)
            result.append({
                "name": doc.Name,
                "fullname": doc.FullName or "",
                "path": doc.Path or "",
            })
        return result

    def activate_document(self, name: str) -> str:
        """
        Activate a document by name.

        Args:
            name: The name of the document to activate.

        Returns:
            The name of the activated document.
        """
        doc = self.documents(name)
        doc.Activate()
        return doc.Name

    # ── Text Operations ──────────────────────────────────────────────

    def get_text(self) -> str:
        """
        Get all text from the active document.

        Returns:
            The full text content of the document.
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        return doc.Content.Text

    def get_selected_text(self) -> str:
        """
        Get the currently selected text.

        Returns:
            The selected text, or empty string.
        """
        sel = self.selection
        if sel is None:
            return ""
        return sel.Text

    def type_text(self, text: str) -> None:
        """
        Type text at the current cursor/selection position.

        Args:
            text: The text to insert.
        """
        sel = self.selection
        if sel is None:
            raise RuntimeError("No selection available.")
        sel.TypeText(text)

    def insert_text_at_end(self, text: str) -> None:
        """
        Append text at the end of the document.

        Args:
            text: The text to append.
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        rng = doc.Content
        rng.Collapse(Direction=0)  # wdCollapseEnd=0
        rng.Text = text

    def insert_text_at_start(self, text: str) -> None:
        """
        Insert text at the beginning of the document.

        Args:
            text: The text to insert.
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        rng = doc.Content
        rng.Collapse(Direction=1)  # wdCollapseStart=1
        rng.Text = text

    def set_text(self, text: str) -> None:
        """
        Replace all text in the document with given text.

        Args:
            text: The full text for the document.
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        doc.Content.Text = text

    # ── Paragraph Operations ─────────────────────────────────────────

    def add_paragraph(self, text: str = "") -> int:
        """
        Add a new paragraph at the end of the document.

        Args:
            text: Optional text for the paragraph.

        Returns:
            The 1-based index of the new paragraph.
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")

        # Move to end of document
        rng = doc.Content
        rng.Collapse(Direction=0)  # wdCollapseEnd=0

        # Insert a paragraph break, which creates a new empty paragraph
        rng.InsertParagraphAfter()

        # The new paragraph is now the last one; set its text if provided
        if text:
            new_idx = doc.Paragraphs.Count
            self.set_paragraph_text(new_idx, text)

        return doc.Paragraphs.Count

    def get_paragraph_count(self) -> int:
        """Get the number of paragraphs in the document."""
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        return doc.Paragraphs.Count

    def get_paragraph_text(self, index: int) -> str:
        """
        Get the text of a specific paragraph (1-based index).

        Args:
            index: 1-based paragraph index.

        Returns:
            The paragraph text (without trailing newline/paragraph mark).
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        para = doc.Paragraphs(index)
        # Get range without the paragraph mark
        rng = para.Range
        # Move end back by one character to exclude the paragraph mark (\r)
        rng.MoveEnd(Unit=1, Count=-1)  # wdCharacter=1
        return rng.Text

    def set_paragraph_text(self, index: int, text: str) -> None:
        """
        Set the text of a specific paragraph.

        Args:
            index: 1-based paragraph index.
            text: New text for the paragraph.
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        para = doc.Paragraphs(index)
        rng = para.Range
        # Exclude the paragraph mark from the range so we don't eat into next paragraph
        rng.MoveEnd(Unit=1, Count=-1)  # wdCharacter=1
        rng.Text = text

    def insert_paragraph_before(self, index: int, text: str = "") -> None:
        """
        Insert a paragraph before the given paragraph index.

        Args:
            index: 1-based paragraph index to insert before.
            text: Optional text for the new paragraph.
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        para = doc.Paragraphs(index)
        para.Range.InsertParagraphBefore()
        if text:
            # After InsertParagraphBefore, the new paragraph is at the same index
            # Use same technique to avoid eating next paragraph
            new_para = doc.Paragraphs(index)
            rng = new_para.Range
            rng.MoveEnd(Unit=1, Count=-1)
            rng.Text = text

    def delete_paragraph(self, index: int) -> None:
        """
        Delete a paragraph by index.

        Args:
            index: 1-based paragraph index.
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        doc.Paragraphs(index).Range.Delete()

    # ── Paragraph Formatting ─────────────────────────────────────────

    def set_paragraph_alignment(self, index: int, alignment: str) -> None:
        """
        Set paragraph alignment.

        Args:
            index: 1-based paragraph index.
            alignment: 'left', 'center', 'right', 'justify'.
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        xl_align = {
            "left": 0,      # wdAlignParagraphLeft
            "center": 1,    # wdAlignParagraphCenter
            "right": 2,     # wdAlignParagraphRight
            "justify": 3,   # wdAlignParagraphJustify
        }
        align_val = xl_align.get(alignment.lower(), 0)
        doc.Paragraphs(index).Alignment = align_val

    def set_paragraph_spacing(
        self,
        index: int,
        before: float | None = None,
        after: float | None = None,
        line_spacing: float | None = None,
    ) -> None:
        """
        Set paragraph spacing.

        Args:
            index: 1-based paragraph index.
            before: Space before in points.
            after: Space after in points.
            line_spacing: Line spacing (e.g., 1.0, 1.5, 2.0).
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        pf = doc.Paragraphs(index).Format
        if before is not None:
            pf.SpaceBefore = before
        if after is not None:
            pf.SpaceAfter = after
        if line_spacing is not None:
            pf.LineSpacingRule = 0  # wdLineSpaceMultiple
            pf.LineSpacing = line_spacing

    # ── Font Formatting ──────────────────────────────────────────────

    def _get_range(self, doc: Any, range_or_selection: str = "content") -> Any:
        """
        Get a range object based on the mode string.

        Args:
            doc: The document object.
            range_or_selection: 'content' (all text), 'selection', or a range
                                string like "start=0,end=100".

        Returns:
            The appropriate Range object.
        """
        if range_or_selection == "selection":
            return self.selection.Range if self.selection else doc.Content
        elif range_or_selection.startswith("start="):
            # Parse range string: "start=0,end=100"
            parts = range_or_selection.split(",")
            start = int(parts[0].split("=")[1])
            end = int(parts[1].split("=")[1]) if len(parts) > 1 else doc.Content.End
            rng = doc.Range(start, end)
            return rng
        else:
            return doc.Content

    def set_font_bold(
        self,
        bold: bool = True,
        range_spec: str = "selection",
    ) -> None:
        """
        Set font bold on a range.

        Args:
            bold: True to make bold, False to remove.
            range_spec: 'content', 'selection', or "start=X,end=Y".
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        rng = self._get_range(doc, range_spec)
        rng.Font.Bold = bold

    def set_font_italic(
        self,
        italic: bool = True,
        range_spec: str = "selection",
    ) -> None:
        """Set font italic on a range."""
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        rng = self._get_range(doc, range_spec)
        rng.Font.Italic = italic

    def set_font_underline(
        self,
        underline: bool = True,
        range_spec: str = "selection",
    ) -> None:
        """
        Set font underline on a range.

        Args:
            underline: True to underline, False to remove.
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        rng = self._get_range(doc, range_spec)
        rng.Font.Underline = 1 if underline else 0  # wdUnderlineSingle=1

    def set_font_name(
        self,
        font_name: str,
        range_spec: str = "selection",
    ) -> None:
        """Set font name on a range."""
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        rng = self._get_range(doc, range_spec)
        rng.Font.Name = font_name

    def set_font_size(
        self,
        size: float,
        range_spec: str = "selection",
    ) -> None:
        """Set font size in points on a range."""
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        rng = self._get_range(doc, range_spec)
        rng.Font.Size = size

    def set_font_color(
        self,
        color: int,
        range_spec: str = "selection",
    ) -> None:
        """
        Set font color on a range.

        Args:
            color: RGB color value (e.g., 0xFF0000 for red).
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        rng = self._get_range(doc, range_spec)
        rng.Font.Color = color

    # ── Find / Replace ───────────────────────────────────────────────

    def find_text(
        self,
        search_text: str,
        match_case: bool = False,
        match_whole_word: bool = False,
    ) -> dict[str, Any] | None:
        """
        Find text in the document. Selects the first match.

        Returns:
            Dict with match info, or None if not found.
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")

        find = self.selection.Find if self.selection else doc.Content.Find
        find.ClearFormatting()
        find.Text = search_text
        find.MatchCase = match_case
        find.MatchWholeWord = match_whole_word

        found = find.Execute()
        if not found:
            return None

        sel = self.selection
        return {
            "found": True,
            "text": sel.Text if sel else search_text,
            "start": sel.Start if sel else -1,
            "end": sel.End if sel else -1,
        }

    def find_replace(
        self,
        find_text: str,
        replace_text: str,
        match_case: bool = False,
        match_whole_word: bool = False,
        replace_all: bool = True,
    ) -> int:
        """
        Find and replace text in the document.

        Args:
            find_text: Text to search for.
            replace_text: Replacement text.
            match_case: If True, match case.
            match_whole_word: If True, match whole words only (not yet implemented
                for whole-word via string replace; uses simple substring matching).
            replace_all: If True, replace all occurrences. If False, replace only first.

        Returns:
            Number of replacements made.
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")

        old_text = doc.Content.Text

        if match_case:
            if replace_all:
                new_text = old_text.replace(find_text, replace_text)
                count = old_text.count(find_text)
            else:
                new_text = old_text.replace(find_text, replace_text, 1)
                count = 1 if find_text in old_text else 0
        else:
            # Case-insensitive: use regex
            pattern = re.escape(find_text)
            flags = 0 if match_case else re.IGNORECASE
            if replace_all:
                new_text, count = re.subn(pattern, replace_text, old_text, flags=flags)
            else:
                new_text, count = re.subn(pattern, replace_text, old_text, count=1, flags=flags)

        if count > 0:
            doc.Content.Text = new_text

        return count

    # ── Table Operations ─────────────────────────────────────────────

    def add_table(
        self,
        rows: int,
        cols: int,
        text: str = "",
    ) -> int:
        """
        Add a table at the end of the document.

        Args:
            rows: Number of rows.
            cols: Number of columns.
            text: Optional text to populate (tab-separated for cells).

        Returns:
            The 1-based index of the new table.
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")

        rng = doc.Content
        rng.Collapse(Direction=0)  # wdCollapseEnd=0

        table = doc.Tables.Add(rng, rows, cols)

        # Populate if text provided
        if text:
            lines = text.split("\n")
            for row_idx, line in enumerate(lines):
                if row_idx >= rows:
                    break
                cells = line.split("\t")
                for col_idx, cell_text in enumerate(cells):
                    if col_idx >= cols:
                        break
                    try:
                        table.Cell(row_idx + 1, col_idx + 1).Range.Text = cell_text
                    except Exception:
                        pass

        return doc.Tables.Count

    def get_table_count(self) -> int:
        """Get the number of tables in the document."""
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        return doc.Tables.Count

    def get_table_data(self, index: int) -> list[list[str]]:
        """
        Get all data from a table as a 2D list.

        Args:
            index: 1-based table index.

        Returns:
            2D list of cell text values.
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        table = doc.Tables(index)
        result = []
        for row_idx in range(1, table.Rows.Count + 1):
            row = []
            for col_idx in range(1, table.Columns.Count + 1):
                try:
                    cell_text = table.Cell(row_idx, col_idx).Range.Text
                    # Remove trailing \r\x07 (cell markers)
                    row.append(cell_text.rstrip("\r\x07"))
                except Exception:
                    row.append("")
            result.append(row)
        return result

    def set_cell_text(
        self,
        table_index: int,
        row: int,
        col: int,
        text: str,
    ) -> None:
        """
        Set text in a table cell.

        Args:
            table_index: 1-based table index.
            row: 1-based row index.
            col: 1-based column index.
            text: Text to set.
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        doc.Tables(table_index).Cell(row, col).Range.Text = text

    def add_table_row(self, table_index: int) -> None:
        """
        Add a row to the end of a table.

        Args:
            table_index: 1-based table index.
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        table = doc.Tables(table_index)
        table.Rows.Add()

    def add_table_column(self, table_index: int) -> None:
        """
        Add a column to the end of a table.

        Args:
            table_index: 1-based table index.
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        table = doc.Tables(table_index)
        table.Columns.Add()

    def delete_table(self, index: int) -> None:
        """
        Delete a table by index.

        Args:
            index: 1-based table index.
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        doc.Tables(index).Delete()

    # ── Section Operations ───────────────────────────────────────────

    def add_section_break(self) -> None:
        """Add a section break at the end of the document."""
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        rng = doc.Content
        rng.Collapse(Direction=0)  # wdCollapseEnd=0
        rng.InsertBreak(Type=2)  # wdSectionBreakNextPage=2

    def get_section_count(self) -> int:
        """Get the number of sections in the document."""
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        return doc.Sections.Count

    # ── Page Setup ───────────────────────────────────────────────────

    def set_page_orientation(self, orientation: str) -> None:
        """
        Set page orientation for the document.

        Args:
            orientation: 'portrait' or 'landscape'.
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        orient_val = 0 if orientation.lower() == "portrait" else 1
        doc.PageSetup.Orientation = orient_val  # wdOrientPortrait=0, wdOrientLandscape=1

    def set_page_margins(
        self,
        left: float | None = None,
        right: float | None = None,
        top: float | None = None,
        bottom: float | None = None,
    ) -> None:
        """
        Set page margins in points.

        Args:
            left, right, top, bottom: Margin values in points.
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        ps = doc.PageSetup
        if left is not None:
            ps.LeftMargin = left
        if right is not None:
            ps.RightMargin = right
        if top is not None:
            ps.TopMargin = top
        if bottom is not None:
            ps.BottomMargin = bottom

    def set_page_size(
        self,
        width: float | None = None,
        height: float | None = None,
    ) -> None:
        """
        Set page size in points (72 points = 1 inch).

        Args:
            width: Page width in points (595 for A4).
            height: Page height in points (842 for A4).
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        ps = doc.PageSetup
        if width is not None:
            ps.PageWidth = width
        if height is not None:
            ps.PageHeight = height

    # ── Header / Footer ──────────────────────────────────────────────

    def add_header(self, text: str) -> None:
        """
        Set the primary header text.

        Args:
            text: Header text.
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        section = doc.Sections(1)
        header = section.Headers(1)  # wdHeaderFooterPrimary=1
        header.Range.Text = text

    def add_footer(self, text: str) -> None:
        """
        Set the primary footer text.

        Args:
            text: Footer text.
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        section = doc.Sections(1)
        footer = section.Footers(1)  # wdHeaderFooterPrimary=1
        footer.Range.Text = text

    # ── Export / Save-As ─────────────────────────────────────────────

    def export_to_pdf(self, filepath: str) -> str:
        """
        Export the active document to PDF.

        Args:
            filepath: Output PDF file path.

        Returns:
            The full path of the exported PDF.
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document to export.")
        full_path = os.path.abspath(filepath)
        doc.ExportAsFixedFormat(
            OutputFileName=full_path,
            ExportFormat=17,  # wdExportFormatPDF=17
        )
        return full_path

    # ── Image / Shape Operations ─────────────────────────────────────

    def insert_picture(
        self,
        filepath: str,
        left: float | None = None,
        top: float | None = None,
        width: float | None = None,
        height: float | None = None,
    ) -> str:
        """
        Insert a picture into the document.

        Args:
            filepath: Full path to the image file.
            left: Left position in points.
            top: Top position in points.
            width: Width in points.
            height: Height in points.

        Returns:
            The name of the inserted shape.
        """
        doc = self.active_document
        if doc is None:
            raise RuntimeError("No active document.")
        full_path = os.path.abspath(filepath)
        if not os.path.exists(full_path):
            raise FileNotFoundError(f"Image not found: {full_path}")

        shape = doc.Shapes.AddPicture(
            FileName=full_path,
            LinkToFile=False,
            SaveWithDocument=True,
        )
        if left is not None:
            shape.Left = left
        if top is not None:
            shape.Top = top
        if width is not None:
            shape.Width = width
        if height is not None:
            shape.Height = height

        return shape.Name

    def insert_page_break(self) -> None:
        """Insert a page break at the current cursor position."""
        sel = self.selection
        if sel is None:
            raise RuntimeError("No selection available.")
        sel.InsertBreak(Type=7)  # wdPageBreak=7

    # ── View / Zoom ──────────────────────────────────────────────────

    def set_zoom(self, percentage: int) -> None:
        """
        Set the zoom level of the active window.

        Args:
            percentage: Zoom percentage (10-500).
        """
        self.app.ActiveWindow.View.Zoom.Percentage = percentage
