"""
WPS Office Excel COM client.

Provides a high-level Python interface to automate WPS Office Excel
using the COM automation API (ET.Application).
"""

from __future__ import annotations

import os
import sys
import time
from typing import Any

import pythoncom
import win32com.client


class WPSExcelClient:
    """Client for interacting with WPS Office Excel via COM automation."""

    # WPS Office Excel ProgIDs to try (in order of preference)
    _PROG_IDS = [
        "ET.Application",  # WPS Excel Table
        "KET.Application",  # Older WPS Excel
    ]

    def __init__(self, visible: bool = False) -> None:
        """
        Initialize the WPS Excel client.

        Args:
            visible: If True, make the WPS Excel window visible.
        """
        self._app: Any = None
        self._visible = visible
        self._connect()

    def _connect(self) -> None:
        """Connect to a running WPS Excel instance or create a new one."""
        pythoncom.CoInitialize()

        existing_instance = False

        # Try connecting with each ProgID, preferring ET.Application
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

            # Verify the connection works by accessing Workbooks
            if self._app is not None:
                try:
                    _ = self._app.Workbooks.Count
                    break  # Connection verified
                except Exception:
                    self._app = None
                    existing_instance = False
                    continue

        if self._app is None:
            raise RuntimeError(
                "Could not connect to WPS Office Excel. "
                "Please ensure WPS Office is installed."
            )

        # Only force visibility on new instances. For existing instances,
        # preserve whatever visibility the user already has set.
        if not existing_instance:
            self._app.Visible = self._visible

    @property
    def app(self) -> Any:
        """Get the underlying COM application object."""
        if self._app is None:
            self._connect()
        return self._app

    @property
    def workbooks(self) -> Any:
        """Get the Workbooks collection."""
        return self.app.Workbooks

    @property
    def active_workbook(self) -> Any | None:
        """Get the active workbook, or None if no workbook is open."""
        try:
            return self.app.ActiveWorkbook
        except Exception:
            return None

    @property
    def active_sheet(self) -> Any | None:
        """Get the active worksheet, or None if no sheet is active."""
        try:
            return self.app.ActiveSheet
        except Exception:
            return None

    # ── Workbook Operations ──────────────────────────────────────────

    def create_workbook(self) -> str:
        """
        Create a new workbook.

        Returns:
            The name of the new workbook.
        """
        wb = self.workbooks.Add()
        return wb.Name

    def open_workbook(self, filepath: str) -> str:
        """
        Open an existing workbook.

        Args:
            filepath: Full path to the workbook file.

        Returns:
            The name of the opened workbook.
        """
        full_path = os.path.abspath(filepath)
        if not os.path.exists(full_path):
            raise FileNotFoundError(f"Workbook not found: {full_path}")
        wb = self.workbooks.Open(full_path)
        return wb.Name

    def save_workbook(self, filepath: str | None = None) -> str:
        """
        Save the active workbook.

        Args:
            filepath: Path to save to. If None, saves to current location.

        Returns:
            The full path where the workbook was saved.
        """
        wb = self.active_workbook
        if wb is None:
            raise RuntimeError("No active workbook to save.")

        if filepath:
            full_path = os.path.abspath(filepath)
            wb.SaveAs(full_path)
        else:
            wb.Save()
            full_path = wb.FullName

        return full_path

    def close_workbook(self, save: bool = True) -> None:
        """
        Close the active workbook.

        Args:
            save: If True, save changes before closing.
        """
        wb = self.active_workbook
        if wb is None:
            raise RuntimeError("No active workbook to close.")
        wb.Close(SaveChanges=save)

    def list_workbooks(self) -> list[dict[str, str]]:
        """
        List all open workbooks.

        Returns:
            List of dicts with 'name', 'fullname', and 'sheets_count'.
        """
        result = []
        for i in range(1, self.workbooks.Count + 1):
            wb = self.workbooks.Item(i)
            result.append({
                "name": wb.Name,
                "fullname": wb.FullName,
                "sheets_count": str(wb.Sheets.Count),
            })
        return result

    # ── Worksheet Operations ─────────────────────────────────────────

    def add_sheet(self, name: str | None = None, before: Any = None, after: Any = None) -> str:
        """
        Add a new worksheet.

        Args:
            name: Optional name for the new sheet.
            before: Sheet object or name to add before.
            after: Sheet object or name to add after.

        Returns:
            The name of the new sheet.
        """
        wb = self.active_workbook
        if wb is None:
            raise RuntimeError("No active workbook.")

        kwargs = {}
        if before is not None:
            kwargs["Before"] = before
        elif after is not None:
            kwargs["After"] = after

        sheet = wb.Sheets.Add(**kwargs) if kwargs else wb.Sheets.Add()
        if name:
            sheet.Name = name
        return sheet.Name

    def activate_sheet(self, name: str) -> str:
        """
        Activate a worksheet by name.

        Args:
            name: The name of the sheet to activate.

        Returns:
            The name of the activated sheet.
        """
        wb = self.active_workbook
        if wb is None:
            raise RuntimeError("No active workbook.")

        sheet = wb.Sheets(name)
        sheet.Activate()
        return sheet.Name

    def rename_sheet(self, old_name: str, new_name: str) -> str:
        """
        Rename a worksheet.

        Args:
            old_name: Current name of the sheet.
            new_name: New name for the sheet.

        Returns:
            The new name of the sheet.
        """
        wb = self.active_workbook
        if wb is None:
            raise RuntimeError("No active workbook.")

        sheet = wb.Sheets(old_name)
        sheet.Name = new_name
        return new_name

    def delete_sheet(self, name: str) -> None:
        """
        Delete a worksheet by name.

        Args:
            name: The name of the sheet to delete.
        """
        wb = self.active_workbook
        if wb is None:
            raise RuntimeError("No active workbook.")

        sheet = wb.Sheets(name)
        sheet.Delete()

    def list_sheets(self) -> list[dict[str, str]]:
        """
        List all sheets in the active workbook.

        Returns:
            List of dicts with 'name', 'index', 'type', and 'visible'.
        """
        wb = self.active_workbook
        if wb is None:
            raise RuntimeError("No active workbook.")

        result = []
        for i in range(1, wb.Sheets.Count + 1):
            sheet = wb.Sheets.Item(i)
            sheet_type = "worksheet"
            try:
                # xlSheetType: -4167 = xlWorksheet, -4116 = xlChart, etc.
                if sheet.Type == -4167:
                    sheet_type = "worksheet"
                elif sheet.Type == -4116:
                    sheet_type = "chart"
                else:
                    sheet_type = "other"
            except Exception:
                pass

            result.append({
                "name": sheet.Name,
                "index": str(i),
                "type": sheet_type,
                "visible": str(sheet.Visible),
            })
        return result

    # ── Cell Operations ──────────────────────────────────────────────

    def get_cell_value(self, cell_ref: str, sheet_name: str | None = None) -> Any:
        """
        Get the value of a cell.

        Args:
            cell_ref: Cell reference like "A1", "B2", etc.
            sheet_name: Optional sheet name (uses active sheet if None).

        Returns:
            The value of the cell.
        """
        sheet = self._get_sheet(sheet_name)
        value = sheet.Range(cell_ref).Value
        return value

    def set_cell_value(self, cell_ref: str, value: Any, sheet_name: str | None = None) -> None:
        """
        Set the value of a cell.

        Args:
            cell_ref: Cell reference like "A1", "B2", etc.
            value: The value to set.
            sheet_name: Optional sheet name (uses active sheet if None).
        """
        sheet = self._get_sheet(sheet_name)
        sheet.Range(cell_ref).Value = value

    def get_range_values(
        self,
        range_ref: str,
        sheet_name: str | None = None,
        as_list: bool = True,
    ) -> Any:
        """
        Get values from a range of cells.

        Args:
            range_ref: Range reference like "A1:B10".
            sheet_name: Optional sheet name.
            as_list: If True, return as list of lists.

        Returns:
            The range values.
        """
        sheet = self._get_sheet(sheet_name)
        rng = sheet.Range(range_ref)
        if as_list:
            values = rng.Value
            if values is None:
                return []
            # If it's a single value, wrap it
            if not isinstance(values, tuple):
                return [[values]]
            # If it's a 1D tuple (single row), wrap in list
            if len(values) > 0 and not isinstance(values[0], tuple):
                return [list(values)]
            return [list(row) for row in values]
        return rng.Value

    def set_range_values(
        self,
        range_ref: str,
        values: list[list[Any]],
        sheet_name: str | None = None,
    ) -> None:
        """
        Set values for a range of cells.

        Args:
            range_ref: Top-left cell reference or range reference.
            values: 2D list of values to set.
            sheet_name: Optional sheet name.
        """
        sheet = self._get_sheet(sheet_name)
        rows = len(values)
        cols = max(len(row) for row in values) if rows > 0 else 0

        # Expand to the full range
        col_letter = self._col_to_letter(self._col_from_ref(range_ref))
        start_row = self._row_from_ref(range_ref)
        end_col_letter = self._col_to_letter(self._col_from_ref(range_ref) + cols - 1)
        end_row = start_row + rows - 1

        full_range = f"{col_letter}{start_row}:{end_col_letter}{end_row}"
        sheet.Range(full_range).Value = values

    def clear_cell(self, cell_ref: str, sheet_name: str | None = None) -> None:
        """
        Clear the contents of a cell.

        Args:
            cell_ref: Cell reference.
            sheet_name: Optional sheet name.
        """
        sheet = self._get_sheet(sheet_name)
        sheet.Range(cell_ref).ClearContents()

    def clear_range(self, range_ref: str, sheet_name: str | None = None) -> None:
        """
        Clear the contents of a range.

        Args:
            range_ref: Range reference.
            sheet_name: Optional sheet name.
        """
        sheet = self._get_sheet(sheet_name)
        sheet.Range(range_ref).ClearContents()

    # ── Formatting Operations ────────────────────────────────────────

    def set_font_bold(
        self,
        cell_ref: str,
        bold: bool = True,
        sheet_name: str | None = None,
    ) -> None:
        """Set font bold for a cell or range."""
        sheet = self._get_sheet(sheet_name)
        sheet.Range(cell_ref).Font.Bold = bold

    def set_font_size(
        self,
        cell_ref: str,
        size: int,
        sheet_name: str | None = None,
    ) -> None:
        """Set font size for a cell or range."""
        sheet = self._get_sheet(sheet_name)
        sheet.Range(cell_ref).Font.Size = size

    def set_font_color(
        self,
        cell_ref: str,
        color: int,
        sheet_name: str | None = None,
    ) -> None:
        """
        Set font color for a cell or range.

        Args:
            color: RGB color value (e.g., 0xFF0000 for red).
        """
        sheet = self._get_sheet(sheet_name)
        sheet.Range(cell_ref).Font.Color = color

    def set_cell_color(
        self,
        cell_ref: str,
        color: int,
        sheet_name: str | None = None,
    ) -> None:
        """
        Set background (interior) color for a cell or range.

        Args:
            color: RGB color value (e.g., 0x00FF00 for green).
        """
        sheet = self._get_sheet(sheet_name)
        sheet.Range(cell_ref).Interior.Color = color

    def set_horizontal_alignment(
        self,
        cell_ref: str,
        alignment: str,
        sheet_name: str | None = None,
    ) -> None:
        """
        Set horizontal alignment.

        Args:
            alignment: 'left', 'center', 'right', 'general', 'justify', etc.
        """
        sheet = self._get_sheet(sheet_name)
        xl_align = {
            "left": -4131,   # xlLeft
            "center": -4108,  # xlCenter
            "right": -4152,   # xlRight
            "general": 1,     # xlGeneral
            "justify": -4130, # xlJustify
        }
        align_val = xl_align.get(alignment.lower(), -4108)
        sheet.Range(cell_ref).HorizontalAlignment = align_val

    def set_number_format(
        self,
        cell_ref: str,
        fmt: str,
        sheet_name: str | None = None,
    ) -> None:
        """
        Set number format for a cell or range.

        Args:
            fmt: Excel number format string (e.g., "0.00", "#,##0", etc.).
        """
        sheet = self._get_sheet(sheet_name)
        sheet.Range(cell_ref).NumberFormat = fmt

    def set_font_name(
        self,
        cell_ref: str,
        font_name: str,
        sheet_name: str | None = None,
    ) -> None:
        """Set font name for a cell or range."""
        sheet = self._get_sheet(sheet_name)
        sheet.Range(cell_ref).Font.Name = font_name

    def set_font_italic(
        self,
        cell_ref: str,
        italic: bool = True,
        sheet_name: str | None = None,
    ) -> None:
        """Set font italic for a cell or range."""
        sheet = self._get_sheet(sheet_name)
        sheet.Range(cell_ref).Font.Italic = italic

    def set_wrap_text(
        self,
        cell_ref: str,
        wrap: bool = True,
        sheet_name: str | None = None,
    ) -> None:
        """Set text wrapping for a cell or range."""
        sheet = self._get_sheet(sheet_name)
        sheet.Range(cell_ref).WrapText = wrap

    def merge_cells(self, range_ref: str, sheet_name: str | None = None) -> None:
        """Merge a range of cells."""
        sheet = self._get_sheet(sheet_name)
        sheet.Range(range_ref).Merge()

    def unmerge_cells(self, range_ref: str, sheet_name: str | None = None) -> None:
        """Unmerge a range of cells."""
        sheet = self._get_sheet(sheet_name)
        sheet.Range(range_ref).UnMerge()

    def set_borders(
        self,
        cell_ref: str,
        border_style: str = "thin",
        border_color: int | None = None,
        outline_only: bool = False,
        sheet_name: str | None = None,
    ) -> None:
        """
        Set borders for a cell or range.

        Args:
            border_style: 'thin', 'medium', 'thick', 'dotted', 'dashed', 'double',
                          'hairline', 'none', 'dashDot', 'dashDotDot', 'slantDashDot'.
            border_color: Optional RGB color for the border.
            outline_only: If True, only set outer border. If False, set all borders.
        """
        sheet = self._get_sheet(sheet_name)
        rng = sheet.Range(cell_ref)

        xl_styles = {
            "thin": 2,          # xlThin
            "medium": -4138,    # xlMedium
            "thick": 4,         # xlThick
            "dotted": -4118,    # xlDot
            "dashed": -4115,    # xlDash
            "double": -4119,    # xlDouble
            "hairline": 7,      # xlHairline
            "none": -4142,      # xlNone
            "dashDot": 4,       # xlDashDot (may vary)
            "dashDotDot": 5,    # xlDashDotDot
            "slantDashDot": 13, # xlSlantDashDot
        }
        style_val = xl_styles.get(border_style.lower(), 2)

        if outline_only:
            edges = [7, 8, 9, 10]  # xlEdgeLeft=7, xlEdgeTop=8, xlEdgeRight=10, xlEdgeBottom=9
            for edge in edges:
                border = rng.Borders(edge)
                border.LineStyle = style_val
                if border_color is not None:
                    border.Color = border_color
        else:
            # Set all borders (inside + outline)
            # xlInsideVertical=11, xlInsideHorizontal=12
            all_edges = [7, 8, 9, 10, 11, 12]
            for edge in all_edges:
                border = rng.Borders(edge)
                border.LineStyle = style_val
                if border_color is not None:
                    border.Color = border_color

    # ── Row / Column Operations ──────────────────────────────────────

    def get_used_range_address(self, sheet_name: str | None = None) -> str:
        """
        Get the address of the used range on a sheet.

        Returns:
            The range address string (e.g., "A1:D20").
        """
        sheet = self._get_sheet(sheet_name)
        return sheet.UsedRange.Address.replace("$", "")

    def get_row_count(self, sheet_name: str | None = None) -> int:
        """Get the number of used rows in a sheet."""
        sheet = self._get_sheet(sheet_name)
        return sheet.UsedRange.Rows.Count

    def get_column_count(self, sheet_name: str | None = None) -> int:
        """Get the number of used columns in a sheet."""
        sheet = self._get_sheet(sheet_name)
        return sheet.UsedRange.Columns.Count

    def insert_row(self, row: int, sheet_name: str | None = None) -> None:
        """Insert a row at the specified position."""
        sheet = self._get_sheet(sheet_name)
        sheet.Rows(row).Insert()

    def insert_column(self, col: int, sheet_name: str | None = None) -> None:
        """Insert a column at the specified position."""
        sheet = self._get_sheet(sheet_name)
        col_letter = self._col_to_letter(col)
        sheet.Columns(f"{col_letter}:{col_letter}").Insert()

    def delete_row(self, row: int, sheet_name: str | None = None) -> None:
        """Delete a row at the specified position."""
        sheet = self._get_sheet(sheet_name)
        sheet.Rows(row).Delete()

    def delete_column(self, col: int, sheet_name: str | None = None) -> None:
        """Delete a column at the specified position."""
        sheet = self._get_sheet(sheet_name)
        col_letter = self._col_to_letter(col)
        sheet.Columns(f"{col_letter}:{col_letter}").Delete()

    def set_row_height(self, row: int, height: float, sheet_name: str | None = None) -> None:
        """Set row height in points."""
        sheet = self._get_sheet(sheet_name)
        sheet.Rows(row).RowHeight = height

    def set_column_width(self, col: int, width: float, sheet_name: str | None = None) -> None:
        """Set column width in characters."""
        sheet = self._get_sheet(sheet_name)
        col_letter = self._col_to_letter(col)
        sheet.Columns(f"{col_letter}:{col_letter}").ColumnWidth = width

    def autofit_columns(
        self,
        sheet_name: str | None = None,
        start_col: int | None = None,
        end_col: int | None = None,
    ) -> None:
        """
        Auto-fit column widths to content.

        Args:
            start_col: Optional first column (1-based). If omitted, fits all used columns.
            end_col: Optional last column (1-based). If omitted, fits start_col or all used columns.
        """
        sheet = self._get_sheet(sheet_name)
        if start_col is not None:
            col_start = self._col_to_letter(start_col)
            if end_col is not None:
                col_end = self._col_to_letter(end_col)
                sheet.Columns(f"{col_start}:{col_end}").AutoFit()
            else:
                sheet.Columns(col_start).AutoFit()
        else:
            sheet.UsedRange.Columns.AutoFit()

    def autofit_rows(
        self,
        sheet_name: str | None = None,
        start_row: int | None = None,
        end_row: int | None = None,
    ) -> None:
        """
        Auto-fit row heights to content.

        Args:
            start_row: Optional first row. If omitted, fits all used rows.
            end_row: Optional last row. If omitted, fits start_row or all used rows.
        """
        sheet = self._get_sheet(sheet_name)
        if start_row is not None:
            if end_row is not None:
                sheet.Rows(f"{start_row}:{end_row}").AutoFit()
            else:
                sheet.Rows(start_row).AutoFit()
        else:
            sheet.UsedRange.Rows.AutoFit()

    def freeze_panes(
        self,
        cell_ref: str = "B2",
        sheet_name: str | None = None,
    ) -> None:
        """
        Freeze panes at a specific cell.

        Rows above and columns to the left of the cell will be frozen.

        Args:
            cell_ref: The cell at which to freeze. Default 'B2' freezes first row and first column.
                      Use 'A2' to freeze only the first row.
                      Use 'B1' to freeze only the first column.
        """
        sheet = self._get_sheet(sheet_name)
        # Activate the sheet first, then select the freeze cell
        sheet.Activate()
        sheet.Range(cell_ref).Select()
        self.app.ActiveWindow.FreezePanes = True

    def unfreeze_panes(self, sheet_name: str | None = None) -> None:
        """Remove frozen panes from a sheet."""
        sheet = self._get_sheet(sheet_name)
        sheet.Activate()
        self.app.ActiveWindow.FreezePanes = False

    def auto_filter(
        self,
        range_ref: str | None = None,
        sheet_name: str | None = None,
    ) -> None:
        """
        Add or toggle AutoFilter dropdowns for a range or the used range.

        Args:
            range_ref: Optional range to apply the filter to (e.g., 'A1:D100').
                       If omitted, uses the used range on the sheet.
        """
        sheet = self._get_sheet(sheet_name)
        if range_ref:
            rng = sheet.Range(range_ref)
        else:
            rng = sheet.UsedRange
        rng.AutoFilter()

    # ── Chart Operations ──────────────────────────────────────────

    def add_chart(
        self,
        chart_type: str = "column",
        range_ref: str = "",
        left: float = 100,
        top: float = 100,
        width: float = 400,
        height: float = 300,
        sheet_name: str | None = None,
    ) -> str:
        """
        Add a chart to a sheet.

        Args:
            chart_type: 'column', 'line', 'pie', 'bar', 'area', 'scatter'.
            range_ref: The data range for the chart.
            left, top, width, height: Position and size.
            sheet_name: Optional sheet name.

        Returns:
            The name of the new chart.
        """
        sheet = self._get_sheet(sheet_name)

        xl_chart_types = {
            "column": 51,    # xlColumnClustered
            "line": 4,       # xlLine
            "pie": 5,        # xlPie
            "bar": 57,       # xlBarClustered
            "area": 1,       # xlArea
            "scatter": -4169,  # xlXYScatter
        }

        ct = xl_chart_types.get(chart_type.lower(), 51)

        chart_obj = sheet.ChartObjects().Add(left, top, width, height)
        chart = chart_obj.Chart

        if range_ref:
            chart.SetSourceData(sheet.Range(range_ref))

        chart.ChartType = ct
        return chart_obj.Name

    # ── Sort / Copy-Paste ────────────────────────────────────────────

    def sort_range(
        self,
        range_ref: str,
        sort_key: str | None = None,
        sort_order: str = "ascending",
        sheet_name: str | None = None,
    ) -> None:
        """
        Sort a range of cells.

        Args:
            range_ref: The range to sort (e.g., 'A2:D100').
            sort_key: The column within the range to sort by (single cell ref, e.g., 'A2').
                      If omitted, sorts by the first column in the range.
            sort_order: 'ascending' or 'descending'.
        """
        sheet = self._get_sheet(sheet_name)
        rng = sheet.Range(range_ref)
        if sort_key is None:
            # Use the first cell in the range
            sort_key = range_ref.split(":")[0]
        order = 1 if sort_order.lower() == "ascending" else 2  # xlAscending=1, xlDescending=2
        rng.Sort(
            Key1=sheet.Range(sort_key),
            Order1=order,
            Header=0,  # xlGuess=0 (Excel guesses whether there's a header)
        )

    def copy_range(
        self,
        range_ref: str,
        sheet_name: str | None = None,
    ) -> None:
        """
        Copy a range to clipboard.

        Args:
            range_ref: The range to copy (e.g., 'A1:D10').
        """
        sheet = self._get_sheet(sheet_name)
        sheet.Range(range_ref).Copy()

    def paste_range(
        self,
        dest_cell: str,
        sheet_name: str | None = None,
        paste_special: str | None = None,
    ) -> None:
        """
        Paste clipboard contents to a destination cell.

        Args:
            dest_cell: The top-left cell to paste to (e.g., 'A1').
            paste_special: Optional: 'values', 'formats', 'formulas', 'all', 'transpose'.
        """
        sheet = self._get_sheet(sheet_name)
        dest = sheet.Range(dest_cell)

        if paste_special:
            xl_paste = {
                "values": -4163,    # xlPasteValues
                "formats": -4122,   # xlPasteFormats
                "formulas": -4123,  # xlPasteFormulas
                "all": -4104,       # xlPasteAll
                "transpose": 8,     # xlPasteAllUsingSourceTheme (then Transpose)
            }
            if paste_special.lower() == "transpose":
                dest.Select()
                sheet.Paste()
                # Transpose has limited COM support; fallback works with Paste
            else:
                paste_val = xl_paste.get(paste_special.lower(), -4104)
                dest.PasteSpecial(Paste=paste_val)
        else:
            dest.Select()
            sheet.Paste()

    # ── Query / Search ───────────────────────────────────────────────

    def find_cell(
        self,
        search_text: str,
        sheet_name: str | None = None,
        look_at: str = "part",
    ) -> dict[str, Any] | None:
        """
        Find a cell containing specific text.

        Args:
            search_text: Text to search for.
            sheet_name: Optional sheet name.
            look_at: 'part' or 'whole' (match part or whole cell).

        Returns:
            Dict with cell address and value, or None if not found.
        """
        sheet = self._get_sheet(sheet_name)
        xl_look_at = 2 if look_at == "part" else 1  # xlPart=2, xlWhole=1
        found = sheet.Cells.Find(search_text, LookAt=xl_look_at)
        if found is None:
            return None
        return {
            "address": found.Address.replace("$", ""),
            "row": str(found.Row),
            "column": str(found.Column),
            "value": found.Value,
        }

    def find_next_cell(
        self,
        sheet_name: str | None = None,
    ) -> dict[str, Any] | None:
        """
        Find the next occurrence after a previous find_cell() call.

        Returns:
            Dict with cell address and value, or None if no more matches.
        """
        sheet = self._get_sheet(sheet_name)
        found = sheet.Cells.FindNext()
        if found is None:
            return None
        return {
            "address": found.Address.replace("$", ""),
            "row": str(found.Row),
            "column": str(found.Column),
            "value": found.Value,
        }

    def add_comment(
        self,
        cell_ref: str,
        text: str,
        sheet_name: str | None = None,
    ) -> None:
        """
        Add a comment to a cell.

        Args:
            cell_ref: Cell reference (e.g., 'A1').
            text: Comment text.
        """
        sheet = self._get_sheet(sheet_name)
        rng = sheet.Range(cell_ref)
        rng.AddComment(text)

    def delete_comment(
        self,
        cell_ref: str,
        sheet_name: str | None = None,
    ) -> None:
        """
        Delete a comment from a cell.

        Args:
            cell_ref: Cell reference (e.g., 'A1').
        """
        sheet = self._get_sheet(sheet_name)
        rng = sheet.Range(cell_ref)
        if rng.Comment is not None:
            rng.Comment.Delete()

    def clear_formats(
        self,
        cell_ref: str,
        sheet_name: str | None = None,
    ) -> None:
        """
        Clear formats (but not content) from a cell or range.

        Args:
            cell_ref: Cell or range reference.
        """
        sheet = self._get_sheet(sheet_name)
        sheet.Range(cell_ref).ClearFormats()

    def clear_all(
        self,
        cell_ref: str,
        sheet_name: str | None = None,
    ) -> None:
        """
        Clear everything (contents, formats, comments) from a cell or range.

        Args:
            cell_ref: Cell or range reference.
        """
        sheet = self._get_sheet(sheet_name)
        sheet.Range(cell_ref).Clear()

    # ── Conditional Formatting ───────────────────────────────────────

    def add_conditional_format(
        self,
        range_ref: str,
        rule_type: str = "cellValue",
        operator: str = "greaterThan",
        formula: str = "0",
        font_color: int | None = None,
        bg_color: int | None = None,
        bold: bool | None = None,
        sheet_name: str | None = None,
    ) -> None:
        """
        Add a conditional formatting rule to a range.

        Args:
            range_ref: Range to apply to (e.g., 'A1:A100').
            rule_type: 'cellValue' for value-based rules.
            operator: 'greaterThan', 'lessThan', 'equal', 'between', 'greaterThanOrEqual',
                      'lessThanOrEqual', 'notEqual', 'notBetween'.
            formula: The threshold formula/value (e.g., '100', '=$B$1').
            font_color: Optional font color (RGB int).
            bg_color: Optional background color (RGB int).
            bold: Optional font bold.
        """
        sheet = self._get_sheet(sheet_name)
        rng = sheet.Range(range_ref)

        xl_ops = {
            "greaterThan": 5,
            "lessThan": 6,
            "equal": 3,
            "between": 1,
            "greaterThanOrEqual": 5,  # approximate with greaterThan + >=
            "lessThanOrEqual": 6,     # approximate
            "notEqual": 4,
            "notBetween": 2,
        }
        op_val = xl_ops.get(operator, 5)

        fc = rng.FormatConditions.Add(
            Type=1,  # xlCellValue
            Operator=op_val,
            Formula1=formula,
        )

        if font_color is not None:
            fc.Font.Color = font_color
        if bg_color is not None:
            fc.Interior.Color = bg_color
        if bold is not None:
            fc.Font.Bold = bold

    # ── Data Validation ──────────────────────────────────────────────

    def add_data_validation(
        self,
        range_ref: str,
        validation_type: str = "list",
        formula1: str = "",
        formula2: str = "",
        ignore_blank: bool = True,
        show_dropdown: bool = True,
        error_title: str = "",
        error_message: str = "",
        sheet_name: str | None = None,
    ) -> None:
        """
        Add data validation to a range.

        Args:
            range_ref: Range to validate (e.g., 'A1:A100').
            validation_type: 'list', 'whole', 'decimal', 'date', 'time', 'textLength', 'custom'.
            formula1: Validation formula (e.g., 'Option1,Option2,Option3' for list,
                      or '=$D$1:$D$10' for range-based list).
            formula2: Second formula (for 'between' / 'notBetween' operators on numeric types).
            ignore_blank: Allow blank cells.
            show_dropdown: Show a dropdown arrow (for list validation).
            error_title: Title for the error dialog.
            error_message: Error message to show.
        """
        sheet = self._get_sheet(sheet_name)
        rng = sheet.Range(range_ref)

        xl_types = {
            "list": 3,
            "whole": 1,
            "decimal": 2,
            "date": 4,
            "time": 5,
            "textLength": 6,
            "custom": 7,
        }
        val_type = xl_types.get(validation_type, 3)

        rng.Validation.Add(
            Type=val_type,
            AlertStyle=1,  # xlValidAlertStop
            Operator=1 if not formula2 else 1,  # xlBetween=1
            Formula1=formula1,
            Formula2=formula2,
        )
        rng.Validation.IgnoreBlank = ignore_blank
        rng.Validation.InCellDropdown = show_dropdown

        if error_title and validation_type == "list":
            rng.Validation.ErrorTitle = error_title
        if error_message and validation_type == "list":
            rng.Validation.ErrorMessage = error_message

    # ── Sheet Protection ─────────────────────────────────────────────

    def protect_sheet(
        self,
        password: str = "",
        allow_sort: bool = False,
        allow_filter: bool = False,
        allow_format_cells: bool = False,
        allow_insert_rows: bool = False,
        allow_delete_rows: bool = False,
        sheet_name: str | None = None,
    ) -> None:
        """
        Protect a worksheet.

        Args:
            password: Optional protection password.
            allow_sort: Allow sorting locked cells.
            allow_filter: Allow using autofilter.
            allow_format_cells: Allow formatting cells.
            allow_insert_rows: Allow inserting rows.
            allow_delete_rows: Allow deleting rows.
        """
        sheet = self._get_sheet(sheet_name)
        sheet.Protect(
            Password=password,
            DrawingObjects=True,
            Contents=True,
            Scenarios=True,
            AllowSorting=allow_sort,
            AllowFiltering=allow_filter,
            AllowFormattingCells=allow_format_cells,
            AllowInsertingRows=allow_insert_rows,
            AllowDeletingRows=allow_delete_rows,
        )

    def unprotect_sheet(
        self,
        password: str = "",
        sheet_name: str | None = None,
    ) -> None:
        """
        Remove protection from a worksheet.

        Args:
            password: Password if the sheet was protected with one.
        """
        sheet = self._get_sheet(sheet_name)
        if password:
            sheet.Unprotect(Password=password)
        else:
            sheet.Unprotect()

    # ── Page Setup ───────────────────────────────────────────────────

    def set_print_area(
        self,
        range_ref: str,
        sheet_name: str | None = None,
    ) -> None:
        """
        Set the print area for a sheet.

        Args:
            range_ref: The range to set as print area (e.g., 'A1:F50').
        """
        sheet = self._get_sheet(sheet_name)
        sheet.PageSetup.PrintArea = range_ref

    def clear_print_area(self, sheet_name: str | None = None) -> None:
        """Clear the print area for a sheet."""
        sheet = self._get_sheet(sheet_name)
        sheet.PageSetup.PrintArea = ""

    def set_page_orientation(
        self,
        orientation: str = "portrait",
        sheet_name: str | None = None,
    ) -> None:
        """
        Set page orientation.

        Args:
            orientation: 'portrait' (1) or 'landscape' (2).
        """
        sheet = self._get_sheet(sheet_name)
        xl_orient = {"portrait": 1, "landscape": 2}
        sheet.PageSetup.Orientation = xl_orient.get(orientation.lower(), 1)

    # ── Application Control ───────────────────────────────────────

    def quit(self, save: bool = False) -> None:
        """Quit the WPS Excel application."""
        if self._app:
            self._app.Quit()
            self._app = None

    def show(self) -> None:
        """Make the WPS Excel window visible."""
        self.app.Visible = True
        self._visible = True

    def hide(self) -> None:
        """Hide the WPS Excel window."""
        self.app.Visible = False
        self._visible = False

    def run_macro(self, macro_name: str, *args: Any) -> Any:
        """
        Run a VBA macro.

        Args:
            macro_name: The name of the macro to run.
            *args: Arguments to pass to the macro.

        Returns:
            The return value of the macro.
        """
        if args:
            return self.app.Run(macro_name, *args)
        return self.app.Run(macro_name)

    def get_app_info(self) -> dict[str, str]:
        """Get information about the WPS Excel application."""
        info: dict[str, str] = {}
        try:
            info["name"] = str(self.app.Name)
        except Exception:
            info["name"] = "WPS Excel (unknown)"
        try:
            info["version"] = str(self.app.Version)
        except Exception:
            info["version"] = "unknown"
        info["visible"] = str(self._visible)
        try:
            info["workbooks_count"] = str(self.workbooks.Count)
        except Exception:
            info["workbooks_count"] = "unknown"
        return info

    # ── Helper Methods ───────────────────────────────────────────────

    def _get_sheet(self, sheet_name: str | None = None) -> Any:
        """Get a sheet by name, or the active sheet if name is None."""
        if sheet_name:
            wb = self.active_workbook
            if wb is None:
                raise RuntimeError("No active workbook.")
            return wb.Sheets(sheet_name)

        sheet = self.active_sheet
        if sheet is None:
            raise RuntimeError("No active sheet.")
        return sheet

    @staticmethod
    def _col_from_ref(ref: str) -> int:
        """Extract column number from a cell reference like 'A1' or 'AB10'."""
        col_str = ""
        for c in ref:
            if c.isalpha():
                col_str += c
            else:
                break
        result = 0
        for i, c in enumerate(reversed(col_str.upper())):
            result += (ord(c) - ord("A") + 1) * (26**i)
        return result

    @staticmethod
    def _row_from_ref(ref: str) -> int:
        """Extract row number from a cell reference like 'A1' or 'AB10'."""
        num_str = ""
        for c in ref:
            if c.isdigit():
                num_str += c
        return int(num_str) if num_str else 1

    @staticmethod
    def _col_to_letter(col: int) -> str:
        """Convert a column number (1-based) to a letter (A, B, ..., Z, AA, etc.)."""
        result = ""
        while col > 0:
            col, remainder = divmod(col - 1, 26)
            result = chr(65 + remainder) + result
        return result
