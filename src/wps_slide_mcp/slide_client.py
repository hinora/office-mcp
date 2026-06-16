"""
WPS Office Slide (Presentation) COM client.

Provides a high-level Python interface to automate WPS Office Slide
using the COM automation API (WPP.Application).
"""

from __future__ import annotations

import os
import time
from typing import Any

import pythoncom
import win32com.client


class WPSSlideClient:
    """Client for interacting with WPS Office Slide via COM automation."""

    # WPS Office Slide ProgIDs to try (in order of preference)
    _PROG_IDS = [
        "WPP.Application",   # WPS Slide (Presentation)
        "KWPP.Application",  # Older WPS Slide
        "PowerPoint.Application",  # Fallback to MS PowerPoint
    ]

    def __init__(self, visible: bool = False) -> None:
        """
        Initialize the WPS Slide client.

        Args:
            visible: If True, make the WPS Slide window visible.
        """
        self._app: Any = None
        self._visible = visible
        self._connect()

    def _connect(self) -> None:
        """Connect to a running WPS Slide instance or create a new one.

        Tries all ProgIDs with GetActiveObject FIRST (reuse) before
        falling back to Dispatch (create new). This prevents spawning
        a duplicate instance when WPS Slide is already open but not
        registered under the first ProgID tried.
        """
        # COM is already initialized on the STA thread by the server's
        # _init_sta(). Calling CoInitialize again is harmless (nested),
        # but we skip it to reduce overhead.

        # Phase 1: Try to reuse an already-running WPS Slide instance
        # across ALL known ProgIDs before creating a new one.
        for prog_id in self._PROG_IDS:
            try:
                self._app = win32com.client.GetActiveObject(prog_id)
            except Exception:
                continue

            # Verify the connection works
            if self._app is not None:
                try:
                    _ = self._app.Presentations.Count
                    break  # Connection verified — reuse this instance
                except Exception:
                    self._app = None

        # Phase 2: No running instance found — create a new one.
        if self._app is None:
            for prog_id in self._PROG_IDS:
                try:
                    self._app = win32com.client.Dispatch(prog_id)
                except Exception:
                    continue

                # Verify the connection works
                if self._app is not None:
                    try:
                        _ = self._app.Presentations.Count
                        break  # Connection verified
                    except Exception:
                        self._app = None

        if self._app is None:
            raise RuntimeError(
                "Could not connect to WPS Office Slide. "
                "Please ensure WPS Office is installed."
            )

        # Always enforce visibility — the MCP server is a user-facing
        # automation tool and users expect to see the app. Also restore
        # from minimized state so it doesn't stay hidden in the taskbar.
        self._app.Visible = self._visible
        if self._visible:
            try:
                self._app.WindowState = 2  # ppWindowNormal (restore)
            except Exception:
                pass

    @property
    def app(self) -> Any:
        """Get the underlying COM application object."""
        if self._app is None:
            self._connect()
        return self._app

    @property
    def presentations(self) -> Any:
        """Get the Presentations collection."""
        return self.app.Presentations

    @property
    def active_presentation(self) -> Any | None:
        """Get the active presentation, or None if no presentation is open."""
        try:
            return self.app.ActivePresentation
        except Exception:
            return None

    @property
    def active_window(self) -> Any | None:
        """Get the active window, or None."""
        try:
            return self.app.ActiveWindow
        except Exception:
            return None

    @property
    def selection(self) -> Any | None:
        """Get the current selection, or None if unavailable."""
        try:
            return self.app.ActiveWindow.Selection
        except Exception:
            return None

    @property
    def slide_selection(self) -> Any | None:
        """Get the current slide selection (SlideRange)."""
        try:
            return self.selection.SlideRange if self.selection else None
        except Exception:
            return None

    # ── Application ──────────────────────────────────────────────────

    def get_app_info(self) -> dict[str, str]:
        """Get information about the WPS Slide application."""
        info: dict[str, str] = {}
        try:
            info["name"] = str(self.app.Name)
        except Exception:
            info["name"] = "WPS Slide (unknown)"
        try:
            info["version"] = str(self.app.Version)
        except Exception:
            info["version"] = "unknown"
        info["visible"] = str(self._visible)
        try:
            info["presentations_count"] = str(self.presentations.Count)
        except Exception:
            info["presentations_count"] = "unknown"
        try:
            pres = self.active_presentation
            info["active_presentation"] = pres.Name if pres else "none"
        except Exception:
            info["active_presentation"] = "none"
        return info

    def show(self) -> None:
        """Make the WPS Slide window visible."""
        self.app.Visible = True
        self._visible = True

    def hide(self) -> None:
        """Hide the WPS Slide window."""
        try:
            self.app.Visible = False
        except Exception:
            # Some PPT versions throw when setting Visible
            pass
        self._visible = False

    def quit_app(self) -> None:
        """Quit the WPS Slide application."""
        try:
            self.app.Quit()
        except Exception:
            pass
        self._app = None

    # ── Presentation Operations ──────────────────────────────────────

    def create_presentation(self) -> str:
        """
        Create a new blank presentation.

        Returns:
            The name of the new presentation.
        """
        pres = self.presentations.Add()
        try:
            return pres.Name
        except AttributeError:
            return self.active_presentation.Name if self.active_presentation else "Presentation1"

    def open_presentation(self, filepath: str) -> str:
        """
        Open an existing presentation from disk.

        Args:
            filepath: Full path to the presentation file (.pptx, .ppt, .dps, etc.).

        Returns:
            The name of the opened presentation.
        """
        full_path = os.path.abspath(filepath)
        if not os.path.exists(full_path):
            raise FileNotFoundError(f"Presentation not found: {full_path}")
        pres = self.presentations.Open(full_path)
        return pres.Name

    def save_presentation(self, filepath: str | None = None) -> str:
        """
        Save the active presentation.

        Args:
            filepath: Path to save to. If None, saves to current location.

        Returns:
            The full path where the presentation was saved.
        """
        pres = self.active_presentation
        if pres is None:
            raise RuntimeError("No active presentation to save.")

        if filepath:
            full_path = os.path.abspath(filepath)
            pres.SaveAs(full_path)
        else:
            pres.Save()
            full_path = pres.FullName

        return full_path

    def close_presentation(self, save: bool = True) -> None:
        """
        Close the active presentation.

        Args:
            save: If True, save changes before closing.
        """
        pres = self.active_presentation
        if pres is None:
            raise RuntimeError("No active presentation to close.")
        pres.Close()
        # Note: PowerPoint/Presentation COM Close doesn't always take SaveChanges

    def list_presentations(self) -> list[dict[str, str]]:
        """
        List all open presentations.

        Returns:
            List of dicts with 'name', 'fullname', and 'slides_count'.
        """
        result = []
        for i in range(1, self.presentations.Count + 1):
            pres = self.presentations.Item(i)
            result.append({
                "name": pres.Name,
                "fullname": pres.FullName or "",
                "slides_count": str(pres.Slides.Count),
            })
        return result

    def activate_presentation(self, name: str) -> str:
        """
        Activate a presentation by name.

        Args:
            name: The name of the presentation to activate.

        Returns:
            The name of the activated presentation.
        """
        for i in range(1, self.presentations.Count + 1):
            pres = self.presentations.Item(i)
            if pres.Name == name:
                try:
                    pres.Activate()
                except Exception:
                    # Some versions don't support Activate on Item
                    # Try activating via the Windows collection
                    try:
                        win = pres.Windows(1)
                        win.Activate()
                    except Exception:
                        pass
                return pres.Name
        raise ValueError(f"Presentation '{name}' not found.")

    # ── Slide Operations ─────────────────────────────────────────────

    def _get_slides(self):
        """Get the Slides collection from the active presentation."""
        pres = self.active_presentation
        if pres is None:
            raise RuntimeError("No active presentation.")
        return pres.Slides

    def get_slide_count(self) -> int:
        """Get the number of slides in the active presentation."""
        return self._get_slides().Count

    def add_slide(self, layout_index: int = 1) -> int:
        """
        Add a new slide with the given layout.

        Args:
            layout_index: 1-based index of the layout. 1=Title Slide, 2=Title and Content, etc.

        Returns:
            The 1-based index of the new slide.
        """
        pres = self.active_presentation
        if pres is None:
            raise RuntimeError("No active presentation.")
        layout = pres.SlideMaster.CustomLayouts(layout_index) if hasattr(pres.SlideMaster, 'CustomLayouts') else pres.SlideMaster.Designs.Item(1)
        slide = pres.Slides.AddSlide(pres.Slides.Count + 1, layout)
        return slide.SlideIndex

    def delete_slide(self, index: int) -> None:
        """
        Delete a slide by index.

        Args:
            index: 1-based slide index.
        """
        self._get_slides()(index).Delete()

    def duplicate_slide(self, index: int) -> int:
        """
        Duplicate a slide.

        Args:
            index: 1-based slide index to duplicate.

        Returns:
            The index of the duplicated slide.
        """
        slide = self._get_slides()(index)
        slide.Duplicate()
        return index + 1

    def move_slide(self, index: int, to_index: int) -> None:
        """
        Move a slide to a new position.

        Args:
            index: 1-based index of the slide to move.
            to_index: Destination 1-based index.
        """
        self._get_slides()(index).MoveTo(to_index)

    def go_to_slide(self, index: int) -> None:
        """
        Go to a specific slide in the active window.

        Args:
            index: 1-based slide index.
        """
        self.app.ActiveWindow.View.GotoSlide(index)

    def list_slides(self) -> list[dict[str, Any]]:
        """
        List all slides with basic info.

        Returns:
            List of dicts with 'index', 'name', 'layout', 'shapes_count'.
        """
        result = []
        slides = self._get_slides()
        for i in range(1, slides.Count + 1):
            slide = slides(i)
            result.append({
                "index": slide.SlideIndex,
                "name": slide.Name or "",
                "shapes_count": slide.Shapes.Count if hasattr(slide, 'Shapes') else 0,
            })
        return result

    def slide_set_background(self, index: int, color_rgb: int) -> None:
        """
        Set slide background color.

        Args:
            index: 1-based slide index.
            color_rgb: RGB color value (e.g., 0x0000FF for red).
        """
        slide = self._get_slides()(index)
        slide.Background.Fill.ForeColor.RGB = color_rgb
        slide.Background.Fill.Solid()

    def slide_set_transition(self, index: int, transition_type: int, duration: float = 1.0) -> None:
        """
        Set slide transition effect.

        Args:
            index: 1-based slide index.
            transition_type: Transition effect ID. Common values:
                1=Fade, 2=Push, 8=Wipe, 13=Random, 25=Zoom.
            duration: Transition duration in seconds.
        """
        slide = self._get_slides()(index)
        slide.SlideShowTransition.EntryEffect = transition_type
        slide.SlideShowTransition.Duration = duration

    # ── Shape Operations ─────────────────────────────────────────────

    def _get_slide_shapes(self, slide_index: int = 0):
        """
        Get shapes from a specific slide or the first slide.

        Args:
            slide_index: 1-based slide index. 0 = active slide.
        """
        if slide_index == 0:
            try:
                return self.app.ActiveWindow.View.Slide.Shapes
            except Exception:
                raise RuntimeError("No active slide view. Use go_to_slide first.")
        else:
            return self._get_slides()(slide_index).Shapes

    def add_text_box(
        self,
        text: str,
        left: float = 50,
        top: float = 50,
        width: float = 400,
        height: float = 100,
        slide_index: int = 0,
    ) -> str:
        """
        Add a text box to a slide.

        Args:
            text: The text to add.
            left, top, width, height: Position and size in points.
            slide_index: 1-based slide index, 0=active slide.

        Returns:
            The name of the new shape.
        """
        shapes = self._get_slide_shapes(slide_index)
        shape = shapes.AddTextbox(
            Orientation=1,  # msoTextOrientationHorizontal
            Left=left,
            Top=top,
            Width=width,
            Height=height,
        )
        shape.TextFrame.TextRange.Text = text
        return shape.Name

    def add_picture(
        self,
        filepath: str,
        left: float = 50,
        top: float = 50,
        width: float = -1,
        height: float = -1,
        slide_index: int = 0,
    ) -> str:
        """
        Add a picture to a slide.

        Args:
            filepath: Full path to the image file.
            left, top: Position in points.
            width, height: Size in points. -1 = keep original.
            slide_index: 1-based slide index, 0=active slide.

        Returns:
            The name of the new shape.
        """
        full_path = os.path.abspath(filepath)
        if not os.path.exists(full_path):
            raise FileNotFoundError(f"Image not found: {full_path}")
        shapes = self._get_slide_shapes(slide_index)
        shape = shapes.AddPicture(
            FileName=full_path,
            LinkToFile=False,
            SaveWithDocument=True,
            Left=left,
            Top=top,
            Width=width,
            Height=height,
        )
        return shape.Name

    def add_rectangle(
        self,
        left: float = 50,
        top: float = 50,
        width: float = 200,
        height: float = 100,
        slide_index: int = 0,
    ) -> str:
        """
        Add a rectangle shape to a slide.

        Returns:
            The name of the new shape.
        """
        shapes = self._get_slide_shapes(slide_index)
        shape = shapes.AddShape(
            Type=1,  # msoShapeRectangle
            Left=left, Top=top,
            Width=width, Height=height,
        )
        return shape.Name

    def add_oval(
        self,
        left: float = 50,
        top: float = 50,
        width: float = 200,
        height: float = 100,
        slide_index: int = 0,
    ) -> str:
        """
        Add an oval/ellipse shape to a slide.

        Returns:
            Name of the new shape.
        """
        shapes = self._get_slide_shapes(slide_index)
        shape = shapes.AddShape(
            Type=9,  # msoShapeOval
            Left=left, Top=top,
            Width=width, Height=height,
        )
        return shape.Name

    def add_arrow(
        self,
        left: float = 50,
        top: float = 50,
        width: float = 200,
        height: float = 50,
        slide_index: int = 0,
    ) -> str:
        """
        Add a right arrow shape to a slide.

        Returns:
            Name of the new shape.
        """
        shapes = self._get_slide_shapes(slide_index)
        shape = shapes.AddShape(
            Type=33,  # msoShapeRightArrow
            Left=left, Top=top,
            Width=width, Height=height,
        )
        return shape.Name

    def add_line(
        self,
        begin_x: float = 50,
        begin_y: float = 50,
        end_x: float = 300,
        end_y: float = 50,
        slide_index: int = 0,
    ) -> str:
        """
        Add a line to a slide.

        Returns:
            Name of the new shape.
        """
        shapes = self._get_slide_shapes(slide_index)
        shape = shapes.AddLine(
            BeginX=begin_x, BeginY=begin_y,
            EndX=end_x, EndY=end_y,
        )
        return shape.Name

    def get_shape_count(self, slide_index: int = 0) -> int:
        """Get the number of shapes on a slide."""
        return self._get_slide_shapes(slide_index).Count

    def list_shapes(self, slide_index: int = 0) -> list[dict[str, Any]]:
        """
        List all shapes on a slide.

        Returns:
            List of dicts with 'index', 'name', 'type', 'has_text'.
        """
        shapes = self._get_slide_shapes(slide_index)
        result = []
        for i in range(1, shapes.Count + 1):
            shape = shapes(i)
            info: dict[str, Any] = {
                "index": i,
                "name": shape.Name,
                "type": shape.Type if hasattr(shape, 'Type') else None,
            }
            try:
                info["has_text"] = shape.HasTextFrame == -1  # msoTrue = -1
            except Exception:
                info["has_text"] = False
            try:
                info["left"] = shape.Left
                info["top"] = shape.Top
                info["width"] = shape.Width
                info["height"] = shape.Height
            except Exception:
                pass
            result.append(info)
        return result

    def delete_shape(self, name_or_index: Any, slide_index: int = 0) -> None:
        """
        Delete a shape by name or 1-based index.

        Args:
            name_or_index: Shape name (str) or 1-based index (int).
            slide_index: 1-based slide index, 0=active slide.
        """
        shapes = self._get_slide_shapes(slide_index)
        shapes(name_or_index).Delete()

    def set_shape_position(
        self,
        name_or_index: Any,
        left: float | None = None,
        top: float | None = None,
        width: float | None = None,
        height: float | None = None,
        slide_index: int = 0,
    ) -> None:
        """
        Set shape position and size.

        Args:
            name_or_index: Shape name or 1-based index.
            left, top: New position in points.
            width, height: New size in points.
        """
        shapes = self._get_slide_shapes(slide_index)
        shape = shapes(name_or_index)
        if left is not None:
            shape.Left = left
        if top is not None:
            shape.Top = top
        if width is not None:
            shape.Width = width
        if height is not None:
            shape.Height = height

    def set_shape_fill(self, name_or_index: Any, color_rgb: int, slide_index: int = 0) -> None:
        """
        Set shape fill color.

        Args:
            name_or_index: Shape name or 1-based index.
            color_rgb: RGB color value (e.g., 0xFF0000 for red).
        """
        shapes = self._get_slide_shapes(slide_index)
        shape = shapes(name_or_index)
        shape.Fill.ForeColor.RGB = color_rgb
        shape.Fill.Solid()

    def set_shape_line(self, name_or_index: Any, color_rgb: int, weight: float = 1.0, slide_index: int = 0) -> None:
        """
        Set shape outline/border color and weight.

        Args:
            name_or_index: Shape name or 1-based index.
            color_rgb: RGB color value.
            weight: Line weight in points.
        """
        shapes = self._get_slide_shapes(slide_index)
        shape = shapes(name_or_index)
        shape.Line.ForeColor.RGB = color_rgb
        shape.Line.Weight = weight

    def set_shape_rotation(self, name_or_index: Any, rotation: float, slide_index: int = 0) -> None:
        """
        Set shape rotation.

        Args:
            name_or_index: Shape name or 1-based index.
            rotation: Rotation angle in degrees.
        """
        shapes = self._get_slide_shapes(slide_index)
        shape = shapes(name_or_index)
        shape.Rotation = rotation

    def set_shape_zorder(self, name_or_index: Any, zorder: str, slide_index: int = 0) -> None:
        """
        Set shape z-order.

        Args:
            name_or_index: Shape name or 1-based index.
            zorder: 'front', 'back', 'forward', 'backward'.
        """
        shapes = self._get_slide_shapes(slide_index)
        shape = shapes(name_or_index)
        if zorder == "front":
            shape.ZOrder(0)  # msoBringToFront
        elif zorder == "back":
            shape.ZOrder(1)  # msoSendToBack
        elif zorder == "forward":
            shape.ZOrder(2)  # msoBringForward
        elif zorder == "backward":
            shape.ZOrder(3)  # msoSendBackward

    def group_shapes(self, names: list[Any], slide_index: int = 0) -> str:
        """
        Group multiple shapes together.

        Args:
            names: List of shape names or indices to group.
            slide_index: 1-based slide index, 0=active slide.

        Returns:
            Name of the new group shape.
        """
        shapes = self._get_slide_shapes(slide_index)
        # Select each shape and build a ShapeRange
        shape_range = shapes.Range(names)
        group = shape_range.Group()
        return group.Name

    def ungroup_shapes(self, name_or_index: Any, slide_index: int = 0) -> None:
        """
        Ungroup a grouped shape.

        Args:
            name_or_index: Shape name or 1-based index of the group.
        """
        shapes = self._get_slide_shapes(slide_index)
        shapes(name_or_index).Ungroup()

    # ── Text Operations ──────────────────────────────────────────────

    def _get_shape_textframe(self, name_or_index: Any, slide_index: int):
        """Get the TextFrame of a shape."""
        shapes = self._get_slide_shapes(slide_index)
        shape = shapes(name_or_index)
        if not shape.HasTextFrame:
            raise RuntimeError(f"Shape '{name_or_index}' does not have a text frame.")
        return shape.TextFrame

    def get_shape_text(self, name_or_index: Any, slide_index: int = 0) -> str:
        """
        Get all text from a shape.

        Args:
            name_or_index: Shape name or 1-based index.

        Returns:
            The text content of the shape.
        """
        tf = self._get_shape_textframe(name_or_index, slide_index)
        return tf.TextRange.Text

    def set_shape_text(self, name_or_index: Any, text: str, slide_index: int = 0) -> None:
        """
        Set the text of a shape.

        Args:
            name_or_index: Shape name or 1-based index.
            text: The text to set.
        """
        tf = self._get_shape_textframe(name_or_index, slide_index)
        tf.TextRange.Text = text

    # ── Font Formatting ──────────────────────────────────────────────

    def _get_textrange(self, name_or_index: Any, slide_index: int = 0, start: int = 0, length: int = 0):
        """Get a TextRange from a shape, optionally a portion of it."""
        tf = self._get_shape_textframe(name_or_index, slide_index)
        tr = tf.TextRange
        if start > 0 and length > 0:
            return tr.Characters(start, length)
        return tr

    def set_font_bold(
        self,
        name_or_index: Any,
        bold: bool = True,
        start: int = 0,
        length: int = 0,
        slide_index: int = 0,
    ) -> None:
        """Set font bold on shape text."""
        tr = self._get_textrange(name_or_index, slide_index, start, length)
        tr.Font.Bold = -1 if bold else 0  # msoTrue=-1

    def set_font_italic(
        self,
        name_or_index: Any,
        italic: bool = True,
        start: int = 0,
        length: int = 0,
        slide_index: int = 0,
    ) -> None:
        """Set font italic on shape text."""
        tr = self._get_textrange(name_or_index, slide_index, start, length)
        tr.Font.Italic = -1 if italic else 0

    def set_font_underline(
        self,
        name_or_index: Any,
        underline: bool = True,
        start: int = 0,
        length: int = 0,
        slide_index: int = 0,
    ) -> None:
        """Set font underline on shape text."""
        tr = self._get_textrange(name_or_index, slide_index, start, length)
        tr.Font.Underline = -1 if underline else 0  # msoTrue=-1

    def set_font_name(
        self,
        name_or_index: Any,
        font_name: str,
        start: int = 0,
        length: int = 0,
        slide_index: int = 0,
    ) -> None:
        """Set font name on shape text."""
        tr = self._get_textrange(name_or_index, slide_index, start, length)
        tr.Font.Name = font_name

    def set_font_size(
        self,
        name_or_index: Any,
        size: float,
        start: int = 0,
        length: int = 0,
        slide_index: int = 0,
    ) -> None:
        """Set font size on shape text."""
        tr = self._get_textrange(name_or_index, slide_index, start, length)
        tr.Font.Size = size

    def set_font_color(
        self,
        name_or_index: Any,
        color_rgb: int,
        start: int = 0,
        length: int = 0,
        slide_index: int = 0,
    ) -> None:
        """Set font color on shape text."""
        tr = self._get_textrange(name_or_index, slide_index, start, length)
        tr.Font.Color.RGB = color_rgb

    def set_text_alignment(
        self,
        name_or_index: Any,
        alignment: str,
        slide_index: int = 0,
    ) -> None:
        """
        Set text alignment in a shape.

        Args:
            name_or_index: Shape name or 1-based index.
            alignment: 'left' (1), 'center' (2), 'right' (3), 'justify' (4).
        """
        tf = self._get_shape_textframe(name_or_index, slide_index)
        pp_align = {
            "left": 1,     # ppAlignLeft
            "center": 2,   # ppAlignCenter
            "right": 3,    # ppAlignRight
            "justify": 4,  # ppAlignJustify
        }
        tf.TextRange.ParagraphFormat.Alignment = pp_align.get(alignment.lower(), 1)

    # ── Table Operations ─────────────────────────────────────────────

    def add_table(
        self,
        rows: int,
        cols: int,
        left: float = 50,
        top: float = 50,
        width: float = 600,
        height: float = 300,
        slide_index: int = 0,
    ) -> str:
        """
        Add a table to a slide.

        Returns:
            Name of the new table shape.
        """
        shapes = self._get_slide_shapes(slide_index)
        shape = shapes.AddTable(
            NumRows=rows,
            NumColumns=cols,
            Left=left,
            Top=top,
            Width=width,
            Height=height,
        )
        return shape.Name

    def set_table_cell(
        self,
        shape_name: Any,
        row: int,
        col: int,
        text: str,
        slide_index: int = 0,
    ) -> None:
        """
        Set text in a table cell.

        Args:
            shape_name: Table shape name or 1-based index.
            row: 1-based row index.
            col: 1-based column index.
            text: Text to set.
        """
        shapes = self._get_slide_shapes(slide_index)
        table = shapes(shape_name).Table
        table.Cell(row, col).Shape.TextFrame.TextRange.Text = text

    def get_table_data(self, shape_name: Any, slide_index: int = 0) -> list[list[str]]:
        """
        Get all data from a table shape.

        Returns:
            2D list of cell text values.
        """
        shapes = self._get_slide_shapes(slide_index)
        table = shapes(shape_name).Table
        result = []
        for r in range(1, table.Rows.Count + 1):
            row_data = []
            for c in range(1, table.Columns.Count + 1):
                try:
                    cell_text = table.Cell(r, c).Shape.TextFrame.TextRange.Text
                    row_data.append(cell_text)
                except Exception:
                    row_data.append("")
            result.append(row_data)
        return result

    # ── Notes ────────────────────────────────────────────────────────

    def get_notes(self, slide_index: int) -> str:
        """
        Get speaker notes for a slide.

        Args:
            slide_index: 1-based slide index.

        Returns:
            Notes text.
        """
        slide = self._get_slides()(slide_index)
        try:
            return slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
        except Exception:
            return ""

    def set_notes(self, slide_index: int, text: str) -> None:
        """
        Set speaker notes for a slide.

        Args:
            slide_index: 1-based slide index.
            text: Notes text.
        """
        slide = self._get_slides()(slide_index)
        try:
            slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = text
        except Exception:
            # Notes page body placeholder may not exist (index 2 varies by locale).
            # Try to find any text-capable shape on the notes page.
            try:
                for i in range(1, slide.NotesPage.Shapes.Count + 1):
                    shp = slide.NotesPage.Shapes(i)
                    if shp.HasTextFrame:
                        shp.TextFrame.TextRange.Text = text
                        break
            except Exception:
                pass

    # ── Export ───────────────────────────────────────────────────────

    def export_to_pdf(self, filepath: str) -> str:
        """
        Export the active presentation to PDF.

        Args:
            filepath: Output PDF file path.

        Returns:
            The full path of the exported PDF.
        """
        pres = self.app.ActivePresentation
        if pres is None:
            raise RuntimeError("No active presentation to export.")
        full_path = os.path.abspath(filepath)
        full_path_com = full_path.replace('/', '\\')
        try:
            # Office 2010+: ExportAsFixedFormat
            pres.ExportAsFixedFormat(full_path_com, 2)
        except Exception:
            try:
                # Office 2007: SaveAs with ppSaveAsPDF = 32
                pres.SaveAs(full_path_com, 32)
            except Exception as e:
                raise RuntimeError(f"PDF export failed: {e}")
        return full_path

    def export_slide_image(
        self,
        filepath: str,
        slide_index: int,
        width: int = 1920,
        height: int = 1080,
    ) -> str:
        """
        Export a single slide as an image.

        Args:
            filepath: Output image file path (.png, .jpg).
            slide_index: 1-based slide index.
            width, height: Image dimensions in pixels.

        Returns:
            The full path of the exported image.
        """
        full_path = os.path.abspath(filepath)
        slide = self._get_slides()(slide_index)
        slide.Select()
        slide.Export(full_path, "PNG" if full_path.lower().endswith(".png") else "JPG", width, height)
        return full_path

    # ── Hyperlinks ───────────────────────────────────────────────────

    def add_hyperlink(
        self,
        address: str,
        text_to_display: str | None = None,
        name_or_index: Any = 0,
        slide_index: int = 0,
    ) -> None:
        """
        Add a hyperlink to a shape's text or selection.

        Args:
            address: URL or email address.
            text_to_display: Display text.
            name_or_index: Shape name or index. 0 = use selection.
        """
        if name_or_index == 0:
            sel = self.selection
            if sel is None:
                raise RuntimeError("No selection available.")
            tr = sel.TextRange
        else:
            tr = self._get_textrange(name_or_index, slide_index)

        # Use ActionSettings(1)=ppMouseClick to add hyperlink on the shape's text.
        # The old SlideShowWindow approach only works during a running slide show.
        tr.ActionSettings(1).Hyperlink.Address = address
        if text_to_display:
            tr.Text = text_to_display

    # ── Slide Show ───────────────────────────────────────────────────

    def start_slideshow(self) -> None:
        """Start the slide show from the beginning."""
        pres = self.active_presentation
        if pres is None:
            raise RuntimeError("No active presentation.")
        pres.SlideShowSettings.Run()

    def start_slideshow_from(self, slide_index: int) -> None:
        """
        Start the slide show from a specific slide.

        Args:
            slide_index: 1-based slide index.
        """
        pres = self.active_presentation
        if pres is None:
            raise RuntimeError("No active presentation.")
        ss = pres.SlideShowSettings
        ss.StartingSlide = slide_index
        ss.Run()

    def stop_slideshow(self) -> None:
        """Stop the running slide show."""
        try:
            self.app.SlideShowWindows(1).View.Exit()
        except Exception:
            pass

    # ── Presentation Properties ─────────────────────────────────────

    def get_presentation_properties(self) -> dict[str, Any]:
        """Get presentation metadata."""
        pres = self.active_presentation
        if pres is None:
            raise RuntimeError("No active presentation.")
        return {
            "name": pres.Name,
            "fullname": pres.FullName or "",
            "slides_count": pres.Slides.Count,
            "slide_width": pres.PageSetup.SlideWidth,
            "slide_height": pres.PageSetup.SlideHeight,
        }

    def set_slide_size(self, width: float, height: float) -> None:
        """
        Set the slide size for the presentation.

        Args:
            width: Slide width in points.
            height: Slide height in points.
        """
        pres = self.active_presentation
        if pres is None:
            raise RuntimeError("No active presentation.")
        pres.PageSetup.SlideWidth = width
        pres.PageSetup.SlideHeight = height

    # ── Shape Copy / Paste ──────────────────────────────────────────

    def copy_shape(self, name_or_index: Any, slide_index: int = 0) -> None:
        """Copy a shape to the clipboard."""
        shapes = self._get_slide_shapes(slide_index)
        shape = shapes(name_or_index)
        shape.Copy()

    def paste_shape(self, slide_index: int = 0) -> str:
        """
        Paste the clipboard content as a shape onto a slide.

        Returns:
            Name of the pasted shape.
        """
        shapes = self._get_slide_shapes(slide_index)
        shape_range = shapes.Paste()
        return shape_range(1).Name if shape_range.Count > 0 else ""

    def duplicate_shape(self, name_or_index: Any, slide_index: int = 0) -> str:
        """
        Duplicate a shape on the same slide (copy + paste offset).

        Returns:
            Name of the duplicated shape.
        """
        self.copy_shape(name_or_index, slide_index)
        new_name = self.paste_shape(slide_index)
        # Offset the duplicate slightly so it's visible
        shapes = self._get_slide_shapes(slide_index)
        try:
            dup = shapes(new_name)
            dup.Left += 20
            dup.Top += 20
        except Exception:
            pass
        return new_name

    # ── Animation ───────────────────────────────────────────────────

    # MSO animation effect IDs (subset of common effects)
    _ANIM_EFFECTS: dict[str, int] = {
        "appear": 1,
        "fly": 2,
        "blinds": 3,
        "box": 4,
        "checkerboard": 6,
        "dissolve": 12,
        "fade": 13,
        "flash_once": 14,
        "peek": 18,
        "random_bars": 21,
        "spiral": 22,
        "split": 23,
        "stretch": 25,
        "strips": 26,
        "swivel": 27,
        "wipe": 28,
        "zoom": 31,
        "random_effects": 32,
        "spin": 33,
        "grow_shrink": 38,
        "float": 41,
    }

    _ANIM_TRIGGERS: dict[str, int] = {
        "on_click": 1,       # msoAnimTriggerOnPageClick
        "with_previous": 2,  # msoAnimTriggerWithPrevious
        "after_previous": 3, # msoAnimTriggerAfterPrevious
    }

    def _get_slide(self, slide_index: int = 0):
        """Get a slide object by index, 0=active slide."""
        if slide_index == 0:
            try:
                return self.app.ActiveWindow.View.Slide
            except Exception:
                raise RuntimeError("No active slide view. Use slide_slide go_to first.")
        return self._get_slides()(slide_index)

    def add_animation(
        self,
        name_or_index: Any,
        effect_type: str = "fade",
        trigger: str = "on_click",
        duration: float = 1.0,
        delay: float = 0.0,
        slide_index: int = 0,
    ) -> None:
        """
        Add an animation effect to a shape.

        Args:
            name_or_index: Shape name or 1-based index.
            effect_type: Animation effect (see _ANIM_EFFECTS keys).
            trigger: 'on_click', 'with_previous', or 'after_previous'.
            duration: Effect duration in seconds.
            delay: Delay before effect starts in seconds.
            slide_index: 1-based slide index, 0=active slide.
        """
        shapes = self._get_slide_shapes(slide_index)
        shape = shapes(name_or_index)
        slide = self._get_slide(slide_index)

        effect_id = self._ANIM_EFFECTS.get(effect_type.lower(), 13)  # default: fade
        trigger_id = self._ANIM_TRIGGERS.get(trigger.lower(), 1)

        sequence = slide.TimeLine.MainSequence
        effect = sequence.AddEffect(Shape=shape, effectId=effect_id, trigger=trigger_id)
        if effect is not None:
            try:
                effect.Timing.Duration = duration
            except Exception:
                pass
            try:
                effect.Timing.TriggerDelayTime = delay
            except Exception:
                pass

    def clear_animations(self, name_or_index: Any = 0, slide_index: int = 0) -> None:
        """
        Clear all animations from a shape or the entire slide.

        Args:
            name_or_index: Shape name/index to clear, or 0 to clear all on slide.
            slide_index: 1-based slide index, 0=active slide.
        """
        slide = self._get_slide(slide_index)
        sequence = slide.TimeLine.MainSequence
        if name_or_index == 0:
            # Clear all animations on the slide by removing from end
            for _ in range(sequence.Count, 0, -1):
                try:
                    sequence(1).Delete()
                except Exception:
                    pass
        else:
            # Clear animations for a specific shape
            shapes = self._get_slide_shapes(slide_index)
            shape = shapes(name_or_index)
            for i in range(sequence.Count, 0, -1):
                try:
                    if sequence(i).Shape.Name == shape.Name:
                        sequence(i).Delete()
                except Exception:
                    pass

    # ── Find ────────────────────────────────────────────────────────

    def find_text(
        self,
        search_text: str,
        match_case: bool = False,
        match_whole_word: bool = False,
    ) -> list[dict[str, Any]]:
        """
        Find text across all slides.

        Args:
            search_text: Text to search for.
            match_case: If True, case-sensitive search.
            match_whole_word: If True, whole-word only.

        Returns:
            List of matches with slide_index, shape_name, text snippet.
        """
        results: list[dict[str, Any]] = []
        search_lower = search_text if match_case else search_text.lower()
        slides = self._get_slides()
        for i in range(1, slides.Count + 1):
            slide = slides(i)
            for j in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(j)
                if not shape.HasTextFrame:
                    continue
                try:
                    text = shape.TextFrame.TextRange.Text
                except Exception:
                    continue
                compare = text if match_case else text.lower()
                if match_whole_word:
                    import re
                    pattern = r'\b' + re.escape(search_text) + r'\b'
                    flags = 0 if match_case else re.IGNORECASE
                    if re.search(pattern, text, flags):
                        results.append({
                            "slide_index": i,
                            "shape_name": shape.Name,
                            "text": text[:200],
                        })
                else:
                    if search_lower in compare:
                        results.append({
                            "slide_index": i,
                            "shape_name": shape.Name,
                            "text": text[:200],
                        })
        return results

    def find_replace(
        self,
        find_text: str,
        replace_text: str,
        match_case: bool = False,
        match_whole_word: bool = False,
        replace_all: bool = True,
    ) -> int:
        """
        Find and replace text across all slides.

        Args:
            find_text: Text to find.
            replace_text: Replacement text.
            match_case: Case-sensitive match.
            match_whole_word: Whole-word match.
            replace_all: If True, replace all occurrences. If False, replace first only.

        Returns:
            Number of replacements made.
        """
        import re
        count = 0
        slides = self._get_slides()
        for i in range(1, slides.Count + 1):
            slide = slides(i)
            for j in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(j)
                if not shape.HasTextFrame:
                    continue
                try:
                    tr = shape.TextFrame.TextRange
                    text = tr.Text
                except Exception:
                    continue
                if match_whole_word:
                    pattern = r'\b' + re.escape(find_text) + r'\b'
                    flags = 0 if match_case else re.IGNORECASE
                    new_text, n = re.subn(pattern, replace_text, text, count=0 if replace_all else 1)
                else:
                    if match_case:
                        n = text.count(find_text) if replace_all else (1 if find_text in text else 0)
                        new_text = text.replace(find_text, replace_text, -1 if replace_all else 1)
                    else:
                        pattern = re.escape(find_text)
                        new_text, n = re.subn(pattern, replace_text, text, count=0 if replace_all else 1, flags=re.IGNORECASE)
                if n > 0:
                    tr.Text = new_text
                    count += n
                    if not replace_all:
                        return count
        return count

    # ── Slide Master ────────────────────────────────────────────────

    def get_master_info(self) -> dict[str, Any]:
        """
        Get slide master and layout information.

        Returns:
            Dict with master info and available layouts.
        """
        pres = self.active_presentation
        if pres is None:
            raise RuntimeError("No active presentation.")
        result: dict[str, Any] = {
            "slide_master_name": "",
            "layouts": [],
        }
        try:
            master = pres.SlideMaster
            result["slide_master_name"] = master.Name
            if hasattr(master, "CustomLayouts"):
                for i in range(1, master.CustomLayouts.Count + 1):
                    layout = master.CustomLayouts(i)
                    result["layouts"].append({
                        "index": i,
                        "name": layout.Name,
                    })
        except Exception:
            pass
        # Try Design approach for older WPS versions
        if not result["layouts"]:
            try:
                for i in range(1, pres.Designs.Count + 1):
                    design = pres.Designs(i)
                    result["layouts"].append({
                        "index": i,
                        "name": design.Name,
                    })
            except Exception:
                pass
        return result

    def apply_layout(self, slide_index: int, layout_index: int) -> None:
        """
        Apply a layout to an existing slide.

        Args:
            slide_index: 1-based slide index.
            layout_index: 1-based layout index (from get_master_info).
        """
        slide = self._get_slides()(slide_index)
        pres = self.active_presentation
        if pres is None:
            raise RuntimeError("No active presentation.")
        try:
            layout = pres.SlideMaster.CustomLayouts(layout_index)
            slide.Layout = layout
        except Exception:
            try:
                design = pres.Designs(layout_index)
                slide.Design = design
            except Exception as e:
                raise RuntimeError(f"Could not apply layout {layout_index}: {e}")

    def set_master_background(self, color_rgb: int, master_index: int = 1) -> None:
        """
        Set the slide master background color (affects all slides).

        Args:
            color_rgb: RGB color value.
            master_index: 1-based master index (usually 1).
        """
        pres = self.active_presentation
        if pres is None:
            raise RuntimeError("No active presentation.")
        try:
            master = pres.SlideMaster
            master.Background.Fill.ForeColor.RGB = color_rgb
            master.Background.Fill.Solid()
        except Exception as e:
            raise RuntimeError(f"Could not set master background: {e}")

    # ── Insert (Headers / Footers / Slide Numbers) ─────────────────

    def insert_slide_number(self) -> None:
        """Insert slide numbers on all slides."""
        pres = self.active_presentation
        if pres is None:
            raise RuntimeError("No active presentation.")
        slides = self._get_slides()
        for i in range(1, slides.Count + 1):
            slide = slides(i)
            try:
                slide.HeadersFooters.SlideNumber.Visible = -1  # msoTrue
            except Exception:
                pass

    def insert_date_time(self) -> None:
        """Insert date/time on all slides."""
        pres = self.active_presentation
        if pres is None:
            raise RuntimeError("No active presentation.")
        slides = self._get_slides()
        for i in range(1, slides.Count + 1):
            slide = slides(i)
            try:
                hf = slide.HeadersFooters
                hf.DateAndTime.Visible = -1
                hf.DateAndTime.UseFormat = True
                hf.DateAndTime.Format = 0  # ppDateTimeFormatDatePresent
            except Exception:
                pass

    def insert_header_footer(
        self,
        header_text: str = "",
        footer_text: str = "",
        slide_number: bool = True,
        date_time: bool = False,
    ) -> None:
        """
        Set header/footer text on all slides.

        Args:
            header_text: Header text (appears on notes/handouts only in most views).
            footer_text: Footer text.
            slide_number: Whether to show slide numbers.
            date_time: Whether to show date/time.
        """
        pres = self.active_presentation
        if pres is None:
            raise RuntimeError("No active presentation.")
        slides = self._get_slides()
        for i in range(1, slides.Count + 1):
            slide = slides(i)
            try:
                hf = slide.HeadersFooters
                if footer_text:
                    hf.Footer.Visible = -1
                    hf.Footer.Text = footer_text
                if slide_number:
                    hf.SlideNumber.Visible = -1
                if date_time:
                    hf.DateAndTime.Visible = -1
            except Exception:
                pass

    # ── Chart ───────────────────────────────────────────────────────

    _CHART_TYPES: dict[str, int] = {
        "column": 51,      # xlColumnClustered
        "column_stacked": 52,
        "line": 4,         # xlLine
        "line_markers": 65,
        "pie": 5,          # xlPie
        "bar": 57,         # xlBarClustered
        "bar_stacked": 58,
        "area": 1,         # xlArea
        "scatter": -4169,  # xlXYScatter
    }

    def add_chart(
        self,
        chart_type: str = "column",
        left: float = 50,
        top: float = 50,
        width: float = 600,
        height: float = 400,
        slide_index: int = 0,
    ) -> str:
        """
        Add a chart to a slide.

        Args:
            chart_type: 'column', 'line', 'pie', 'bar', 'area', 'scatter', etc.
            left, top, width, height: Position and size in points.
            slide_index: 1-based slide index, 0=active slide.

        Returns:
            Name of the new chart shape.
        """
        shapes = self._get_slide_shapes(slide_index)
        ct = self._CHART_TYPES.get(chart_type.lower(), 51)
        # Try named param first (newer PPT), fallback to positional
        try:
            shape = shapes.AddChart(Type=ct, Left=left, Top=top, Width=width, Height=height)
        except Exception:
            shape = shapes.AddChart(ct, left, top, width, height)
        return shape.Name

    def set_chart_data(
        self,
        shape_name: Any,
        categories: list[str],
        series_list: list[dict[str, Any]],
        slide_index: int = 0,
    ) -> None:
        """
        Set chart data.

        Args:
            shape_name: Chart shape name or 1-based index.
            categories: List of category labels (e.g., ['Q1', 'Q2', 'Q3', 'Q4']).
            series_list: List of dicts, each with 'name' (str) and 'values' (list of numbers).
                Example: [{"name": "Sales", "values": [100, 200, 150, 300]}]
            slide_index: 1-based slide index, 0=active slide.
        """
        shapes = self._get_slide_shapes(slide_index)
        shape = shapes(shape_name)
        chart = shape.Chart
        chart_data = chart.ChartData

        # Determine data dimensions
        num_categories = len(categories)
        num_series = len(series_list)

        # Write categories in row 1 starting at column 2
        for c_idx, cat in enumerate(categories):
            chart_data.Workbook.Worksheets(1).Cells(1, c_idx + 2).Value = cat

        # Write series names in column 1 starting at row 2, and values
        for s_idx, series in enumerate(series_list):
            row = s_idx + 2
            chart_data.Workbook.Worksheets(1).Cells(row, 1).Value = series.get("name", f"Series {s_idx + 1}")
            for v_idx, val in enumerate(series.get("values", [])):
                chart_data.Workbook.Worksheets(1).Cells(row, v_idx + 2).Value = val

        # Apply chart data changes; some PPT versions need RefreshData
        try:
            chart.Apply()
        except Exception:
            try:
                chart.Refresh()
            except Exception:
                pass

    # ── Media (Video / Audio) ───────────────────────────────────────

    def add_video(
        self,
        filepath: str,
        left: float = 50,
        top: float = 50,
        width: float = 640,
        height: float = 480,
        slide_index: int = 0,
    ) -> str:
        """
        Add a video to a slide.

        Args:
            filepath: Full path to video file (.mp4, .wmv, .avi, etc.).
            left, top, width, height: Position and size in points.
            slide_index: 1-based slide index, 0=active slide.

        Returns:
            Name of the new media shape.
        """
        full_path = os.path.abspath(filepath)
        if not os.path.exists(full_path):
            raise FileNotFoundError(f"Video file not found: {full_path}")
        shapes = self._get_slide_shapes(slide_index)
        shape = shapes.AddMediaObject2(
            FileName=full_path,
            LinkToFile=False,
            SaveWithDocument=True,
            Left=left,
            Top=top,
            Width=width,
            Height=height,
        )
        return shape.Name

    def add_audio(
        self,
        filepath: str,
        slide_index: int = 0,
    ) -> str:
        """
        Add audio to a slide.

        Args:
            filepath: Full path to audio file (.mp3, .wav, .wma, etc.).
            slide_index: 1-based slide index, 0=active slide.

        Returns:
            Name of the new media shape.
        """
        full_path = os.path.abspath(filepath)
        if not os.path.exists(full_path):
            raise FileNotFoundError(f"Audio file not found: {full_path}")
        shapes = self._get_slide_shapes(slide_index)
        shape = shapes.AddMediaObject2(
            FileName=full_path,
            LinkToFile=False,
            SaveWithDocument=True,
            Left=0,
            Top=0,
            Width=50,
            Height=50,
        )
        return shape.Name
