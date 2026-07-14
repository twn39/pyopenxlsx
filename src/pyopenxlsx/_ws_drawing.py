from typing import Any

import asyncio


class WorksheetDrawingMixin:
    # Provided by Worksheet via mixin composition (for type checkers).
    _sheet: Any
    _workbook: Any
    _closed: bool
    max_row: int
    max_column: int

    def add_shape(self, row=1, col=1, shape_type="Rectangle", **kwargs):
        """
        Add a vector shape to the worksheet.

        Args:
            row (int): The 1-based row index to place the top-left corner.
            col (int): The 1-based column index to place the top-left corner.
            shape_type (str): The type of the shape (e.g., "Rectangle", "Ellipse", "Arrow").
            **kwargs: Shape options including:
                - name (str): Shape name.
                - text (str): Text inside the shape.
                - fill_color (str): ARGB fill color.
                - line_color (str): ARGB line color.
                - line_width (float): Line width.
                - width (int): Width in pixels.
                - height (int): Height in pixels.
                - offset_x (int): Offset X in pixels.
                - offset_y (int): Offset Y in pixels.
                - end_row (int): End row for two-cell anchor.
                - end_col (int): End column for two-cell anchor.
                - end_offset_x (int): End offset X.
                - end_offset_y (int): End offset Y.
                - rotation (int): Rotation in degrees.
                - flip_h (bool): Flip horizontally.
                - flip_v (bool): Flip vertically.
                - line_dash (str): Line dash style (e.g., "dash", "sysDash").
                - arrow_start (str): Arrow start style.
                - arrow_end (str): Arrow end style.
                - horz_align (str): Horizontal text alignment ("l", "ctr", "r").
                - vert_align (str): Vertical text alignment ("t", "ctr", "b").
        """
        from . import _openxlsx

        options = _openxlsx.XLVectorShapeOptions()

        # Resolve shape type enum
        shape_enum = getattr(_openxlsx.XLVectorShapeType, shape_type, None)
        if shape_enum is None:
            raise ValueError(f"Unknown shape type: {shape_type}")
        options.type = shape_enum

        for k, v in kwargs.items():
            if hasattr(options, k):
                setattr(options, k, v)
            else:
                raise ValueError(f"Unknown shape option: {k}")

        drawing = self._sheet.drawing()
        drawing.add_shape(row, col, options)

    def add_image(self, img_path, anchor="A1", width=None, height=None):
        """
        Add an image to the worksheet.

        :param img_path: Path to the image file.
        :param anchor: Cell reference for the top-left corner of the image (e.g., 'A1').
        :param width: Width of the image in pixels. If None, it will try to get it from the image.
        :param height: Height of the image in pixels. If None, it will try to get it from the image.
        """
        from pathlib import Path

        img_path = Path(img_path)
        if not img_path.exists():
            raise FileNotFoundError(f"Image file not found: {img_path}")

        extension = img_path.suffix.lower().lstrip(".")
        if extension not in ["png", "jpg", "jpeg", "gif"]:
            raise ValueError(f"Unsupported image format: {extension}")

        # Normalize extension for OOXML
        if extension == "jpeg":
            extension = "jpg"

        with open(img_path, "rb") as f:
            img_data = f.read()

        if width is None or height is None:
            try:
                from PIL import Image

                with Image.open(img_path) as img:
                    w, h = img.size
                    if width is None and height is None:
                        width = w
                        height = h
                    elif width is not None and height is None:
                        height = int(h * (width / w))
                    elif width is None and height is not None:
                        width = int(w * (height / h))
            except ImportError:
                if width is None or height is None:
                    raise ImportError(
                        "Pillow is required to automatically detect image dimensions. "
                        "Please install it or provide width and height manually."
                    )

        # Parse anchor
        from ._openxlsx import XLCellReference

        ref = XLCellReference(anchor)

        if width is None or height is None:
            raise ValueError("Width and height must be provided or detected.")

        self._sheet.add_image(
            img_data, extension, ref.row(), ref.column(), int(width), int(height)
        )

    async def add_image_async(self, img_path, anchor="A1", width=None, height=None):
        await asyncio.to_thread(self.add_image, img_path, anchor, width, height)

    def add_hyperlink(self, cell_ref, url, tooltip=""):
        """
        Add an external hyperlink to a cell.

        :param cell_ref: Cell reference (e.g., 'A1').
        :param url: URL of the hyperlink.
        :param tooltip: Optional tooltip text.
        """
        self._sheet.add_hyperlink(cell_ref, url, tooltip)

    def add_internal_hyperlink(self, cell_ref, location, tooltip=""):
        """
        Add an internal hyperlink (to another sheet or range) to a cell.

        :param cell_ref: Cell reference (e.g., 'A1').
        :param location: Destination in the workbook (e.g., 'Sheet2!A1').
        :param tooltip: Optional tooltip text.
        """
        self._sheet.add_internal_hyperlink(cell_ref, location, tooltip)

    def link(
        self,
        cell_ref,
        target,
        *,
        text=None,
        tooltip="",
        internal=None,
    ):
        """Add a hyperlink with optional display text (external or internal).

        See :func:`pyopenxlsx.hyperlink.link`. When *internal* is omitted,
        URLs are treated as external and other targets as internal locations.
        """
        from .hyperlink import link as link_cell

        link_cell(
            self,
            cell_ref,
            target,
            text=text,
            tooltip=tooltip,
            internal=internal,
        )

    def has_hyperlink(self, cell_ref):
        """Check if a cell has a hyperlink."""
        return self._sheet.has_hyperlink(cell_ref)

    def get_hyperlink(self, cell_ref):
        """Get the hyperlink target for a cell."""
        return self._sheet.get_hyperlink(cell_ref)

    def remove_hyperlink(self, cell_ref):
        """Remove a hyperlink from a cell."""
        self._sheet.remove_hyperlink(cell_ref)

    def images(self):
        return self._sheet.images()

    def insert_image_bytes(self, cell_ref: str, image_data: bytes, options=None):
        """Insert an image from raw bytes at the given cell reference."""
        if options is None:
            self._sheet.insert_image_bytes(cell_ref, image_data)
        else:
            self._sheet.insert_image_bytes(cell_ref, image_data, options)
