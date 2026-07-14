import asyncio
import tempfile
import os
from weakref import WeakValueDictionary

from . import _openxlsx
from .properties import CustomProperties, DocumentProperties
from .worksheet import Worksheet

# Re-export for historical ``from pyopenxlsx.workbook import DocumentProperties``.
__all__ = [
    "Workbook",
    "DocumentProperties",
    "CustomProperties",
    "load_workbook",
    "load_workbook_async",
]


class Workbook:
    """
    Represents an Excel workbook.

    Uses WeakValueDictionary for worksheet caching to allow garbage collection
    of Worksheet objects when they are no longer referenced elsewhere.
    """

    def __init__(self, filename=None, force_overwrite=True, password=None):
        self._doc = _openxlsx.XLDocument()
        self._temp_file = None  # Track temp file for cleanup
        if filename:
            if password is not None:
                self._doc.open(str(filename), password)
            else:
                self._doc.open(str(filename))
            self._filename = str(filename)
        else:
            # Use a temporary file to avoid polluting the current directory
            # OpenXLSX's create() writes to disk immediately
            fd, temp_path = tempfile.mkstemp(suffix=".xlsx", prefix="pyopenxlsx_")
            os.close(fd)  # Close the file descriptor, XLDocument will manage the file
            self._temp_file = temp_path
            self._doc.create(temp_path, force_overwrite)
            self._filename = None
        self._wb = self._doc.workbook()
        # Use WeakValueDictionary to avoid keeping Worksheet objects alive indefinitely
        # Worksheets will be garbage collected when no external references remain
        self._sheets = WeakValueDictionary()
        self._styles = None
        self._date_format_cache = {}
        # Cached style indices for auto date/datetime number formats
        self._auto_date_style_idx = None
        self._auto_datetime_style_idx = None
        # When True, assigning date/datetime via Cell.value, set_cell_value,
        # write_row(s), set_cells, or append_row applies a default number format
        # if the cell is not already date-formatted.
        self.auto_date_formats = True
        self._closed = False

    @property
    def has_macro(self):
        """Check if the loaded document contains a VBA macro project."""
        return self._doc.has_macro()

    def _get_auto_date_style(self, is_datetime: bool = True) -> int:
        """Return (and cache) a style index for automatic date/datetime formats."""
        if is_datetime:
            if self._auto_datetime_style_idx is None:
                self._auto_datetime_style_idx = self.add_style(
                    number_format="yyyy-mm-dd hh:mm:ss"
                )
            return self._auto_datetime_style_idx
        if self._auto_date_style_idx is None:
            self._auto_date_style_idx = self.add_style(number_format="yyyy-mm-dd")
        return self._auto_date_style_idx

    def save(self, filename=None, force_overwrite=True, password=None):
        if filename:
            if password is not None:
                self._doc.save_as(str(filename), force_overwrite, password)
            else:
                self._doc.save_as(str(filename), force_overwrite)
        elif self._filename:
            # OpenXLSX's save() doesn't take a password, but save_as does
            if password is not None:
                self._doc.save_as(self._filename, force_overwrite, password)
            else:
                self._doc.save()
        else:
            raise ValueError("No filename specified")

    def validate_package_invariants(self):
        """Validate package-level OOXML invariants (also run automatically on save)."""
        self._doc.validate_package_invariants()

    def sheet_names(self):
        return self._wb.sheet_names()

    def worksheet_names(self):
        return self._wb.worksheet_names()

    def chartsheet_names(self):
        return self._wb.chartsheet_names()

    def add_chartsheet(self, name: str):
        self._wb.add_chartsheet(name)
        return self[name] if name in self.sheet_names() else None

    def set_sheet_index(self, name: str, index: int):
        self._wb.set_sheet_index(name, index)

    def protect(self, lock_structure=True, lock_windows=False, password=""):
        self._wb.protect(lock_structure, lock_windows, password)

    def unprotect(self):
        self._wb.unprotect()

    def is_protected(self) -> bool:
        return self._wb.is_protected()

    def set_full_calculation_on_load(self):
        self._wb.set_full_calculation_on_load()

    def cleanup_shared_strings(self):
        self._doc.cleanup_shared_strings()

    async def save_async(self, filename=None, force_overwrite=True, password=None):
        await asyncio.to_thread(self.save, filename, force_overwrite, password)

    def close(self):
        if getattr(self, "_closed", False):
            return
        self._doc.close()
        # Clean up temporary file if it was created
        if self._temp_file and os.path.exists(self._temp_file):
            try:
                os.unlink(self._temp_file)
            except OSError:
                pass  # Ignore errors during cleanup
            self._temp_file = None
        self._closed = True

    async def close_async(self):
        await asyncio.to_thread(self.close)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    async def __aenter__(self):
        return self

    async def __aexit__(self, exc_type, exc_val, exc_tb):
        await self.close_async()

    @property
    def styles(self):
        if self._styles is None:
            self._styles = self._doc.styles()
        return self._styles

    @property
    def defined_names(self):
        """
        Access defined names (named ranges) via :class:`~pyopenxlsx.defined_names.DefinedNames`.
        """
        from .defined_names import DefinedNames

        return DefinedNames(self._wb.defined_names(), self)

    @property
    def properties(self):
        if not hasattr(self, "_properties"):
            self._properties = DocumentProperties(self._doc)
        return self._properties

    @property
    def custom_properties(self):
        """
        Get the custom document properties.
        """
        if not hasattr(self, "_custom_properties"):
            self._custom_properties = CustomProperties(self._doc)
        return self._custom_properties

    def add_style(
        self,
        font=None,
        fill=None,
        border=None,
        alignment=None,
        number_format=None,
        protection=None,
    ):
        """Register a cell style and return its style index.

        Implementation lives in :mod:`pyopenxlsx._style_registry` so this
        method stays a thin public entry point.
        """
        from ._style_registry import register_cell_style

        return register_cell_style(
            self.styles,
            font=font,
            fill=fill,
            border=border,
            alignment=alignment,
            number_format=number_format,
            protection=protection,
        )

    async def add_style_async(
        self,
        font=None,
        fill=None,
        border=None,
        alignment=None,
        number_format=None,
        protection=None,
    ):
        return await asyncio.to_thread(
            self.add_style, font, fill, border, alignment, number_format, protection
        )

    @property
    def workbook(self):
        if getattr(self, "_closed", False):
            raise ValueError("I/O operation on closed Workbook/Worksheet.")
        return self._wb

    @property
    def active(self):
        try:
            names = self.sheetnames
            for name in names:
                ws = self.workbook.worksheet(name)
                if ws.is_active():
                    return self[name]  # Use cache via __getitem__
        except Exception:
            pass

        if self.sheetnames:
            return self[self.sheetnames[0]]
        return None

    @active.setter
    def active(self, ws):
        if not isinstance(ws, Worksheet):
            raise TypeError("Must be a Worksheet object")
        self.workbook.clear_active_tab()
        ws._sheet.set_active()

    def create_sheet(self, title=None, index=None):
        if title is None:
            i = 1
            while f"Sheet{i}" in self.sheetnames:
                i += 1
            title = f"Sheet{i}"
        self.workbook.add_worksheet(title)
        ws = self[title]
        if index is not None:
            ws._sheet.set_index(index + 1)
        return ws

    async def create_sheet_async(self, title=None, index=None):
        return await asyncio.to_thread(self.create_sheet, title, index)

    def remove(self, worksheet):
        self.workbook.delete_sheet(worksheet.title)

    async def remove_async(self, worksheet):
        await asyncio.to_thread(self.remove, worksheet)

    def copy_worksheet(self, from_worksheet):
        new_name = f"{from_worksheet.title} Copy"
        i = 1
        while new_name in self.sheetnames:
            new_name = f"{from_worksheet.title} Copy{i}"
            i += 1
        self.workbook.clone_sheet(from_worksheet.title, new_name)
        return self[new_name]

    async def copy_worksheet_async(self, from_worksheet):
        return await asyncio.to_thread(self.copy_worksheet, from_worksheet)

    @property
    def sheetnames(self):
        return list(self.workbook.worksheet_names())

    def __getitem__(self, key):
        if key in self._sheets:
            return self._sheets[key]
        if self.workbook.sheet_exists(key):
            ws = Worksheet(self.workbook.worksheet(key), self)
            self._sheets[key] = ws
            return ws
        raise KeyError(f"Worksheet {key} does not exist")

    def __delitem__(self, key):
        if self.workbook.sheet_exists(key):
            self.workbook.delete_sheet(key)
            if key in self._sheets:
                del self._sheets[key]
        else:
            raise KeyError(f"Worksheet {key} does not exist")

    def __iter__(self):
        for name in self.sheetnames:
            yield self[name]

    def __len__(self):
        return self.workbook.sheet_count()

    def __contains__(self, key):
        return self.workbook.sheet_exists(key)

    def get_archive_entries(self):
        """
        Get a list of all files/directories in the underlying zip archive.

        Returns:
            list[str]: List of entry paths.
        """
        return self._doc.get_archive_entries()

    def has_archive_entry(self, path):
        """
        Check if the underlying zip archive contains an entry with the given path.

        Args:
            path: Path in the archive (e.g., 'xl/workbook.xml').

        Returns:
            bool: True if the entry exists.
        """
        return self._doc.has_archive_entry(path)

    def get_archive_entry(self, path):
        """
        Get the raw bytes of an entry from the underlying zip archive.

        Args:
            path: Path in the archive (e.g., 'xl/workbook.xml').

        Returns:
            bytes: The raw binary data of the entry.

        Raises:
            RuntimeError: If the entry is not found in the archive.
        """
        return self._doc.get_archive_entry(path)

    def get_embedded_images(self):
        """
        Get a list of all embedded images in the workbook.

        Returns:
            list[ImageInfo]: List of ImageInfo objects with name, path, and extension attributes.

        Example:
            >>> wb = Workbook("test.xlsx")
            >>> images = wb.get_embedded_images()
            >>> for img in images:
            ...     print(f"Found image: {img.name}")
        """
        return self._doc.get_embedded_images()

    def get_image_data(self, image_path_or_name):
        """
        Get the binary data of an embedded image.

        Args:
            image_path_or_name: Either the full path (e.g., 'xl/media/image1.png')
                               or just the filename (e.g., 'image1.png')

        Returns:
            bytes: The raw binary data of the image.

        Raises:
            RuntimeError: If the image is not found in the archive.

        Example:
            >>> wb = Workbook("test.xlsx")
            >>> images = wb.get_embedded_images()
            >>> if images:
            ...     data = wb.get_image_data(images[0].name)
            ...     with open("extracted_image.png", "wb") as f:
            ...         f.write(data)
        """
        return self._doc.get_image_data(image_path_or_name)

    def extract_images(self, output_dir):
        """
        Extract all embedded images to a directory.

        Args:
            output_dir: Directory path where images will be saved.
                       Will be created if it doesn't exist.

        Returns:
            list[str]: List of paths to the extracted image files.

        Example:
            >>> wb = Workbook("test.xlsx")
            >>> extracted = wb.extract_images("./images/")
            >>> print(f"Extracted {len(extracted)} images")
        """
        import os

        os.makedirs(output_dir, exist_ok=True)

        images = self.get_embedded_images()
        extracted_paths = []

        for img in images:
            data = self.get_image_data(img.path)
            output_path = os.path.join(output_dir, img.name)
            with open(output_path, "wb") as f:
                f.write(data)
            extracted_paths.append(output_path)

        return extracted_paths

    async def extract_images_async(self, output_dir):
        """Async version of extract_images."""
        return await asyncio.to_thread(self.extract_images, output_dir)

    def __del__(self):
        # Ensure temporary file is cleaned up even if close() was not called
        if hasattr(self, "_temp_file") and self._temp_file:
            try:
                os.unlink(self._temp_file)
            except (OSError, FileNotFoundError):
                pass


def load_workbook(filename, password=None):
    return Workbook(filename, password=password)


async def load_workbook_async(filename, password=None):
    return await asyncio.to_thread(load_workbook, filename, password)
