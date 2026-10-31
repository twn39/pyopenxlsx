"""Internal physical document session and lifecycle coordinator.

This module encapsulates the lower-level file lifecycle, temporary scratchpad
tracking, and raw package I/O (such as ZIP entries and embedded image extraction),
keeping the domain Workbook focused on sheet orchestration, styles, and names.
"""

from __future__ import annotations

import os
import tempfile
import weakref
from typing import Any, List, Optional

from . import _openxlsx


def _cleanup_temp_file(path: Optional[str]) -> None:
    """Safely unlink a temporary scratchpad file without raising."""
    if path and os.path.exists(path):
        try:
            os.unlink(path)
        except OSError:
            pass


class _DocumentSession:
    """Manages physical document lifecycle, temporary files, and low-level I/O.

    Single Source of Truth (SSOT) for:
    - The underlying ``_openxlsx.XLDocument`` handle
    - Target file path and temporary file lifetime
    - Guaranteed scratchpad cleanup via ``weakref.finalize``
    - Authoritative ``is_closed`` state and defensive guards
    - Raw ZIP archive inspection and image extraction
    """

    __slots__ = (
        "_doc",
        "_filename",
        "_temp_file",
        "_closed",
        "_finalizer",
        "__weakref__",
    )

    def __init__(
        self,
        filename: Optional[str] = None,
        force_overwrite: bool = True,
        password: Optional[str] = None,
    ) -> None:
        self._doc = _openxlsx.XLDocument()
        self._temp_file: Optional[str] = None
        self._closed: bool = False

        if filename:
            target = str(filename)
            if password is not None:
                self._doc.open(target, password)
            else:
                self._doc.open(target)
            self._filename: Optional[str] = target
            self._finalizer = None
        else:
            fd, temp_path = tempfile.mkstemp(suffix=".xlsx", prefix="pyopenxlsx_")
            os.close(fd)
            self._temp_file = temp_path
            self._finalizer = weakref.finalize(self, _cleanup_temp_file, temp_path)
            self._doc.create(temp_path, force_overwrite)
            self._filename = None

    @property
    def doc(self) -> _openxlsx.XLDocument:
        """The underlying OpenXLSX XLDocument instance."""
        return self._doc

    @property
    def filename(self) -> Optional[str]:
        return self._filename

    @filename.setter
    def filename(self, value: Optional[str]) -> None:
        self._filename = str(value) if value is not None else None

    @property
    def temp_file(self) -> Optional[str]:
        return self._temp_file

    @property
    def is_closed(self) -> bool:
        return self._closed

    def check_closed(self) -> None:
        """Raise ValueError if the session has been closed."""
        if self._closed:
            raise ValueError("I/O operation on closed Workbook/Worksheet.")

    def save(
        self,
        filename: Optional[str] = None,
        force_overwrite: bool = True,
        password: Optional[str] = None,
    ) -> str:
        """Save the document to target or the open filename."""
        self.check_closed()
        target = str(filename) if filename is not None else self._filename
        if not target:
            raise ValueError("No filename specified")

        if filename is not None:
            if password is not None:
                self._doc.save_as(target, force_overwrite, password)
            else:
                self._doc.save_as(target, force_overwrite)
            self._filename = target
        elif self._filename:
            if password is not None:
                self._doc.save_as(self._filename, force_overwrite, password)
            else:
                self._doc.save()
        else:
            raise ValueError("No filename specified")

        return target

    def close(self) -> None:
        """Close the underlying document and detach/clean temporary files."""
        if self._closed:
            return
        try:
            self._doc.close()
        finally:
            self._closed = True
            if self._finalizer is not None:
                self._finalizer()
                self._finalizer = None
            elif self._temp_file and os.path.exists(self._temp_file):
                _cleanup_temp_file(self._temp_file)
            self._temp_file = None

    def compact_styles(self) -> None:
        self.check_closed()
        self._doc.compact_styles()

    def validate_package_invariants(self) -> None:
        self.check_closed()
        self._doc.validate_package_invariants()

    def cleanup_shared_strings(self) -> None:
        self.check_closed()
        self._doc.cleanup_shared_strings()

    def has_macro(self) -> bool:
        self.check_closed()
        return self._doc.has_macro()

    def get_archive_entries(self) -> List[str]:
        self.check_closed()
        return self._doc.get_archive_entries()

    def has_archive_entry(self, path: str) -> bool:
        self.check_closed()
        return self._doc.has_archive_entry(path)

    def get_archive_entry(self, path: str) -> bytes:
        self.check_closed()
        return self._doc.get_archive_entry(path)

    def get_embedded_images(self) -> List[Any]:
        self.check_closed()
        return self._doc.get_embedded_images()

    def get_image_data(self, image_path_or_name: str) -> bytes:
        self.check_closed()
        return self._doc.get_image_data(image_path_or_name)

    def extract_images(self, output_dir: str) -> List[str]:
        self.check_closed()
        os.makedirs(output_dir, exist_ok=True)
        images = self.get_embedded_images()
        extracted_paths: List[str] = []
        for img in images:
            data = self.get_image_data(img.path)
            output_path = os.path.join(output_dir, img.name)
            with open(output_path, "wb") as f:
                f.write(data)
            extracted_paths.append(output_path)
        return extracted_paths
