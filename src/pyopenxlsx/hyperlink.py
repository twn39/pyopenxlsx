"""Helpers for worksheet hyperlinks (external and internal)."""

from __future__ import annotations

from typing import Any, Optional
from urllib.parse import urlparse


def is_external_url(target: str) -> bool:
    """Return True if *target* looks like an external URL or file link."""
    if not target:
        return False
    lower = target.lower()
    if lower.startswith(("http://", "https://", "mailto:", "ftp://", "file:")):
        return True
    # Windows / UNC paths often used as hyperlinks
    if lower.startswith("\\\\") or (len(target) > 2 and target[1] == ":"):
        return True
    parsed = urlparse(target)
    return bool(parsed.scheme)


def link(
    worksheet: Any,
    cell_ref: str,
    target: str,
    *,
    text: Optional[str] = None,
    tooltip: str = "",
    internal: Optional[bool] = None,
) -> None:
    """Add a hyperlink and optionally set the cell display text.

    :param worksheet: High-level ``Worksheet``.
    :param cell_ref: Anchor cell (e.g. ``\"A1\"``).
    :param target: URL or internal location (``\"Sheet2!A1\"``).
    :param text: If given, write this as the cell value.
    :param tooltip: Optional tooltip string.
    :param internal: Force internal vs external. When ``None``, infer from
        *target* (URLs → external, otherwise internal).
    """
    if text is not None:
        worksheet[cell_ref].value = text

    use_internal = internal if internal is not None else not is_external_url(target)
    if use_internal:
        worksheet.add_internal_hyperlink(cell_ref, target, tooltip)
    else:
        worksheet.add_hyperlink(cell_ref, target, tooltip)


def unlink(worksheet: Any, cell_ref: str) -> None:
    """Remove a hyperlink from *cell_ref* if present."""
    worksheet.remove_hyperlink(cell_ref)
