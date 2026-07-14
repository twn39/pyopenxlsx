"""Shared cell-value coercion for Cell and bulk write paths.

Keeps date/datetime serialisation and auto number-format application
consistent across:

- ``Cell.value``
- ``Worksheet.set_cell_value``
- ``Worksheet.write_row`` / ``write_rows`` / ``set_cells`` / ``append_row``

C++ ``CellData::from_python`` also converts date/datetime to Excel serials;
Python-side conversion here is still required so auto date styles can be
applied before the native write (and so paths stay behaviourally aligned).
"""

from __future__ import annotations

from datetime import date, datetime, timedelta
from typing import Any, Iterable, List, Optional, Sequence, Tuple, Union

# date_kind markers returned by coerce helpers
DATE_KIND_DATE = "date"
DATE_KIND_DATETIME = "datetime"

DateKind = Optional[str]  # None | DATE_KIND_DATE | DATE_KIND_DATETIME
Coerced = Tuple[Any, DateKind]


def datetime_to_serial(val: Union[date, datetime]) -> float:
    """Convert ``date`` / ``datetime`` to Excel serial (days since 1899-12-30)."""
    if isinstance(val, date) and not isinstance(val, datetime):
        val = datetime.combine(val, datetime.min.time())
    delta = val - datetime(1899, 12, 30)
    return delta.total_seconds() / 86400.0


def serial_to_datetime(serial: float) -> datetime:
    """Convert Excel serial to ``datetime``."""
    return datetime(1899, 12, 30) + timedelta(days=serial)


def coerce_cell_value(value: Any) -> Coerced:
    """Coerce a single Python value for storage.

    Returns ``(storage_value, date_kind)`` where ``date_kind`` is
    ``\"date\"``, ``\"datetime\"``, or ``None`` if not a date type.
    """
    if isinstance(value, datetime):
        return datetime_to_serial(value), DATE_KIND_DATETIME
    if isinstance(value, date):
        return datetime_to_serial(value), DATE_KIND_DATE
    return value, None


def coerce_row_values(values: Sequence[Any]) -> Tuple[List[Any], List[Tuple[int, str]]]:
    """Coerce a 1D row.

    Returns ``(coerced_row, date_hits)`` where each hit is
    ``(0-based column offset, date_kind)``.
    """
    out: List[Any] = []
    hits: List[Tuple[int, str]] = []
    for i, val in enumerate(values):
        coerced, kind = coerce_cell_value(val)
        out.append(coerced)
        if kind is not None:
            hits.append((i, kind))
    return out, hits


def coerce_rows_data(
    data: Sequence[Sequence[Any]],
) -> Tuple[List[List[Any]], List[Tuple[int, int, str]]]:
    """Coerce a 2D grid.

    Returns ``(coerced_rows, date_hits)`` where each hit is
    ``(0-based row offset, 0-based col offset, date_kind)``.
    """
    out: List[List[Any]] = []
    hits: List[Tuple[int, int, str]] = []
    for r, row in enumerate(data):
        coerced_row, row_hits = coerce_row_values(row)
        out.append(coerced_row)
        for c, kind in row_hits:
            hits.append((r, c, kind))
    return out, hits


def workbook_wants_auto_date(workbook: Any) -> bool:
    return workbook is not None and bool(getattr(workbook, "auto_date_formats", False))


def style_is_date_format(workbook: Any, style_index: int) -> bool:
    """Whether *style_index* is a date/time number format (uses workbook cache)."""
    if workbook is None or style_index is None or style_index < 0:
        return False

    cache = getattr(workbook, "_date_format_cache", None)
    if cache is not None and style_index in cache:
        return cache[style_index]

    # Lazy import to avoid circular imports at module load.
    from .styles import is_date_format

    try:
        styles = workbook.styles
        cfs = styles.cell_formats()
        if style_index >= cfs.count():
            if cache is not None:
                cache[style_index] = False
            return False

        cf = cfs.cell_format_by_index(style_index)
        nf_id = cf.number_format_id()
        if is_date_format(nf_id):
            if cache is not None:
                cache[style_index] = True
            return True

        nfs = styles.number_formats()
        try:
            nf = nfs.number_format_by_id(nf_id)
            if nf:
                res = is_date_format(nf.format_code())
                if cache is not None:
                    cache[style_index] = res
                return res
        except Exception:
            pass
    except Exception:
        pass

    if cache is not None:
        cache[style_index] = False
    return False


def apply_auto_date_style(
    workbook: Any,
    sheet: Any,
    row: int,
    column: int,
    date_kind: str,
    *,
    current_style_index: Optional[int] = None,
) -> None:
    """Apply auto date/datetime number format when enabled and not already date.

    :param current_style_index: If known, skip a raw cell format lookup.
    """
    if not workbook_wants_auto_date(workbook) or date_kind is None:
        return

    if current_style_index is None:
        try:
            current_style_index = sheet.cell(row, column).cell_format()
        except Exception:
            current_style_index = 0

    if style_is_date_format(workbook, current_style_index):
        return

    is_datetime = date_kind == DATE_KIND_DATETIME
    style_idx = workbook._get_auto_date_style(is_datetime=is_datetime)
    sheet.cell(row, column).set_cell_format(style_idx)


def apply_auto_date_styles_batch(
    workbook: Any,
    sheet: Any,
    hits: Iterable[Tuple[int, int, str]],
    *,
    start_row: int = 1,
    start_col: int = 1,
) -> None:
    """Apply auto date styles for relative ``(r_off, c_off, kind)`` hits."""
    if not workbook_wants_auto_date(workbook):
        return

    date_style = None
    datetime_style = None

    for r_off, c_off, kind in hits:
        row = start_row + r_off
        col = start_col + c_off
        try:
            raw = sheet.cell(row, col)
            cur = raw.cell_format()
        except Exception:
            continue

        if style_is_date_format(workbook, cur):
            continue

        if kind == DATE_KIND_DATETIME:
            if datetime_style is None:
                datetime_style = workbook._get_auto_date_style(is_datetime=True)
            raw.set_cell_format(datetime_style)
        else:
            if date_style is None:
                date_style = workbook._get_auto_date_style(is_datetime=False)
            raw.set_cell_format(date_style)
