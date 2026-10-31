"""High-level stream I/O wrappers with shared value coercion.

Native ``XLStreamWriter`` already accepts many Python types via C++
``CellData::from_python``. This wrapper additionally converts
``date`` / ``datetime`` to Excel serials *and* records style indices when
``Workbook.auto_date_formats`` is enabled (styles applied as ``(value, style)``
tuples understood by the stream writer).
"""

from __future__ import annotations

from typing import Any, List, Optional, Sequence, Union


class StreamWriter:
    """Thin façade over the native worksheet stream writer."""

    __slots__ = ("_writer", "_workbook", "_has_started_rows")

    def __init__(self, native_writer: Any, workbook: Any = None):
        self._writer = native_writer
        self._workbook = workbook
        self._has_started_rows = False

    def __enter__(self) -> "StreamWriter":
        self._writer.__enter__()
        return self

    def __exit__(self, *args):
        return self._writer.__exit__(*args)

    @property
    def is_active(self) -> bool:
        return self._writer.is_active

    @property
    def last_row(self) -> int:
        return self._writer.last_row

    @property
    def max_column(self) -> int:
        return self._writer.max_column

    def close(self) -> None:
        self._writer.close()

    def flush(self) -> None:
        self._writer.flush()

    def _check_phase0(self, op: str) -> None:
        if self._has_started_rows:
            raise RuntimeError(
                f"StreamWriter: '{op}' must be configured before writing rows "
                "(SpreadsheetML schema requires <cols> and <sheetViews> to precede <sheetData>)."
            )

    # ── Phase 0: Column and pane configuration ────────────────────────────

    def set_column_width(
        self, min_col: int, max_col: int, width: float
    ) -> "StreamWriter":
        """Set column width for range [min_col, max_col] (1-based, inclusive)."""
        self._check_phase0("set_column_width")
        self._writer.set_col_width(min_col, max_col, float(width))
        return self

    def set_column_visible(
        self, min_col: int, max_col: int, visible: bool
    ) -> "StreamWriter":
        """Show or hide columns [min_col, max_col]."""
        self._check_phase0("set_column_visible")
        self._writer.set_col_visible(min_col, max_col, visible)
        return self

    def set_column_outline_level(self, col: int, level: int) -> "StreamWriter":
        """Set outline level for column (1-7)."""
        self._check_phase0("set_column_outline_level")
        self._writer.set_col_outline_level(col, level)
        return self

    def set_column_style(
        self, min_col: int, max_col: int, style: int
    ) -> "StreamWriter":
        """Apply style index to columns [min_col, max_col]."""
        self._check_phase0("set_column_style")
        self._writer.set_col_style(min_col, max_col, style)
        return self

    def freeze_panes(
        self, top_left_or_cols: Union[str, int], rows: Optional[int] = None
    ) -> "StreamWriter":
        """Freeze panes before row writing starts (e.g. 'B2' or cols=1, rows=1)."""
        self._check_phase0("freeze_panes")
        if isinstance(top_left_or_cols, str):
            self._writer.freeze_panes(top_left_or_cols)
        else:
            self._writer.freeze_panes(top_left_or_cols, rows or 0)
        return self

    # ── Phase 1: Row writing ──────────────────────────────────────────────

    def append_row(self, values: Sequence[Any], row_opts: Any = None) -> None:
        self._has_started_rows = True
        from . import _coercion

        coerced = self._coerce_stream_row(values, _coercion)
        if row_opts is None:
            self._writer.append_row(coerced)
        else:
            self._writer.append_row(coerced, row_opts)

    def append_rows(self, rows_iterable: Any) -> None:
        self._has_started_rows = True
        for row in rows_iterable:
            self.append_row(row)

    def set_row(
        self, row: int, start_col: int, values: Sequence[Any], row_opts: Any = None
    ) -> None:
        self._has_started_rows = True
        from . import _coercion

        coerced = self._coerce_stream_row(values, _coercion)
        if row_opts is None:
            self._writer.set_row(row, start_col, coerced)
        else:
            self._writer.set_row(row, start_col, coerced, row_opts)

    def set_row_ref(
        self, ref: str, values: Sequence[Any], row_opts: Any = None
    ) -> None:
        self._has_started_rows = True
        from . import _coercion

        coerced = self._coerce_stream_row(values, _coercion)
        if row_opts is None:
            self._writer.set_row_ref(ref, coerced)
        else:
            self._writer.set_row_ref(ref, coerced, row_opts)

    # ── Phase 2: Post-row decoration (written on close) ───────────────────

    def merge_cells(
        self, top_left: str, bottom_right: Optional[str] = None
    ) -> "StreamWriter":
        """Merge a cell range (e.g. 'A1:C1' or top_left='A1', bottom_right='C1')."""
        if bottom_right is None and ":" in top_left:
            top_left, bottom_right = top_left.split(":", 1)
        elif bottom_right is None:
            raise ValueError(
                "merge_cells requires either a range string ('A1:C1') or both top_left and bottom_right"
            )
        self._writer.merge_cells(top_left, bottom_right)
        return self

    def auto_filter(self, range_str: str) -> "StreamWriter":
        """Set AutoFilter range (e.g. 'A1:E100'). Injected on close."""
        self._writer.set_auto_filter(range_str)
        return self

    def add_hyperlink(
        self,
        cell_ref: str,
        target: str,
        tooltip: str = "",
        is_external: bool = True,
    ) -> "StreamWriter":
        """Add hyperlink to cell. Injected on close."""
        if is_external:
            self._writer.add_hyperlink(cell_ref, target, tooltip)
        else:
            self._writer.add_internal_hyperlink(cell_ref, target, tooltip)
        return self

    def insert_page_break(self, cell_ref: str) -> "StreamWriter":
        """Insert page break at cell reference (e.g. 'A10'). Injected on close."""
        self._writer.insert_page_break(cell_ref)
        return self

    def _coerce_stream_row(self, values: Sequence[Any], _coercion) -> List[Any]:
        wb = self._workbook
        wants = _coercion.workbook_wants_auto_date(wb)
        date_style = None
        datetime_style = None
        out: List[Any] = []
        for val in values:
            # Preserve explicit (value, style_id) pairs.
            if isinstance(val, tuple) and len(val) == 2 and isinstance(val[1], int):
                storage, kind = _coercion.coerce_cell_value(val[0])
                out.append((storage, val[1]))
                continue
            storage, kind = _coercion.coerce_cell_value(val)
            if kind is not None and wants:
                if kind == _coercion.DATE_KIND_DATETIME:
                    if datetime_style is None:
                        datetime_style = wb._get_auto_date_style(is_datetime=True)
                    out.append((storage, datetime_style))
                else:
                    if date_style is None:
                        date_style = wb._get_auto_date_style(is_datetime=False)
                    out.append((storage, date_style))
            else:
                out.append(storage)
        return out

    @property
    def raw(self) -> Any:
        """Access the underlying native XLStreamWriter handle."""
        return self._writer

    def is_stream_active(self) -> bool:
        """Return True if the native row stream is currently active."""
        return self._writer.is_stream_active()

    def set_auto_filter(self, range_str: str) -> "StreamWriter":
        """Alias for auto_filter()."""
        return self.auto_filter(range_str)


class StreamReader:
    """Thin façade over the native worksheet stream reader."""

    __slots__ = ("_reader",)

    def __init__(self, native_reader: Any):
        self._reader = native_reader

    @property
    def raw(self) -> Any:
        """Access the underlying native XLStreamReader handle."""
        return self._reader

    def __iter__(self):
        return self

    def __next__(self):
        return next(self._reader)

    def has_next(self) -> bool:
        """Check if there are more rows to read."""
        return self._reader.has_next()

    def next_row(self) -> List[Any]:
        """Read the next row of cell values."""
        return self._reader.next_row()

    def next_row_detailed(self) -> List[dict]:
        """Read the next row with detailed cell attributes (formula, style_index, etc.)."""
        return self._reader.next_row_detailed()

    def next_row_sparse(self) -> List[dict]:
        """Read the next row with only non-empty cells (as dicts with column, value, etc.)."""
        return self._reader.next_row_sparse()

    def next_row_strings(self) -> List[str]:
        """Read the next row as display strings (formatting applied if configured)."""
        return self._reader.next_row_strings()

    def current_row(self) -> int:
        """Get 1-based index of the last returned row."""
        return self._reader.current_row()

    @property
    def current_row_index(self) -> int:
        """Property alias for current_row()."""
        return self._reader.current_row_index

    def current_row_opts(self) -> dict:
        """Row attributes (height, hidden, outline_level, etc.) for the last returned row."""
        return self._reader.current_row_opts()

    @property
    def last_error(self) -> Optional[str]:
        """Last parse error or None."""
        return self._reader.last_error

    @property
    def options(self) -> Any:
        """Read options configured on this stream reader."""
        return self._reader.options

    def close(self) -> None:
        """Close the stream reader."""
        self._reader.close()

    def iter_sparse(self):
        """Iterate yielding non-empty cells as (col_index, value) tuples for zero-overhead ETL."""
        while self._reader.has_next():
            cells = self._reader.next_row_sparse()
            yield [(c["column"], c["value"]) for c in cells]

    def __enter__(self) -> "StreamReader":
        if hasattr(self._reader, "__enter__"):
            self._reader.__enter__()
        return self

    def __exit__(self, *args):
        if hasattr(self._reader, "__exit__"):
            return self._reader.__exit__(*args)
        return None
