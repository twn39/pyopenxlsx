"""High-level stream I/O wrappers with shared value coercion.

Native ``XLStreamWriter`` already accepts many Python types via C++
``CellData::from_python``. This wrapper additionally converts
``date`` / ``datetime`` to Excel serials *and* records style indices when
``Workbook.auto_date_formats`` is enabled (styles applied as ``(value, style)``
tuples understood by the stream writer).
"""

from __future__ import annotations

from typing import Any, List, Sequence


class StreamWriter:
    """Thin façade over the native worksheet stream writer."""

    __slots__ = ("_writer", "_workbook")

    def __init__(self, native_writer: Any, workbook: Any = None):
        self._writer = native_writer
        self._workbook = workbook

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

    def append_row(self, values: Sequence[Any], row_opts: Any = None) -> None:
        from . import _coercion

        coerced = self._coerce_stream_row(values, _coercion)
        if row_opts is None:
            self._writer.append_row(coerced)
        else:
            self._writer.append_row(coerced, row_opts)

    def set_row(
        self, row: int, start_col: int, values: Sequence[Any], row_opts: Any = None
    ) -> None:
        from . import _coercion

        coerced = self._coerce_stream_row(values, _coercion)
        if row_opts is None:
            self._writer.set_row(row, start_col, coerced)
        else:
            self._writer.set_row(row, start_col, coerced, row_opts)

    def set_row_ref(
        self, ref: str, values: Sequence[Any], row_opts: Any = None
    ) -> None:
        from . import _coercion

        coerced = self._coerce_stream_row(values, _coercion)
        if row_opts is None:
            self._writer.set_row_ref(ref, coerced)
        else:
            self._writer.set_row_ref(ref, coerced, row_opts)

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

    def __getattr__(self, name: str) -> Any:
        return getattr(self._writer, name)


class StreamReader:
    """Thin façade over the native worksheet stream reader."""

    __slots__ = ("_reader",)

    def __init__(self, native_reader: Any):
        self._reader = native_reader

    def __iter__(self):
        return self

    def __next__(self):
        return next(self._reader)

    def __enter__(self) -> "StreamReader":
        if hasattr(self._reader, "__enter__"):
            self._reader.__enter__()
        return self

    def __exit__(self, *args):
        if hasattr(self._reader, "__exit__"):
            return self._reader.__exit__(*args)
        return None

    def __getattr__(self, name: str) -> Any:
        return getattr(self._reader, name)
