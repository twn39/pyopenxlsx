import asyncio
from weakref import ref as weakref

from .cell import Cell


class Range:
    """
    Lightweight view over a rectangular block of cells.

    Architectural role
    ------------------
    ``Range`` is a **geometry + iteration** façade: address, dimensions,
    cell iteration, and clear. Prefer delegating multi-cell style, bulk
    write, drawing, or workbook I/O to ``Worksheet`` / bulk APIs rather
    than growing kitchen-sink methods on this type.

    Uses a weak reference to the parent worksheet to avoid cycles.
    """

    __slots__ = ("_range", "_worksheet_ref")

    def __init__(self, raw_range, worksheet=None):
        self._range = raw_range
        # Use weak reference to avoid circular references
        self._worksheet_ref = weakref(worksheet) if worksheet else None

    @property
    def _worksheet(self):
        """Get the worksheet, or None if it has been garbage collected."""
        return self._worksheet_ref() if self._worksheet_ref else None

    def __iter__(self):
        for c in self._range:
            if self._worksheet:
                yield self._worksheet._get_cached_cell(c)
            else:
                yield Cell(c)

    @property
    def address(self):
        return self._range.address()

    @property
    def num_rows(self):
        return self._range.num_rows()

    @property
    def num_columns(self):
        return self._range.num_columns()

    def clear(self):
        self._range.clear()

    async def clear_async(self):
        await asyncio.to_thread(self.clear)
