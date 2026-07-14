from typing import Any

from .page_setup import PageMargins, PrintOptions, PageSetup


class WorksheetPageMixin:
    # Provided by Worksheet via mixin composition (for type checkers).
    _sheet: Any
    _workbook: Any
    _closed: bool
    max_row: int
    max_column: int

    @property
    def page_margins(self):
        """
        Get the PageMargins object for this worksheet.
        """
        return PageMargins(self._sheet.page_margins(), self)

    @property
    def print_options(self):
        """
        Get the PrintOptions object for this worksheet.
        """
        return PrintOptions(self._sheet.print_options(), self)

    @property
    def page_setup(self):
        """
        Get the PageSetup object for this worksheet.
        """
        return PageSetup(self._sheet.page_setup(), self)

    def set_print_area(self, sqref: str):
        """Set the print area for the worksheet."""
        self._sheet.set_print_area(sqref)

    def set_print_title_rows(self, first_row: int, last_row: int):
        """Set the rows to repeat at top on printed pages."""
        self._sheet.set_print_title_rows(first_row, last_row)

    def set_print_title_cols(self, first_col: int, last_col: int):
        """Set the columns to repeat at left on printed pages."""
        self._sheet.set_print_title_cols(first_col, last_col)
