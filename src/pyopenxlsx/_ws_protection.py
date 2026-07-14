from typing import Any

import asyncio


class WorksheetProtectionMixin:
    # Provided by Worksheet via mixin composition (for type checkers).
    _sheet: Any
    _workbook: Any
    _closed: bool
    max_row: int
    max_column: int

    def protect(
        self,
        password=None,
        sheet=True,
        objects=False,
        scenarios=False,
        insert_columns=False,
        insert_rows=False,
        insert_hyperlinks=False,
        delete_columns=False,
        delete_rows=False,
        select_locked_cells=True,
        select_unlocked_cells=True,
        auto_filter=False,
        sort=False,
        pivot_tables=False,
        format_cells=False,
        format_columns=False,
        format_rows=False,
    ):
        """
        Protect the worksheet.
        """
        from . import _openxlsx

        options = _openxlsx.XLSheetProtectionOptions()
        options.sheet = sheet
        options.objects = objects
        options.scenarios = scenarios
        options.format_cells = format_cells
        options.format_columns = format_columns
        options.format_rows = format_rows
        options.insert_columns = insert_columns
        options.insert_rows = insert_rows
        options.insert_hyperlinks = insert_hyperlinks
        options.delete_columns = delete_columns
        options.delete_rows = delete_rows
        options.sort = sort
        options.auto_filter = auto_filter
        options.pivot_tables = pivot_tables
        options.select_locked_cells = select_locked_cells
        options.select_unlocked_cells = select_unlocked_cells

        return self._sheet.protect(options, password or "")

    async def protect_async(
        self,
        password=None,
        sheet=True,
        objects=False,
        scenarios=False,
        insert_columns=False,
        insert_rows=False,
        insert_hyperlinks=False,
        delete_columns=False,
        delete_rows=False,
        select_locked_cells=True,
        select_unlocked_cells=True,
        auto_filter=False,
        sort=False,
        pivot_tables=False,
        format_cells=False,
        format_columns=False,
        format_rows=False,
    ):
        return await asyncio.to_thread(
            self.protect,
            password=password,
            sheet=sheet,
            objects=objects,
            scenarios=scenarios,
            insert_columns=insert_columns,
            insert_rows=insert_rows,
            insert_hyperlinks=insert_hyperlinks,
            delete_columns=delete_columns,
            delete_rows=delete_rows,
            select_locked_cells=select_locked_cells,
            select_unlocked_cells=select_unlocked_cells,
            auto_filter=auto_filter,
            sort=sort,
            pivot_tables=pivot_tables,
            format_cells=format_cells,
            format_columns=format_columns,
            format_rows=format_rows,
        )

    def unprotect(self):
        """
        Unprotect the worksheet.
        """
        self._sheet.protect_sheet(False)
        self._sheet.clear_password()

    async def unprotect_async(self):
        await asyncio.to_thread(self.unprotect)

    @property
    def protection(self):
        """
        Get the protection status of the worksheet.
        """
        return {
            "protected": self._sheet.sheet_protected(),
            "password_set": self._sheet.password_is_set(),
            "objects": self._sheet.objects_protected(),
            "scenarios": self._sheet.scenarios_protected(),
            "insert_columns": self._sheet.insert_columns_allowed(),
            "insert_rows": self._sheet.insert_rows_allowed(),
            "insert_hyperlinks": self._sheet.insert_hyperlinks_allowed(),
            "delete_columns": self._sheet.delete_columns_allowed(),
            "delete_rows": self._sheet.delete_rows_allowed(),
            "select_locked_cells": self._sheet.select_locked_cells_allowed(),
            "select_unlocked_cells": self._sheet.select_unlocked_cells_allowed(),
            "auto_filter": self._sheet.auto_filter_allowed(),
            "sort": self._sheet.sort_allowed(),
            "pivot_tables": self._sheet.pivot_tables_allowed(),
            "format_cells": self._sheet.format_cells_allowed(),
            "format_columns": self._sheet.format_columns_allowed(),
            "format_rows": self._sheet.format_rows_allowed(),
        }

    def freeze_panes(self, row_or_ref, col=None):
        """
        Freeze the worksheet panes.

        :param row_or_ref: Row number (1-indexed) or a cell reference string (e.g., 'B2').
        :param col: Column number (1-indexed). Only used if row_or_ref is an int.
        """
        if isinstance(row_or_ref, str):
            self._sheet.freeze_panes(row_or_ref)
        elif isinstance(row_or_ref, int):
            if col is None:
                self._sheet.freeze_panes(0, row_or_ref)
            else:
                self._sheet.freeze_panes(col, row_or_ref)
        else:
            raise TypeError("row_or_ref must be an int or a string reference")

    def split_panes(
        self, x_split, y_split, top_left_cell="", active_pane="bottomRight"
    ):
        """
        Split the worksheet panes at given pixel coordinates.

        :param x_split: Horizontal split position in 1/20th of a point.
        :param y_split: Vertical split position in 1/20th of a point.
        :param top_left_cell: Cell address of the top-left cell in the bottom-right pane.
        :param active_pane: The pane that is active ('bottomRight', 'topRight', 'bottomLeft', 'topLeft').
        """
        from ._openxlsx import XLPane

        pane_map = {
            "bottomRight": XLPane.BottomRight,
            "topRight": XLPane.TopRight,
            "bottomLeft": XLPane.BottomLeft,
            "topLeft": XLPane.TopLeft,
        }
        active_pane_enum = pane_map.get(active_pane, XLPane.BottomRight)
        self._sheet.split_panes(x_split, y_split, top_left_cell, active_pane_enum)

    def clear_panes(self):
        """Clear all panes (frozen or split) from the worksheet."""
        self._sheet.clear_panes()

    @property
    def has_panes(self):
        """Check if the worksheet has frozen or split panes."""
        return self._sheet.has_panes()
