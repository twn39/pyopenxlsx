"""pyopenxlsx — complete Python bindings for OpenXLSX-NX.

API layers
----------
**Recommended high-level API** (stable, Pythonic)::

    Workbook, Worksheet, Cell, Range, Column, load_workbook,
    Font, Fill, Border, Side, Alignment, Style, Protection,
    Table, Formula, FormulaEngine, CalculationEngine,
    DataValidation, DataValidations, PageMargins, PrintOptions, PageSetup,
    chart helpers (Chart, chart_type, add_chart), pivot builders,
    conditional_formatting builders, StreamWriter / StreamReader,
    DefinedNames / link_cell

**Advanced / native surface** (``pyopenxlsx._openxlsx`` and star-export of
``XL*`` types): 1:1 with the C++ OpenXLSX-NX API. Prefer for charts/pivot
options enums and low-level rule objects; names may track upstream C++
more closely than the high-level façade.

Internal modules named ``_ws_*`` are Worksheet mixins and are not part of
the public API.
"""

from __future__ import annotations

import pyopenxlsx._openxlsx as _ox

# Re-export the entire native surface (no BC filtering).
from pyopenxlsx._openxlsx import *  # noqa: F403

from .styles import (
    Font,
    Fill,
    Alignment,
    Border,
    Side,
    Style,
    Protection,
    is_date_format,
)
from .cell import Cell
from .formula import Formula
from .formula_engine import CalculationEngine, FormulaEngine, calculation_options
from .range import Range
from .worksheet import Worksheet
from .column import Column
from .table import Table
from .page_setup import PageMargins, PrintOptions, PageSetup
from .workbook import Workbook, load_workbook, load_workbook_async
from .merge import MergeCells as PythonMergeCells
from .data_validation import DataValidation, DataValidations
from .chart import Chart, add_chart, chart_type
from .pivot import PivotTableBuilder, pivot_subtotal, pivot_table
from .streams import StreamReader, StreamWriter
from .defined_names import DefinedName, DefinedNames
from .hyperlink import link as link_cell
from . import conditional_formatting

# Constant shortcuts
XLPatternNone = getattr(_ox.XLPatternType, "None")
XLPatternSolid = _ox.XLPatternType.Solid
XLAlignGeneral = _ox.XLAlignmentStyle.General
XLAlignLeft = _ox.XLAlignmentStyle.Left
XLAlignRight = _ox.XLAlignmentStyle.Right
XLAlignCenter = _ox.XLAlignmentStyle.Center
XLAlignTop = _ox.XLAlignmentStyle.Top
XLAlignBottom = _ox.XLAlignmentStyle.Bottom
XLAlignVCenter = _ox.XLAlignmentStyle.Center

__version__ = getattr(_ox, "__version__", "1.4.2")

__all__ = [name for name in dir(_ox) if not name.startswith("_")]
__all__ += [
    "Workbook",
    "Worksheet",
    "PythonMergeCells",
    "DataValidation",
    "DataValidations",
    "Table",
    "PageMargins",
    "PrintOptions",
    "PageSetup",
    "Formula",
    "FormulaEngine",
    "CalculationEngine",
    "calculation_options",
    "Cell",
    "Range",
    "Column",
    "load_workbook",
    "load_workbook_async",
    "Font",
    "Fill",
    "Alignment",
    "Border",
    "Side",
    "Style",
    "Protection",
    "is_date_format",
    "chart_type",
    "add_chart",
    "Chart",
    "PivotTableBuilder",
    "pivot_table",
    "pivot_subtotal",
    "StreamWriter",
    "StreamReader",
    "DefinedName",
    "DefinedNames",
    "link_cell",
    "conditional_formatting",
    "XLPatternNone",
    "XLPatternSolid",
    "XLAlignGeneral",
    "XLAlignLeft",
    "XLAlignRight",
    "XLAlignCenter",
    "XLAlignTop",
    "XLAlignBottom",
    "XLAlignVCenter",
]
