#!/usr/bin/env python3
"""Generate Excel samples for manual validation in Microsoft Excel.

Usage (repo root)::

    uv run python scripts/generate_manual_samples.py

Output: ``samples/manual_validation/*.xlsx`` + ``CHECKLIST.md``
"""

from __future__ import annotations

from datetime import date, datetime, timedelta
from pathlib import Path

from pyopenxlsx import (
    Alignment,
    Border,
    Fill,
    Font,
    FormulaEngine,
    PivotTableBuilder,
    Side,
    Workbook,
    conditional_formatting as cf,
)

OUT = Path(__file__).resolve().parents[1] / "samples" / "manual_validation"


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)

    # 01 basics
    with Workbook() as wb:
        ws = wb.active
        ws.title = "Basics"
        ws["A1"].value = "pyopenxlsx manual sample"
        ws["A2"].value = "Hello"
        ws["B2"].value = 42
        ws["C2"].value = 3.14159
        ws["D2"].value = True
        style = wb.add_style(
            font=Font(name="Arial", size=12, bold=True, color="FFFFFF"),
            fill=Fill(pattern_type="solid", color="4472C4"),
            border=Border(
                left=Side(style="thin"),
                right=Side(style="thin"),
                top=Side(style="thin"),
                bottom=Side(style="thin"),
            ),
            alignment=Alignment(horizontal="center", vertical="center"),
        )
        ws["A1"].style_index = style
        ws.merge_cells("A1:E1")
        ws.write_rows(
            4,
            [
                ["Name", "Qty", "Price"],
                ["Apple", 10, 1.5],
                ["Banana", 20, 0.8],
                ["Cherry", 5, 3.2],
            ],
        )
        hs = wb.add_style(font=Font(bold=True))
        for c in range(1, 4):
            ws.cell(4, c).style_index = hs
        try:
            ws.column(1).width = 18
        except Exception:
            pass
        wb.properties.title = "Manual Validation 01 Basics"
        wb.properties.creator = "pyopenxlsx"
        try:
            wb.custom_properties["Sample"] = "01_basics"
        except Exception:
            pass
        wb.save(OUT / "01_basics_styles_merge.xlsx")

    # 02 dates
    with Workbook() as wb:
        ws = wb.active
        ws.title = "Dates"
        ws.write_row(1, ["path", "date", "datetime", "note"])
        ws.cell(2, 1).value = "cell"
        ws.cell(2, 2).value = date(2024, 3, 15)
        ws.cell(2, 3).value = datetime(2024, 3, 15, 14, 30, 0)
        ws.cell(2, 4).value = "auto_date_formats=True"
        ws.write_rows(
            3,
            [
                [
                    "write_rows",
                    date(2024, 6, 1),
                    datetime(2024, 6, 1, 9, 0, 0),
                    "bulk",
                ],
                [
                    "write_rows",
                    date(2024, 12, 25),
                    datetime(2024, 12, 25, 23, 59, 0),
                    "bulk",
                ],
            ],
        )
        ws.set_cell_value(5, 1, "set_cell_value")
        ws.set_cell_value(5, 2, date(2025, 1, 1))
        ws.set_cell_value(5, 3, datetime(2025, 1, 1, 12, 0, 0))
        ws.set_cell_value(5, 4, "fast path")
        ws2 = wb.create_sheet("NoAutoDate")
        wb.auto_date_formats = False
        ws2["A1"].value = "auto_date_formats=False"
        ws2["A2"].value = date(2024, 7, 4)
        wb.save(OUT / "02_dates_bulk.xlsx")

    # 03 table + filter (separate sheets)
    # IMPORTANT: Do not set worksheet autoFilter on the same range as an Excel
    # Table — OOXML then emits both sheet <autoFilter> and table <autoFilter>,
    # which Microsoft Excel rejects as a corrupt workbook.
    data_rows = [
        ["Region", "Product", "Units", "Status"],
        ["North", "Laptop", 5, "Open"],
        ["South", "Mouse", 12, "Open"],
        ["East", "Keyboard", 3, "Closed"],
        ["West", "Monitor", 7, "Open"],
        ["North", "Laptop", 2, "Closed"],
    ]
    with Workbook() as wb:
        ws = wb.active
        ws.title = "WithTable"
        ws.write_rows(1, data_rows)
        t = ws.add_table("SalesTable", "A1:D6")
        try:
            t.style = "TableStyleMedium2"
            t.show_row_stripes = True
        except Exception:
            pass
        # Table already owns autoFilter; do not call ws.auto_filter here.

        ws2 = wb.create_sheet("FilterOnly")
        ws2.write_rows(1, data_rows)
        ws2.auto_filter = "A1:D6"
        try:
            if hasattr(ws2, "apply_auto_filter"):
                ws2.apply_auto_filter()
        except Exception:
            pass
        wb.save(OUT / "03_table_filter_validation.xlsx")

    # 04 charts
    with Workbook() as wb:
        ws = wb.active
        ws.title = "ChartData"
        ws.write_rows(
            1,
            [
                ["Category", "Q1", "Q2"],
                ["A", 10, 15],
                ["B", 20, 18],
                ["C", 15, 25],
                ["D", 30, 22],
            ],
        )
        (
            ws.add_chart("column", "SalesChart", row=2, col=5, width=480, height=320)
            .title("Quarterly Sales")
            .legend("bottom")
            .series(
                "ChartData!$B$2:$B$5",
                name="Q1",
                categories_ref="ChartData!$A$2:$A$5",
            )
            .series(
                "ChartData!$C$2:$C$5",
                name="Q2",
                categories_ref="ChartData!$A$2:$A$5",
            )
        )
        (
            ws.add_chart("line", "TrendChart", row=18, col=5, width=480, height=280)
            .title("Q1 Trend")
            .series(
                "ChartData!$B$2:$B$5",
                name="Q1",
                categories_ref="ChartData!$A$2:$A$5",
            )
        )
        wb.save(OUT / "04_charts.xlsx")

    # 05 pivot
    with Workbook() as wb:
        src = wb.active
        src.title = "Source"
        src.write_rows(
            1,
            [
                ["Region", "Product", "Sales"],
                ["North", "Apples", 100],
                ["South", "Bananas", 300],
                ["North", "Bananas", 150],
                ["South", "Apples", 200],
                ["East", "Apples", 80],
                ["East", "Bananas", 120],
            ],
        )
        piv = wb.create_sheet("Pivot")
        (
            PivotTableBuilder("SalesPivot", "Source!A1:C7", "B2")
            .rows("Region")
            .columns("Product")
            .data("Sales", name="Total Sales", subtotal="sum")
            .style("PivotStyleMedium9")
            .add_to(piv)
        )
        wb.save(OUT / "05_pivot.xlsx")

    # 06 CF
    with Workbook() as wb:
        ws = wb.active
        ws.title = "CF"
        ws.write_rows(
            1,
            [
                [10, 20, 30, 40, 50],
                [5, 15, 25, 35, 45],
                [1, 2, 3, 4, 5],
                [100, 80, 60, 40, 20],
            ],
        )
        ws.add_conditional_formatting(
            "A1:E1", cf.color_scale("#F8696B", "#63BE7B")
        )
        ws.add_conditional_formatting("A2:E2", cf.data_bar((0, 112, 192)))
        ws.add_conditional_formatting("A3:E3", cf.cell_is(">", "3"))
        ws.add_conditional_formatting("A4:E4", cf.top10(2))
        ws["A6"].value = "Row1 color scale; Row2 data bar; Row3 >3; Row4 top2"
        wb.save(OUT / "06_conditional_formatting.xlsx")

    # 07 links + names
    with Workbook() as wb:
        ws = wb.active
        ws.title = "Links"
        ws.link(
            "A1",
            "https://github.com/twn39/pyopenxlsx",
            text="pyopenxlsx GitHub",
            tooltip="Project",
        )
        ws.link("A2", "https://www.python.org", text="Python.org")
        dest = wb.create_sheet("Target")
        dest["A1"].value = "You arrived via internal link"
        dest["B1"].value = 123
        ws.link("A3", "Target!A1", text="Jump to Target!A1", internal=True)
        ws.write_rows(5, [["Named range block"], [10], [20], [30]])
        wb.defined_names.define("SampleBlock", "Links!$A$6:$A$8")
        wb.defined_names.define("LocalTitle", "Links!$A$1", sheet=ws)
        ws["A10"].value = "Name Manager: SampleBlock, LocalTitle"
        wb.save(OUT / "07_hyperlinks_defined_names.xlsx")

    # 08 comments
    with Workbook() as wb:
        ws = wb.active
        ws.title = "Comments"
        ws["A1"].value = "Hover for comment"
        ws["A1"].comment = "Legacy comment from pyopenxlsx\nSecond line"
        try:
            ws.add_comment("B1", "Simple comment via add_comment", "tester")
            ws["B1"].value = "add_comment"
        except Exception:
            pass
        try:
            ws.add_threaded_comment("C1", "Threaded root message", "alice")
            ws["C1"].value = "threaded"
        except Exception:
            pass
        wb.create_sheet("Extra")["A1"].value = "Second sheet"
        wb.save(OUT / "08_comments.xlsx")

    # 09 formulas
    with Workbook() as wb:
        ws = wb.active
        ws.title = "Formulas"
        ws.write_rows(1, [["A", "B", "C"], [10, 20, 30], [1, 2, 3]])
        ws["A4"].formula = "SUM(A2:C2)"
        ws["B4"].formula = "AVERAGE(A3:C3)"
        ws["C4"].formula = "A2*B2"
        ws["A5"].value = "Excel recalculates A4:C4 on open"
        eng = FormulaEngine()
        ws["A6"].value = "Precomputed engine SUM(A2:C2)"
        ws["B6"].value = eng.evaluate("SUM(A2:C2)", ws)
        wb.save(OUT / "09_formulas.xlsx")

    # 10 page setup
    with Workbook() as wb:
        ws = wb.active
        ws.title = "Page"
        ws.write_rows(
            1, [[f"R{r}C{c}" for c in range(1, 6)] for r in range(1, 21)]
        )
        try:
            from pyopenxlsx._openxlsx import XLPageOrientation

            ws.page_setup.orientation = XLPageOrientation.Landscape
        except Exception:
            pass
        try:
            ws.page_margins.left = 0.5
            ws.page_margins.right = 0.5
        except Exception:
            pass
        try:
            ws.set_print_area("A1:E20")
        except Exception:
            pass
        wb.create_sheet("Notes")["A1"].value = "Check Page Layout on sheet Page"
        wb.properties.title = "Manual Validation Page Setup"
        wb.save(OUT / "10_page_setup_multisheet.xlsx")

    # 11 stream
    with Workbook() as wb:
        ws = wb.active
        ws.title = "Stream"
        with ws.stream_writer() as writer:
            writer.append_row(["id", "when", "value", "label"])
            for i in range(1, 201):
                writer.append_row(
                    [
                        i,
                        date(2024, 1, 1) + timedelta(days=i % 30),
                        i * 1.5,
                        f"row-{i}",
                    ]
                )
        wb.save(OUT / "11_stream_200rows.xlsx")

    # 12 kitchen sink — add charts BEFORE external https:// links
    # (backend path builder can break on "//" from prior hyperlinks).
    with Workbook() as wb:
        ws = wb.active
        ws.title = "Overview"
        ws.write_rows(
            1,
            [
                ["Feature", "How to verify"],
                ["CF", "Color scale on A11:E11"],
                ["Link", "Click A13"],
                ["Dates", "A15 date, B15 datetime"],
                ["Chart", "Column chart"],
                ["Named range", "OverviewTotal"],
                ["Table", "OverviewTable A17:B20"],
            ],
        )
        ws.write_row(11, [12, 45, 78, 33, 90])
        ws.add_conditional_formatting(
            "A11:E11", cf.color_scale("#F8696B", "#63BE7B")
        )
        ws["A15"].value = date(2024, 10, 1)
        ws["B15"].value = datetime(2024, 10, 1, 8, 30)
        ws.write_rows(
            17, [["Item", "Amount"], ["X", 10], ["Y", 20], ["Z", 15]]
        )
        (
            ws.add_chart("column", "AmtChart", row=17, col=4, width=420, height=260)
            .title("Amounts")
            .legend("bottom")
            .series(
                "Overview!$B$18:$B$20",
                name="Amount",
                categories_ref="Overview!$A$18:$A$20",
            )
        )
        try:
            ws.add_table("OverviewTable", "A17:B20")
        except Exception:
            pass
        ws.link(
            "A13",
            "https://pypi.org/project/pyopenxlsx/",
            text="PyPI: pyopenxlsx",
        )
        wb.defined_names.define("OverviewTotal", "Overview!$B$18:$B$20")
        ws["A22"].value = "Try in Excel: =SUM(OverviewTotal)"
        wb.properties.title = "pyopenxlsx Kitchen Sink"
        wb.properties.creator = "pyopenxlsx generator"
        try:
            wb.custom_properties["Purpose"] = "manual Excel validation"
        except Exception:
            pass
        wb.save(OUT / "12_kitchen_sink.xlsx")

    (OUT / "CHECKLIST.md").write_text(
        """# Manual Excel validation checklist

Open each file in **Microsoft Excel** and confirm **no repair dialog**.

| # | File | What to check |
|---|------|----------------|
| 01 | `01_basics_styles_merge.xlsx` | Merged blue header; data rows |
| 02 | `02_dates_bulk.xlsx` | Dates/times on **Dates**; **NoAutoDate** may be serial |
| 03 | `03_table_filter_validation.xlsx` | Table / AutoFilter |
| 04 | `04_charts.xlsx` | Column + line charts |
| 05 | `05_pivot.xlsx` | Pivot on **Pivot** (refresh if prompted) |
| 06 | `06_conditional_formatting.xlsx` | Color scale / data bars / rules |
| 07 | `07_hyperlinks_defined_names.xlsx` | Links + Name Manager |
| 08 | `08_comments.xlsx` | Comments / notes |
| 09 | `09_formulas.xlsx` | Formulas calculate; B6 = 60 |
| 10 | `10_page_setup_multisheet.xlsx` | Multi-sheet + page layout |
| 11 | `11_stream_200rows.xlsx` | ~200 rows with dates |
| 12 | `12_kitchen_sink.xlsx` | Combined features |

```bash
uv run python scripts/generate_manual_samples.py
```
""",
        encoding="utf-8",
    )

    print(f"Wrote samples to {OUT}")
    for p in sorted(OUT.glob("*.xlsx")):
        print(f"  {p.name:40s} {p.stat().st_size:8d} bytes")


if __name__ == "__main__":
    main()
