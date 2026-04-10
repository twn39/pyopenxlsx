#include "bindings.hpp"

void init_pivot_table(py::module_& m) {
    py::enum_<XLPivotSubtotal>(m, "XLPivotSubtotal")
        .value("Sum", XLPivotSubtotal::Sum)
        .value("Average", XLPivotSubtotal::Average)
        .value("Count", XLPivotSubtotal::Count)
        .value("Max", XLPivotSubtotal::Max)
        .value("Min", XLPivotSubtotal::Min)
        .value("Product", XLPivotSubtotal::Product);

    py::class_<XLPivotField>(m, "XLPivotField")
        .def(py::init<>())
        .def_rw("name", &XLPivotField::name)
        .def_rw("subtotal", &XLPivotField::subtotal)
        .def_rw("custom_name", &XLPivotField::customName)
        .def_rw("num_fmt_id", &XLPivotField::numFmtId);

    py::class_<XLPivotTableOptions>(m, "XLPivotTableOptions")
        .def(py::init<>())
        .def_rw("name", &XLPivotTableOptions::name)
        .def_rw("source_range", &XLPivotTableOptions::sourceRange)
        .def_rw("target_cell", &XLPivotTableOptions::targetCell)
        .def_rw("rows", &XLPivotTableOptions::rows)
        .def_rw("columns", &XLPivotTableOptions::columns)
        .def_rw("data", &XLPivotTableOptions::data)
        .def_rw("filters", &XLPivotTableOptions::filters)
        .def_rw("data_on_rows", &XLPivotTableOptions::dataOnRows)
        .def_rw("row_grand_totals", &XLPivotTableOptions::rowGrandTotals)
        .def_rw("col_grand_totals", &XLPivotTableOptions::colGrandTotals)
        .def_rw("show_drill", &XLPivotTableOptions::showDrill)
        .def_rw("use_auto_formatting", &XLPivotTableOptions::useAutoFormatting)
        .def_rw("page_over_then_down", &XLPivotTableOptions::pageOverThenDown)
        .def_rw("merge_item", &XLPivotTableOptions::mergeItem)
        .def_rw("compact_data", &XLPivotTableOptions::compactData)
        .def_rw("show_error", &XLPivotTableOptions::showError)
        .def_rw("show_row_headers", &XLPivotTableOptions::showRowHeaders)
        .def_rw("show_col_headers", &XLPivotTableOptions::showColHeaders)
        .def_rw("show_row_stripes", &XLPivotTableOptions::showRowStripes)
        .def_rw("show_col_stripes", &XLPivotTableOptions::showColStripes)
        .def_rw("show_last_column", &XLPivotTableOptions::showLastColumn)
        .def_rw("pivot_table_style_name", &XLPivotTableOptions::pivotTableStyleName);

    py::class_<XLPivotTable>(m, "XLPivotTable").def("set_name", &XLPivotTable::setName);
}
