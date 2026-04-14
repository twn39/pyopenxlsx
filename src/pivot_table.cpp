#include "bindings.hpp"

void init_pivot_table(py::module_& m) {
    // Bind XLPivotField
    py::class_<XLPivotField>(m, "XLPivotField")
        .def(py::init<>())
        .def_rw("name", &XLPivotField::name)
        .def_rw("custom_name", &XLPivotField::customName)
        .def_rw("subtotal", &XLPivotField::subtotal)
        .def_rw("num_fmt_id", &XLPivotField::numFmtId);

    // Bind XLPivotSubtotal enum
    py::enum_<XLPivotSubtotal>(m, "XLPivotSubtotal")
        .value("Average", XLPivotSubtotal::Average)
        .value("Count", XLPivotSubtotal::Count)
        .value("Max", XLPivotSubtotal::Max)
        .value("Min", XLPivotSubtotal::Min)
        .value("Product", XLPivotSubtotal::Product)
        .value("Sum", XLPivotSubtotal::Sum)
        .export_values();

    // Bind XLPivotTableOptions
    py::class_<XLPivotTableOptions>(m, "XLPivotTableOptions")
        .def("__init__", [](XLPivotTableOptions *t, std::string name, std::string source_range, std::string target_cell) {
            new (t) XLPivotTableOptions(std::move(name), std::move(source_range), std::move(target_cell));
        }, py::arg("name"), py::arg("source_range"), py::arg("target_cell"))
        .def_prop_ro("name", &XLPivotTableOptions::name)
        .def_prop_ro("source_range", &XLPivotTableOptions::sourceRange)
        .def_prop_ro("target_cell", &XLPivotTableOptions::targetCell)
        .def_prop_ro("rows", &XLPivotTableOptions::rows)
        .def_prop_ro("columns", &XLPivotTableOptions::columns)
        .def_prop_ro("data", &XLPivotTableOptions::data)
        .def_prop_ro("filters", &XLPivotTableOptions::filters)
        .def_prop_ro("pivot_table_style_name", &XLPivotTableOptions::pivotTableStyleName)
        .def("add_row_field", &XLPivotTableOptions::addRowField)
        .def("add_column_field", &XLPivotTableOptions::addColumnField)
        .def("add_data_field", &XLPivotTableOptions::addDataField, py::arg("field_name"), py::arg("custom_name") = "", py::arg("subtotal") = XLPivotSubtotal::Sum, py::arg("num_fmt_id") = 0)
        .def("add_filter_field", &XLPivotTableOptions::addFilterField)
        .def("set_pivot_table_style", &XLPivotTableOptions::setPivotTableStyle)
        .def("set_data_on_rows", &XLPivotTableOptions::setDataOnRows)
        .def("set_row_grand_totals", &XLPivotTableOptions::setRowGrandTotals)
        .def("set_col_grand_totals", &XLPivotTableOptions::setColGrandTotals)
        .def("set_show_drill", &XLPivotTableOptions::setShowDrill)
        .def("set_use_auto_formatting", &XLPivotTableOptions::setUseAutoFormatting)
        .def("set_page_over_then_down", &XLPivotTableOptions::setPageOverThenDown)
        .def("set_merge_item", &XLPivotTableOptions::setMergeItem)
        .def("set_compact_data", &XLPivotTableOptions::setCompactData)
        .def("set_show_error", &XLPivotTableOptions::setShowError)
        .def("set_show_row_headers", &XLPivotTableOptions::setShowRowHeaders)
        .def("set_show_col_headers", &XLPivotTableOptions::setShowColHeaders)
        .def("set_show_row_stripes", &XLPivotTableOptions::setShowRowStripes)
        .def("set_show_col_stripes", &XLPivotTableOptions::setShowColStripes)
        .def("set_show_last_column", &XLPivotTableOptions::setShowLastColumn);

    // Bind XLPivotTable
    py::class_<XLPivotTable>(m, "XLPivotTable")
        .def("name", &XLPivotTable::name)
        .def("source_range", &XLPivotTable::sourceRange)
        .def("target_cell", &XLPivotTable::targetCell)
        .def("set_name", &XLPivotTable::setName)
;
}
