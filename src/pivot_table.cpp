#include "bindings.hpp"

using namespace OpenXLSX;
using namespace nanobind::literals;


void init_pivot_table(py::module_& m) {
    m.def("create_pivot_options", [](std::string name, std::string source_range, std::string target_cell) {
        return XLPivotTableOptions(std::move(name), std::move(source_range), std::move(target_cell));
    }, "name"_a, "source_range"_a, "target_cell"_a);

    // Bind XLPivotField
    py::class_<XLPivotField>(m, "XLPivotField")
        .def_ro("name", &XLPivotField::name)
        .def_ro("custom_name", &XLPivotField::customName)
        .def_ro("subtotal", &XLPivotField::subtotal)
        .def_ro("num_fmt_id", &XLPivotField::numFmtId);

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
        .def(py::init<std::string, std::string, std::string>(), 
             "name"_a, "source_range"_a, "target_cell"_a)
        .def_prop_ro("name", &XLPivotTableOptions::name)
        .def_prop_ro("source_range", &XLPivotTableOptions::sourceRange)
        .def_prop_ro("target_cell", &XLPivotTableOptions::targetCell)
        .def_prop_ro("rows", &XLPivotTableOptions::rows)
        .def_prop_ro("columns", &XLPivotTableOptions::columns)
        .def_prop_ro("data", &XLPivotTableOptions::data)
        .def_prop_ro("filters", &XLPivotTableOptions::filters)
        .def_prop_ro("pivot_table_style_name", &XLPivotTableOptions::pivotTableStyleName)
        .def("add_row_field", &XLPivotTableOptions::addRowField, py::rv_policy::reference)
        .def("add_column_field", &XLPivotTableOptions::addColumnField, py::rv_policy::reference)
        .def("add_data_field", &XLPivotTableOptions::addDataField, 
             "field_name"_a, "custom_name"_a = "", 
             "subtotal"_a = XLPivotSubtotal::Sum, "num_fmt_id"_a = 0, py::rv_policy::reference)
        .def("add_filter_field", &XLPivotTableOptions::addFilterField, py::rv_policy::reference)
        .def("set_pivot_table_style", &XLPivotTableOptions::setPivotTableStyle, py::rv_policy::reference)
        .def("set_data_on_rows", &XLPivotTableOptions::setDataOnRows, py::rv_policy::reference)
        .def("set_row_grand_totals", &XLPivotTableOptions::setRowGrandTotals, py::rv_policy::reference)
        .def("set_col_grand_totals", &XLPivotTableOptions::setColGrandTotals, py::rv_policy::reference)
        .def("set_show_drill", &XLPivotTableOptions::setShowDrill, py::rv_policy::reference)
        .def("set_use_auto_formatting", &XLPivotTableOptions::setUseAutoFormatting, py::rv_policy::reference)
        .def("set_page_over_then_down", &XLPivotTableOptions::setPageOverThenDown, py::rv_policy::reference)
        .def("set_merge_item", &XLPivotTableOptions::setMergeItem, py::rv_policy::reference)
        .def("set_compact_data", &XLPivotTableOptions::setCompactData, py::rv_policy::reference)
        .def("set_show_error", &XLPivotTableOptions::setShowError, py::rv_policy::reference)
        .def("set_show_row_headers", &XLPivotTableOptions::setShowRowHeaders, py::rv_policy::reference)
        .def("set_show_col_headers", &XLPivotTableOptions::setShowColHeaders, py::rv_policy::reference)
        .def("set_show_row_stripes", &XLPivotTableOptions::setShowRowStripes, py::rv_policy::reference)
        .def("set_show_col_stripes", &XLPivotTableOptions::setShowColStripes, py::rv_policy::reference)
        .def("set_show_last_column", &XLPivotTableOptions::setShowLastColumn, py::rv_policy::reference);

    // Bind XLPivotTable
    py::class_<XLPivotTable>(m, "XLPivotTable")
        .def("name", &XLPivotTable::name)
        .def("source_range", &XLPivotTable::sourceRange)
        .def("target_cell", &XLPivotTable::targetCell)
        .def("set_name", &XLPivotTable::setName)
;
}
