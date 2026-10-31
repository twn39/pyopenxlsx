#include "bindings.hpp"

using namespace OpenXLSX;
using namespace nanobind::literals;

void init_pivot_table(py::module_& m) {
    m.def(
        "create_pivot_options",
        [](std::string name, std::string source_range, std::string target_cell) {
            return XLPivotTableOptions(std::move(name), std::move(source_range),
                                       std::move(target_cell));
        },
        "name"_a, "source_range"_a, "target_cell"_a);

    // Bind XLPivotShowValuesAs enum
    py::enum_<XLPivotShowValuesAs>(m, "XLPivotShowValuesAs")
        .value("Normal", XLPivotShowValuesAs::Normal)
        .value("PercentOfGrandTotal", XLPivotShowValuesAs::PercentOfGrandTotal)
        .value("PercentOfColumnTotal", XLPivotShowValuesAs::PercentOfColumnTotal)
        .value("PercentOfRowTotal", XLPivotShowValuesAs::PercentOfRowTotal)
        .value("PercentOf", XLPivotShowValuesAs::PercentOf)
        .value("PercentOfParentRowTotal", XLPivotShowValuesAs::PercentOfParentRowTotal)
        .value("PercentOfParentColumnTotal", XLPivotShowValuesAs::PercentOfParentColumnTotal)
        .value("PercentOfParentTotal", XLPivotShowValuesAs::PercentOfParentTotal)
        .value("DifferenceFrom", XLPivotShowValuesAs::DifferenceFrom)
        .value("PercentDifferenceFrom", XLPivotShowValuesAs::PercentDifferenceFrom)
        .value("RunningTotalIn", XLPivotShowValuesAs::RunningTotalIn)
        .value("PercentRunningTotalIn", XLPivotShowValuesAs::PercentRunningTotalIn)
        .value("RankSmallestToLargest", XLPivotShowValuesAs::RankSmallestToLargest)
        .value("RankLargestToSmallest", XLPivotShowValuesAs::RankLargestToSmallest)
        .value("Index", XLPivotShowValuesAs::Index)
        .export_values();

    // Bind XLPivotShowValuesAsOptions struct
    py::class_<XLPivotShowValuesAsOptions>(m, "XLPivotShowValuesAsOptions")
        .def(py::init<XLPivotShowValuesAs, std::string, std::string>(),
             "type"_a = XLPivotShowValuesAs::Normal, "base_field"_a = "", "base_item"_a = "")
        .def_rw("type", &XLPivotShowValuesAsOptions::type)
        .def_rw("base_field", &XLPivotShowValuesAsOptions::baseField)
        .def_rw("base_item", &XLPivotShowValuesAsOptions::baseItem)
        .freeze();

    // Bind XLPivotField
    py::class_<XLPivotField>(m, "XLPivotField")
        .def(py::init<>())
        .def_rw("name", &XLPivotField::name)
        .def_rw("custom_name", &XLPivotField::customName)
        .def_rw("subtotal", &XLPivotField::subtotal)
        .def_rw("num_fmt_id", &XLPivotField::numFmtId)
        .def_rw("compact", &XLPivotField::compact)
        .def_rw("outline", &XLPivotField::outline)
        .def_rw("show_all", &XLPivotField::showAll)
        .def_rw("insert_blank_row", &XLPivotField::insertBlankRow)
        .def_rw("default_subtotal", &XLPivotField::defaultSubtotal)
        .def_rw("show_values_as", &XLPivotField::showValuesAs)
        .freeze();

    // Bind XLPivotSubtotal enum
    py::enum_<XLPivotSubtotal>(m, "XLPivotSubtotal")
        .value("Average", XLPivotSubtotal::Average)
        .value("Count", XLPivotSubtotal::Count)
        .value("Max", XLPivotSubtotal::Max)
        .value("Min", XLPivotSubtotal::Min)
        .value("Product", XLPivotSubtotal::Product)
        .value("Sum", XLPivotSubtotal::Sum)
        .value("CountNums", XLPivotSubtotal::CountNums)
        .value("StdDev", XLPivotSubtotal::StdDev)
        .value("StdDevP", XLPivotSubtotal::StdDevP)
        .value("Var", XLPivotSubtotal::Var)
        .value("VarP", XLPivotSubtotal::VarP)
        .export_values();

    // Bind XLPivotTableOptions
    py::class_<XLPivotTableOptions>(m, "XLPivotTableOptions")
        .def(py::init<std::string, std::string, std::string>(), "name"_a, "source_range"_a,
             "target_cell"_a)
        .def_prop_ro("name", &XLPivotTableOptions::name)
        .def_prop_ro("source_range", &XLPivotTableOptions::sourceRange)
        .def_prop_ro("target_cell", &XLPivotTableOptions::targetCell)
        .def_prop_ro("rows", &XLPivotTableOptions::rows)
        .def_prop_ro("columns", &XLPivotTableOptions::columns)
        .def_prop_ro("data", &XLPivotTableOptions::data)
        .def_prop_ro("filters", &XLPivotTableOptions::filters)
        .def_prop_ro("pivot_table_style_name", &XLPivotTableOptions::pivotTableStyleName)
        .def(
            "add_row_field",
            [](XLPivotTableOptions& self, std::string field_name,
               std::vector<std::string> selected_items, bool compact, bool outline, bool show_all,
               bool insert_blank_row, bool default_subtotal) -> XLPivotTableOptions& {
                XLPivotField f;
                f.name = std::move(field_name);
                f.selectedItems = std::move(selected_items);
                f.compact = compact;
                f.outline = outline;
                f.showAll = show_all;
                f.insertBlankRow = insert_blank_row;
                f.defaultSubtotal = default_subtotal;
                return self.addRowField(std::move(f));
            },
            "field_name"_a, "selected_items"_a = std::vector<std::string>{}, "compact"_a = false,
            "outline"_a = false, "show_all"_a = false, "insert_blank_row"_a = false,
            "default_subtotal"_a = true, py::rv_policy::reference)
        .def(
            "add_column_field",
            [](XLPivotTableOptions& self, std::string field_name,
               std::vector<std::string> selected_items, bool compact, bool outline, bool show_all,
               bool insert_blank_row, bool default_subtotal) -> XLPivotTableOptions& {
                XLPivotField f;
                f.name = std::move(field_name);
                f.selectedItems = std::move(selected_items);
                f.compact = compact;
                f.outline = outline;
                f.showAll = show_all;
                f.insertBlankRow = insert_blank_row;
                f.defaultSubtotal = default_subtotal;
                return self.addColumnField(std::move(f));
            },
            "field_name"_a, "selected_items"_a = std::vector<std::string>{}, "compact"_a = false,
            "outline"_a = false, "show_all"_a = false, "insert_blank_row"_a = false,
            "default_subtotal"_a = true, py::rv_policy::reference)
        .def(
            "add_data_field",
            [](XLPivotTableOptions& self, std::string field_name, std::string custom_name,
               XLPivotSubtotal subtotal, uint32_t num_fmt_id,
               const XLPivotShowValuesAsOptions& show_values_as) -> XLPivotTableOptions& {
                return self.addDataField(std::move(field_name), std::move(custom_name), subtotal,
                                         num_fmt_id, show_values_as);
            },
            "field_name"_a, "custom_name"_a = "", "subtotal"_a = XLPivotSubtotal::Sum,
            "num_fmt_id"_a = 0, "show_values_as"_a = XLPivotShowValuesAsOptions{},
            py::rv_policy::reference)
        .def(
            "add_filter_field",
            [](XLPivotTableOptions& self, std::string field_name,
               std::vector<std::string> selected_items) -> XLPivotTableOptions& {
                return self.addFilterField(std::move(field_name), std::move(selected_items));
            },
            "field_name"_a, "selected_items"_a = std::vector<std::string>{},
            py::rv_policy::reference)
        .def("set_pivot_table_style", &XLPivotTableOptions::setPivotTableStyle,
             py::rv_policy::reference)
        .def("set_data_on_rows", &XLPivotTableOptions::setDataOnRows, py::rv_policy::reference)
        .def("set_row_grand_totals", &XLPivotTableOptions::setRowGrandTotals,
             py::rv_policy::reference)
        .def("set_col_grand_totals", &XLPivotTableOptions::setColGrandTotals,
             py::rv_policy::reference)
        .def("set_show_drill", &XLPivotTableOptions::setShowDrill, py::rv_policy::reference)
        .def("set_use_auto_formatting", &XLPivotTableOptions::setUseAutoFormatting,
             py::rv_policy::reference)
        .def("set_page_over_then_down", &XLPivotTableOptions::setPageOverThenDown,
             py::rv_policy::reference)
        .def("set_merge_item", &XLPivotTableOptions::setMergeItem, py::rv_policy::reference)
        .def("set_compact_data", &XLPivotTableOptions::setCompactData, py::rv_policy::reference)
        .def("set_show_error", &XLPivotTableOptions::setShowError, py::rv_policy::reference)
        .def("set_show_row_headers", &XLPivotTableOptions::setShowRowHeaders,
             py::rv_policy::reference)
        .def("set_show_col_headers", &XLPivotTableOptions::setShowColHeaders,
             py::rv_policy::reference)
        .def("set_show_row_stripes", &XLPivotTableOptions::setShowRowStripes,
             py::rv_policy::reference)
        .def("set_show_col_stripes", &XLPivotTableOptions::setShowColStripes,
             py::rv_policy::reference)
        .def("set_show_last_column", &XLPivotTableOptions::setShowLastColumn,
             py::rv_policy::reference)
        .freeze();

    // Bind XLPivotTable
    py::class_<XLPivotTable>(m, "XLPivotTable")
        .def("name", &XLPivotTable::name)
        .def("source_range", &XLPivotTable::sourceRange)
        .def("target_cell", &XLPivotTable::targetCell)
        .def("set_name", &XLPivotTable::setName)
        .freeze();
}
