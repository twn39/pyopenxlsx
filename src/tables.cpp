#include "bindings.hpp"

void init_tables(py::module_& m) {
    py::enum_<XLTotalsRowFunction>(m, "XLTotalsRowFunction")
        .value("None", XLTotalsRowFunction::None)
        .value("Sum", XLTotalsRowFunction::Sum)
        .value("Min", XLTotalsRowFunction::Min)
        .value("Max", XLTotalsRowFunction::Max)
        .value("Average", XLTotalsRowFunction::Average)
        .value("Count", XLTotalsRowFunction::Count)
        .value("CountNums", XLTotalsRowFunction::CountNums)
        .value("StdDev", XLTotalsRowFunction::StdDev)
        .value("Var", XLTotalsRowFunction::Var)
        .value("Custom", XLTotalsRowFunction::Custom);

    py::class_<XLTableColumn>(m, "XLTableColumn")
        .def("id", &XLTableColumn::id)
        .def("name", &XLTableColumn::name)
        .def("set_name", &XLTableColumn::setName)
        .def("totals_row_function", &XLTableColumn::totalsRowFunction)
        .def("set_totals_row_function", &XLTableColumn::setTotalsRowFunction)
        .def("totals_row_label", &XLTableColumn::totalsRowLabel)
        .def("set_totals_row_label", &XLTableColumn::setTotalsRowLabel)
        .def("calculated_column_formula", &XLTableColumn::calculatedColumnFormula)
        .def("set_calculated_column_formula", &XLTableColumn::setCalculatedColumnFormula)
        .def("totals_row_formula", &XLTableColumn::totalsRowFormula)
        .def("set_totals_row_formula", &XLTableColumn::setTotalsRowFormula);

    py::class_<XLTable>(m, "XLTable")
        .def(py::init<>())
        .def("name", &XLTable::name)
        .def("set_name", &XLTable::setName)
        .def("display_name", &XLTable::displayName)
        .def("set_display_name", &XLTable::setDisplayName)
        .def("range_reference", &XLTable::rangeReference)
        .def("set_range_reference", &XLTable::setRangeReference)
        .def("style_name", &XLTable::styleName)
        .def("set_style_name", &XLTable::setStyleName)
        .def("comment", &XLTable::comment)
        .def("set_comment", &XLTable::setComment)
        .def("show_row_stripes", &XLTable::showRowStripes)
        .def("set_show_row_stripes", &XLTable::setShowRowStripes)
        .def("show_column_stripes", &XLTable::showColumnStripes)
        .def("set_show_column_stripes", &XLTable::setShowColumnStripes)
        .def("show_first_column", &XLTable::showFirstColumn)
        .def("set_show_first_column", &XLTable::setShowFirstColumn)
        .def("show_last_column", &XLTable::showLastColumn)
        .def("set_show_last_column", &XLTable::setShowLastColumn)
        .def("show_header_row", &XLTable::showHeaderRow)
        .def("set_show_header_row", &XLTable::setShowHeaderRow)
        .def("show_totals_row", &XLTable::showTotalsRow)
        .def("set_show_totals_row", &XLTable::setShowTotalsRow)
        .def("append_column", &XLTable::appendColumn)
        .def("column", py::overload_cast<std::string_view>(&XLTable::column, py::const_))
        .def("column", py::overload_cast<uint32_t>(&XLTable::column, py::const_));

    py::class_<XLSlicerOptions>(m, "XLSlicerOptions")
        .def(py::init<>())
        .def_rw("name", &XLSlicerOptions::name)
        .def_rw("caption", &XLSlicerOptions::caption)
        .def_rw("width", &XLSlicerOptions::width)
        .def_rw("height", &XLSlicerOptions::height)
        .def_rw("offset_x", &XLSlicerOptions::offsetX)
        .def_rw("offset_y", &XLSlicerOptions::offsetY);

    py::class_<XLTables>(m, "XLTables")
        .def(py::init<>())
        .def("count", &XLTables::count)
        .def("__len__", &XLTables::count)
        .def("__getitem__",
             [](const XLTables& self, size_t index) {
                 if (index >= self.count()) throw py::index_error();
                 return self[index];
             })
        .def("get_table", &XLTables::table)
        .def("add", py::overload_cast<std::string_view, std::string_view>(&XLTables::add),
             py::arg("name"), py::arg("range"))
        .def("add_range", py::overload_cast<std::string_view, const XLCellRange&>(&XLTables::add),
             py::arg("name"), py::arg("range"));
}
