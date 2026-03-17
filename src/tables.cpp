#include "bindings.hpp"

void init_tables(py::module_& m) {
    py::class_<XLTables>(m, "XLTables")
        .def(py::init<>())
        .def("name", &XLTables::name)
        .def("set_name", &XLTables::setName)
        .def("display_name", &XLTables::displayName)
        .def("set_display_name", &XLTables::setDisplayName)
        .def("range_reference", &XLTables::rangeReference)
        .def("set_range_reference", &XLTables::setRangeReference)
        .def("style_name", &XLTables::styleName)
        .def("set_style_name", &XLTables::setStyleName)
        .def("show_row_stripes", &XLTables::showRowStripes)
        .def("set_show_row_stripes", &XLTables::setShowRowStripes)
        .def("show_column_stripes", &XLTables::showColumnStripes)
        .def("set_show_column_stripes", &XLTables::setShowColumnStripes)
        .def("show_first_column", &XLTables::showFirstColumn)
        .def("set_show_first_column", &XLTables::setShowFirstColumn)
        .def("show_last_column", &XLTables::showLastColumn)
        .def("set_show_last_column", &XLTables::setShowLastColumn)
        .def("append_column", &XLTables::appendColumn);
}
