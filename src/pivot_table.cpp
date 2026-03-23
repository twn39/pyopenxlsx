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
        .def_rw("custom_name", &XLPivotField::customName);

    py::class_<XLPivotTableOptions>(m, "XLPivotTableOptions")
        .def(py::init<>())
        .def_rw("name", &XLPivotTableOptions::name)
        .def_rw("source_range", &XLPivotTableOptions::sourceRange)
        .def_rw("target_cell", &XLPivotTableOptions::targetCell)
        .def_rw("rows", &XLPivotTableOptions::rows)
        .def_rw("columns", &XLPivotTableOptions::columns)
        .def_rw("data", &XLPivotTableOptions::data)
        .def_rw("filters", &XLPivotTableOptions::filters);

    py::class_<XLPivotTable>(m, "XLPivotTable").def("set_name", &XLPivotTable::setName);
}
