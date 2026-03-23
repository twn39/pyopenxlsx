#include <XLAutoFilter.hpp>

#include "bindings.hpp"

void init_autofilter(py::module_& m) {
    py::enum_<XLFilterLogic>(m, "XLFilterLogic")
        .value("And", XLFilterLogic::And)
        .value("Or", XLFilterLogic::Or);

    py::class_<XLFilterColumn>(m, "XLFilterColumn")
        .def("add_filter", &XLFilterColumn::addFilter)
        .def("clear_filters", &XLFilterColumn::clearFilters)
        .def("set_custom_filter", py::overload_cast<const std::string&, const std::string&>(
                                      &XLFilterColumn::setCustomFilter))
        .def("set_custom_filter",
             py::overload_cast<const std::string&, const std::string&, XLFilterLogic,
                               const std::string&, const std::string&>(
                 &XLFilterColumn::setCustomFilter))
        .def("set_top10", &XLFilterColumn::setTop10, py::arg("value"), py::arg("percent") = false,
             py::arg("top") = true)
        .def("col_id", &XLFilterColumn::colId);

    py::class_<XLAutoFilter>(m, "XLAutoFilter")
        .def("__bool__", [](const XLAutoFilter& self) { return static_cast<bool>(self); })
        .def("ref", &XLAutoFilter::ref)
        .def("set_ref", py::overload_cast<const std::string&>(&XLAutoFilter::setRef),
             py::arg("ref"))
        .def("set_ref_range", py::overload_cast<const XLCellRange&>(&XLAutoFilter::setRef),
             py::arg("range"))
        .def("filter_column", &XLAutoFilter::filterColumn);
}
