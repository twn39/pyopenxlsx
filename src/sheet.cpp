#include <headers/XLChartsheet.hpp>
#include <headers/XLSheet.hpp>
#include <headers/XLWorkbook.hpp>

#include "bindings.hpp"

using namespace OpenXLSX;
using namespace nanobind::literals;

void init_sheet(py::module_& m) {
    py::enum_<XLSheetType>(m, "XLSheetType")
        .value("Worksheet", XLSheetType::Worksheet)
        .value("Chartsheet", XLSheetType::Chartsheet)
        .value("Dialogsheet", XLSheetType::Dialogsheet)
        .value("Macrosheet", XLSheetType::Macrosheet);

    py::class_<XLSheet>(m, "XLSheet")
        .def("name", &XLSheet::name)
        .def("set_name", &XLSheet::setName, "name"_a)
        .def("index", &XLSheet::index)
        .def("set_index", &XLSheet::setIndex, "index"_a)
        .def("visibility", &XLSheet::visibility)
        .def("set_visibility", &XLSheet::setVisibility, "state"_a)
        .def("color", &XLSheet::color)
        .def("set_color", &XLSheet::setColor, "color"_a)
        .def("is_selected", &XLSheet::isSelected)
        .def("set_selected", &XLSheet::setSelected, "selected"_a)
        .def("is_active", &XLSheet::isActive)
        .def("set_active", &XLSheet::setActive)
        .def(
            "is_worksheet",
            [](const XLSheet& self) { return self.isType<XLWorksheet>(); })
        .def(
            "is_chartsheet",
            [](const XLSheet& self) { return self.isType<XLChartsheet>(); })
        .def(
            "as_worksheet",
            [](XLSheet& self) -> XLWorksheet { return self.get<XLWorksheet>(); })
        .def(
            "as_chartsheet",
            [](XLSheet& self) -> XLChartsheet { return self.get<XLChartsheet>(); });

    py::class_<XLChartsheet>(m, "XLChartsheet")
        .def("name", &XLChartsheet::name)
        .def("set_name", &XLChartsheet::setName, "name"_a)
        .def("index", &XLChartsheet::index)
        .def("set_index", &XLChartsheet::setIndex, "index"_a)
        .def("visibility", &XLChartsheet::visibility)
        .def("set_visibility", &XLChartsheet::setVisibility, "state"_a)
        .def("color", &XLChartsheet::color)
        .def("set_color", &XLChartsheet::setColor, "color"_a)
        .def("is_selected", &XLChartsheet::isSelected)
        .def("set_selected", &XLChartsheet::setSelected, "selected"_a)
        .def("is_active", &XLChartsheet::isActive)
        .def("set_active", &XLChartsheet::setActive);
}
