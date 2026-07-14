#include <headers/XLRow.hpp>
#include <headers/XLRowData.hpp>

#include "bindings.hpp"
#include "internal_access.hpp"

using namespace OpenXLSX;
using namespace nanobind::literals;

void init_row(py::module_& m) {
    py::class_<XLRow>(m, "XLRow")
        .def(py::init<>())
        .def("empty", &XLRow::empty)
        .def("height", &XLRow::height)
        .def("set_height", &XLRow::setHeight, "height"_a)
        .def("descent", &XLRow::descent)
        .def("set_descent", &XLRow::setDescent, "descent"_a)
        .def("is_hidden", &XLRow::isHidden)
        .def("set_hidden", &XLRow::setHidden, "state"_a)
        .def("outline_level", &XLRow::outlineLevel)
        .def("set_outline_level", &XLRow::setOutlineLevel, "level"_a)
        .def("is_collapsed", &XLRow::isCollapsed)
        .def("set_collapsed", &XLRow::setCollapsed, "state"_a)
        .def("row_number", &XLRow::rowNumber)
        .def("cell_count", &XLRow::cellCount)
        .def(
            "values",
            [](XLRow& self) {
                py::list result;
                auto range = self.cells();
                for (auto cell : range) {
                    result.append(CellData::from(cell.value()).to_python());
                }
                return result;
            },
            "Return all cell values in this row as a list.")
        .def(
            "set_values",
            [](XLRow& self, py::sequence values) {
                std::vector<XLCellValue> vals;
                vals.reserve(py::len(values));
                for (auto v : values) {
                    vals.push_back(CellData::from_python(v).to_xlcellvalue());
                }
                self.values() = vals;
            },
            "values"_a)
        .def("find_cell", &XLRow::findCell, "column_number"_a)
        .def("format", &XLRow::format)
        .def("set_format", &XLRow::setFormat, "cell_format_index"_a);

    py::class_<XLRowRange>(m, "XLRowRange")
        .def("row_count", &XLRowRange::rowCount)
        .def("__len__", &XLRowRange::rowCount)
        .def(
            "__iter__",
            [m = py::module_(m)](XLRowRange& self) {
                return py::make_iterator(m, "XLRowRangeIterator", self.begin(), self.end());
            },
            py::keep_alive<0, 1>());
}
