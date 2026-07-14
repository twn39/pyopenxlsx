#include "internal_access.hpp"

using namespace OpenXLSX;
using namespace nanobind::literals;


void init_cell(py::module_& m) {
    // Bind XLMergeCells
    py::class_<XLMergeCells>(m, "XLMergeCells")
        .def("count", &XLMergeCells::count)
        .def("find_merge", &XLMergeCells::findMerge)
        .def("merge_exists", &XLMergeCells::mergeExists)
        .def("append_merge", &XLMergeCells::appendMerge)
        .def("delete_merge", &XLMergeCells::deleteMerge)
        .def("__getitem__", [](const XLMergeCells& self, int index) {
            if (index < 0 || index >= self.count()) throw py::index_error();
            return self.merge(index);
        });

    // Bind XLCellReference
    py::class_<XLCellReference>(m, "XLCellReference")
        .def(py::init<>())
        .def(py::init<std::string_view>(), "cell_address"_a)
        .def(py::init<uint32_t, uint16_t>(), "row"_a, "column"_a)
        .def(py::init<uint32_t, std::string_view>(), "row"_a, "column"_a)
        .def("address", &XLCellReference::address)
        .def("set_address", &XLCellReference::setAddress, "address"_a)
        .def("row", &XLCellReference::row)
        .def("set_row", &XLCellReference::setRow, "row"_a)
        .def("column", &XLCellReference::column)
        .def("set_column", &XLCellReference::setColumn, "column"_a)
        .def("set_row_and_column", &XLCellReference::setRowAndColumn, "row"_a, "column"_a)
        .def("__str__", &XLCellReference::address)
        .def("__eq__", [](const XLCellReference& a, const XLCellReference& b) { return a == b; })
        .def("__lt__", [](const XLCellReference& a, const XLCellReference& b) { return a < b; });

    // Bind XLCellRange
    py::class_<XLCellRange>(m, "XLCellRange")
        .def("address", &XLCellRange::address)
        .def("top_left", &XLCellRange::topLeft)
        .def("bottom_right", &XLCellRange::bottomRight)
        .def("num_rows", &XLCellRange::numRows)
        .def("num_columns", &XLCellRange::numColumns)
        .def("empty", &XLCellRange::empty)
        .def("clear", &XLCellRange::clear)
        .def("set_format", &XLCellRange::setFormat, "cell_format_index"_a,
             py::rv_policy::reference_internal)
        .def("apply_style", &XLCellRange::applyStyle, "style"_a)
        .def("set_border_outline", &XLCellRange::setBorderOutline, "style"_a, "color"_a)
        .def("intersect", &XLCellRange::intersect, "other"_a)
        .def(
            "__iter__",
            [m = py::module_(m)](const XLCellRange& self) {
                return py::make_iterator(m, "XLCellRangeIterator", self.begin(), self.end());
            },
            py::keep_alive<0, 1>());

    // Bind XLFormula
    py::class_<XLFormula>(m, "XLFormula")
        .def(py::init<>())
        .def(py::init<const std::string&>())
        .def("get", &XLFormula::get)
        .def("clear", &XLFormula::clear)
        .def("__str__", [](const XLFormula& self) { return self.get(); })
        .def("__eq__", [](const XLFormula& self, const XLFormula& other) { return self == other; })
        .def("__eq__",
             [](const XLFormula& self, const std::string& other) { return self.get() == other; });

    // Bind XLCell
    py::class_<XLCell>(m, "XLCell")
        .def_prop_rw(
            "value",
            [](const XLCell& self) -> py::object {
                CellData data;
                {
                    py::gil_scoped_release release;
                    data = CellData::from(self.value());
                }
                return data.to_python();
            },
            [](XLCell& self, py::object value) {
                CellData data = CellData::from_python(value);
                py::gil_scoped_release release;
                data.apply_to(self);
            })
        .def("empty", &XLCell::empty)
        .def("clear", &XLCell::clear, "keep"_a = 0)
        .def("copy_from", &XLCell::copyFrom, "other"_a)
        .def("offset", &XLCell::offset, "row_offset"_a, "col_offset"_a)
        .def("get_string", &XLCell::getString)
        .def("get_formula", [](XLCell& self) { return self.formula().get(); })
        .def("set_formula",
             [](XLCell& self, const py::object& value) {
                 if (py::isinstance<py::str>(value)) {
                     self.formula() = py::cast<std::string>(value);
                 } else if (py::isinstance<XLFormula>(value)) {
                     self.formula() = py::cast<XLFormula>(value);
                 } else {
                     throw py::type_error("Unsupported type for formula assignment");
                 }
             })
        .def("clear_formula", [](XLCell& self) { self.formula().clear(); })
        .def("has_formula", [](const XLCell& self) { return self.hasFormula(); })
        .def("cell_reference", &XLCell::cellReference)
        .def("cell_format", &XLCell::cellFormat)
        .def("set_cell_format", &XLCell::setCellFormat, py::rv_policy::reference_internal)
        .def("set_style", &XLCell::setStyle, "style"_a, py::rv_policy::reference_internal)
        .def("add_comment", &XLCell::addComment, "text"_a, "author"_a = "")
        .def(
            "add_note",
            [](XLCell& self, std::string_view text, std::string_view author) -> XLCell& {
                return self.addNote(text, author);
            },
            "text"_a, "author"_a = "", py::rv_policy::reference_internal);
}
