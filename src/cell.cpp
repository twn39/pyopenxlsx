#include "internal_access.hpp"

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
        .def(py::init<const std::string&>())
        .def(py::init<uint32_t, uint16_t>())
        .def("address", &XLCellReference::address)
        .def("row", &XLCellReference::row)
        .def("column", &XLCellReference::column);

    // Bind XLCellRange
    py::class_<XLCellRange>(m, "XLCellRange")
        .def("address", &XLCellRange::address)
        .def("num_rows", &XLCellRange::numRows)
        .def("num_columns", &XLCellRange::numColumns)
        .def("clear", &XLCellRange::clear)
        .def(
            "__iter__",
            [](const XLCellRange& self) {
                return py::make_iterator(py::type<XLCellRange>(), "iterator", self.begin(),
                                         self.end());
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
            // FIX: Use CellData intermediate struct for safe GIL management.
            // Previous code had nested gil_scoped_release/acquire which was fragile.
            [](const XLCell& self) -> py::object {
                CellData data;
                {
                    py::gil_scoped_release release;
                    data = CellData::from(self.value());
                }
                return data.to_python();
            },
            [](XLCell& self, py::object value) {
                if (value.is_none()) {
                    py::gil_scoped_release release;
                    self.value().clear();
                } else if (py::isinstance<py::bool_>(value)) {
                    bool val = py::cast<bool>(value);
                    py::gil_scoped_release release;
                    self.value() = val;
                } else if (py::isinstance<py::int_>(value)) {
                    int64_t val = py::cast<int64_t>(value);
                    py::gil_scoped_release release;
                    self.value() = val;
                } else if (py::isinstance<py::float_>(value)) {
                    double val = py::cast<double>(value);
                    py::gil_scoped_release release;
                    self.value() = val;
                } else if (py::isinstance<py::str>(value)) {
                    std::string val = py::cast<std::string>(value);
                    py::gil_scoped_release release;
                    self.value() = val;
                } else {
                    throw py::type_error("Unsupported type for cell value");
                }
            })
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
        .def("has_formula", [](XLCell& self) { return self.hasFormula(); })
        .def("cell_reference", &XLCell::cellReference)
        .def("cell_format", &XLCell::cellFormat)
        .def("set_cell_format", &XLCell::setCellFormat);
}
