#include "bindings.hpp"

void init_cell(py::module& m) {
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
                return py::make_iterator<py::return_value_policy::copy>(self.begin(), self.end());
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
        .def_property(
            "value",
            [](const XLCell& self) -> py::object {
                py::gil_scoped_release release;
                const auto& valProxy = self.value();
                XLValueType type = valProxy.type();

                if (type == XLValueType::Boolean) {
                    bool val = valProxy.get<bool>();
                    py::gil_scoped_acquire acquire;
                    return py::cast(val);
                } else if (type == XLValueType::Integer) {
                    int64_t val = valProxy.get<int64_t>();
                    py::gil_scoped_acquire acquire;
                    return py::cast(val);
                } else if (type == XLValueType::Float) {
                    double val = valProxy.get<double>();
                    py::gil_scoped_acquire acquire;
                    return py::cast(val);
                } else if (type == XLValueType::String) {
                    std::string val = valProxy.get<std::string>();
                    py::gil_scoped_acquire acquire;
                    return py::cast(val);
                } else {
                    py::gil_scoped_acquire acquire;
                    return py::none();
                }
            },
            [](XLCell& self, py::object value) {
                if (value.is_none()) {
                    py::gil_scoped_release release;
                    self.value().clear();
                } else if (py::isinstance<py::bool_>(value)) {
                    bool val = value.cast<bool>();
                    py::gil_scoped_release release;
                    self.value() = val;
                } else if (py::isinstance<py::int_>(value)) {
                    int64_t val = value.cast<int64_t>();
                    py::gil_scoped_release release;
                    self.value() = val;
                } else if (py::isinstance<py::float_>(value)) {
                    double val = value.cast<double>();
                    py::gil_scoped_release release;
                    self.value() = val;
                } else if (py::isinstance<py::str>(value)) {
                    std::string val = value.cast<std::string>();
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
                     self.formula() = value.cast<std::string>();
                 } else if (py::isinstance<XLFormula>(value)) {
                     self.formula() = value.cast<XLFormula>();
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
