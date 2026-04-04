#include "bindings.hpp"
#include "internal_access.hpp"

void init_streams(py::module_& m) {
    py::class_<XLStreamWriter>(m, "XLStreamWriter")
        .def_prop_ro("is_active", &XLStreamWriter::isStreamActive)
        .def(
            "append_row",
            [](XLStreamWriter& self, py::list values) {
                std::vector<XLStreamCell> data;
                data.reserve(py::len(values));
                for (auto val : values) {
                    if (py::isinstance<py::tuple>(val)) {
                        py::tuple t = py::cast<py::tuple>(val);
                        if (py::len(t) == 2) {
                            CellData cd = CellData::from_python(t[0]);
                            uint32_t styleIndex = py::cast<uint32_t>(t[1]);
                            data.push_back(XLStreamCell(cd.to_xlcellvalue(), styleIndex));
                            continue;
                        }
                    }
                    CellData cd = CellData::from_python(val);
                    data.push_back(XLStreamCell(cd.to_xlcellvalue()));
                }
                {
                    py::gil_scoped_release release;
                    self.appendRow(data);
                }
            },
            py::arg("values"))
        .def(
            "append_rows",
            [](XLStreamWriter& self, py::iterable rows) {
                for (auto row : rows) {
                    std::vector<XLStreamCell> data;
                    py::list                  values = py::cast<py::list>(row);
                    data.reserve(py::len(values));
                    for (auto val : values) {
                        if (py::isinstance<py::tuple>(val)) {
                            py::tuple t = py::cast<py::tuple>(val);
                            if (py::len(t) == 2) {
                                CellData cd = CellData::from_python(t[0]);
                                uint32_t styleIndex = py::cast<uint32_t>(t[1]);
                                data.push_back(XLStreamCell(cd.to_xlcellvalue(), styleIndex));
                                continue;
                            }
                        }
                        CellData cd = CellData::from_python(val);
                        data.push_back(XLStreamCell(cd.to_xlcellvalue()));
                    }
                    {
                        py::gil_scoped_release release;
                        self.appendRow(data);
                    }
                }
            },
            py::arg("rows"))
        .def("close", &XLStreamWriter::close)
        // FIX: __enter__ must return the *same* Python object (via py::borrow), not a raw C++
        // pointer. Returning &self caused nanobind to create a second Python wrapper around the
        // same C++ object, resulting in double cleanup() / double-free when either wrapper was GC'd.
        .def("__enter__", [](py::handle self) -> py::object { return py::borrow(self); })
        .def("__exit__", [](XLStreamWriter& self, py::object, py::object, py::object) { self.close(); },
             py::arg("exc_type") = py::none(), py::arg("exc_value") = py::none(), py::arg("traceback") = py::none());

    py::class_<XLStreamReader>(m, "XLStreamReader")
        .def("has_next", &XLStreamReader::hasNext)
        .def("next_row",
             [](XLStreamReader& self) {
                 std::vector<XLCellValue> row = self.nextRow();
                 py::list                 result;
                 for (auto& cell : row) {
                     CellData cd = CellData::from(cell);
                     result.append(cd.to_python());
                 }
                 return result;
             })
        .def("current_row", &XLStreamReader::currentRow)
        .def_prop_ro("current_row_index", &XLStreamReader::currentRow)
        .def("close", &XLStreamReader::close)
        // FIX: same as XLStreamWriter — return py::borrow(self) so Python gets the *existing*
        // wrapper object, not a second C++ wrapper that would double-free on destruction.
        .def("__enter__", [](py::handle self) -> py::object { return py::borrow(self); })
        .def("__exit__", [](XLStreamReader& self, py::object, py::object, py::object) { self.close(); },
             py::arg("exc_type") = py::none(), py::arg("exc_value") = py::none(), py::arg("traceback") = py::none())
        // FIX: __iter__ must also return the same Python object, not a new C++ reference wrapper.
        .def("__iter__", [](py::handle self) -> py::object { return py::borrow(self); })
        .def("__next__", [](XLStreamReader& self) {
            if (!self.hasNext()) {
                throw py::stop_iteration();
            }
            std::vector<XLCellValue> row = self.nextRow();
            py::list                 result;
            for (auto& cell : row) {
                CellData cd = CellData::from(cell);
                result.append(cd.to_python());
            }
            return result;
        });
}
