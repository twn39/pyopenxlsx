#include "bindings.hpp"
#include "internal_access.hpp"

void init_streams(py::module_& m) {
    py::class_<XLStreamWriter>(m, "XLStreamWriter")
        .def("is_stream_active", &XLStreamWriter::isStreamActive)
        .def("append_row", [](XLStreamWriter& self, py::list values) {
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
        }, py::arg("values"))
        .def("close", &XLStreamWriter::close);

    py::class_<XLStreamReader>(m, "XLStreamReader")
        .def("has_next", &XLStreamReader::hasNext)
        .def("next_row", [](XLStreamReader& self) {
            std::vector<XLCellValue> row = self.nextRow();
            py::list result;
            for (auto& cell : row) {
                CellData cd = CellData::from(cell);
                result.append(cd.to_python());
            }
            return result;
        })
        .def("current_row", &XLStreamReader::currentRow)
        // Make it an iterable in Python
        .def("__iter__", [](XLStreamReader& self) -> XLStreamReader& { return self; })
        .def("__next__", [](XLStreamReader& self) {
            if (!self.hasNext()) {
                throw py::stop_iteration();
            }
            std::vector<XLCellValue> row = self.nextRow();
            py::list result;
            for (auto& cell : row) {
                CellData cd = CellData::from(cell);
                result.append(cd.to_python());
            }
            return result;
        });
}
