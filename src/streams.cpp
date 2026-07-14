#include "bindings.hpp"
#include "internal_access.hpp"

#include <optional>

using namespace OpenXLSX;
using namespace nanobind::literals;

namespace {

/**
 * @brief Convert a Python cell item into an XLStreamCell.
 *
 * Accepted forms:
 * - value
 * - (value, style_index)
 * - (value, style_index, formula)  — formula without leading '='
 * - dict with keys: value, style / style_index, formula
 */
XLStreamCell stream_cell_from_python(py::handle val) {
    if (py::isinstance<py::dict>(val)) {
        py::dict d = py::cast<py::dict>(val);
        XLCellValue value;
        if (d.contains("value")) {
            value = CellData::from_python(d["value"]).to_xlcellvalue();
        }
        std::optional<XLStyleIndex> style;
        if (d.contains("style_index")) {
            style = py::cast<XLStyleIndex>(d["style_index"]);
        } else if (d.contains("style")) {
            style = py::cast<XLStyleIndex>(d["style"]);
        }
        std::optional<std::string> formula;
        if (d.contains("formula") && !d["formula"].is_none()) {
            formula = py::cast<std::string>(d["formula"]);
        }
        if (style && formula) {
            return XLStreamCell(std::move(value), *style, std::move(*formula));
        }
        if (style) {
            return XLStreamCell(std::move(value), *style);
        }
        if (formula) {
            return XLStreamCell::withFormula(std::move(*formula), std::move(value));
        }
        return XLStreamCell(std::move(value));
    }

    if (py::isinstance<py::tuple>(val) || py::isinstance<py::list>(val)) {
        py::sequence seq = py::cast<py::sequence>(val);
        const size_t n = py::len(seq);
        if (n == 2) {
            CellData cd = CellData::from_python(seq[0]);
            auto styleIndex = py::cast<XLStyleIndex>(seq[1]);
            return XLStreamCell(cd.to_xlcellvalue(), styleIndex);
        }
        if (n == 3) {
            CellData cd = CellData::from_python(seq[0]);
            auto styleIndex = py::cast<XLStyleIndex>(seq[1]);
            std::string formula = py::cast<std::string>(seq[2]);
            return XLStreamCell(cd.to_xlcellvalue(), styleIndex, std::move(formula));
        }
    }

    CellData cd = CellData::from_python(val);
    return XLStreamCell(cd.to_xlcellvalue());
}

std::vector<XLStreamCell> stream_row_from_python(py::handle row) {
    py::sequence values = py::cast<py::sequence>(row);
    std::vector<XLStreamCell> data;
    data.reserve(py::len(values));
    for (auto val : values) {
        data.push_back(stream_cell_from_python(val));
    }
    return data;
}

py::dict row_opts_to_dict(const XLStreamRowOptsView& opts) {
    py::dict d;
    if (opts.height) d["height"] = *opts.height;
    if (opts.hidden) d["hidden"] = *opts.hidden;
    if (opts.outlineLevel) d["outline_level"] = *opts.outlineLevel;
    if (opts.styleIndex) d["style_index"] = *opts.styleIndex;
    d["is_synthetic_empty"] = opts.isSyntheticEmpty;
    return d;
}

py::dict cell_view_to_dict(const XLStreamCellView& cell) {
    py::dict d;
    d["value"] = CellData::from(cell.value).to_python();
    d["column"] = cell.column;
    if (cell.formula) d["formula"] = *cell.formula;
    if (cell.styleIndex) d["style_index"] = *cell.styleIndex;
    return d;
}

}  // namespace

void init_streams(py::module_& m) {
    py::enum_<XLStreamEmptyRowPolicy>(m, "XLStreamEmptyRowPolicy")
        .value("SkipMissingRows", XLStreamEmptyRowPolicy::SkipMissingRows)
        .value("EmitEmptyRows", XLStreamEmptyRowPolicy::EmitEmptyRows);

    py::class_<XLStreamReadOptions>(m, "XLStreamReadOptions")
        .def(py::init<>())
        .def_rw("empty_rows", &XLStreamReadOptions::emptyRows)
        .def_rw("apply_number_formats", &XLStreamReadOptions::applyNumberFormats);

    py::class_<XLStreamRowOpts>(m, "XLStreamRowOpts")
        .def(py::init<>())
        .def_prop_rw(
            "height",
            [](const XLStreamRowOpts& self) -> py::object {
                if (self.height) return py::cast(*self.height);
                return py::none();
            },
            [](XLStreamRowOpts& self, py::object val) {
                if (val.is_none()) self.height = std::nullopt;
                else self.height = py::cast<double>(val);
            })
        .def_prop_rw(
            "hidden",
            [](const XLStreamRowOpts& self) -> py::object {
                if (self.hidden) return py::cast(*self.hidden);
                return py::none();
            },
            [](XLStreamRowOpts& self, py::object val) {
                if (val.is_none()) self.hidden = std::nullopt;
                else self.hidden = py::cast<bool>(val);
            })
        .def_prop_rw(
            "outline_level",
            [](const XLStreamRowOpts& self) -> py::object {
                if (self.outlineLevel) return py::cast(*self.outlineLevel);
                return py::none();
            },
            [](XLStreamRowOpts& self, py::object val) {
                if (val.is_none()) self.outlineLevel = std::nullopt;
                else self.outlineLevel = py::cast<uint8_t>(val);
            })
        .def_prop_rw(
            "style_index",
            [](const XLStreamRowOpts& self) -> py::object {
                if (self.styleIndex) return py::cast(*self.styleIndex);
                return py::none();
            },
            [](XLStreamRowOpts& self, py::object val) {
                if (val.is_none()) self.styleIndex = std::nullopt;
                else self.styleIndex = py::cast<XLStyleIndex>(val);
            });

    py::class_<XLStreamWriter>(m, "XLStreamWriter")
        .def_prop_ro("is_active", &XLStreamWriter::isStreamActive)
        .def("is_stream_active", &XLStreamWriter::isStreamActive)
        .def_prop_ro("last_row", &XLStreamWriter::lastRow)
        .def_prop_ro("max_column", &XLStreamWriter::maxColumn)
        .def(
            "append_row",
            [](XLStreamWriter& self, py::object values, py::object opts) {
                auto data = stream_row_from_python(values);
                {
                    py::gil_scoped_release release;
                    if (opts.is_none()) {
                        self.appendRow(data);
                    } else {
                        self.appendRow(data, py::cast<XLStreamRowOpts&>(opts));
                    }
                }
            },
            "values"_a, "opts"_a = py::none(),
            "Append a row. Items may be values, (value, style), (value, style, formula), or dicts.")
        .def(
            "append_rows",
            [](XLStreamWriter& self, py::iterable rows) {
                constexpr size_t kChunkSize = 256;
                std::vector<std::vector<XLStreamCell>> chunk;
                chunk.reserve(kChunkSize);

                auto flush_chunk = [&]() {
                    if (chunk.empty()) return;
                    {
                        py::gil_scoped_release release;
                        for (auto& rowData : chunk) self.appendRow(rowData);
                    }
                    chunk.clear();
                };

                for (auto row : rows) {
                    chunk.push_back(stream_row_from_python(row));
                    if (chunk.size() >= kChunkSize) flush_chunk();
                }
                flush_chunk();
            },
            "rows"_a)
        .def(
            "set_row",
            [](XLStreamWriter& self, uint32_t row, uint16_t start_col, py::object values,
               py::object opts) {
                auto data = stream_row_from_python(values);
                {
                    py::gil_scoped_release release;
                    if (opts.is_none()) {
                        self.setRow(row, start_col, data);
                    } else {
                        self.setRow(row, start_col, data, py::cast<XLStreamRowOpts&>(opts));
                    }
                }
            },
            "row"_a, "start_col"_a, "values"_a, "opts"_a = py::none(),
            "Write a row at an explicit 1-based row index (strictly increasing).")
        .def(
            "set_row_ref",
            [](XLStreamWriter& self, const std::string& cell_ref, py::object values,
               py::object opts) {
                auto data = stream_row_from_python(values);
                {
                    py::gil_scoped_release release;
                    if (opts.is_none()) {
                        self.setRow(cell_ref, data);
                    } else {
                        self.setRow(cell_ref, data, py::cast<XLStreamRowOpts&>(opts));
                    }
                }
            },
            "cell_ref"_a, "values"_a, "opts"_a = py::none(),
            "Write a row starting at a cell reference such as 'C10'.")
        .def("flush", &XLStreamWriter::flush, "Alias for close().")
        .def("close", &XLStreamWriter::close)
        .def("__enter__", [](py::handle self) -> py::object { return py::borrow(self); })
        .def(
            "__exit__",
            [](XLStreamWriter& self, py::object, py::object, py::object) { self.close(); },
            "exc_type"_a = py::none(), "exc_value"_a = py::none(), "traceback"_a = py::none());

    py::class_<XLStreamReader>(m, "XLStreamReader")
        .def("has_next", &XLStreamReader::hasNext)
        .def(
            "next_row",
            [](XLStreamReader& self) {
                std::vector<XLCellValue> row;
                {
                    py::gil_scoped_release release;
                    row = self.nextRow();
                }
                py::list result;
                for (auto& cell : row) {
                    result.append(CellData::from(cell).to_python());
                }
                return result;
            })
        .def(
            "next_row_detailed",
            [](XLStreamReader& self) {
                std::vector<XLStreamCellView> row;
                {
                    py::gil_scoped_release release;
                    row = self.nextRowDetailed();
                }
                py::list result;
                for (const auto& cell : row) {
                    result.append(cell_view_to_dict(cell));
                }
                return result;
            },
            "Next row as list of dicts: value, column, optional formula/style_index.")
        .def(
            "next_row_strings",
            [](XLStreamReader& self) {
                std::vector<std::string> row;
                {
                    py::gil_scoped_release release;
                    row = self.nextRowStrings();
                }
                py::list result;
                for (auto& s : row) result.append(s);
                return result;
            },
            "Next row as display strings (respects apply_number_formats when set).")
        .def("current_row", &XLStreamReader::currentRow)
        .def_prop_ro("current_row_index", &XLStreamReader::currentRow)
        .def(
            "current_row_opts",
            [](const XLStreamReader& self) { return row_opts_to_dict(self.currentRowOpts()); },
            "Row attributes for the last returned row.")
        .def_prop_ro(
            "last_error",
            [](const XLStreamReader& self) -> py::object {
                const auto& err = self.lastError();
                if (err.empty()) return py::none();
                return py::cast(err);
            })
        .def_prop_ro(
            "options",
            [](const XLStreamReader& self) { return self.options(); })
        .def("close", &XLStreamReader::close)
        .def("__enter__", [](py::handle self) -> py::object { return py::borrow(self); })
        .def(
            "__exit__",
            [](XLStreamReader& self, py::object, py::object, py::object) { self.close(); },
            "exc_type"_a = py::none(), "exc_value"_a = py::none(), "traceback"_a = py::none())
        .def("__iter__", [](py::handle self) -> py::object { return py::borrow(self); })
        .def("__next__", [](XLStreamReader& self) {
            if (!self.hasNext()) {
                throw py::stop_iteration();
            }
            std::vector<XLCellValue> row;
            {
                py::gil_scoped_release release;
                row = self.nextRow();
            }
            py::list result;
            for (auto& cell : row) {
                result.append(CellData::from(cell).to_python());
            }
            return result;
        });
}
