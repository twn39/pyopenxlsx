#include <nanobind/ndarray.h>

#include <headers/XLContentTypes.hpp>
#include <headers/XLDrawing.hpp>
#include <variant>
#include <vector>

#include "bindings.hpp"

// Helper to access protected members
class XLXmlFilePublic : public XLXmlFile {
   public:
    using XLXmlFile::getXmlPath;
    using XLXmlFile::parentDoc;
    using XLXmlFile::xmlDocument;
};

namespace {
// Template trick to access private members
template <typename Tag, typename Tag::type M>
struct Rob {
    friend typename Tag::type get(Tag) { return M; }
};

struct XLDocumentContentTypes {
    typedef XLContentTypes XLDocument::* type;
};
template struct Rob<XLDocumentContentTypes, &XLDocument::m_contentTypes>;
XLContentTypes XLDocument::* get(XLDocumentContentTypes);

struct XLDocumentArchive {
    typedef IZipArchive XLDocument::* type;
};
template struct Rob<XLDocumentArchive, &XLDocument::m_archive>;
IZipArchive XLDocument::* get(XLDocumentArchive);

struct XLDocumentData {
    typedef std::list<XLXmlData> XLDocument::* type;
};
template struct Rob<XLDocumentData, &XLDocument::m_data>;
std::list<XLXmlData> XLDocument::* get(XLDocumentData);
}  // namespace

void add_image_to_worksheet(XLWorksheet& ws, py::bytes imageData, const std::string& extension,
                            uint32_t row, uint16_t col, double width, double height) {
    auto& ws_public = reinterpret_cast<XLXmlFilePublic&>(ws);
    XLDocument& doc = ws_public.parentDoc();

    // 1. Add image to document package
    std::string imgName = "image" + std::to_string(std::time(nullptr)) + "." + extension;
    std::string imgPath;
    {
        py::gil_scoped_release release;
        imgPath =
            doc.addImage(imgName, std::string(static_cast<const char*>(imageData.data()),
                                              imageData.size()));
    }

    // 2. Get worksheet drawing
    XLDrawing& drawing = ws.drawing();

    // 3. Add relationship from drawing to image
    // Note: drawing is in xl/drawings/, images are in xl/media/
    // Path should be ../media/imageN.ext
    std::string relPath = "../media/" + imgPath.substr(imgPath.find_last_of('/') + 1);

    std::string relId;
    {
        py::gil_scoped_release release;
        auto rel =
            drawing.relationships().addRelationship(XLRelationshipType::Image, relPath);
        relId = rel.id();

        // 4. Add image to drawing
        drawing.addImage(relId, imgName, "Image", row - 1, col - 1, (uint32_t)width,
                         (uint32_t)height);
    }
}


// Helper function to convert XLCellValue to py::object efficiently
// Note: GIL must be held when calling this function
inline py::object cell_value_to_pyobject(const XLCellValue& val) {
    XLValueType type = val.type();
    switch (type) {
        case XLValueType::Boolean:
            return py::cast(val.get<bool>());
        case XLValueType::Integer:
            return py::cast(val.get<int64_t>());
        case XLValueType::Float:
            return py::cast(val.get<double>());
        case XLValueType::String:
            return py::cast(val.get<std::string>());
        default:
            return py::none();
    }
}

// Internal structure to hold cell value data without Python objects
struct CellValueData {
    enum class Type { Empty, Boolean, Integer, Float, String };
    Type type = Type::Empty;
    bool boolVal = false;
    int64_t intVal = 0;
    double floatVal = 0.0;
    std::string strVal;

    static CellValueData from(const XLCellValue& val) {
        CellValueData data;
        switch (val.type()) {
            case XLValueType::Boolean:
                data.type = Type::Boolean;
                data.boolVal = val.get<bool>();
                break;
            case XLValueType::Integer:
                data.type = Type::Integer;
                data.intVal = val.get<int64_t>();
                break;
            case XLValueType::Float:
                data.type = Type::Float;
                data.floatVal = val.get<double>();
                break;
            case XLValueType::String:
                data.type = Type::String;
                data.strVal = val.get<std::string>();
                break;
            default:
                data.type = Type::Empty;
                break;
        }
        return data;
    }

    py::object to_python() const {
        switch (type) {
            case Type::Boolean:
                return py::cast(boolVal);
            case Type::Integer:
                return py::cast(intVal);
            case Type::Float:
                return py::cast(floatVal);
            case Type::String:
                return py::cast(strVal);
            default:
                return py::none();
        }
    }
};

// Get a single cell's value directly without creating a Cell object
py::object get_cell_value(XLWorksheet& ws, uint32_t row, uint16_t col) {
    CellValueData data;
    {
        py::gil_scoped_release release;
        XLCell cell = ws.cell(row, col);
        data = CellValueData::from(cell.value());
    }
    return data.to_python();
}

// Bulk read a specific range of cells - returns list[list[Any]]
py::list get_range_data(XLWorksheet& ws, uint32_t startRow, uint16_t startCol, uint32_t endRow,
                        uint16_t endCol) {
    // First, read all data without GIL
    std::vector<CellValueData> data;
    uint32_t numRows = endRow - startRow + 1;
    uint16_t numCols = endCol - startCol + 1;

    {
        py::gil_scoped_release release;

        // Pre-allocate everything in one block
        data.resize(numRows * numCols);

        // Iterate over the specified range
        for (uint32_t r = startRow; r <= endRow; ++r) {
            uint32_t baseIdx = (r - startRow) * numCols;
            XLRow row = ws.row(r);
            if (!row.empty()) {
                // Get all values from the row first
                std::vector<XLCellValue> values = row.values();

                // Extract values for the specified column range
                for (uint16_t c = startCol; c <= endCol; ++c) {
                    uint16_t colIdx = c - 1;  // values is 0-indexed
                    if (colIdx < values.size()) {
                        data[baseIdx + (c - startCol)] = CellValueData::from(values[colIdx]);
                    }
                    // For missing cells, data[baseIdx + offset] is already initialized as
                    // Type::Empty
                }
            }
        }
    }

    // Now convert to Python with GIL held
    py::list result;
    for (uint32_t r = 0; r < numRows; ++r) {
        py::list pyRow;
        uint32_t baseIdx = r * numCols;
        for (uint16_t c = 0; c < numCols; ++c) {
            pyRow.append(data[baseIdx + c].to_python());
        }
        result.append(pyRow);
    }

    return result;
}

// Bulk read all rows data - returns list[list[Any]]
py::list get_rows_data(XLWorksheet& ws) {
    // First, read all data without GIL
    std::vector<CellValueData> data;
    uint32_t rowCount = 0;
    uint16_t colCount = 0;

    {
        py::gil_scoped_release release;

        rowCount = ws.rowCount();
        colCount = ws.columnCount();

        // Pre-allocate everything in one shot
        data.resize(rowCount * colCount);

        // Iterate over rows
        for (uint32_t r = 1; r <= rowCount; ++r) {
            uint32_t baseIdx = (r - 1) * colCount;
            XLRow row = ws.row(r);
            if (!row.empty()) {
                // Use the implicit conversion to std::vector<XLCellValue>
                std::vector<XLCellValue> values = row.values();
                uint32_t valCount =
                    std::min(static_cast<uint32_t>(values.size()), static_cast<uint32_t>(colCount));
                for (uint32_t i = 0; i < valCount; ++i) {
                    data[baseIdx + i] = CellValueData::from(values[i]);
                }
            }
        }
    }

    // Now convert to Python with GIL held
    py::list result;
    for (uint32_t r = 0; r < rowCount; ++r) {
        py::list pyRow;
        uint32_t baseIdx = r * colCount;
        for (uint16_t c = 0; c < colCount; ++c) {
            pyRow.append(data[baseIdx + c].to_python());
        }
        result.append(pyRow);
    }

    return result;
}

// Get a single row's data as list[Any] - more efficient for row iteration
py::list get_row_values(XLWorksheet& ws, uint32_t rowNumber) {
    // First, read data without GIL
    std::vector<CellValueData> rowData;
    uint16_t colCount;

    {
        py::gil_scoped_release release;

        colCount = ws.columnCount();
        rowData.reserve(colCount);

        XLRow row = ws.row(rowNumber);
        if (!row.empty()) {
            // Use the implicit conversion to std::vector<XLCellValue>
            std::vector<XLCellValue> values = row.values();
            for (const auto& val : values) {
                rowData.push_back(CellValueData::from(val));
            }
        }

        // Pad with empty values if needed
        while (rowData.size() < colCount) {
            rowData.emplace_back();
        }
    }

    // Convert to Python with GIL held
    py::list result;
    for (const auto& cellData : rowData) {
        result.append(cellData.to_python());
    }

    return result;
}

// Optimized rows iterator - yields row values directly as list[Any]
class RowValuesIterator {
   public:
    RowValuesIterator(XLWorksheet& ws)
        : m_ws(ws), m_currentRow(1), m_maxRow(ws.rowCount()), m_colCount(ws.columnCount()) {}

    py::list next() {
        if (m_currentRow > m_maxRow) {
            throw py::stop_iteration();
        }

        // Read data without GIL
        std::vector<CellValueData> rowData;
        {
            py::gil_scoped_release release;

            rowData.reserve(m_colCount);
            XLRow row = m_ws.row(m_currentRow);
            if (!row.empty()) {
                // Use the implicit conversion to std::vector<XLCellValue>
                std::vector<XLCellValue> values = row.values();
                for (const auto& val : values) {
                    rowData.push_back(CellValueData::from(val));
                }
            }

            // Pad with empty values if needed
            while (rowData.size() < m_colCount) {
                rowData.emplace_back();
            }
        }

        // Convert to Python with GIL held
        py::list result;
        for (const auto& cellData : rowData) {
            result.append(cellData.to_python());
        }

        ++m_currentRow;
        return result;
    }

   private:
    XLWorksheet& m_ws;
    uint32_t m_currentRow;
    uint32_t m_maxRow;
    uint16_t m_colCount;
};

// Internal structure to hold cell value for writing
struct WriteCellData {
    enum class Type { Empty, Boolean, Integer, Float };
    Type type = Type::Empty;
    bool boolVal = false;
    int64_t intVal = 0;
    double floatVal = 0.0;
};

// Write a numpy array to a worksheet range cleanly using nanobind's ndarray
template <typename T>
void write_range_typed(XLWorksheet& ws, uint32_t startRow, uint16_t startCol,
                       py::ndarray<T, py::c_contig, py::device::cpu> b) {
    if (b.ndim() != 2) {
        throw std::runtime_error("Incompatible buffer dimension! Expected 2D array.");
    }

    uint32_t numRows = static_cast<uint32_t>(b.shape(0));
    uint16_t numCols = static_cast<uint16_t>(b.shape(1));

    const T* ptr = static_cast<const T*>(b.data());
    std::vector<T> data(ptr, ptr + numRows * numCols);

    // Now release GIL and write to worksheet using our copied data
    {
        py::gil_scoped_release release;
        for (uint32_t r = 0; r < numRows; ++r) {
            for (uint16_t c = 0; c < numCols; ++c) {
                T val = data[r * numCols + c];
                ws.cell(startRow + r, startCol + c).value() = val;
            }
        }
    }
}

// Read numeric data into a numpy array
py::ndarray<py::numpy, double, py::shape<-1, -1>> get_range_values(
    XLWorksheet& ws, uint32_t startRow, uint16_t startCol, uint32_t endRow, uint16_t endCol) {
    uint32_t numRows = endRow - startRow + 1;
    uint32_t numCols = endCol - startCol + 1;

    double* ptr = new double[numRows * numCols];

    {
        py::gil_scoped_release release;
        for (uint32_t r = 0; r < numRows; ++r) {
            XLRow row = ws.row(startRow + r);
            if (row.empty()) {
                for (uint32_t c = 0; c < numCols; ++c) {
                    ptr[r * numCols + c] = 0.0;
                }
                continue;
            }

            std::vector<XLCellValue> values = row.values();
            for (uint32_t c = 0; c < numCols; ++c) {
                uint32_t colIdx = startCol + c - 1;
                if (colIdx < values.size()) {
                    const auto& val = values[colIdx];
                    if (val.type() == XLValueType::Float) {
                        ptr[r * numCols + c] = val.get<double>();
                    } else if (val.type() == XLValueType::Integer) {
                        ptr[r * numCols + c] = (double)val.get<int64_t>();
                    } else {
                        ptr[r * numCols + c] = 0.0;
                    }
                } else {
                    ptr[r * numCols + c] = 0.0;
                }
            }
        }
    }

    py::capsule owner(ptr, [](void* p) noexcept { delete[] (double*)p; });
    size_t shape[2] = {numRows, numCols};
    return py::ndarray<py::numpy, double, py::shape<-1, -1>>(ptr, 2, shape, owner);
}

// Direct cell value setter - bypasses Python Cell object creation
// This is much faster for bulk writes as it avoids:
// 1. Creating Python Cell wrapper objects
// 2. WeakValueDictionary cache operations
// 3. Multiple Python/C++ boundary crossings
void set_cell_value(XLWorksheet& ws, uint32_t row, uint16_t col, py::object value) {
    if (value.is_none()) {
        py::gil_scoped_release release;
        ws.cell(row, col).value().clear();
    } else if (py::isinstance<py::bool_>(value)) {
        bool val = py::cast<bool>(value);
        py::gil_scoped_release release;
        ws.cell(row, col).value() = val;
    } else if (py::isinstance<py::int_>(value)) {
        int64_t val = py::cast<int64_t>(value);
        py::gil_scoped_release release;
        ws.cell(row, col).value() = val;
    } else if (py::isinstance<py::float_>(value)) {
        double val = py::cast<double>(value);
        py::gil_scoped_release release;
        ws.cell(row, col).value() = val;
    } else if (py::isinstance<py::str>(value)) {
        std::string val = py::cast<std::string>(value);
        py::gil_scoped_release release;
        ws.cell(row, col).value() = val;
    } else {
        throw py::type_error("Unsupported type for cell value");
    }
}

// Internal structure to hold any cell value for batch operations
struct BatchCellValue {
    enum class Type { Empty, Boolean, Integer, Float, String };
    Type type = Type::Empty;
    bool boolVal = false;
    int64_t intVal = 0;
    double floatVal = 0.0;
    std::string strVal;

    static BatchCellValue from_python(py::handle obj) {
        BatchCellValue val;
        if (obj.is_none()) {
            val.type = Type::Empty;
        } else if (py::isinstance<py::bool_>(obj)) {
            val.type = Type::Boolean;
            val.boolVal = py::cast<bool>(obj);
        } else if (py::isinstance<py::int_>(obj)) {
            val.type = Type::Integer;
            val.intVal = py::cast<int64_t>(obj);
        } else if (py::isinstance<py::float_>(obj)) {
            val.type = Type::Float;
            val.floatVal = py::cast<double>(obj);
        } else if (py::isinstance<py::str>(obj)) {
            val.type = Type::String;
            val.strVal = py::cast<std::string>(obj);
        } else {
            // Try to convert to string as fallback
            val.type = Type::String;
            val.strVal = py::cast<std::string>(py::str(obj));
        }
        return val;
    }

    void apply_to(XLCell& cell) const {
        switch (type) {
            case Type::Empty:
                cell.value().clear();
                break;
            case Type::Boolean:
                cell.value() = boolVal;
                break;
            case Type::Integer:
                cell.value() = intVal;
                break;
            case Type::Float:
                cell.value() = floatVal;
                break;
            case Type::String:
                cell.value() = strVal;
                break;
        }
    }
};

// Convert BatchCellValue to XLCellValue for use with OpenXLSX row assignment
inline XLCellValue to_xlcellvalue(const BatchCellValue& val) {
    switch (val.type) {
        case BatchCellValue::Type::Boolean:
            return XLCellValue(val.boolVal);
        case BatchCellValue::Type::Integer:
            return XLCellValue(val.intVal);
        case BatchCellValue::Type::Float:
            return XLCellValue(val.floatVal);
        case BatchCellValue::Type::String:
            return XLCellValue(val.strVal);
        default:
            return XLCellValue();
    }
}

// Write a 2D Python list to a worksheet range
// This is optimized for any Python data (strings, mixed types, etc.)
// Uses OpenXLSX's row batch assignment for better performance
void write_rows_data(XLWorksheet& ws, uint32_t startRow, uint16_t startCol, py::list rows) {
    // First pass: extract all data while holding GIL
    std::vector<std::vector<XLCellValue>> data;
    data.reserve(py::len(rows));

    for (auto row : rows) {
        std::vector<XLCellValue> rowData;
        py::list rowList = py::cast<py::list>(row);
        rowData.reserve(py::len(rowList));

        for (auto cell : rowList) {
            BatchCellValue bv = BatchCellValue::from_python(cell);
            rowData.push_back(to_xlcellvalue(bv));
        }
        data.push_back(std::move(rowData));
    }

    // Second pass: write to worksheet without GIL using row-level batch assignment
    {
        py::gil_scoped_release release;

        for (size_t r = 0; r < data.size(); ++r) {
            // Use OpenXLSX's optimized row assignment
            XLRow xlRow = ws.row(startRow + r);
            xlRow.values() = data[r];
        }
    }
}

// Write a single row of Python data
void write_row_data(XLWorksheet& ws, uint32_t row, uint16_t startCol, py::list values) {
    // Extract data while holding GIL
    std::vector<XLCellValue> data;
    data.reserve(py::len(values));

    for (auto val : values) {
        BatchCellValue bv = BatchCellValue::from_python(val);
        data.push_back(to_xlcellvalue(bv));
    }

    // Write without GIL using row-level batch assignment
    {
        py::gil_scoped_release release;

        XLRow xlRow = ws.row(row);
        xlRow.values() = data;
    }
}

// Batch set multiple cell values: [(row, col, value), ...]
void set_cells_batch(XLWorksheet& ws, py::list cells) {
    // Structure to hold row, col, value
    struct CellWrite {
        uint32_t row;
        uint16_t col;
        BatchCellValue value;
    };

    // Extract all data while holding GIL
    std::vector<CellWrite> writes;
    writes.reserve(py::len(cells));

    for (auto item : cells) {
        py::tuple t = py::cast<py::tuple>(item);
        if (py::len(t) != 3) {
            throw py::value_error("Each item must be a tuple of (row, col, value)");
        }
        CellWrite cw;
        cw.row = py::cast<uint32_t>(t[0]);
        cw.col = py::cast<uint16_t>(t[1]);
        cw.value = BatchCellValue::from_python(t[2]);
        writes.push_back(std::move(cw));
    }

    // Write without GIL
    {
        py::gil_scoped_release release;

        for (const auto& cw : writes) {
            XLCell cell = ws.cell(cw.row, cw.col);
            cw.value.apply_to(cell);
        }
    }
}

void init_worksheet(py::module_& m) {
    // Bind XLDrawingItem
    py::class_<XLDrawingItem>(m, "XLDrawingItem")
        .def("name", &XLDrawingItem::name)
        .def("description", &XLDrawingItem::description)
        .def("row", &XLDrawingItem::row)
        .def("col", &XLDrawingItem::col)
        .def("width", &XLDrawingItem::width)
        .def("height", &XLDrawingItem::height)
        .def("relationship_id", &XLDrawingItem::relationshipId);

    // Bind XLDrawing
    py::class_<XLDrawing>(m, "XLDrawing")
        .def("image_count", &XLDrawing::imageCount)
        .def("image", &XLDrawing::image, py::arg("index"))
        .def("add_image", &XLDrawing::addImage, py::arg("r_id"), py::arg("name"),
             py::arg("description"), py::arg("row"), py::arg("col"), py::arg("width"),
             py::arg("height"))
        .def("add_scaled_image", &XLDrawing::addScaledImage, py::arg("r_id"), py::arg("name"),
             py::arg("description"), py::arg("data"), py::arg("row"), py::arg("col"),
             py::arg("scaling_factor") = 1.0);

    // Bind XLColumn
    py::class_<XLColumn>(m, "XLColumn")
        .def("width", &XLColumn::width)
        .def("set_width", &XLColumn::setWidth, py::arg("width"))
        .def("is_hidden", &XLColumn::isHidden)
        .def("set_hidden", &XLColumn::setHidden, py::arg("state"))
        .def("format", &XLColumn::format)
        .def("set_format", &XLColumn::setFormat, py::arg("cellFormatIndex"));

    // Bind XLWorksheet
    py::class_<XLWorksheet>(m, "XLWorksheet")
        .def("name", &XLWorksheet::name)
        .def("set_name", &XLWorksheet::setName)
        .def("index", &XLWorksheet::index)
        .def("set_index", &XLWorksheet::setIndex)
        .def("visibility", &XLWorksheet::visibility)
        .def("set_visibility", &XLWorksheet::setVisibility)
        .def("is_active", &XLWorksheet::isActive)
        .def("set_active", &XLWorksheet::setActive)
        .def("row_count", &XLWorksheet::rowCount)
        .def("column_count", &XLWorksheet::columnCount)
        .def("has_drawing", &XLWorksheet::hasDrawing)
        .def("drawing", &XLWorksheet::drawing, py::rv_policy::reference_internal)
        .def("add_hyperlink", &XLWorksheet::addHyperlink, py::arg("cellRef"), py::arg("url"),
             py::arg("tooltip") = "")
        .def("add_internal_hyperlink", &XLWorksheet::addInternalHyperlink, py::arg("cellRef"),
             py::arg("location"), py::arg("tooltip") = "")
        .def(
            "cell",
            [](XLWorksheet& self, const std::string& ref) {
                py::gil_scoped_release release;
                return (XLCell)self.cell(ref);
            },
            py::keep_alive<0, 1>())
        .def(
            "cell",
            [](XLWorksheet& self, int row, int col) {
                py::gil_scoped_release release;
                return (XLCell)self.cell(row, col);
            },
            py::keep_alive<0, 1>())
        .def(
            "range",
            [](XLWorksheet& self, const std::string& address) {
                py::gil_scoped_release release;
                return self.range(address);
            },
            py::keep_alive<0, 1>())
        .def(
            "range",
            [](XLWorksheet& self, const std::string& topLeft, const std::string& bottomRight) {
                py::gil_scoped_release release;
                return self.range(XLCellReference(topLeft), XLCellReference(bottomRight));
            },
            py::keep_alive<0, 1>())
        .def("column", py::overload_cast<uint16_t>(&XLWorksheet::column, py::const_),
             py::keep_alive<0, 1>())
        .def("column", py::overload_cast<const std::string&>(&XLWorksheet::column, py::const_),
             py::keep_alive<0, 1>())
        .def(
            "merge_cells",
            [](XLWorksheet& self, const std::string& rangeReference, bool emptyHiddenCells) {
                py::gil_scoped_release release;
                self.mergeCells(rangeReference, emptyHiddenCells);
            },
            py::arg("rangeReference"), py::arg("emptyHiddenCells") = false)
        .def(
            "unmerge_cells",
            [](XLWorksheet& self, const std::string& rangeReference) {
                py::gil_scoped_release release;
                self.unmergeCells(rangeReference);
            },
            py::arg("rangeReference"))
        .def("column_format",
             py::overload_cast<const std::string&>(&XLWorksheet::getColumnFormat, py::const_))
        .def("merges", &XLWorksheet::merges, py::rv_policy::reference_internal)
        .def("set_column_format",
             py::overload_cast<const std::string&, XLStyleIndex>(&XLWorksheet::setColumnFormat),
             py::arg("column"), py::arg("cellFormatIndex"))
        .def("set_column_format",
             py::overload_cast<uint16_t, XLStyleIndex>(&XLWorksheet::setColumnFormat),
             py::arg("column"), py::arg("cellFormatIndex"))
        .def("row_format", &XLWorksheet::getRowFormat)
        .def("set_row_format", &XLWorksheet::setRowFormat, py::arg("row"),
             py::arg("cellFormatIndex"))
        .def(
            "protect_sheet",
            [](XLWorksheet& self, bool set) {
                py::gil_scoped_release release;
                self.protectSheet(set);
            },
            py::arg("set") = true)
        .def(
            "protect_objects",
            [](XLWorksheet& self, bool set) {
                py::gil_scoped_release release;
                self.protectObjects(set);
            },
            py::arg("set") = true)
        .def(
            "protect_scenarios",
            [](XLWorksheet& self, bool set) {
                py::gil_scoped_release release;
                self.protectScenarios(set);
            },
            py::arg("set") = true)
        .def("sheet_protected", &XLWorksheet::sheetProtected)
        .def("objects_protected", &XLWorksheet::objectsProtected)
        .def("scenarios_protected", &XLWorksheet::scenariosProtected)
        .def(
            "set_password",
            [](XLWorksheet& self, const std::string& password) {
                py::gil_scoped_release release;
                self.setPassword(password);
            },
            py::arg("password"))
        .def("clear_password",
             [](XLWorksheet& self) {
                 py::gil_scoped_release release;
                 self.clearPassword();
             })
        .def("password_hash", &XLWorksheet::passwordHash)
        .def("password_is_set", &XLWorksheet::passwordIsSet)
        .def("insert_columns_allowed", &XLWorksheet::insertColumnsAllowed)
        .def("insert_rows_allowed", &XLWorksheet::insertRowsAllowed)
        .def("delete_columns_allowed", &XLWorksheet::deleteColumnsAllowed)
        .def("delete_rows_allowed", &XLWorksheet::deleteRowsAllowed)
        .def("select_locked_cells_allowed", &XLWorksheet::selectLockedCellsAllowed)
        .def("select_unlocked_cells_allowed", &XLWorksheet::selectUnlockedCellsAllowed)
        .def(
            "set_insert_columns_allowed",
            [](XLWorksheet& self, bool set) {
                py::gil_scoped_release release;
                self.allowInsertColumns(set);
            },
            py::arg("set") = true)
        .def(
            "set_insert_rows_allowed",
            [](XLWorksheet& self, bool set) {
                py::gil_scoped_release release;
                self.allowInsertRows(set);
            },
            py::arg("set") = true)
        .def(
            "set_delete_columns_allowed",
            [](XLWorksheet& self, bool set) {
                py::gil_scoped_release release;
                self.allowDeleteColumns(set);
            },
            py::arg("set") = true)
        .def(
            "set_delete_rows_allowed",
            [](XLWorksheet& self, bool set) {
                py::gil_scoped_release release;
                self.allowDeleteRows(set);
            },
            py::arg("set") = true)
        .def(
            "set_select_locked_cells_allowed",
            [](XLWorksheet& self, bool set) {
                py::gil_scoped_release release;
                self.allowSelectLockedCells(set);
            },
            py::arg("set") = true)
        .def(
            "set_select_unlocked_cells_allowed",
            [](XLWorksheet& self, bool set) {
                py::gil_scoped_release release;
                self.allowSelectUnlockedCells(set);
            },
            py::arg("set") = true)
        .def("comments", &XLWorksheet::comments, py::rv_policy::reference_internal)
        .def("add_image", &add_image_to_worksheet, py::arg("image_data"), py::arg("extension"),
             py::arg("row") = 1, py::arg("col") = 1, py::arg("width") = 0, py::arg("height") = 0)
        // Bulk read APIs for performance optimization
        .def("get_rows_data", &get_rows_data,
             "Get all rows data as list[list[Any]] - optimized for bulk read")
        .def("get_row_values", &get_row_values, py::arg("row"),
             "Get a single row's values as list[Any]")
        .def("get_range_data", &get_range_data, py::arg("start_row"), py::arg("start_col"),
             py::arg("end_row"), py::arg("end_col"),
             "Get a range of cells as list[list[Any]] - optimized bulk read for specific range")
        .def("get_cell_value", &get_cell_value, py::arg("row"), py::arg("col"),
             "Get a single cell's value directly without creating a Cell object")
        .def(
            "iter_row_values", [](XLWorksheet& self) { return RowValuesIterator(self); },
            py::keep_alive<0, 1>(), "Iterate over rows, yielding each row's values as list[Any]")
        .def("write_range_data", &write_range_typed<double>, py::arg("start_row"),
             py::arg("start_col"), py::arg("data"),
             "Write a 2D numpy array or buffer to a worksheet range")
        .def("write_range_data", &write_range_typed<int64_t>, py::arg("start_row"),
             py::arg("start_col"), py::arg("data"))
        .def("write_range_data", &write_range_typed<bool>, py::arg("start_row"),
             py::arg("start_col"), py::arg("data"))
        .def("get_range_values", &get_range_values, py::arg("start_row"), py::arg("start_col"),
             py::arg("end_row"), py::arg("end_col"),
             "Read a range of numeric cells into a 2D numpy array of doubles")
        // Performance-optimized write APIs - bypass Python Cell object creation
        .def("set_cell_value", &set_cell_value, py::arg("row"), py::arg("col"), py::arg("value"),
             "Set a cell's value directly without creating a Cell object. "
             "10-20x faster than ws.cell(row, col).value = val for bulk operations")
        .def("write_rows_data", &write_rows_data, py::arg("start_row"), py::arg("start_col"),
             py::arg("rows"),
             "Write a 2D Python list to a worksheet range. "
             "Optimized for any Python data (strings, mixed types). "
             "For pure numeric data, use write_range_data with numpy for best performance")
        .def("write_row_data", &write_row_data, py::arg("row"), py::arg("start_col"),
             py::arg("values"), "Write a single row of Python data")
        .def("set_cells_batch", &set_cells_batch, py::arg("cells"),
             "Batch set multiple cell values: [(row, col, value), ...]. "
             "Efficient for non-contiguous cell updates");

    // Bind the RowValuesIterator
    py::class_<RowValuesIterator>(m, "RowValuesIterator")
        .def("__iter__", [](RowValuesIterator& self) -> RowValuesIterator& { return self; })
        .def("__next__", &RowValuesIterator::next);
}
