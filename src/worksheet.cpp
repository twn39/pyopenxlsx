#include <nanobind/ndarray.h>

#include <headers/XLSlicerCollection.hpp>
#include <variant>
#include <vector>

#include "internal_access.hpp"

using namespace OpenXLSX;
using namespace nanobind::literals;


void add_image_to_worksheet(XLWorksheet& ws, py::bytes imageData, const std::string& extension,
                            uint32_t row, uint16_t col, double width, double height) {
    // Use public APIs instead of Rob hack / reinterpret_cast
    XLDocument& doc = get_parent_doc(ws);

    // 1. Add image to document package
    auto& archive = get_archive(doc);
    int imgNum = 1;
    while (archive.hasEntry("xl/media/image" + std::to_string(imgNum) + "." + extension)) {
        ++imgNum;
    }
    std::string imgName = "image" + std::to_string(imgNum) + "." + extension;

    // FIX: Copy py::bytes data BEFORE releasing GIL (accessing Python buffer requires GIL)
    std::string imgDataStr(static_cast<const char*>(imageData.data()), imageData.size());
    std::string imgPath;
    {
        py::gil_scoped_release release;
        imgPath = doc.addImage(imgName, std::move(imgDataStr));
    }

    // 2. Get worksheet drawing
    XLDrawing& drawing = ws.drawing();

    // 3. Add relationship from drawing to image
    std::string relPath = "../media/" + imgPath.substr(imgPath.find_last_of('/') + 1);

    std::string relId;
    {
        py::gil_scoped_release release;
        auto rel = drawing.relationships().addRelationship(XLRelationshipType::Image, relPath);
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

// Get a single cell's value directly without creating a Cell object
py::object get_cell_value(XLWorksheet& ws, uint32_t row, uint16_t col) {
    Expects(row >= 1 && row <= kExcelMaxRows);
    Expects(col >= 1 && col <= kExcelMaxCols);

    CellData data;
    {
        py::gil_scoped_release release;
        XLCell cell = ws.cell(row, col);
        data = CellData::from(cell.value());
    }
    return data.to_python();
}

// Bulk read a specific range of cells - returns list[list[Any]]
py::list get_range_data(XLWorksheet& ws, uint32_t startRow, uint16_t startCol, uint32_t endRow,
                        uint16_t endCol) {
    Expects(startRow >= 1 && startRow <= kExcelMaxRows);
    Expects(endRow >= startRow && endRow <= kExcelMaxRows);
    Expects(startCol >= 1 && startCol <= kExcelMaxCols);
    Expects(endCol >= startCol && endCol <= kExcelMaxCols);

    auto numRows = gsl::narrow<uint32_t>(endRow - startRow + 1);
    auto numCols = gsl::narrow<uint16_t>(endCol - startCol + 1);

    // First, read all data without GIL
    std::vector<CellData> data;

    {
        py::gil_scoped_release release;

        data.resize(static_cast<size_t>(numRows) * numCols);

        for (uint32_t r = startRow; r <= endRow; ++r) {
            size_t baseIdx = static_cast<size_t>(r - startRow) * numCols;
            XLRow row = ws.row(r);
            if (!row.empty()) {
                std::vector<XLCellValue> values = row.values();

                for (uint16_t c = startCol; c <= endCol; ++c) {
                    auto colIdx = gsl::narrow<size_t>(c - 1);  // values is 0-indexed
                    if (colIdx < values.size()) {
                        data[baseIdx + (c - startCol)] = CellData::from(values[colIdx]);
                    }
                }
            }
        }
    }

    // Now convert to Python with GIL held
    py::list result;
    for (uint32_t r = 0; r < numRows; ++r) {
        py::list pyRow;
        size_t baseIdx = static_cast<size_t>(r) * numCols;
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
    std::vector<CellData> data;
    uint32_t rowCount = 0;
    uint16_t colCount = 0;

    {
        py::gil_scoped_release release;

        rowCount = ws.rowCount();
        colCount = ws.columnCount();

        data.resize(static_cast<size_t>(rowCount) * colCount);

        for (uint32_t r = 1; r <= rowCount; ++r) {
            size_t baseIdx = static_cast<size_t>(r - 1) * colCount;
            XLRow row = ws.row(r);
            if (!row.empty()) {
                std::vector<XLCellValue> values = row.values();
                auto valCount =
                    std::min(static_cast<uint32_t>(values.size()), static_cast<uint32_t>(colCount));
                for (uint32_t i = 0; i < valCount; ++i) {
                    data[baseIdx + i] = CellData::from(values[i]);
                }
            }
        }
    }

    // Now convert to Python with GIL held
    // Note: nanobind py::list has no size-based constructor, so we use append
    py::list result;
    for (uint32_t r = 0; r < rowCount; ++r) {
        py::list pyRow;
        size_t baseIdx = static_cast<size_t>(r) * colCount;
        for (uint16_t c = 0; c < colCount; ++c) {
            pyRow.append(data[baseIdx + c].to_python());
        }
        result.append(std::move(pyRow));
    }

    return result;
}

// Get a single row's data as list[Any] - more efficient for row iteration
py::list get_row_values(XLWorksheet& ws, uint32_t rowNumber) {
    Expects(rowNumber >= 1 && rowNumber <= kExcelMaxRows);

    // First, read data without GIL
    std::vector<CellData> rowData;
    uint16_t colCount;

    {
        py::gil_scoped_release release;

        colCount = ws.columnCount();
        rowData.reserve(colCount);

        XLRow row = ws.row(rowNumber);
        if (!row.empty()) {
            std::vector<XLCellValue> values = row.values();
            for (const auto& val : values) {
                rowData.push_back(CellData::from(val));
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

// Write a numpy array to a worksheet range cleanly using nanobind's ndarray
template <typename T>
void write_range_typed(XLWorksheet& ws, uint32_t startRow, uint16_t startCol,
                       py::ndarray<T, py::c_contig, py::device::cpu> b) {
    if (b.ndim() != 2) {
        throw std::runtime_error("Incompatible buffer dimension! Expected 2D array.");
    }

    auto numRows = gsl::narrow<uint32_t>(b.shape(0));
    auto numCols = gsl::narrow<uint16_t>(b.shape(1));

    Expects(startRow >= 1 && startRow + numRows - 1 <= kExcelMaxRows);
    Expects(startCol >= 1 && startCol + numCols - 1 <= kExcelMaxCols);

    const T* ptr = static_cast<const T*>(b.data());
    std::vector<T> data(ptr, ptr + static_cast<size_t>(numRows) * numCols);

    // Now release GIL and write to worksheet using our copied data
    {
        py::gil_scoped_release release;
        for (uint32_t r = 0; r < numRows; ++r) {
            for (uint16_t c = 0; c < numCols; ++c) {
                T val = data[static_cast<size_t>(r) * numCols + c];
                ws.cell(startRow + r, startCol + c).value() = val;
            }
        }
    }
}

// Read numeric data into a numpy array
// FIX: Use unique_ptr for exception-safe memory management (was: raw new with delayed capsule)
py::ndarray<py::numpy, double, py::shape<-1, -1>> get_range_values(
    XLWorksheet& ws, uint32_t startRow, uint16_t startCol, uint32_t endRow, uint16_t endCol) {
    Expects(startRow >= 1 && startRow <= kExcelMaxRows);
    Expects(endRow >= startRow && endRow <= kExcelMaxRows);
    Expects(startCol >= 1 && startCol <= kExcelMaxCols);
    Expects(endCol >= startCol && endCol <= kExcelMaxCols);

    auto numRows = gsl::narrow<size_t>(endRow - startRow + 1);
    auto numCols = gsl::narrow<size_t>(endCol - startCol + 1);

    // FIX: Use unique_ptr so memory is freed on exception before capsule takes ownership
    auto uptr = std::make_unique<double[]>(numRows * numCols);
    gsl::span<double> buf(uptr.get(), numRows * numCols);

    {
        py::gil_scoped_release release;
        for (size_t r = 0; r < numRows; ++r) {
            auto rowSpan = buf.subspan(r * numCols, numCols);
            XLRow row = ws.row(gsl::narrow<uint32_t>(startRow + r));
            if (row.empty()) {
                std::fill(rowSpan.begin(), rowSpan.end(), 0.0);
                continue;
            }

            std::vector<XLCellValue> values = row.values();
            for (size_t c = 0; c < numCols; ++c) {
                auto colIdx = gsl::narrow<size_t>(startCol + c - 1);
                if (colIdx < values.size()) {
                    const auto& val = values[colIdx];
                    if (val.type() == XLValueType::Float) {
                        rowSpan[c] = val.get<double>();
                    } else if (val.type() == XLValueType::Integer) {
                        rowSpan[c] = static_cast<double>(val.get<int64_t>());
                    } else {
                        rowSpan[c] = 0.0;
                    }
                } else {
                    rowSpan[c] = 0.0;
                }
            }
        }
    }

    // Transfer ownership from unique_ptr to capsule
    double* ptr = uptr.release();
    py::capsule owner(ptr, [](void* p) noexcept { delete[] (double*)p; });
    size_t shape[2] = {numRows, numCols};
    return py::ndarray<py::numpy, double, py::shape<-1, -1>>(ptr, 2, shape, owner);
}

// Direct cell value setter - bypasses Python Cell object creation
void set_cell_value(XLWorksheet& ws, uint32_t row, uint16_t col, py::object value) {
    Expects(row >= 1 && row <= kExcelMaxRows);
    Expects(col >= 1 && col <= kExcelMaxCols);

    // Convert Python value while holding the GIL (all Python API calls happen here).
    // FIX (P9): previously each branch had its own gil_scoped_release, causing up to
    // 5 unnecessary lock/unlock round-trips per call.  Using CellData::from_python()
    // also fixes a functional gap: datetime, numpy scalars, and XLRichText were
    // silently rejected here with TypeError while Cell.value accepted them.
    CellData cd = CellData::from_python(value);

    // One contiguous GIL-free window for the C++ DOM write.
    // Note: ws.cell() is stored in a named variable — apply_to() takes XLCell&
    // and cannot bind directly to the temporary returned by ws.cell().
    {
        py::gil_scoped_release release;
        XLCell cell = ws.cell(row, col);
        cd.apply_to(cell);
    }
}

// Write a 2D Python list to a worksheet range
// Uses OpenXLSX's row batch assignment for better performance
void write_rows_data(XLWorksheet& ws, uint32_t startRow, uint16_t startCol, py::list rows) {
    Expects(startRow >= 1 && startRow <= kExcelMaxRows);
    Expects(startCol >= 1 && startCol <= kExcelMaxCols);

    // First pass: extract all data while holding GIL, convert to XLCellValue directly
    std::vector<std::vector<XLCellValue>> data;
    data.reserve(py::len(rows));

    for (auto row : rows) {
        std::vector<XLCellValue> rowData;
        py::list rowList = py::cast<py::list>(row);
        rowData.reserve(py::len(rowList));

        for (auto cell : rowList) {
            CellData cd = CellData::from_python(cell);
            rowData.push_back(cd.to_xlcellvalue());
        }
        data.push_back(std::move(rowData));
    }

    // Second pass: write to worksheet without GIL using row-level batch assignment
    {
        py::gil_scoped_release release;

        for (size_t r = 0; r < data.size(); ++r) {
            XLRow xlRow = ws.row(gsl::narrow<uint32_t>(startRow + r));
            xlRow.values() = data[r];
        }
    }
}

// Write a single row of Python data
void write_row_data(XLWorksheet& ws, uint32_t row, uint16_t startCol, py::list values) {
    Expects(row >= 1 && row <= kExcelMaxRows);
    Expects(startCol >= 1 && startCol <= kExcelMaxCols);

    // Extract data while holding GIL
    std::vector<XLCellValue> data;
    data.reserve(py::len(values));

    for (auto val : values) {
        CellData cd = CellData::from_python(val);
        data.push_back(cd.to_xlcellvalue());
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
        CellData value;
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
        Expects(cw.row >= 1 && cw.row <= kExcelMaxRows);
        Expects(cw.col >= 1 && cw.col <= kExcelMaxCols);
        cw.value = CellData::from_python(t[2]);
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
    // Bind XLVectorShapeType
    py::enum_<XLVectorShapeType>(m, "XLVectorShapeType")
        .value("Rectangle", XLVectorShapeType::Rectangle)
        .value("Ellipse", XLVectorShapeType::Ellipse)
        .value("Line", XLVectorShapeType::Line)
        .value("Triangle", XLVectorShapeType::Triangle)
        .value("RightTriangle", XLVectorShapeType::RightTriangle)
        .value("Arrow", XLVectorShapeType::Arrow)
        .value("Diamond", XLVectorShapeType::Diamond)
        .value("Parallelogram", XLVectorShapeType::Parallelogram)
        .value("Hexagon", XLVectorShapeType::Hexagon)
        .value("Star4", XLVectorShapeType::Star4)
        .value("Star5", XLVectorShapeType::Star5)
        .value("Star16", XLVectorShapeType::Star16)
        .value("Star24", XLVectorShapeType::Star24)
        .value("Heart", XLVectorShapeType::Heart)
        .value("SmileyFace", XLVectorShapeType::SmileyFace)
        .value("Cloud", XLVectorShapeType::Cloud)
        .value("Donut", XLVectorShapeType::Donut)
        .value("Ribbon", XLVectorShapeType::Ribbon)
        .value("Sun", XLVectorShapeType::Sun)
        .value("Moon", XLVectorShapeType::Moon)
        .value("LightningBolt", XLVectorShapeType::LightningBolt)
        .value("FlowChartProcess", XLVectorShapeType::FlowChartProcess)
        .value("FlowChartDecision", XLVectorShapeType::FlowChartDecision)
        .value("FlowChartDocument", XLVectorShapeType::FlowChartDocument)
        .value("FlowChartData", XLVectorShapeType::FlowChartData);

    py::class_<XLVectorShapeOptions>(m, "XLVectorShapeOptions")
        .def(py::init<>())
        .def_rw("type", &XLVectorShapeOptions::type)
        .def_rw("name", &XLVectorShapeOptions::name)
        .def_rw("text", &XLVectorShapeOptions::text)
        .def_rw("fill_color", &XLVectorShapeOptions::fillColor)
        .def_rw("line_color", &XLVectorShapeOptions::lineColor)
        .def_rw("line_width", &XLVectorShapeOptions::lineWidth)
        .def_rw("width", &XLVectorShapeOptions::width)
        .def_rw("height", &XLVectorShapeOptions::height)
        .def_rw("offset_x", &XLVectorShapeOptions::offsetX)
        .def_rw("offset_y", &XLVectorShapeOptions::offsetY)
        .def_rw("end_row", &XLVectorShapeOptions::endRow)
        .def_rw("end_col", &XLVectorShapeOptions::endCol)
        .def_rw("end_offset_x", &XLVectorShapeOptions::endOffsetX)
        .def_rw("end_offset_y", &XLVectorShapeOptions::endOffsetY)
        .def_rw("rotation", &XLVectorShapeOptions::rotation)
        .def_rw("flip_h", &XLVectorShapeOptions::flipH)
        .def_rw("flip_v", &XLVectorShapeOptions::flipV)
        .def_rw("line_dash", &XLVectorShapeOptions::lineDash)
        .def_rw("arrow_start", &XLVectorShapeOptions::arrowStart)
        .def_rw("arrow_end", &XLVectorShapeOptions::arrowEnd)
        .def_rw("rich_text", &XLVectorShapeOptions::richText)
        .def_rw("horz_align", &XLVectorShapeOptions::horzAlign)
        .def_rw("vert_align", &XLVectorShapeOptions::vertAlign)
        .def_rw("macro", &XLVectorShapeOptions::macro);

    // Bind XLDrawingItem
    py::class_<XLDrawingItem>(m, "XLDrawingItem")
        .def("name", &XLDrawingItem::name)
        .def("description", &XLDrawingItem::description)
        .def("row", &XLDrawingItem::row)
        .def("col", &XLDrawingItem::col)
        .def("width", &XLDrawingItem::width)
        .def("height", &XLDrawingItem::height)
        .def("relationship_id", &XLDrawingItem::relationshipId)
        .def("image_binary", [](const XLDrawingItem& self) { auto data = self.imageBinary(); return py::bytes(reinterpret_cast<const char*>(data.data()), data.size()); })
        ;



    // Bind XLVmlDrawing
    py::class_<XLVmlDrawing>(m, "XLVmlDrawing")
        .def("shape_count", &XLVmlDrawing::shapeCount)
        .def("shape", &XLVmlDrawing::shape, "index"_a)
        .def("delete_shape", py::overload_cast<uint32_t>(&XLVmlDrawing::deleteShape), "index"_a)
        .def("delete_shape_by_ref", py::overload_cast<std::string_view>(&XLVmlDrawing::deleteShape), "cell_ref"_a)
        .def("create_shape", &XLVmlDrawing::createShape, "shape_template"_a = XLShape());
    // Bind XLDrawing
    py::class_<XLDrawing>(m, "XLDrawing")
        .def("image_count", &XLDrawing::imageCount)
        .def("image", &XLDrawing::image, "index"_a)
        .def("add_image", &XLDrawing::addImage, "r_id"_a, "name"_a,
             "description"_a, "row"_a, "col"_a, "width"_a,
             "height"_a, "options"_a = XLImageOptions())
        .def("add_scaled_image", &XLDrawing::addScaledImage, "r_id"_a, "name"_a,
             "description"_a, "data"_a, "row"_a, "col"_a,
             "scaling_factor"_a = 1.0)
        .def("add_shape", &XLDrawing::addShape, "row"_a, "col"_a,
             "options"_a = XLVectorShapeOptions());

    // Bind XLColumn
    py::class_<XLColumn>(m, "XLColumn")
        .def("width", &XLColumn::width)
        .def("set_width", &XLColumn::setWidth, "width"_a)
        .def("is_hidden", &XLColumn::isHidden)
        .def("set_hidden", &XLColumn::setHidden, "state"_a)
        .def("format", &XLColumn::format)
        .def("set_format", &XLColumn::setFormat, "cellFormatIndex"_a);

    // XLPaneState Enum
    py::enum_<XLPaneState>(m, "XLPaneState")
        .value("Split", XLPaneState::Split)
        .value("Frozen", XLPaneState::Frozen)
        .value("FrozenSplit", XLPaneState::FrozenSplit);

    // XLPane Enum
    py::enum_<XLPane>(m, "XLPane")
        .value("BottomRight", XLPane::BottomRight)
        .value("TopRight", XLPane::TopRight)
        .value("BottomLeft", XLPane::BottomLeft)
        .value("TopLeft", XLPane::TopLeft);

    // Bind XLSheetProtectionOptions
    py::class_<XLSheetProtectionOptions>(m, "XLSheetProtectionOptions")
        .def(py::init<>())
        .def_rw("sheet", &XLSheetProtectionOptions::sheet)
        .def_rw("objects", &XLSheetProtectionOptions::objects)
        .def_rw("scenarios", &XLSheetProtectionOptions::scenarios)
        .def_rw("format_cells", &XLSheetProtectionOptions::formatCells)
        .def_rw("format_columns", &XLSheetProtectionOptions::formatColumns)
        .def_rw("format_rows", &XLSheetProtectionOptions::formatRows)
        .def_rw("insert_columns", &XLSheetProtectionOptions::insertColumns)
        .def_rw("insert_rows", &XLSheetProtectionOptions::insertRows)
        .def_rw("insert_hyperlinks", &XLSheetProtectionOptions::insertHyperlinks)
        .def_rw("delete_columns", &XLSheetProtectionOptions::deleteColumns)
        .def_rw("delete_rows", &XLSheetProtectionOptions::deleteRows)
        .def_rw("sort", &XLSheetProtectionOptions::sort)
        .def_rw("auto_filter", &XLSheetProtectionOptions::autoFilter)
        .def_rw("pivot_tables", &XLSheetProtectionOptions::pivotTables)
        .def_rw("select_locked_cells", &XLSheetProtectionOptions::selectLockedCells)
        .def_rw("select_unlocked_cells", &XLSheetProtectionOptions::selectUnlockedCells);

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
        .def("vml_drawing", &XLWorksheet::vmlDrawing, py::rv_policy::reference_internal)
        .def("has_vml_drawing", &XLWorksheet::hasVmlDrawing)
        .def("has_panes", &XLWorksheet::hasPanes)
        .def("freeze_panes", py::overload_cast<uint16_t, uint32_t>(&XLWorksheet::freezePanes),
             "column"_a, "row"_a)
        .def("freeze_panes", py::overload_cast<const std::string&>(&XLWorksheet::freezePanes),
             "cellRef"_a)
        .def("split_panes", &XLWorksheet::splitPanes, "xSplit"_a, "ySplit"_a,
             "topLeftCell"_a = "", "activePane"_a = XLPane::BottomRight)
        .def("clear_panes", &XLWorksheet::clearPanes)
        .def("has_auto_filter", &XLWorksheet::hasAutoFilter)
        .def("auto_filter", &XLWorksheet::autoFilter)
        .def(
            "set_auto_filter",
            [](XLWorksheet& self, const std::string& range) {
                self.setAutoFilter(self.range(range));
            },
            "range"_a)
        .def("autofilter_object", &XLWorksheet::autofilterObject)
        .def("clear_auto_filter", &XLWorksheet::clearAutoFilter)
        .def("set_zoom", &XLWorksheet::setZoom, "scale"_a)
        .def("zoom", &XLWorksheet::zoom)
        .def("add_hyperlink", &XLWorksheet::addHyperlink, "cellRef"_a, "url"_a,
             "tooltip"_a = "")
        .def("add_internal_hyperlink", &XLWorksheet::addInternalHyperlink, "cellRef"_a,
             "location"_a, "tooltip"_a = "")
        .def("has_hyperlink", &XLWorksheet::hasHyperlink, "cellRef"_a)
        .def("get_hyperlink", &XLWorksheet::getHyperlink, "cellRef"_a)
        .def("remove_hyperlink", &XLWorksheet::removeHyperlink, "cellRef"_a)
        .def("data_validations", &XLWorksheet::dataValidations, py::rv_policy::reference_internal)
        .def("page_setup", &XLWorksheet::pageSetup)
        .def("page_margins", &XLWorksheet::pageMargins)
        .def("print_options", &XLWorksheet::printOptions)
        .def("tables", &XLWorksheet::tables, py::rv_policy::reference_internal)
        .def("has_tables", &XLWorksheet::hasTables)
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
            "rangeReference"_a, "emptyHiddenCells"_a = false)
        .def("insert_row", &XLWorksheet::insertRow, "row_number"_a, "count"_a = 1)
        .def("delete_row", py::overload_cast<uint32_t>(&XLWorksheet::deleteRow),
             "row_number"_a)
        .def("delete_row", py::overload_cast<uint32_t, uint32_t>(&XLWorksheet::deleteRow),
             "row_number"_a, "count"_a)
        .def("insert_column", &XLWorksheet::insertColumn, "col_number"_a,
             "count"_a = 1)
        .def("delete_column", &XLWorksheet::deleteColumn, "col_number"_a,
             "count"_a = 1)
        .def(
            "unmerge_cells",
            [](XLWorksheet& self, const std::string& rangeReference) {
                py::gil_scoped_release release;
                self.unmergeCells(rangeReference);
            },
            "rangeReference"_a)
        .def("column_format",
             py::overload_cast<const std::string&>(&XLWorksheet::getColumnFormat, py::const_))
        .def("merges", &XLWorksheet::merges, py::rv_policy::reference_internal)
        .def("set_column_format",
             py::overload_cast<const std::string&, XLStyleIndex>(&XLWorksheet::setColumnFormat),
             "column"_a, "cellFormatIndex"_a)
        .def("set_column_format",
             py::overload_cast<uint16_t, XLStyleIndex>(&XLWorksheet::setColumnFormat),
             "column"_a, "cellFormatIndex"_a)
        .def("row_format", &XLWorksheet::getRowFormat)
        .def("set_row_format", &XLWorksheet::setRowFormat, "row"_a,
             "cellFormatIndex"_a)
        .def(
            "protect",
            [](XLWorksheet& self, const XLSheetProtectionOptions& options,
               const std::string& password) {
                py::gil_scoped_release release;
                return self.protect(options, password);
            },
            "options"_a, "password"_a = "")
        .def(
            "protect_sheet",
            [](XLWorksheet& self, bool set) {
                py::gil_scoped_release release;
                self.protectSheet(set);
            },
            "set"_a = true)
        .def(
            "protect_objects",
            [](XLWorksheet& self, bool set) {
                py::gil_scoped_release release;
                self.protectObjects(set);
            },
            "set"_a = true)
        .def(
            "protect_scenarios",
            [](XLWorksheet& self, bool set) {
                py::gil_scoped_release release;
                self.protectScenarios(set);
            },
            "set"_a = true)
        .def("sheet_protected", &XLWorksheet::sheetProtected)
        .def("objects_protected", &XLWorksheet::objectsProtected)
        .def("scenarios_protected", &XLWorksheet::scenariosProtected)
        .def("threaded_comments", &XLWorksheet::threadedComments, py::rv_policy::reference_internal)
        .def(
            "set_password",
            [](XLWorksheet& self, const std::string& password) {
                py::gil_scoped_release release;
                self.setPassword(password);
            },
            "password"_a)
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
        .def("insert_hyperlinks_allowed", &XLWorksheet::insertHyperlinksAllowed)
        .def("auto_filter_allowed", &XLWorksheet::autoFilterAllowed)
        .def("sort_allowed", &XLWorksheet::sortAllowed)
        .def("pivot_tables_allowed", &XLWorksheet::pivotTablesAllowed)
        .def("format_cells_allowed", &XLWorksheet::formatCellsAllowed)
        .def("format_columns_allowed", &XLWorksheet::formatColumnsAllowed)
        .def("format_rows_allowed", &XLWorksheet::formatRowsAllowed)
        .def(
            "set_insert_columns_allowed",
            [](XLWorksheet& self, bool set) {
                py::gil_scoped_release release;
                self.allowInsertColumns(set);
            },
            "set"_a = true)
        .def(
            "set_insert_rows_allowed",
            [](XLWorksheet& self, bool set) {
                py::gil_scoped_release release;
                self.allowInsertRows(set);
            },
            "set"_a = true)
        .def(
            "set_insert_hyperlinks_allowed",
            [](XLWorksheet& self, bool set) {
                py::gil_scoped_release release;
                self.allowInsertHyperlinks(set);
            },
            "set"_a = true)
        .def(
            "set_delete_columns_allowed",
            [](XLWorksheet& self, bool set) {
                py::gil_scoped_release release;
                self.allowDeleteColumns(set);
            },
            "set"_a = true)
        .def(
            "set_delete_rows_allowed",
            [](XLWorksheet& self, bool set) {
                py::gil_scoped_release release;
                self.allowDeleteRows(set);
            },
            "set"_a = true)
        .def(
            "set_select_locked_cells_allowed",
            [](XLWorksheet& self, bool set) {
                py::gil_scoped_release release;
                self.allowSelectLockedCells(set);
            },
            "set"_a = true)
        .def(
            "set_select_unlocked_cells_allowed",
            [](XLWorksheet& self, bool set) {
                py::gil_scoped_release release;
                self.allowSelectUnlockedCells(set);
            },
            "set"_a = true)
        .def(
            "set_auto_filter_allowed",
            [](XLWorksheet& self, bool set) {
                py::gil_scoped_release release;
                self.allowAutoFilter(set);
            },
            "set"_a = true)
        .def(
            "set_sort_allowed",
            [](XLWorksheet& self, bool set) {
                py::gil_scoped_release release;
                self.allowSort(set);
            },
            "set"_a = true)
        .def(
            "set_pivot_tables_allowed",
            [](XLWorksheet& self, bool set) {
                py::gil_scoped_release release;
                self.allowPivotTables(set);
            },
            "set"_a = true)
        .def(
            "set_format_cells_allowed",
            [](XLWorksheet& self, bool set) {
                py::gil_scoped_release release;
                self.allowFormatCells(set);
            },
            "set"_a = true)
        .def(
            "set_format_columns_allowed",
            [](XLWorksheet& self, bool set) {
                py::gil_scoped_release release;
                self.allowFormatColumns(set);
            },
            "set"_a = true)
        .def(
            "set_format_rows_allowed",
            [](XLWorksheet& self, bool set) {
                py::gil_scoped_release release;
                self.allowFormatRows(set);
            },
            "set"_a = true)
        .def("comments", &XLWorksheet::comments, py::rv_policy::reference_internal)
        .def("add_image", &add_image_to_worksheet, "image_data"_a, "extension"_a,
             "row"_a = 1, "col"_a = 1, "width"_a = 0, "height"_a = 0)
        // Bulk read APIs for performance optimization
        .def("get_rows_data", &get_rows_data,
             "Get all rows data as list[list[Any]] - optimized for bulk read")
        .def("get_row_values", &get_row_values, "row"_a,
             "Get a single row's values as list[Any]")
        .def("get_range_data", &get_range_data, "start_row"_a, "start_col"_a,
             "end_row"_a, "end_col"_a,
             "Get a range of cells as list[list[Any]] - optimized bulk read for specific range")
        .def("get_cell_value", &get_cell_value, "row"_a, "col"_a,
             "Get a single cell's value directly without creating a Cell object")
        .def("write_range_data", &write_range_typed<double>, "start_row"_a,
             "start_col"_a, "data"_a,
             "Write a 2D numpy array or buffer to a worksheet range")
        .def("write_range_data", &write_range_typed<int64_t>, "start_row"_a,
             "start_col"_a, "data"_a)
        .def("write_range_data", &write_range_typed<bool>, "start_row"_a,
             "start_col"_a, "data"_a)
        .def("get_range_values", &get_range_values, "start_row"_a, "start_col"_a,
             "end_row"_a, "end_col"_a,
             "Read a range of numeric cells into a 2D numpy array of doubles")
        // Performance-optimized write APIs - bypass Python Cell object creation
        .def("set_cell_value", &set_cell_value, "row"_a, "col"_a, "value"_a,
             "Set a cell's value directly without creating a Cell object. "
             "10-20x faster than ws.cell(row, col).value = val for bulk operations")
        .def("write_rows_data", &write_rows_data, "start_row"_a, "start_col"_a,
             "rows"_a,
             "Write a 2D Python list to a worksheet range. "
             "Optimized for any Python data (strings, mixed types). "
             "For pure numeric data, use write_range_data with numpy for best performance")
        .def("write_row_data", &write_row_data, "row"_a, "start_col"_a,
             "values"_a, "Write a single row of Python data")
        .def("set_cells_batch", &set_cells_batch, "cells"_a,
             "Batch set multiple cell values: [(row, col, value), ...]. "
             "Efficient for non-contiguous cell updates")
        .def(
            "stream_writer",
            [](XLWorksheet& self, bool use_shared_strings, size_t max_unique_strings) {
                return self.streamWriter(use_shared_strings, max_unique_strings);
            },
            "use_shared_strings"_a = false, "max_unique_strings"_a = size_t{100000},
            "Start a stream writer. Optionally enable shared-string caching.")
        .def(
            "stream_reader",
            [](const XLWorksheet& self, py::object options) {
                if (options.is_none()) {
                    return self.streamReader();
                }
                return self.streamReader(py::cast<XLStreamReadOptions&>(options));
            },
            "options"_a = py::none(),
            "Create a stream reader. Pass XLStreamReadOptions for empty-row / number-format policy.")
        .def("peek_cell", py::overload_cast<const std::string&>(&XLWorksheet::peekCell, py::const_),
             "ref"_a)
        .def("peek_cell", py::overload_cast<uint32_t, uint16_t>(&XLWorksheet::peekCell, py::const_),
             "row"_a, "col"_a)
        .def("auto_fit_column", &XLWorksheet::autoFitColumn, "column_number"_a)
        .def("add_sort_condition", &XLWorksheet::addSortCondition, "ref"_a,
             "col_id"_a, "descending"_a = false)
        .def("apply_auto_filter", &XLWorksheet::applyAutoFilter)
        .def("add_conditional_formatting",
             py::overload_cast<const std::string&, const XLCfRule&>(
                 &XLWorksheet::addConditionalFormatting),
             "sqref"_a, "rule"_a)
        .def("add_conditional_formatting_dxf",
             py::overload_cast<const std::string&, const XLCfRule&, const XLDxf&>(
                 &XLWorksheet::addConditionalFormatting),
             "sqref"_a, "rule"_a, "dxf"_a)
        .def("remove_conditional_formatting",
             py::overload_cast<const std::string&>(&XLWorksheet::removeConditionalFormatting),
             "sqref"_a)
        .def("clear_all_conditional_formatting", &XLWorksheet::clearAllConditionalFormatting)
        .def("header_footer", &XLWorksheet::headerFooter)
        .def("set_print_area", &XLWorksheet::setPrintArea, "sqref"_a)
        .def("set_print_title_rows", &XLWorksheet::setPrintTitleRows, "first_row"_a,
             "last_row"_a)
        .def("set_print_title_cols", &XLWorksheet::setPrintTitleCols, "first_col"_a,
             "last_col"_a)
        .def("add_sparkline",
             py::overload_cast<const std::string&, const std::string&, XLSparklineType>(
                 &XLWorksheet::addSparkline),
             "location"_a, "data_range"_a, "type"_a = XLSparklineType::Line)
        .def("add_sparkline",
             py::overload_cast<const std::string&, const std::string&, const XLSparklineOptions&>(
                 &XLWorksheet::addSparkline),
             "location"_a, "data_range"_a, "options"_a)
        .def("insert_image",
             py::overload_cast<const std::string&, const std::string&>(&XLWorksheet::insertImage),
             "cell_reference"_a, "image_path"_a)
        .def("insert_image_opt",
             py::overload_cast<const std::string&, const std::string&, const XLImageOptions&>(
                 &XLWorksheet::insertImage),
             "cell_reference"_a, "image_path"_a, "options"_a)
        .def("add_chart",
             py::overload_cast<XLChartType, std::string_view, uint32_t, uint32_t, uint32_t,
                               uint32_t>(&XLWorksheet::addChart),
             "type"_a, "name"_a, "row"_a, "col"_a, "width"_a,
             "height"_a)
        .def("add_chart_anchor",
             py::overload_cast<XLChartType, const XLChartAnchor&>(&XLWorksheet::addChart),
             "type"_a, "anchor"_a)
        .def("add_pivot_table", &XLWorksheet::addPivotTable, "options"_a)
        .def("add_table_slicer", &XLWorksheet::addTableSlicer, "cell_reference"_a,
             "table"_a, "column_name"_a, "options"_a = XLSlicerOptions())
        .def("add_pivot_slicer", &XLWorksheet::addPivotSlicer, "cell_reference"_a,
             "pivot_table"_a, "column_name"_a, "options"_a = XLSlicerOptions())
        .def(
            "slicers", [](XLWorksheet& self) -> XLSlicerCollection& { return self.slicers(); },
            py::rv_policy::reference_internal, "Slicer collection for this worksheet.")
        .def("delete_slicer", &XLWorksheet::deleteSlicer, "name"_a,
             "Delete a slicer by name and clean up orphan caches.")
        .def(
            "insert_image_bytes",
            [](XLWorksheet& self, const std::string& cell_reference, py::bytes data,
               py::object options) {
                const auto* ptr = static_cast<const uint8_t*>(data.data());
                gsl::span<const uint8_t> span(ptr, data.size());
                if (options.is_none()) {
                    self.insertImage(cell_reference, span);
                } else {
                    self.insertImage(cell_reference, span, py::cast<XLImageOptions&>(options));
                }
            },
            "cell_reference"_a, "image_data"_a, "options"_a = py::none(),
            "Insert an image from raw bytes at the given cell.")
        .def(
            "add_comment",
            [](XLWorksheet& self, std::string_view cellRef, std::string_view text,
               std::string_view author) -> XLThreadedComment {
                return self.addComment(cellRef, text, author);
            },
            "cell_ref"_a, "text"_a, "author"_a = "",
            "Add a modern threaded comment; returns XLThreadedComment.")
        .def("add_note", &XLWorksheet::addNote, "cell_ref"_a, "text"_a, "author"_a = "")
        .def("delete_comment", &XLWorksheet::deleteComment, "cell_ref"_a)
        .def("delete_note", &XLWorksheet::deleteNote, "cell_ref"_a)
        .def(
            "add_reply",
            [](XLWorksheet& self, const std::string& parentId, const std::string& text,
               const std::string& author) -> XLThreadedComment {
                return self.addReply(parentId, text, author);
            },
            "parent_id"_a, "text"_a, "author"_a = "")
        // Compatibility aliases used by higher-level Python wrappers
        .def(
            "add_threaded_comment",
            [](XLWorksheet& self, std::string_view cellRef, std::string_view text,
               std::string_view author) -> XLThreadedComment {
                return self.addComment(cellRef, text, author);
            },
            "cell_ref"_a, "text"_a, "author"_a = "")
        .def(
            "add_threaded_reply",
            [](XLWorksheet& self, const std::string& parentId, const std::string& text,
               const std::string& author) -> XLThreadedComment {
                return self.addReply(parentId, text, author);
            },
            "parent_id"_a, "text"_a, "author"_a = "")
        .def("find_cell", py::overload_cast<const std::string&>(&XLWorksheet::findCell, py::const_),
             "ref"_a)
        .def("find_cell",
             py::overload_cast<uint32_t, uint16_t>(&XLWorksheet::findCell, py::const_), "row"_a,
             "col"_a)
        .def("last_cell", &XLWorksheet::lastCell)
        .def("row", &XLWorksheet::row, "row_number"_a, py::keep_alive<0, 1>())
        .def("rows", py::overload_cast<>(&XLWorksheet::rows, py::const_), py::keep_alive<0, 1>())
        .def("rows", py::overload_cast<uint32_t>(&XLWorksheet::rows, py::const_), "row_count"_a,
             py::keep_alive<0, 1>())
        .def("rows", py::overload_cast<uint32_t, uint32_t>(&XLWorksheet::rows, py::const_),
             "first_row"_a, "last_row"_a, py::keep_alive<0, 1>())
        .def(
            "append_row",
            [](XLWorksheet& self, py::sequence values) {
                std::vector<XLCellValue> vals;
                vals.reserve(py::len(values));
                for (auto v : values) vals.push_back(CellData::from_python(v).to_xlcellvalue());
                py::gil_scoped_release release;
                self.appendRow(vals);
            },
            "values"_a)
        .def("group_rows", &XLWorksheet::groupRows, "row_first"_a, "row_last"_a,
             "outline_level"_a = 1, "collapsed"_a = false)
        .def("group_columns", &XLWorksheet::groupColumns, "col_first"_a, "col_last"_a,
             "outline_level"_a = 1, "collapsed"_a = false)
        .def("update_sheet_name", &XLWorksheet::updateSheetName, "old_name"_a, "new_name"_a)
        .def("update_dimension", &XLWorksheet::updateDimension)
        .def("is_streamed_sheet", &XLWorksheet::isStreamedSheet)
        .def("conditional_formats", &XLWorksheet::conditionalFormats)
        .def("clear_sheet_protection", &XLWorksheet::clearSheetProtection)
        .def("sheet_protection_summary", &XLWorksheet::sheetProtectionSummary)
        .def("images", &XLWorksheet::images)
        .def("pivot_tables", &XLWorksheet::pivotTables)
        .def("delete_pivot_table", &XLWorksheet::deletePivotTable, "name"_a)
        .def("has_relationships", &XLWorksheet::hasRelationships)
        .def("has_comments", &XLWorksheet::hasComments)
        .def("has_threaded_comments", &XLWorksheet::hasThreadedComments)
        .def("insert_row_break", &XLWorksheet::insertRowBreak, "row"_a)
        .def("insert_col_break", &XLWorksheet::insertColBreak, "col"_a)
        .def("remove_row_break", &XLWorksheet::removeRowBreak, "row"_a)
        .def("remove_col_break", &XLWorksheet::removeColBreak, "col"_a)
        .def("set_sheet_view_mode", &XLWorksheet::setSheetViewMode, "mode"_a)
        .def("sheet_view_mode", &XLWorksheet::sheetViewMode)
        .def("set_show_grid_lines", &XLWorksheet::setShowGridLines, "show"_a)
        .def("show_grid_lines", &XLWorksheet::showGridLines)
        .def("set_show_row_col_headers", &XLWorksheet::setShowRowColHeaders, "show"_a)
        .def("show_row_col_headers", &XLWorksheet::showRowColHeaders)
        .def("fit_to_pages", &XLWorksheet::fitToPages, "fit_to_width"_a, "fit_to_height"_a)
        .def("add_shape", &XLWorksheet::addShape, "cell_reference"_a, "options"_a)
        .def("add_scaled_image", &XLWorksheet::addScaledImage, "name"_a, "data"_a, "row"_a, "col"_a,
             "scaling_factor"_a = 1.0)
        .def("range_used", py::overload_cast<>(&XLWorksheet::range, py::const_),
             py::keep_alive<0, 1>(), "Used range of the worksheet.")
        ;
}
