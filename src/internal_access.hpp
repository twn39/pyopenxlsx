#ifndef PYOPENXLSX_INTERNAL_ACCESS_HPP
#define PYOPENXLSX_INTERNAL_ACCESS_HPP

/**
 * @file internal_access.hpp
 * @brief Shared internal utilities for pyopenxlsx binding layer.
 *
 * Contains:
 * - Unified CellData structure for read/write operations
 * - Excel limits and precondition helpers
 *
 * Note: The Rob template hack has been removed. All access to OpenXLSX
 * internals is now through public APIs added to the fork:
 * - XLDocument::archive(), appProperties(), coreProperties(), contentTypes()
 * - XLXmlFile::parentDoc(), xmlDocument(), getXmlPath()
 */

#include <gsl/gsl>
#include <headers/XLContentTypes.hpp>
#include <headers/XLDrawing.hpp>

#include "bindings.hpp"

// ============================================================
// Excel Limits (for precondition checks)
// ============================================================
constexpr uint32_t kExcelMaxRows = 1048576;
constexpr uint16_t kExcelMaxCols = 16384;

// ============================================================
// Unified CellData structure for read/write operations
// Merges the former CellValueData (read) and BatchCellValue (write)
// ============================================================

struct CellData {
    enum class Type { Empty, Boolean, Integer, Float, String };
    Type type = Type::Empty;
    bool boolVal = false;
    int64_t intVal = 0;
    double floatVal = 0.0;
    std::string strVal;

    // -- Read from C++ XLCellValue (no GIL needed) --
    static CellData from(const XLCellValue& val) {
        CellData data;
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

    // -- Read from Python object (GIL must be held) --
    static CellData from_python(py::handle obj) {
        CellData val;
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
            // Fallback: convert to string
            val.type = Type::String;
            val.strVal = py::cast<std::string>(py::str(obj));
        }
        return val;
    }

    // -- Convert to Python object (GIL must be held) --
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

    // -- Convert to XLCellValue for writing (no GIL needed) --
    XLCellValue to_xlcellvalue() const {
        switch (type) {
            case Type::Boolean:
                return XLCellValue(boolVal);
            case Type::Integer:
                return XLCellValue(intVal);
            case Type::Float:
                return XLCellValue(floatVal);
            case Type::String:
                return XLCellValue(strVal);
            default:
                return XLCellValue();
        }
    }

    // -- Apply to an XLCell directly (no GIL needed) --
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

#endif  // PYOPENXLSX_INTERNAL_ACCESS_HPP
