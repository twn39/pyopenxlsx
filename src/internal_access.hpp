#ifndef PYOPENXLSX_INTERNAL_ACCESS_HPP
#define PYOPENXLSX_INTERNAL_ACCESS_HPP

/**
 * @file internal_access.hpp
 * @brief Shared internal utilities for pyopenxlsx binding layer.
 *
 * Contains:
 * - Unified CellData structure for read/write operations
 * - Excel limits and precondition helpers
 * - Zero-overhead datetime → Excel serial conversion via CPython datetime C API
 *
 * Note: The Rob template hack has been removed. All access to OpenXLSX
 * internals is now through public APIs added to the fork:
 * - XLDocument::archive(), appProperties(), coreProperties(), contentTypes()
 * - XLXmlFile::parentDoc(), xmlDocument(), getXmlPath()
 */

#include <IZipArchive.hpp>
#include <cstdint>
#include <gsl/gsl>
#include <headers/XLContentTypes.hpp>
#include <headers/XLDrawing.hpp>

#include "bindings.hpp"

using namespace OpenXLSX;
using namespace nanobind::literals;

// ============================================================
// CPython datetime C API — per-TU lazy initialisation
// ============================================================
// IMPORTANT: PyDateTime_IMPORT expands to:
//   PyDateTimeAPI = (PyDateTime_CAPI*)PyCapsule_Import(PyDateTime_CAPSULE_NAME, 0)
// PyDateTimeAPI is declared as `static PyDateTime_CAPI*` inside datetime.h, meaning
// EACH translation unit that includes datetime.h has its OWN copy of the pointer.
// Calling PyDateTime_IMPORT in bindings.cpp only initialises THAT TU's copy;
// the copies in cell.cpp, worksheet.cpp, streams.cpp etc. remain nullptr, which
// causes a Segfault when PyDate_Check() dereferences the null pointer.
//
// Solution: ensure_datetime_api() uses std::call_once with a function-local static
// so that each TU initialises its own PyDateTimeAPI pointer exactly once, on the
// first call to from_python() that reaches the datetime branch.
#include <mutex>

inline void ensure_datetime_api() noexcept {
    static std::once_flag s_flag;
    std::call_once(s_flag, []() {
        PyDateTime_IMPORT;   // NOLINT(cppcoreguidelines-pro-type-cstyle-cast)
    });
}

// ============================================================
// Pure C++ Excel serial-date helpers  (no GIL, no Python calls)
// ============================================================

/**
 * @brief Convert a proleptic Gregorian calendar date to a "civil day number".
 *
 * Uses Howard Hinnant's date-library algorithm (public domain).
 * The reference epoch is the Unix epoch (1970-01-01 = 0).
 *
 * @param y  Full year (e.g. 2024).
 * @param m  Month in [1, 12].
 * @param d  Day   in [1, 31].
 * @return   Signed day count relative to 1970-01-01.
 */
[[nodiscard]] inline int32_t days_from_civil(int32_t y, int32_t m, int32_t d) noexcept {
    y -= (m <= 2) ? 1 : 0;
    const int32_t era = (y >= 0 ? y : y - 399) / 400;
    const auto    yoe = static_cast<uint32_t>(y - era * 400);              // [0, 399]
    const auto    doy =
        (153u * static_cast<uint32_t>(m > 2 ? m - 3 : m + 9) + 2u) / 5u
        + static_cast<uint32_t>(d) - 1u;                                   // [0, 365]
    const uint32_t doe = yoe * 365u + yoe / 4u - yoe / 100u + doy;        // [0, 146096]
    return era * 146097 + static_cast<int32_t>(doe) - 719468;
}

/**
 * @brief Convert a Python datetime.date / datetime.datetime object to an
 *        Excel serial date (floating-point days since 1899-12-30).
 *
 * This function is GIL-safe to call (it uses only C-macro field accessors
 * and no Python bytecode execution).
 *
 * Excel serial arithmetic:
 *   - 1900-01-01 = 1.0 (Lotus 1-2-3 compat: 1900 is falsely treated as a leap year,
 *     so serial 60 = 1900-02-29 is a phantom day; all dates from 1900-03-01
 *     onward are off-by-one from a pure Gregorian count, which is exactly what
 *     OpenXLSX and Excel expect.  We replicate Python's datetime_to_serial:
 *     delta = val - datetime(1899, 12, 30), so the phantom day is naturally
 *     included because we anchor on 1899-12-30, not 1900-01-01.)
 *   - Epoch in civil days: days_from_civil(1899, 12, 30) == -25569
 *
 * @param obj  A PyObject* that has already been validated by PyDate_Check().
 *             If PyDateTime_Check() is also true the time sub-day fraction is
 *             included; otherwise the fractional part is 0.0.
 * @return     Excel serial value as double.
 */
[[nodiscard]] inline double date_to_excel_serial(PyObject* obj) noexcept {
    // Ensure this TU's PyDateTimeAPI pointer has been initialised before
    // using any PyDate_Check / field-accessor macros.
    ensure_datetime_api();

    // Day-of-epoch constants (compile-time)
    constexpr int32_t kExcelEpochCivil = -25569; // days_from_civil(1899, 12, 30)

    const int32_t y = PyDateTime_GET_YEAR(obj);
    const int32_t m = PyDateTime_GET_MONTH(obj);
    const int32_t d = PyDateTime_GET_DAY(obj);

    const int32_t serial_day = days_from_civil(y, m, d) - kExcelEpochCivil;

    double time_frac = 0.0;
    if (PyDateTime_Check(obj)) {
        const int32_t hour = PyDateTime_DATE_GET_HOUR(obj);
        const int32_t min  = PyDateTime_DATE_GET_MINUTE(obj);
        const int32_t sec  = PyDateTime_DATE_GET_SECOND(obj);
        const int32_t usec = PyDateTime_DATE_GET_MICROSECOND(obj);
        time_frac = (hour * 3600.0 + min * 60.0 + sec + usec * 1e-6) / 86400.0;
    }

    return static_cast<double>(serial_day) + time_frac;
}


// ============================================================
// OpenXLSX Internal Access (Using exposed APIs)
// ============================================================

// Helper functions for easy access
inline IZipArchive& get_archive(XLDocument& doc) { return doc.archive(); }
inline XLAppProperties& get_app_properties(XLDocument& doc) { return doc.appProperties(); }
inline XLProperties& get_core_properties(XLDocument& doc) { return doc.coreProperties(); }

inline XMLDocument& get_xml_doc(XLXmlFile& file) { return file.xmlDocument(); }
inline const XMLDocument& get_xml_doc(const XLXmlFile& file) { return file.xmlDocument(); }
inline XLDocument& get_parent_doc(XLXmlFile& file) { return file.parentDoc(); }
inline std::string get_xml_path(const XLXmlFile& file) { return file.getXmlPath(); }

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
    enum class Type { Empty, Boolean, Integer, Float, String, RichText };
    Type type = Type::Empty;
    bool boolVal = false;
    int64_t intVal = 0;
    double floatVal = 0.0;
    std::string strVal;
    XLRichText richTextVal;

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
            case XLValueType::RichText:
                data.type = Type::RichText;
                data.richTextVal = val.get<XLRichText>();
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
        } else if (py::isinstance<XLRichText>(obj)) {
            val.type = Type::RichText;
            val.richTextVal = py::cast<XLRichText>(obj);
        } else {
            // Handle numpy scalars like np.int64, np.float64, np.bool_ efficiently
            // They expose an .item() method that converts them to Python native types
            if (py::hasattr(obj, "item")) {
                py::object native_obj = obj.attr("item")();
                return from_python(native_obj);
            }

            // Fast path: datetime.date / datetime.datetime
            // Uses CPython datetime C API macros — zero Python function calls,
            // zero module imports.  PyDate_Check is a strict subtype check
            // (accepts both date and datetime); PyDateTime_Check tests datetime
            // specifically so date_to_excel_serial() can include the time fraction.
            // ensure_datetime_api() initialises this TU's PyDateTimeAPI on first use.
            ensure_datetime_api();
            if (PyDate_Check(obj.ptr())) {
                val.type     = Type::Float;
                val.floatVal = date_to_excel_serial(obj.ptr());
            } else {
                throw py::type_error("Unsupported type for cell value");
            }
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
            case Type::RichText:
                return py::cast(richTextVal);
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
            case Type::RichText:
                return XLCellValue(richTextVal);
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
            case Type::RichText:
                cell.value() = richTextVal;
                break;
        }
    }
};

#endif  // PYOPENXLSX_INTERNAL_ACCESS_HPP
