#include <headers/XLCalculationEngine.hpp>
#include <headers/XLCellReference.hpp>
#include <headers/XLFormulaEngine.hpp>
#include <nanobind/stl/optional.h>

#include "internal_access.hpp"

using namespace OpenXLSX;
using namespace nanobind::literals;

void init_formula_engine(py::module_& m) {
    py::class_<XLFormulaDiagnostic>(m, "XLFormulaDiagnostic")
        .def_ro("message", &XLFormulaDiagnostic::message)
        .def_ro("offset", &XLFormulaDiagnostic::offset);

    py::class_<XLFormulaDiagnosticReporter>(m, "XLFormulaDiagnosticReporter")
        .def(py::init<>())
        .def(py::init<std::string>(), "formula"_a)
        .def("report_error", &XLFormulaDiagnosticReporter::reportError, "message"_a, "offset"_a)
        .def("has_errors", &XLFormulaDiagnosticReporter::hasErrors)
        .def("clear", &XLFormulaDiagnosticReporter::clear)
        .def("get_full_report", &XLFormulaDiagnosticReporter::getFullReport)
        .def("diagnostics",
             [](const XLFormulaDiagnosticReporter& self) { return self.diagnostics(); });

    py::class_<XLCellKey>(m, "XLCellKey")
        .def(py::init<>())
        .def_rw("sheet", &XLCellKey::sheet)
        .def_rw("row", &XLCellKey::row)
        .def_rw("col", &XLCellKey::col)
        .def("valid", &XLCellKey::valid)
        .def("address", &XLCellKey::address)
        .def("qualified_address", &XLCellKey::qualifiedAddress)
        .def_static("parse", &XLCellKey::parse, "ref"_a, "default_sheet"_a = "",
                    "force_default_sheet"_a = false);

    py::class_<XLEvalSession>(m, "XLEvalSession")
        .def(py::init<>())
        .def(
            "set_current_cell",
            [](XLEvalSession& self, uint32_t row, uint16_t col) -> XLEvalSession& {
                return self.setCurrentCell(row, col);
            },
            "row"_a, "col"_a, py::rv_policy::reference_internal)
        .def(
            "set_current_sheet",
            [](XLEvalSession& self, std::string sheet) -> XLEvalSession& {
                return self.setCurrentSheet(std::move(sheet));
            },
            "sheet_name"_a, py::rv_policy::reference_internal)
        .def(
            "set_name_resolver",
            [](XLEvalSession& self, py::object callback) -> XLEvalSession& {
                if (callback.is_none()) {
                    return self.setNameResolver({});
                }
                // Keep the Python callable alive by capturing a heap-allocated handle owned by the
                // lambda; session lifetime is managed by Python, so this is acceptable.
                auto* cb = new py::object(std::move(callback));
                return self.setNameResolver([cb](std::string_view name) -> std::optional<std::string> {
                    py::gil_scoped_acquire acquire;
                    py::object result = (*cb)(std::string(name));
                    if (result.is_none()) return std::nullopt;
                    return py::cast<std::string>(result);
                });
            },
            "callback"_a, py::rv_policy::reference_internal)
        .def_prop_ro("has_current_cell", &XLEvalSession::hasCurrentCell)
        .def_prop_ro("current_row", &XLEvalSession::currentRow)
        .def_prop_ro("current_col", &XLEvalSession::currentCol)
        .def_prop_ro("current_sheet", &XLEvalSession::currentSheet);

    py::class_<XLFormulaEngine>(m, "XLFormulaEngine")
        .def(py::init<>())
        .def(
            "evaluate",
            [](const XLFormulaEngine& self, std::string_view formula, const XLWorksheet* wks,
               XLEvalSession* session, XLFormulaDiagnosticReporter* reporter) -> py::object {
                XLCellValue result;
                if (session != nullptr) {
                    // XLEvalSession stores a pointer to the resolver; keep it alive for this call.
                    std::optional<XLCellResolver> owned_resolver;
                    if (wks != nullptr) {
                        owned_resolver = XLFormulaEngine::makeResolver(*wks);
                        session->setResolver(*owned_resolver);
                    }
                    result = self.evaluate(formula, *session, reporter);
                } else if (wks != nullptr) {
                    auto resolver = XLFormulaEngine::makeResolver(*wks);
                    result = self.evaluate(formula, resolver, reporter);
                } else {
                    result = self.evaluate(formula, XLCellResolver{}, reporter);
                }
                return CellData::from(result).to_python();
            },
            "formula"_a, "wks"_a = py::none(), "session"_a = py::none(), "reporter"_a = py::none(),
            "Evaluate a formula. Optionally provide a worksheet, XLEvalSession, and/or diagnostic "
            "reporter.");

    py::enum_<XLCalcStatus>(m, "XLCalcStatus")
        .value("Ok", XLCalcStatus::Ok)
        .value("Circular", XLCalcStatus::Circular)
        .value("Error", XLCalcStatus::Error)
        .value("Empty", XLCalcStatus::Empty);

    py::class_<XLCalculationOptions>(m, "XLCalculationOptions")
        .def(py::init<>())
        .def_rw("write_back", &XLCalculationOptions::writeBack)
        .def_rw("max_depth", &XLCalculationOptions::maxDepth)
        .def_rw("max_expanded_deps", &XLCalculationOptions::maxExpandedDeps)
        .def_rw("propagate_dirty", &XLCalculationOptions::propagateDirty)
        .def_rw("set_full_calc_on_load", &XLCalculationOptions::setFullCalcOnLoad)
        .def_rw("circular_error_token", &XLCalculationOptions::circularErrorToken)
        .def_rw("auto_track_changes", &XLCalculationOptions::autoTrackChanges)
        .def_rw("use_defined_names", &XLCalculationOptions::useDefinedNames);

    py::class_<XLCalculationEngine>(m, "XLCalculationEngine")
        .def(
            "__init__",
            [](XLCalculationEngine* self, XLWorksheet& worksheet, py::object options) {
                if (options.is_none()) {
                    new (self) XLCalculationEngine(worksheet);
                } else {
                    new (self)
                        XLCalculationEngine(worksheet, py::cast<XLCalculationOptions&>(options));
                }
            },
            "worksheet"_a, "options"_a = py::none(), "Sheet-scoped calculation engine.")
        .def(
            "__init__",
            [](XLCalculationEngine* self, XLDocument& document, py::object options) {
                if (options.is_none()) {
                    new (self) XLCalculationEngine(document);
                } else {
                    new (self)
                        XLCalculationEngine(document, py::cast<XLCalculationOptions&>(options));
                }
            },
            "document"_a, "options"_a = py::none(), "Workbook-scoped calculation engine.")
        .def("rebuild", &XLCalculationEngine::rebuild)
        .def_prop_ro("formula_count", &XLCalculationEngine::formulaCount)
        .def_prop_ro("dirty_count", &XLCalculationEngine::dirtyCount)
        .def_prop_ro("is_multi_sheet", &XLCalculationEngine::isMultiSheet)
        .def_prop_ro("last_status", &XLCalculationEngine::lastStatus)
        .def(
            "calc_cell_value",
            [](XLCalculationEngine& self, std::string_view a1) {
                return CellData::from(self.calcCellValue(a1)).to_python();
            },
            "a1"_a)
        .def(
            "dependencies",
            [](const XLCalculationEngine& self, std::string_view a1) {
                return self.dependencies(a1);
            },
            "a1"_a)
        .def(
            "dependents",
            [](const XLCalculationEngine& self, std::string_view a1) {
                return self.dependents(a1);
            },
            "a1"_a)
        .def("recalculate", &XLCalculationEngine::recalculate)
        .def("recalculate_all", &XLCalculationEngine::recalculateAll)
        .def(
            "mark_dirty",
            [](XLCalculationEngine& self, std::string_view a1, bool propagate) {
                self.markDirty(a1, propagate);
            },
            "a1"_a, "propagate"_a = true)
        .def("mark_all_dirty", &XLCalculationEngine::markAllDirty)
        .def("clear_cache", &XLCalculationEngine::clearCache)
        .def("reload_defined_names", &XLCalculationEngine::reloadDefinedNames)
        .def(
            "update_formula_cell",
            [](XLCalculationEngine& self, std::string_view a1) {
                return self.updateFormulaCell(a1);
            },
            "a1"_a)
        .def(
            "set_input_value",
            [](XLCalculationEngine& self, std::string_view a1, py::object value) {
                self.setInputValue(a1, CellData::from_python(value).to_xlcellvalue());
            },
            "a1"_a, "value"_a)
        .def(
            "notify_changed",
            [](XLCalculationEngine& self, std::string_view a1) { self.notifyChanged(a1); },
            "a1"_a)
        .def_static(
            "extract_dependencies",
            [](std::string_view formula, std::string_view default_sheet, size_t max_expanded,
               bool force_default_sheet) {
                return XLCalculationEngine::extractDependencies(formula, default_sheet, max_expanded,
                                                                force_default_sheet, nullptr);
            },
            "formula"_a, "default_sheet"_a = "", "max_expanded"_a = size_t{4096},
            "force_default_sheet"_a = false);
}
