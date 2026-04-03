#include <headers/XLCellReference.hpp>
#include <headers/XLFormulaEngine.hpp>

#include "internal_access.hpp"

void init_formula_engine(py::module_& m) {
    py::class_<XLFormulaEngine>(m, "XLFormulaEngine")
        .def(py::init<>())
        .def(
            "evaluate",
            [](const XLFormulaEngine& self, std::string_view formula,
               const XLWorksheet* wks) -> py::object {
                XLCellValue result;
                if (wks) {
                    auto resolver = XLFormulaEngine::makeResolver(*wks);
                    result = self.evaluate(formula, resolver);
                } else {
                    result = self.evaluate(formula);
                }
                CellData cd = CellData::from(result);
                return cd.to_python();
            },
            py::arg("formula"), py::arg("wks") = py::none(),
            "Evaluate a formula string. Optionally provide an XLWorksheet to resolve cell "
            "references.");
}
