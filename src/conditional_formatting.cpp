#include "bindings.hpp"

void init_conditional_formatting(py::module_& m) {
    py::enum_<XLCfOperator>(m, "XLCfOperator")
        .value("LessThan", XLCfOperator::LessThan)
        .value("LessThanOrEqual", XLCfOperator::LessThanOrEqual)
        .value("Equal", XLCfOperator::Equal)
        .value("NotEqual", XLCfOperator::NotEqual)
        .value("GreaterThanOrEqual", XLCfOperator::GreaterThanOrEqual)
        .value("GreaterThan", XLCfOperator::GreaterThan)
        .value("Between", XLCfOperator::Between)
        .value("NotBetween", XLCfOperator::NotBetween)
        .value("ContainsText", XLCfOperator::ContainsText)
        .value("NotContains", XLCfOperator::NotContains)
        .value("BeginsWith", XLCfOperator::BeginsWith)
        .value("EndsWith", XLCfOperator::EndsWith)
        .value("Invalid", XLCfOperator::Invalid);

    py::class_<XLCfRule>(m, "XLCfRule").def(py::init<>()).def("summary", &XLCfRule::summary);

    m.def("XLColorScaleRule", py::overload_cast<const XLColor&, const XLColor&>(&XLColorScaleRule));
    m.def("XLColorScaleRule",
          py::overload_cast<const XLColor&, const XLColor&, const XLColor&>(&XLColorScaleRule));
    m.def("XLDataBarRule", &XLDataBarRule, py::arg("color"), py::arg("show_value") = true);
    m.def("XLCellIsRule", py::overload_cast<XLCfOperator, const std::string&>(&XLCellIsRule));
    m.def("XLCellIsRule", py::overload_cast<const std::string&, const std::string&>(&XLCellIsRule));
    m.def("XLFormulaRule", &XLFormulaRule);
    m.def("XLIconSetRule", &XLIconSetRule, py::arg("icon_set_name") = "3TrafficLights1",
          py::arg("show_value") = true, py::arg("reverse") = false);
    m.def("XLTop10Rule", &XLTop10Rule, py::arg("rank") = 10, py::arg("percent") = false,
          py::arg("bottom") = false);
    m.def("XLAboveAverageRule", &XLAboveAverageRule, py::arg("above_average") = true,
          py::arg("equal_average") = false, py::arg("std_dev") = 0);
    m.def("XLDuplicateValuesRule", &XLDuplicateValuesRule, py::arg("unique") = false);
    m.def("XLContainsTextRule", &XLContainsTextRule);
    m.def("XLNotContainsTextRule", &XLNotContainsTextRule);
    m.def("XLContainsBlanksRule", &XLContainsBlanksRule);
    m.def("XLNotContainsBlanksRule", &XLNotContainsBlanksRule);
    m.def("XLContainsErrorsRule", &XLContainsErrorsRule);
    m.def("XLNotContainsErrorsRule", &XLNotContainsErrorsRule);
}
