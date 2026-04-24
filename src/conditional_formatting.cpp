#include "bindings.hpp"

using namespace OpenXLSX;
using namespace nanobind::literals;


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
    m.def("XLDataBarRule", &XLDataBarRule, "color"_a, "show_value"_a = true);
    m.def("XLCellIsRule", py::overload_cast<XLCfOperator, const std::string&>(&XLCellIsRule));
    m.def("XLCellIsRule", py::overload_cast<const std::string&, const std::string&>(&XLCellIsRule));
    m.def("XLFormulaRule", &XLFormulaRule);
    m.def("XLIconSetRule", &XLIconSetRule, "icon_set_name"_a = "3TrafficLights1",
          "show_value"_a = true, "reverse"_a = false);
    m.def("XLTop10Rule", &XLTop10Rule, "rank"_a = 10, "percent"_a = false,
          "bottom"_a = false);
    m.def("XLAboveAverageRule", &XLAboveAverageRule, "above_average"_a = true,
          "equal_average"_a = false, "std_dev"_a = 0);
    m.def("XLDuplicateValuesRule", &XLDuplicateValuesRule, "unique"_a = false);
    m.def("XLContainsTextRule", &XLContainsTextRule);
    m.def("XLNotContainsTextRule", &XLNotContainsTextRule);
    m.def("XLContainsBlanksRule", &XLContainsBlanksRule);
    m.def("XLNotContainsBlanksRule", &XLNotContainsBlanksRule);
    m.def("XLContainsErrorsRule", &XLContainsErrorsRule);
    m.def("XLNotContainsErrorsRule", &XLNotContainsErrorsRule);
}
