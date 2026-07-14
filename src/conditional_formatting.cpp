#include "bindings.hpp"

using namespace OpenXLSX;
using namespace nanobind::literals;

void init_conditional_formatting(py::module_& m) {
    py::enum_<XLCfType>(m, "XLCfType")
        .value("Expression", XLCfType::Expression)
        .value("CellIs", XLCfType::CellIs)
        .value("ColorScale", XLCfType::ColorScale)
        .value("DataBar", XLCfType::DataBar)
        .value("IconSet", XLCfType::IconSet)
        .value("Top10", XLCfType::Top10)
        .value("UniqueValues", XLCfType::UniqueValues)
        .value("DuplicateValues", XLCfType::DuplicateValues)
        .value("ContainsText", XLCfType::ContainsText)
        .value("NotContainsText", XLCfType::NotContainsText)
        .value("BeginsWith", XLCfType::BeginsWith)
        .value("EndsWith", XLCfType::EndsWith)
        .value("ContainsBlanks", XLCfType::ContainsBlanks)
        .value("NotContainsBlanks", XLCfType::NotContainsBlanks)
        .value("ContainsErrors", XLCfType::ContainsErrors)
        .value("NotContainsErrors", XLCfType::NotContainsErrors)
        .value("TimePeriod", XLCfType::TimePeriod)
        .value("AboveAverage", XLCfType::AboveAverage);

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
        .value("EndsWith", XLCfOperator::EndsWith);

    py::enum_<XLCfTimePeriod>(m, "XLCfTimePeriod")
        .value("Today", XLCfTimePeriod::Today)
        .value("Yesterday", XLCfTimePeriod::Yesterday)
        .value("Tomorrow", XLCfTimePeriod::Tomorrow)
        .value("Last7Days", XLCfTimePeriod::Last7Days)
        .value("ThisMonth", XLCfTimePeriod::ThisMonth)
        .value("LastMonth", XLCfTimePeriod::LastMonth)
        .value("NextMonth", XLCfTimePeriod::NextMonth)
        .value("ThisWeek", XLCfTimePeriod::ThisWeek)
        .value("LastWeek", XLCfTimePeriod::LastWeek)
        .value("NextWeek", XLCfTimePeriod::NextWeek);

    py::enum_<XLCfvoType>(m, "XLCfvoType")
        .value("Min", XLCfvoType::Min)
        .value("Max", XLCfvoType::Max)
        .value("Number", XLCfvoType::Number)
        .value("Percent", XLCfvoType::Percent)
        .value("Formula", XLCfvoType::Formula)
        .value("Percentile", XLCfvoType::Percentile)
        .value("Invalid", XLCfvoType::Invalid);

    py::class_<XLCfRule>(m, "XLCfRule")
        .def(py::init<>())
        .def("empty", &XLCfRule::empty)
        .def("formula", &XLCfRule::formula)
        .def("formulas", &XLCfRule::formulas)
        .def("type", &XLCfRule::type)
        .def("dxf_id", &XLCfRule::dxfId)
        .def("priority", &XLCfRule::priority)
        .def("stop_if_true", &XLCfRule::stopIfTrue)
        .def("above_average", &XLCfRule::aboveAverage)
        .def("percent", &XLCfRule::percent)
        .def("bottom", &XLCfRule::bottom)
        .def("operator_", &XLCfRule::Operator)
        .def("text", &XLCfRule::text)
        .def("time_period", &XLCfRule::timePeriod)
        .def("rank", &XLCfRule::rank)
        .def("std_dev", &XLCfRule::stdDev)
        .def("equal_average", &XLCfRule::equalAverage)
        .def("set_formula", &XLCfRule::setFormula, "new_formula"_a, py::rv_policy::reference_internal)
        .def("add_formula", &XLCfRule::addFormula, "new_formula"_a)
        .def("clear_formulas", &XLCfRule::clearFormulas)
        .def("set_type", &XLCfRule::setType, "new_type"_a, py::rv_policy::reference_internal)
        .def("set_dxf_id", &XLCfRule::setDxfId, "new_dxf_id"_a, py::rv_policy::reference_internal)
        .def("set_priority", &XLCfRule::setPriority, "new_priority"_a,
             py::rv_policy::reference_internal)
        .def("set_stop_if_true", &XLCfRule::setStopIfTrue, "set"_a = true,
             py::rv_policy::reference_internal)
        .def("set_above_average", &XLCfRule::setAboveAverage, "set"_a = true,
             py::rv_policy::reference_internal)
        .def("set_percent", &XLCfRule::setPercent, "set"_a = true, py::rv_policy::reference_internal)
        .def("set_bottom", &XLCfRule::setBottom, "set"_a = true, py::rv_policy::reference_internal)
        .def("set_operator", &XLCfRule::setOperator, "new_operator"_a,
             py::rv_policy::reference_internal)
        .def("set_text", &XLCfRule::setText, "new_text"_a, py::rv_policy::reference_internal)
        .def("set_time_period", &XLCfRule::setTimePeriod, "new_time_period"_a,
             py::rv_policy::reference_internal)
        .def("set_rank", &XLCfRule::setRank, "new_rank"_a, py::rv_policy::reference_internal)
        .def("set_std_dev", &XLCfRule::setStdDev, "new_std_dev"_a, py::rv_policy::reference_internal)
        .def("set_equal_average", &XLCfRule::setEqualAverage, "set"_a = true,
             py::rv_policy::reference_internal)
        .def("summary", &XLCfRule::summary);

    py::class_<XLCfRules>(m, "XLCfRules")
        .def(py::init<>())
        .def("empty", &XLCfRules::empty)
        .def("count", &XLCfRules::count)
        .def("__len__", &XLCfRules::count)
        .def("max_priority_value", &XLCfRules::maxPriorityValue)
        .def("set_priority", &XLCfRules::setPriority, "cf_rule_index"_a, "new_priority"_a)
        .def("renumber_priorities", &XLCfRules::renumberPriorities, "increment"_a = 1)
        .def("cf_rule_by_index", &XLCfRules::cfRuleByIndex, "index"_a)
        .def("__getitem__", &XLCfRules::operator[], "index"_a)
        .def("create", &XLCfRules::create, "copy_from"_a = XLCfRule{},
             "cf_rule_prefix"_a = XLDefaultCfRulePrefix)
        .def("summary", &XLCfRules::summary);

    py::class_<XLConditionalFormat>(m, "XLConditionalFormat")
        .def(py::init<>())
        .def("empty", &XLConditionalFormat::empty)
        .def("sqref", &XLConditionalFormat::sqref)
        .def("cf_rules", &XLConditionalFormat::cfRules)
        .def("set_sqref", &XLConditionalFormat::setSqref, "new_sqref"_a)
        .def("summary", &XLConditionalFormat::summary);

    py::class_<XLConditionalFormats>(m, "XLConditionalFormats")
        .def(py::init<>())
        .def("empty", &XLConditionalFormats::empty)
        .def("count", &XLConditionalFormats::count)
        .def("__len__", &XLConditionalFormats::count)
        .def(
            "__getitem__",
            [](const XLConditionalFormats& self, size_t index) {
                if (index >= self.count()) throw py::index_error();
                return self[index];
            });

    m.def("XLColorScaleRule", py::overload_cast<const XLColor&, const XLColor&>(&XLColorScaleRule));
    m.def("XLColorScaleRule",
          py::overload_cast<const XLColor&, const XLColor&, const XLColor&>(&XLColorScaleRule));
    m.def("XLDataBarRule", &XLDataBarRule, "color"_a, "show_value"_a = true);
    m.def("XLCellIsRule", py::overload_cast<XLCfOperator, const std::string&>(&XLCellIsRule));
    m.def("XLCellIsRule", py::overload_cast<const std::string&, const std::string&>(&XLCellIsRule));
    m.def("XLFormulaRule", &XLFormulaRule);
    m.def("XLIconSetRule", &XLIconSetRule, "icon_set_name"_a = "3TrafficLights1",
          "show_value"_a = true, "reverse"_a = false);  // matches OpenXLSX signature
    m.def("XLTop10Rule", &XLTop10Rule, "rank"_a = 10, "percent"_a = false, "bottom"_a = false);
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
