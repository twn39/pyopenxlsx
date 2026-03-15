#include "bindings.hpp"

void init_constants(py::module_& m) {
    // Bind XLContentType enum
    py::enum_<XLContentType>(m, "XLContentType")
        .value("Workbook", XLContentType::Workbook)
        .value("Relationships", XLContentType::Relationships)
        .value("WorkbookMacroEnabled", XLContentType::WorkbookMacroEnabled)
        .value("Worksheet", XLContentType::Worksheet)
        .value("Chartsheet", XLContentType::Chartsheet)
        .value("ExternalLink", XLContentType::ExternalLink)
        .value("Theme", XLContentType::Theme)
        .value("Styles", XLContentType::Styles)
        .value("SharedStrings", XLContentType::SharedStrings)
        .value("Drawing", XLContentType::Drawing)
        .value("Chart", XLContentType::Chart)
        .value("ChartStyle", XLContentType::ChartStyle)
        .value("ChartColorStyle", XLContentType::ChartColorStyle)
        .value("ControlProperties", XLContentType::ControlProperties)
        .value("CalculationChain", XLContentType::CalculationChain)
        .value("VBAProject", XLContentType::VBAProject)
        .value("CoreProperties", XLContentType::CoreProperties)
        .value("ExtendedProperties", XLContentType::ExtendedProperties)
        .value("CustomProperties", XLContentType::CustomProperties)
        .value("Comments", XLContentType::Comments)
        .value("Table", XLContentType::Table)
        .value("VMLDrawing", XLContentType::VMLDrawing)
        .value("Hyperlink", XLContentType::Hyperlink)
        .value("Unknown", XLContentType::Unknown);

    // Bind XLProperty enum
    py::enum_<XLProperty>(m, "XLProperty")
        .value("Title", XLProperty::Title)
        .value("Subject", XLProperty::Subject)
        .value("Creator", XLProperty::Creator)
        .value("Keywords", XLProperty::Keywords)
        .value("Description", XLProperty::Description)
        .value("LastModifiedBy", XLProperty::LastModifiedBy)
        .value("LastPrinted", XLProperty::LastPrinted)
        .value("CreationDate", XLProperty::CreationDate)
        .value("ModificationDate", XLProperty::ModificationDate)
        .value("Category", XLProperty::Category)
        .value("Application", XLProperty::Application)
        .value("DocSecurity", XLProperty::DocSecurity)
        .value("ScaleCrop", XLProperty::ScaleCrop)
        .value("Manager", XLProperty::Manager)
        .value("Company", XLProperty::Company)
        .value("LinksUpToDate", XLProperty::LinksUpToDate)
        .value("SharedDoc", XLProperty::SharedDoc)
        .value("HyperlinkBase", XLProperty::HyperlinkBase)
        .value("HyperlinksChanged", XLProperty::HyperlinksChanged)
        .value("AppVersion", XLProperty::AppVersion);

    // Bind XLSheetState
    py::enum_<XLSheetState>(m, "XLSheetState")
        .value("Visible", XLSheetState::Visible)
        .value("Hidden", XLSheetState::Hidden)
        .value("VeryHidden", XLSheetState::VeryHidden);

    // Bind Data Validation enums
    py::enum_<XLDataValidationType>(m, "XLDataValidationType")
        .value("None", XLDataValidationType::None)
        .value("Custom", XLDataValidationType::Custom)
        .value("Date", XLDataValidationType::Date)
        .value("Decimal", XLDataValidationType::Decimal)
        .value("List", XLDataValidationType::List)
        .value("TextLength", XLDataValidationType::TextLength)
        .value("Time", XLDataValidationType::Time)
        .value("Whole", XLDataValidationType::Whole);

    py::enum_<XLDataValidationOperator>(m, "XLDataValidationOperator")
        .value("Between", XLDataValidationOperator::Between)
        .value("Equal", XLDataValidationOperator::Equal)
        .value("GreaterThan", XLDataValidationOperator::GreaterThan)
        .value("GreaterThanOrEqual", XLDataValidationOperator::GreaterThanOrEqual)
        .value("LessThan", XLDataValidationOperator::LessThan)
        .value("LessThanOrEqual", XLDataValidationOperator::LessThanOrEqual)
        .value("NotBetween", XLDataValidationOperator::NotBetween)
        .value("NotEqual", XLDataValidationOperator::NotEqual);

    py::enum_<XLDataValidationErrorStyle>(m, "XLDataValidationErrorStyle")
        .value("Stop", XLDataValidationErrorStyle::Stop)
        .value("Warning", XLDataValidationErrorStyle::Warning)
        .value("Information", XLDataValidationErrorStyle::Information);

    py::enum_<XLIMEMode>(m, "XLIMEMode")
        .value("NoControl", XLIMEMode::NoControl)
        .value("Off", XLIMEMode::Off)
        .value("On", XLIMEMode::On)
        .value("Disabled", XLIMEMode::Disabled)
        .value("Hiragana", XLIMEMode::Hiragana)
        .value("FullKatakana", XLIMEMode::FullKatakana)
        .value("HalfKatakana", XLIMEMode::HalfKatakana)
        .value("FullAlpha", XLIMEMode::FullAlpha)
        .value("HalfAlpha", XLIMEMode::HalfAlpha)
        .value("FullHangul", XLIMEMode::FullHangul)
        .value("HalfHangul", XLIMEMode::HalfHangul);
}
