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
}
