#include <headers/XLRelationships.hpp>

#include "bindings.hpp"

using namespace OpenXLSX;
using namespace nanobind::literals;

void init_relationships(py::module_& m) {
    py::enum_<XLRelationshipType>(m, "XLRelationshipType")
        .value("CoreProperties", XLRelationshipType::CoreProperties)
        .value("ExtendedProperties", XLRelationshipType::ExtendedProperties)
        .value("CustomProperties", XLRelationshipType::CustomProperties)
        .value("Workbook", XLRelationshipType::Workbook)
        .value("Worksheet", XLRelationshipType::Worksheet)
        .value("Chartsheet", XLRelationshipType::Chartsheet)
        .value("Dialogsheet", XLRelationshipType::Dialogsheet)
        .value("Macrosheet", XLRelationshipType::Macrosheet)
        .value("CalculationChain", XLRelationshipType::CalculationChain)
        .value("ExternalLink", XLRelationshipType::ExternalLink)
        .value("ExternalLinkPath", XLRelationshipType::ExternalLinkPath)
        .value("Theme", XLRelationshipType::Theme)
        .value("Styles", XLRelationshipType::Styles)
        .value("Chart", XLRelationshipType::Chart)
        .value("ChartStyle", XLRelationshipType::ChartStyle)
        .value("ChartColorStyle", XLRelationshipType::ChartColorStyle)
        .value("Image", XLRelationshipType::Image)
        .value("Drawing", XLRelationshipType::Drawing)
        .value("VMLDrawing", XLRelationshipType::VMLDrawing)
        .value("SharedStrings", XLRelationshipType::SharedStrings)
        .value("PrinterSettings", XLRelationshipType::PrinterSettings)
        .value("VBAProject", XLRelationshipType::VBAProject)
        .value("ControlProperties", XLRelationshipType::ControlProperties)
        .value("Comments", XLRelationshipType::Comments)
        .value("Table", XLRelationshipType::Table)
        .value("Hyperlink", XLRelationshipType::Hyperlink)
        .value("Unknown", XLRelationshipType::Unknown)
        .value("PivotTable", XLRelationshipType::PivotTable)
        .value("Slicer", XLRelationshipType::Slicer)
        .value("SlicerCache", XLRelationshipType::SlicerCache)
        .value("PivotCacheDefinition", XLRelationshipType::PivotCacheDefinition)
        .value("PivotCacheRecords", XLRelationshipType::PivotCacheRecords)
        .value("ThreadedComments", XLRelationshipType::ThreadedComments)
        .value("Person", XLRelationshipType::Person);

    py::class_<XLRelationshipItem>(m, "XLRelationshipItem")
        .def(py::init<>())
        .def("type", &XLRelationshipItem::type)
        .def("target", &XLRelationshipItem::target)
        .def("id", &XLRelationshipItem::id)
        .def("empty", &XLRelationshipItem::empty);

    py::class_<XLRelationships>(m, "XLRelationships")
        .def("relationship_by_id", &XLRelationships::relationshipById, "id"_a)
        .def("relationship_by_target", &XLRelationships::relationshipByTarget, "target"_a,
             "throw_if_not_found"_a = true)
        .def("relationships", &XLRelationships::relationships)
        .def(
            "delete_relationship",
            [](XLRelationships& self, std::string_view rel_id) {
                self.deleteRelationship(rel_id);
            },
            "rel_id"_a)
        .def(
            "delete_relationship_item",
            [](XLRelationships& self, const XLRelationshipItem& item) {
                self.deleteRelationship(item);
            },
            "item"_a)
        .def("add_relationship", &XLRelationships::addRelationship, "type"_a, "target"_a,
             "is_external"_a = false)
        .def("target_exists", &XLRelationships::targetExists, "target"_a)
        .def("id_exists", &XLRelationships::idExists, "id"_a);

    m.def("use_random_ids", &UseRandomIDs);
    m.def("use_sequential_ids", &UseSequentialIDs);
    m.def("init_random", &InitRandom, "pseudo_random"_a = false);
}
