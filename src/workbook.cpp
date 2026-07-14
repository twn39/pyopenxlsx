#include <headers/XLChartsheet.hpp>
#include <headers/XLSheet.hpp>

#include "internal_access.hpp"

using namespace OpenXLSX;
using namespace nanobind::literals;

void init_workbook(py::module_& m) {
    py::class_<XLWorkbook>(m, "XLWorkbook")
        .def("sheet", py::overload_cast<uint16_t>(&XLWorkbook::sheet), "index"_a,
             py::keep_alive<0, 1>())
        .def("sheet", py::overload_cast<std::string_view>(&XLWorkbook::sheet), "sheet_name"_a,
             py::keep_alive<0, 1>())
        .def(
            "worksheet",
            [](XLWorkbook& self, const std::string& name) { return self.worksheet(name); },
            "sheet_name"_a, py::keep_alive<0, 1>())
        .def(
            "worksheet",
            [](XLWorkbook& self, uint16_t index) { return self.worksheet(index); },
            "index"_a, py::keep_alive<0, 1>())
        .def("chartsheet", py::overload_cast<std::string_view>(&XLWorkbook::chartsheet),
             "sheet_name"_a, py::keep_alive<0, 1>())
        .def("chartsheet", py::overload_cast<uint16_t>(&XLWorkbook::chartsheet), "index"_a,
             py::keep_alive<0, 1>())
        .def("add_worksheet",
             [](XLWorkbook& self, const std::string& name) {
                 py::gil_scoped_release release;
                 self.addWorksheet(name);
             },
             "sheet_name"_a)
        .def("add_chartsheet",
             [](XLWorkbook& self, const std::string& name) {
                 py::gil_scoped_release release;
                 self.addChartsheet(name);
             },
             "sheet_name"_a)
        .def("delete_sheet",
             [](XLWorkbook& self, const std::string& name) {
                 py::gil_scoped_release release;
                 self.deleteSheet(name);
             },
             "sheet_name"_a)
        .def("clone_sheet",
             [](XLWorkbook& self, const std::string& existingName, const std::string& newName) {
                 py::gil_scoped_release release;
                 self.cloneSheet(existingName, newName);
             },
             "existing_name"_a, "new_name"_a)
        .def("set_sheet_index", &XLWorkbook::setSheetIndex, "sheet_name"_a, "index"_a)
        .def("index_of_sheet", &XLWorkbook::indexOfSheet, "sheet_name"_a)
        .def("type_of_sheet",
             py::overload_cast<std::string_view>(&XLWorkbook::typeOfSheet, py::const_),
             "sheet_name"_a)
        .def("type_of_sheet",
             py::overload_cast<unsigned int>(&XLWorkbook::typeOfSheet, py::const_), "index"_a)
        .def("sheet_count", &XLWorkbook::sheetCount)
        .def("worksheet_count", &XLWorkbook::worksheetCount)
        .def("chartsheet_count", &XLWorkbook::chartsheetCount)
        .def("sheet_names", &XLWorkbook::sheetNames)
        .def("worksheet_names", &XLWorkbook::worksheetNames)
        .def("chartsheet_names", &XLWorkbook::chartsheetNames)
        .def("sheet_exists", &XLWorkbook::sheetExists, "sheet_name"_a)
        .def("worksheet_exists", &XLWorkbook::worksheetExists, "sheet_name"_a)
        .def("chartsheet_exists", &XLWorkbook::chartsheetExists, "sheet_name"_a)
        .def("defined_names", &XLWorkbook::definedNames)
        .def("update_sheet_references", &XLWorkbook::updateSheetReferences, "old_name"_a,
             "new_name"_a)
        .def("delete_named_ranges", &XLWorkbook::deleteNamedRanges)
        .def("update_worksheet_dimensions", &XLWorkbook::updateWorksheetDimensions)
        .def("set_full_calculation_on_load", &XLWorkbook::setFullCalculationOnLoad)
        .def(
            "protect",
            [](XLWorkbook& self, bool lock_structure, bool lock_windows, std::string_view password) {
                self.protect(lock_structure, lock_windows, password);
            },
            "lock_structure"_a = true, "lock_windows"_a = false, "password"_a = "")
        .def("unprotect", &XLWorkbook::unprotect)
        .def("is_protected", &XLWorkbook::isProtected)
        .def("clear_active_tab", [](XLWorkbook& self) {
            auto bookViews = get_xml_doc(self).document_element().child("bookViews");
            if (!bookViews.empty()) {
                auto view = bookViews.first_child_of_type(node_element);
                if (!view.empty()) {
                    view.remove_attribute("activeTab");
                }
            }
        });
}
