#include "internal_access.hpp"

void init_workbook(py::module_& m) {
    // Bind XLWorkbook
    py::class_<XLWorkbook>(m, "XLWorkbook")
        .def(
            "worksheet",
            [](XLWorkbook& self, const std::string& name) { return self.worksheet(name); },
            py::keep_alive<0, 1>())
        .def("add_worksheet",
             [](XLWorkbook& self, const std::string& name) {
                 py::gil_scoped_release release;
                 return self.addWorksheet(name);
             })
        .def("delete_sheet",
             [](XLWorkbook& self, const std::string& name) {
                 py::gil_scoped_release release;
                 self.deleteSheet(name);
             })
        .def("clone_sheet",
             [](XLWorkbook& self, const std::string& existingName, const std::string& newName) {
                 py::gil_scoped_release release;
                 self.cloneSheet(existingName, newName);
             })
        .def("sheet_count", &XLWorkbook::sheetCount)
        .def("worksheet_names", &XLWorkbook::worksheetNames)
        .def("sheet_exists", &XLWorkbook::sheetExists)
        .def("clear_active_tab", [](XLWorkbook& self) {
            auto bookViews = self.xmlDocument().document_element().child("bookViews");
            if (!bookViews.empty()) {
                auto view = bookViews.first_child_of_type(pugi::node_element);
                if (!view.empty()) {
                    view.remove_attribute("activeTab");
                }
            }
        });
}
