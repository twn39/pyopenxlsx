#include <nanobind/stl/optional.h>

#include "bindings.hpp"

namespace {
struct XLDefinedNameIterator {
    std::vector<XLDefinedName> names;
    size_t index = 0;

    XLDefinedNameIterator(std::vector<XLDefinedName>&& n) : names(std::move(n)) {}

    XLDefinedName* next() {
        if (index < names.size()) return &names[index++];
        return nullptr;
    }
};
}  // namespace

void init_defined_names(py::module_& m) {
    // Bind XLDefinedNameIterator
    py::class_<XLDefinedNameIterator>(m, "XLDefinedNameIterator")
        .def("__next__",
             [](XLDefinedNameIterator& self) -> XLDefinedName {
                 auto* res = self.next();
                 if (!res) throw py::stop_iteration();
                 return *res;
             })
        .def("__iter__", [](XLDefinedNameIterator& self) { return self; });

    // Bind XLDefinedName
    py::class_<XLDefinedName>(m, "XLDefinedName")
        .def(py::init<>())
        .def("name", &XLDefinedName::name)
        .def("set_name", &XLDefinedName::setName)
        .def("refers_to", &XLDefinedName::refersTo)
        .def("set_refers_to", &XLDefinedName::setRefersTo)
        .def("local_sheet_id",
             [](const XLDefinedName& self) -> py::object {
                 auto val = self.localSheetId();
                 return val ? py::cast(*val) : py::none();
             })
        .def("set_local_sheet_id", &XLDefinedName::setLocalSheetId)
        .def("hidden", &XLDefinedName::hidden)
        .def("set_hidden", &XLDefinedName::setHidden)
        .def("comment", &XLDefinedName::comment)
        .def("set_comment", &XLDefinedName::setComment)
        .def("valid", &XLDefinedName::valid);

    // Bind XLDefinedNames
    py::class_<XLDefinedNames>(m, "XLDefinedNames")
        .def(
            "append",
            [](XLDefinedNames& self, std::string_view name, std::string_view formula) {
                return self.append(name, formula, std::nullopt);
            },
            py::arg("name"), py::arg("formula"))
        .def(
            "append",
            [](XLDefinedNames& self, std::string_view name, std::string_view formula,
               uint32_t localSheetId) { return self.append(name, formula, localSheetId); },
            py::arg("name"), py::arg("formula"), py::arg("local_sheet_id"))
        .def(
            "remove",
            [](XLDefinedNames& self, std::string_view name) { self.remove(name, std::nullopt); },
            py::arg("name"))
        .def(
            "remove",
            [](XLDefinedNames& self, std::string_view name, uint32_t localSheetId) {
                self.remove(name, localSheetId);
            },
            py::arg("name"), py::arg("local_sheet_id"))
        .def(
            "get",
            [](const XLDefinedNames& self, std::string_view name) {
                return self.get(name, std::nullopt);
            },
            py::arg("name"))
        .def(
            "get",
            [](const XLDefinedNames& self, std::string_view name, uint32_t localSheetId) {
                return self.get(name, localSheetId);
            },
            py::arg("name"), py::arg("local_sheet_id"))
        .def("all", &XLDefinedNames::all)
        .def(
            "exists",
            [](const XLDefinedNames& self, std::string_view name) {
                return self.exists(name, std::nullopt);
            },
            py::arg("name"))
        .def(
            "exists",
            [](const XLDefinedNames& self, std::string_view name, uint32_t localSheetId) {
                return self.exists(name, localSheetId);
            },
            py::arg("name"), py::arg("local_sheet_id"))
        .def("count", &XLDefinedNames::count)
        .def("__len__", &XLDefinedNames::count)
        .def("__getitem__",
             [](const XLDefinedNames& self, const std::string& name) {
                 auto val = self.get(name);
                 if (!val.valid())
                     throw py::key_error(("Defined name '" + name + "' not found").c_str());
                 return val;
             })
        .def(
            "__iter__",
            [](const XLDefinedNames& self) { return XLDefinedNameIterator(self.all()); },
            py::keep_alive<0, 1>());
}
