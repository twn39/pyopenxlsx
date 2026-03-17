#include "bindings.hpp"

void init_rich_text(py::module_& m) {
    // Bind XLRichTextRun
    py::class_<XLRichTextRun>(m, "XLRichTextRun")
        .def(py::init<>())
        .def(py::init<const std::string&>())
        .def_prop_rw("text", &XLRichTextRun::text, &XLRichTextRun::setText)
        .def_prop_rw("font_name", 
            [](const XLRichTextRun& self) -> py::object {
                auto val = self.fontName();
                return val ? py::cast(*val) : py::none();
            },
            [](XLRichTextRun& self, const py::object& name) {
                if (name.is_none()) self.setFontName(""); // OpenXLSX doesn't have a clear font name, but we can set empty
                else self.setFontName(py::cast<std::string>(name));
            })
        .def_prop_rw("font_size",
            [](const XLRichTextRun& self) -> py::object {
                auto val = self.fontSize();
                return val ? py::cast(*val) : py::none();
            },
            [](XLRichTextRun& self, const py::object& size) {
                if (size.is_none()) {} // No-op or handle appropriately
                else self.setFontSize(py::cast<size_t>(size));
            })
        .def_prop_rw("font_color",
            [](const XLRichTextRun& self) -> py::object {
                auto val = self.fontColor();
                return val ? py::cast(*val) : py::none();
            },
            [](XLRichTextRun& self, const py::object& color) {
                if (color.is_none()) {}
                else self.setFontColor(py::cast<XLColor>(color));
            })
        .def_prop_rw("bold",
            [](const XLRichTextRun& self) -> py::object {
                auto val = self.bold();
                return val ? py::cast(*val) : py::none();
            },
            [](XLRichTextRun& self, const py::object& b) {
                if (b.is_none()) {}
                else self.setBold(py::cast<bool>(b));
            })
        .def_prop_rw("italic",
            [](const XLRichTextRun& self) -> py::object {
                auto val = self.italic();
                return val ? py::cast(*val) : py::none();
            },
            [](XLRichTextRun& self, const py::object& i) {
                if (i.is_none()) {}
                else self.setItalic(py::cast<bool>(i));
            })
        .def_prop_rw("underline",
            [](const XLRichTextRun& self) -> py::object {
                auto val = self.underline();
                return val ? py::cast(*val) : py::none();
            },
            [](XLRichTextRun& self, const py::object& u) {
                if (u.is_none()) {}
                else self.setUnderline(py::cast<bool>(u));
            })
        .def_prop_rw("strikethrough",
            [](const XLRichTextRun& self) -> py::object {
                auto val = self.strikethrough();
                return val ? py::cast(*val) : py::none();
            },
            [](XLRichTextRun& self, const py::object& s) {
                if (s.is_none()) {}
                else self.setStrikethrough(py::cast<bool>(s));
            });

    // Bind XLRichText
    py::class_<XLRichText>(m, "XLRichText")
        .def(py::init<>())
        .def(py::init<const std::string&>())
        .def("add_run", &XLRichText::addRun)
        .def_prop_ro("runs", [](XLRichText& self) {
            return py::make_iterator(py::type<XLRichText>(), "iterator", self.runs().begin(), self.runs().end());
        }, py::keep_alive<0, 1>())
        .def("get_runs", [](XLRichText& self) { return self.runs(); }) // Vector return
        .def_prop_ro("plain_text", &XLRichText::plainText)
        .def("empty", &XLRichText::empty)
        .def("clear", &XLRichText::clear)
        .def("__str__", &XLRichText::plainText);
}
