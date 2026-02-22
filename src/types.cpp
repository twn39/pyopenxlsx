#include <ctime>

#include "bindings.hpp"

void init_types(py::module_& m) {
    // Bind XLContentItem
    py::class_<XLContentItem>(m, "XLContentItem")
        .def(py::init<>())
        .def("type", &XLContentItem::type)
        .def("path", &XLContentItem::path);

    // Bind XLContentTypes
    py::class_<XLContentTypes>(m, "XLContentTypes")
        .def("add_override", &XLContentTypes::addOverride)
        .def("delete_override",
             py::overload_cast<const std::string&>(&XLContentTypes::deleteOverride))
        .def("delete_override",
             py::overload_cast<const XLContentItem&>(&XLContentTypes::deleteOverride))
        .def("content_item", &XLContentTypes::contentItem)
        .def("get_content_items", &XLContentTypes::getContentItems);

    // Bind XLComment
    py::class_<XLComment>(m, "XLComment")
        .def("valid", &XLComment::valid)
        .def("ref", &XLComment::ref)
        .def("text", &XLComment::text)
        .def("author_id", &XLComment::authorId)
        .def("set_text", &XLComment::setText)
        .def("set_author_id", &XLComment::setAuthorId);

    // Bind XLShapeStyle
    py::class_<XLShapeStyle>(m, "XLShapeStyle")
        .def("width", &XLShapeStyle::width)
        .def("height", &XLShapeStyle::height)
        .def("set_width", &XLShapeStyle::setWidth)
        .def("set_height", &XLShapeStyle::setHeight)
        .def("set_raw", &XLShapeStyle::setRaw)
        .def("show", &XLShapeStyle::show)
        .def("hide", &XLShapeStyle::hide);

    // Bind XLShapeClientData
    py::class_<XLShapeClientData>(m, "XLShapeClientData")
        .def("auto_fill", &XLShapeClientData::autoFill)
        .def("set_auto_fill", &XLShapeClientData::setAutoFill)
        .def("set_anchor", &XLShapeClientData::setAnchor);

    // Bind XLShape
    py::class_<XLShape>(m, "XLShape")
        .def("style", &XLShape::style)
        .def("client_data", &XLShape::clientData);

    // Bind XLComments
    py::class_<XLComments>(m, "XLComments")
        .def("count", &XLComments::count)
        .def("get", py::overload_cast<size_t>(&XLComments::get, py::const_))
        .def("get", py::overload_cast<const std::string&>(&XLComments::get, py::const_))
        .def("set", &XLComments::set, py::arg("cellRef"), py::arg("comment"),
             py::arg("author_id") = 0)
        .def("shape", &XLComments::shape)
        .def("delete_comment", &XLComments::deleteComment)
        .def("author_count", &XLComments::authorCount)
        .def("author", &XLComments::author)
        .def("add_author", &XLComments::addAuthor);

    // Bind XLDateTime
    py::class_<XLDateTime>(m, "XLDateTime")
        .def(py::init<>())
        .def("__init__", [](XLDateTime* t, double serial) { new (t) XLDateTime(serial); })
        .def("__init__",
             [](XLDateTime* t, long long timestamp) {
                 new (t) XLDateTime((time_t)timestamp);
             })  // support unix timestamp
        .def("serial", &XLDateTime::serial)
        .def("as_datetime", [](const XLDateTime& self) {
            std::tm t = self.tm();
            auto datetime = py::module_::import_("datetime").attr("datetime");
            // Note: std::tm_year is years since 1900, tm_mon is 0-11
            return datetime(t.tm_year + 1900, t.tm_mon + 1, t.tm_mday, t.tm_hour, t.tm_min,
                            t.tm_sec);
        });
}
