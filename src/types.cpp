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
             py::overload_cast<std::string_view>(&XLContentTypes::deleteOverride))
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

    // Bind XLShapeText Alignments
    py::enum_<XLShapeTextVAlign>(m, "XLShapeTextVAlign")
        .value("Center", XLShapeTextVAlign::Center)
        .value("Top", XLShapeTextVAlign::Top)
        .value("Bottom", XLShapeTextVAlign::Bottom)
        .value("Invalid", XLShapeTextVAlign::Invalid);

    py::enum_<XLShapeTextHAlign>(m, "XLShapeTextHAlign")
        .value("Left", XLShapeTextHAlign::Left)
        .value("Right", XLShapeTextHAlign::Right)
        .value("Center", XLShapeTextHAlign::Center)
        .value("Invalid", XLShapeTextHAlign::Invalid);

    // Bind XLShapeStyle
    py::class_<XLShapeStyle>(m, "XLShapeStyle")
        .def("position", &XLShapeStyle::position)
        .def("set_position", &XLShapeStyle::setPosition)
        .def("margin_left", &XLShapeStyle::marginLeft)
        .def("set_margin_left", &XLShapeStyle::setMarginLeft)
        .def("margin_top", &XLShapeStyle::marginTop)
        .def("set_margin_top", &XLShapeStyle::setMarginTop)
        .def("width", &XLShapeStyle::width)
        .def("set_width", &XLShapeStyle::setWidth)
        .def("height", &XLShapeStyle::height)
        .def("set_height", &XLShapeStyle::setHeight)
        .def("mso_wrap_style", &XLShapeStyle::msoWrapStyle)
        .def("set_mso_wrap_style", &XLShapeStyle::setMsoWrapStyle)
        .def("v_text_anchor", &XLShapeStyle::vTextAnchor)
        .def("set_v_text_anchor", &XLShapeStyle::setVTextAnchor)
        .def("hidden", &XLShapeStyle::hidden)
        .def("visible", &XLShapeStyle::visible)
        .def("hide", &XLShapeStyle::hide)
        .def("show", &XLShapeStyle::show)
        .def("raw", &XLShapeStyle::raw)
        .def("set_raw", &XLShapeStyle::setRaw);

    // Bind XLShapeClientData
    py::class_<XLShapeClientData>(m, "XLShapeClientData")
        .def("object_type", &XLShapeClientData::objectType)
        .def("set_object_type", &XLShapeClientData::setObjectType)
        .def("move_with_cells", &XLShapeClientData::moveWithCells)
        .def("set_move_with_cells", &XLShapeClientData::setMoveWithCells)
        .def("size_with_cells", &XLShapeClientData::sizeWithCells)
        .def("set_size_with_cells", &XLShapeClientData::setSizeWithCells)
        .def("anchor", &XLShapeClientData::anchor)
        .def("set_anchor", &XLShapeClientData::setAnchor)
        .def("auto_fill", &XLShapeClientData::autoFill)
        .def("set_auto_fill", &XLShapeClientData::setAutoFill)
        .def("text_v_align", &XLShapeClientData::textVAlign)
        .def("set_text_v_align", &XLShapeClientData::setTextVAlign)
        .def("text_h_align", &XLShapeClientData::textHAlign)
        .def("set_text_h_align", &XLShapeClientData::setTextHAlign)
        .def("row", &XLShapeClientData::row)
        .def("set_row", &XLShapeClientData::setRow)
        .def("column", &XLShapeClientData::column)
        .def("set_column", &XLShapeClientData::setColumn);

    // Bind XLShape
    py::class_<XLShape>(m, "XLShape")
        .def("shape_id", &XLShape::shapeId)
        .def("fill_color", &XLShape::fillColor)
        .def("set_fill_color", &XLShape::setFillColor)
        .def("stroked", &XLShape::stroked)
        .def("set_stroked", &XLShape::setStroked)
        .def("type", &XLShape::type)
        .def("set_type", &XLShape::setType)
        .def("allow_in_cell", &XLShape::allowInCell)
        .def("set_allow_in_cell", &XLShape::setAllowInCell)
        .def("style", &XLShape::style)
        .def("set_style", py::overload_cast<std::string_view>(&XLShape::setStyle))
        .def("set_style_obj", py::overload_cast<const XLShapeStyle&>(&XLShape::setStyle))
        .def("client_data", &XLShape::clientData);

    // Bind XLComments
    py::class_<XLComments>(m, "XLComments")
        .def("count", &XLComments::count)
        .def("get", py::overload_cast<size_t>(&XLComments::get, py::const_))
        .def("get", py::overload_cast<const std::string&>(&XLComments::get, py::const_))
        .def(
            "set",
            py::overload_cast<const std::string&, const std::string&, uint16_t, uint16_t, uint16_t>(
                &XLComments::set),
            py::arg("cellRef"), py::arg("comment"), py::arg("author_id") = 0,
            py::arg("widthCols") = 4, py::arg("heightRows") = 6)
        .def("shape", py::overload_cast<const std::string&>(&XLComments::shape))
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

    // Bind XLSparklineType
    py::enum_<XLSparklineType>(m, "XLSparklineType")
        .value("Line", XLSparklineType::Line)
        .value("Column", XLSparklineType::Column)
        .value("Stacked", XLSparklineType::Stacked);

    // Bind XLImagePositioning
    py::enum_<XLImagePositioning>(m, "XLImagePositioning")
        .value("OneCell", XLImagePositioning::OneCell)
        .value("TwoCell", XLImagePositioning::TwoCell)
        .value("Absolute", XLImagePositioning::Absolute);

    // Bind XLImageOptions
    py::class_<XLImageOptions>(m, "XLImageOptions")
        .def(py::init<>())
        .def_rw("scale_x", &XLImageOptions::scaleX)
        .def_rw("scale_y", &XLImageOptions::scaleY)
        .def_rw("offset_x", &XLImageOptions::offsetX)
        .def_rw("offset_y", &XLImageOptions::offsetY)
        .def_rw("positioning", &XLImageOptions::positioning)
        .def_rw("bottom_right_cell", &XLImageOptions::bottomRightCell)
        .def_rw("print_with_sheet", &XLImageOptions::printWithSheet)
        .def_rw("locked", &XLImageOptions::locked);
}
