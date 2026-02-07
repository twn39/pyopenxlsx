#include "bindings.hpp"

void init_styles(py::module& m) {
    // Bind Style Enums
    py::enum_<XLUnderlineStyle>(m, "XLUnderlineStyle")
        .value("None", XLUnderlineNone)
        .value("Single", XLUnderlineSingle)
        .value("Double", XLUnderlineDouble)
        .export_values();

    py::enum_<XLFontSchemeStyle>(m, "XLFontSchemeStyle")
        .value("None", XLFontSchemeNone)
        .value("Major", XLFontSchemeMajor)
        .value("Minor", XLFontSchemeMinor)
        .export_values();

    py::enum_<XLVerticalAlignRunStyle>(m, "XLVerticalAlignRunStyle")
        .value("Baseline", XLBaseline)
        .value("Subscript", XLSubscript)
        .value("Superscript", XLSuperscript)
        .export_values();

    py::enum_<XLFillType>(m, "XLFillType")
        .value("Gradient", XLFillType::XLGradientFill)
        .value("Pattern", XLFillType::XLPatternFill)
        .export_values();

    py::enum_<XLPatternType>(m, "XLPatternType")
        .value("None", XLPatternNone)
        .value("Solid", XLPatternSolid)
        .value("MediumGray", XLPatternMediumGray)
        .value("DarkGray", XLPatternDarkGray)
        .value("LightGray", XLPatternLightGray)
        .value("DarkHorizontal", XLPatternDarkHorizontal)
        .value("DarkVertical", XLPatternDarkVertical)
        .value("DarkDown", XLPatternDarkDown)
        .value("DarkUp", XLPatternDarkUp)
        .value("DarkGrid", XLPatternDarkGrid)
        .value("DarkTrellis", XLPatternDarkTrellis)
        .value("LightHorizontal", XLPatternLightHorizontal)
        .value("LightVertical", XLPatternLightVertical)
        .value("LightDown", XLPatternLightDown)
        .value("LightUp", XLPatternLightUp)
        .value("LightGrid", XLPatternLightGrid)
        .value("LightTrellis", XLPatternLightTrellis)
        .value("Gray125", XLPatternGray125)
        .value("Gray0625", XLPatternGray0625)
        .export_values();

    py::enum_<XLLineStyle>(m, "XLLineStyle")
        .value("None", XLLineStyleNone)
        .value("Thin", XLLineStyleThin)
        .value("Medium", XLLineStyleMedium)
        .value("Dashed", XLLineStyleDashed)
        .value("Dotted", XLLineStyleDotted)
        .value("Thick", XLLineStyleThick)
        .value("Double", XLLineStyleDouble)
        .value("Hair", XLLineStyleHair)
        .value("MediumDashed", XLLineStyleMediumDashed)
        .value("DashDot", XLLineStyleDashDot)
        .value("MediumDashDot", XLLineStyleMediumDashDot)
        .value("DashDotDot", XLLineStyleDashDotDot)
        .value("MediumDashDotDot", XLLineStyleMediumDashDotDot)
        .value("SlantDashDot", XLLineStyleSlantDashDot)
        .export_values();

    py::enum_<XLAlignmentStyle>(m, "XLAlignmentStyle")
        .value("General", XLAlignGeneral)
        .value("Left", XLAlignLeft)
        .value("Right", XLAlignRight)
        .value("Center", XLAlignCenter)
        .value("Top", XLAlignTop)
        .value("Bottom", XLAlignBottom)
        .value("Fill", XLAlignFill)
        .value("Justify", XLAlignJustify)
        .value("CenterContinuous", XLAlignCenterContinuous)
        .value("Distributed", XLAlignDistributed)
        .export_values();

    // Bind XLColor
    py::class_<XLColor>(m, "XLColor")
        .def(py::init<>())
        .def(py::init<uint8_t, uint8_t, uint8_t, uint8_t>())
        .def(py::init<uint8_t, uint8_t, uint8_t>())
        .def(py::init<const std::string&>())
        .def("set", py::overload_cast<uint8_t, uint8_t, uint8_t, uint8_t>(&XLColor::set))
        .def("set", py::overload_cast<uint8_t, uint8_t, uint8_t>(&XLColor::set))
        .def("set", py::overload_cast<const std::string&>(&XLColor::set))
        .def("alpha", &XLColor::alpha)
        .def("red", &XLColor::red)
        .def("green", &XLColor::green)
        .def("blue", &XLColor::blue)
        .def("hex", &XLColor::hex);

    // Bind XLFont
    py::class_<XLFont>(m, "XLFont")
        .def(py::init<>())
        .def("name", &XLFont::fontName)
        .def("set_name", &XLFont::setFontName)
        .def("size", &XLFont::fontSize)
        .def("set_size", &XLFont::setFontSize)
        .def("color", &XLFont::fontColor)
        .def("set_color", &XLFont::setFontColor)
        .def("bold", &XLFont::bold)
        .def("set_bold", &XLFont::setBold, py::arg("set") = true)
        .def("italic", &XLFont::italic)
        .def("set_italic", &XLFont::setItalic, py::arg("set") = true)
        .def("strikethrough", &XLFont::strikethrough)
        .def("set_strikethrough", &XLFont::setStrikethrough, py::arg("set") = true)
        .def("underline", &XLFont::underline)
        .def("set_underline", &XLFont::setUnderline, py::arg("style") = XLUnderlineSingle)
        .def("scheme", &XLFont::scheme)
        .def("set_scheme", &XLFont::setScheme)
        .def("vert_align", &XLFont::vertAlign)
        .def("set_vert_align", &XLFont::setVertAlign);

    // Bind XLFill
    py::class_<XLFill>(m, "XLFill")
        .def(py::init<>())
        .def("fill_type", &XLFill::fillType)
        .def("set_fill_type", &XLFill::setFillType, py::arg("newFillType"),
             py::arg("force") = false)
        .def("pattern_type", &XLFill::patternType)
        .def("set_pattern_type", &XLFill::setPatternType)
        .def("color", &XLFill::color)
        .def("set_color", &XLFill::setColor)
        .def("background_color", &XLFill::backgroundColor)
        .def("set_background_color", &XLFill::setBackgroundColor);

    // Bind XLLine
    py::class_<XLLine>(m, "XLLine")
        .def(py::init<>())
        .def("style", &XLLine::style)
        .def("color", &XLLine::color)
        .def("__bool__", &XLLine::operator bool);

    // Bind XLBorder
    py::class_<XLBorder>(m, "XLBorder")
        .def(py::init<>())
        .def("left", &XLBorder::left)
        .def("right", &XLBorder::right)
        .def("top", &XLBorder::top)
        .def("bottom", &XLBorder::bottom)
        .def("diagonal", &XLBorder::diagonal)
        .def("set_left", &XLBorder::setLeft, py::arg("lineStyle"), py::arg("lineColor"),
             py::arg("lineTint") = 0.0)
        .def("set_right", &XLBorder::setRight, py::arg("lineStyle"), py::arg("lineColor"),
             py::arg("lineTint") = 0.0)
        .def("set_top", &XLBorder::setTop, py::arg("lineStyle"), py::arg("lineColor"),
             py::arg("lineTint") = 0.0)
        .def("set_bottom", &XLBorder::setBottom, py::arg("lineStyle"), py::arg("lineColor"),
             py::arg("lineTint") = 0.0)
        .def("set_diagonal", &XLBorder::setDiagonal, py::arg("lineStyle"), py::arg("lineColor"),
             py::arg("lineTint") = 0.0);

    // Bind XLAlignment
    py::class_<XLAlignment>(m, "XLAlignment")
        .def(py::init<>())
        .def("horizontal", &XLAlignment::horizontal)
        .def("set_horizontal", &XLAlignment::setHorizontal)
        .def("vertical", &XLAlignment::vertical)
        .def("set_vertical", &XLAlignment::setVertical)
        .def("rotation", &XLAlignment::textRotation)
        .def("set_rotation", &XLAlignment::setTextRotation)
        .def("wrap_text", &XLAlignment::wrapText)
        .def("set_wrap_text", &XLAlignment::setWrapText, py::arg("set") = true)
        .def("indent", &XLAlignment::indent)
        .def("set_indent", &XLAlignment::setIndent)
        .def("shrink_to_fit", &XLAlignment::shrinkToFit)
        .def("set_shrink_to_fit", &XLAlignment::setShrinkToFit, py::arg("set") = true);

    // Bind XLCellFormat
    py::class_<XLCellFormat>(m, "XLCellFormat")
        .def(py::init<>())
        .def("font_index", &XLCellFormat::fontIndex)
        .def("set_font_index", &XLCellFormat::setFontIndex)
        .def("fill_index", &XLCellFormat::fillIndex)
        .def("set_fill_index", &XLCellFormat::setFillIndex)
        .def("border_index", &XLCellFormat::borderIndex)
        .def("set_border_index", &XLCellFormat::setBorderIndex)
        .def("number_format_id", &XLCellFormat::numberFormatId)
        .def("set_number_format_id", &XLCellFormat::setNumberFormatId)
        .def("apply_number_format", &XLCellFormat::applyNumberFormat)
        .def("set_apply_number_format", &XLCellFormat::setApplyNumberFormat, py::arg("set") = true)
        .def("alignment", &XLCellFormat::alignment, py::arg("createIfMissing") = false)
        .def("apply_font", &XLCellFormat::applyFont)
        .def("set_apply_font", &XLCellFormat::setApplyFont, py::arg("set") = true)
        .def("apply_fill", &XLCellFormat::applyFill)
        .def("set_apply_fill", &XLCellFormat::setApplyFill, py::arg("set") = true)
        .def("apply_border", &XLCellFormat::applyBorder)
        .def("set_apply_border", &XLCellFormat::setApplyBorder, py::arg("set") = true)
        .def("apply_alignment", &XLCellFormat::applyAlignment)
        .def("set_apply_alignment", &XLCellFormat::setApplyAlignment, py::arg("set") = true)
        .def("locked", &XLCellFormat::locked)
        .def("set_locked", &XLCellFormat::setLocked, py::arg("set") = true)
        .def("hidden", &XLCellFormat::hidden)
        .def("set_hidden", &XLCellFormat::setHidden, py::arg("set") = true)
        .def("apply_protection", &XLCellFormat::applyProtection)
        .def("set_apply_protection", &XLCellFormat::setApplyProtection, py::arg("set") = true);

    // Bind XLFonts
    py::class_<XLFonts>(m, "XLFonts")
        .def("count", &XLFonts::count)
        .def("font_by_index", &XLFonts::fontByIndex, py::keep_alive<0, 1>())
        .def("__getitem__", &XLFonts::operator[], py::keep_alive<0, 1>())
        .def("create", &XLFonts::create, py::arg("copyFrom") = XLFont{},
             py::arg("styleEntriesPrefix") = XLDefaultStyleEntriesPrefix);

    // Bind XLFills
    py::class_<XLFills>(m, "XLFills")
        .def("count", &XLFills::count)
        .def("fill_by_index", &XLFills::fillByIndex, py::keep_alive<0, 1>())
        .def("__getitem__", &XLFills::operator[], py::keep_alive<0, 1>())
        .def("create", &XLFills::create, py::arg("copyFrom") = XLFill{},
             py::arg("styleEntriesPrefix") = XLDefaultStyleEntriesPrefix);

    // Bind XLBorders
    py::class_<XLBorders>(m, "XLBorders")
        .def("count", &XLBorders::count)
        .def("border_by_index", &XLBorders::borderByIndex, py::keep_alive<0, 1>())
        .def("__getitem__", &XLBorders::operator[], py::keep_alive<0, 1>())
        .def("create", &XLBorders::create, py::arg("copyFrom") = XLBorder{},
             py::arg("styleEntriesPrefix") = XLDefaultStyleEntriesPrefix);

    // Bind XLCellFormats
    py::class_<XLCellFormats>(m, "XLCellFormats")
        .def("count", &XLCellFormats::count)
        .def("cell_format_by_index", &XLCellFormats::cellFormatByIndex, py::keep_alive<0, 1>())
        .def("__getitem__", &XLCellFormats::operator[], py::keep_alive<0, 1>())
        .def("create", &XLCellFormats::create, py::arg("copyFrom") = XLCellFormat{},
             py::arg("styleEntriesPrefix") = XLDefaultStyleEntriesPrefix);

    // Bind XLStyles
    py::class_<XLStyles>(m, "XLStyles")
        .def("fonts", &XLStyles::fonts, py::return_value_policy::reference_internal)
        .def("fills", &XLStyles::fills, py::return_value_policy::reference_internal)
        .def("borders", &XLStyles::borders, py::return_value_policy::reference_internal)
        .def("cell_formats", &XLStyles::cellFormats, py::return_value_policy::reference_internal)
        .def("number_formats", &XLStyles::numberFormats,
             py::return_value_policy::reference_internal);

    // Bind XLNumberFormat
    py::class_<XLNumberFormat>(m, "XLNumberFormat")
        .def(py::init<>())
        .def("number_format_id", &XLNumberFormat::numberFormatId)
        .def("set_number_format_id", &XLNumberFormat::setNumberFormatId)
        .def("format_code", &XLNumberFormat::formatCode)
        .def("set_format_code", &XLNumberFormat::setFormatCode);

    // Bind XLNumberFormats
    py::class_<XLNumberFormats>(m, "XLNumberFormats")
        .def("count", &XLNumberFormats::count)
        .def("number_format_by_index", &XLNumberFormats::numberFormatByIndex,
             py::keep_alive<0, 1>())
        .def("number_format_by_id", &XLNumberFormats::numberFormatById, py::keep_alive<0, 1>())
        .def("__getitem__", &XLNumberFormats::operator[], py::keep_alive<0, 1>())
        .def("create", &XLNumberFormats::create, py::arg("copyFrom") = XLNumberFormat{},
             py::arg("styleEntriesPrefix") = XLDefaultStyleEntriesPrefix);
}
