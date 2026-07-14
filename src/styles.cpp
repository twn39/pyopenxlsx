#include <headers/XLStyle.hpp>
#include <nanobind/stl/optional.h>

#include "bindings.hpp"
#include "internal_access.hpp"

using namespace OpenXLSX;
using namespace nanobind::literals;


void init_styles(py::module_& m) {
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
        .def(py::init<std::string_view>())
        .def("set", py::overload_cast<uint8_t, uint8_t, uint8_t, uint8_t>(&XLColor::set))
        .def("set", py::overload_cast<uint8_t, uint8_t, uint8_t>(&XLColor::set))
        .def("set", py::overload_cast<std::string_view>(&XLColor::set))
        .def("alpha", &XLColor::alpha)
        .def("red", &XLColor::red)
        .def("green", &XLColor::green)
        .def("blue", &XLColor::blue)
        .def("hex", &XLColor::hex);

    // High-level XLStyle builder (requires XLColor / pattern / line / alignment enums above)
    py::class_<XLStyle> xl_style(m, "XLStyle");
    py::class_<XLStyle::Font>(xl_style, "Font")
        .def(py::init<>())
        .def_rw("name", &XLStyle::Font::name)
        .def_rw("size", &XLStyle::Font::size)
        .def_rw("color", &XLStyle::Font::color)
        .def_rw("bold", &XLStyle::Font::bold)
        .def_rw("italic", &XLStyle::Font::italic)
        .def_rw("underline", &XLStyle::Font::underline)
        .def_rw("strikethrough", &XLStyle::Font::strikethrough);
    py::class_<XLStyle::Fill>(xl_style, "Fill")
        .def(py::init<>())
        .def_rw("pattern", &XLStyle::Fill::pattern)
        .def_rw("fg_color", &XLStyle::Fill::fgColor)
        .def_rw("bg_color", &XLStyle::Fill::bgColor);
    py::class_<XLStyle::BorderElement>(xl_style, "BorderElement")
        .def(py::init<>())
        .def_rw("style", &XLStyle::BorderElement::style)
        .def_rw("color", &XLStyle::BorderElement::color);
    py::class_<XLStyle::Border>(xl_style, "Border")
        .def(py::init<>())
        .def_rw("left", &XLStyle::Border::left)
        .def_rw("right", &XLStyle::Border::right)
        .def_rw("top", &XLStyle::Border::top)
        .def_rw("bottom", &XLStyle::Border::bottom)
        .def_rw("diagonal", &XLStyle::Border::diagonal)
        .def_rw("diagonal_up", &XLStyle::Border::diagonalUp)
        .def_rw("diagonal_down", &XLStyle::Border::diagonalDown);
    py::class_<XLStyle::Alignment>(xl_style, "Alignment")
        .def(py::init<>())
        .def_rw("horizontal", &XLStyle::Alignment::horizontal)
        .def_rw("vertical", &XLStyle::Alignment::vertical)
        .def_rw("wrap_text", &XLStyle::Alignment::wrapText)
        .def_rw("text_rotation", &XLStyle::Alignment::textRotation)
        .def_rw("indent", &XLStyle::Alignment::indent);
    xl_style.def(py::init<>())
        .def_rw("font", &XLStyle::font)
        .def_rw("fill", &XLStyle::fill)
        .def_rw("border", &XLStyle::border)
        .def_rw("alignment", &XLStyle::alignment)
        .def_rw("number_format", &XLStyle::numberFormat);

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
        .def("set_bold", &XLFont::setBold, "set"_a = true)
        .def("italic", &XLFont::italic)
        .def("set_italic", &XLFont::setItalic, "set"_a = true)
        .def("strikethrough", &XLFont::strikethrough)
        .def("set_strikethrough", &XLFont::setStrikethrough, "set"_a = true)
        .def("underline", &XLFont::underline)
        .def("set_underline", &XLFont::setUnderline, "style"_a = XLUnderlineSingle)
        .def("scheme", &XLFont::scheme)
        .def("set_scheme", &XLFont::setScheme)
        .def("vert_align", &XLFont::vertAlign)
        .def("set_vert_align", &XLFont::setVertAlign);

    // Bind XLFill
    py::class_<XLFill>(m, "XLFill")
        .def(py::init<>())
        .def("fill_type", &XLFill::fillType)
        .def("set_fill_type", &XLFill::setFillType, "newFillType"_a,
             "force"_a = false)
        .def("pattern_type", &XLFill::patternType)
        .def("set_pattern_type", &XLFill::setPatternType)
        .def("color", &XLFill::color)
        .def("set_color", &XLFill::setColor)
        .def("background_color", &XLFill::backgroundColor)
        .def("set_background_color", &XLFill::setBackgroundColor);

    // Bind XLDataBarColor
    // Must be registered before XLLine, as XLLine::color() returns XLDataBarColor.
    py::class_<XLDataBarColor>(m, "XLDataBarColor")
        .def(py::init<>())
        .def("rgb", &XLDataBarColor::rgb,
             "Get the line color from the rgb attribute as an XLColor.")
        .def("tint", &XLDataBarColor::tint,
             "Get the color tint value (between -1.0 and +1.0).")
        .def("automatic", &XLDataBarColor::automatic)
        .def("indexed", &XLDataBarColor::indexed)
        .def("theme", &XLDataBarColor::theme)
        .def("set_rgb", &XLDataBarColor::setRgb, "color"_a,
             "Set the color from an XLColor (rgb attribute).")
        .def("set", &XLDataBarColor::set, "color"_a,
             "Alias for set_rgb.")
        .def("set_tint", &XLDataBarColor::setTint, "tint"_a,
             "Set the color tint (value between -1.0 and +1.0).")
        .def("set_automatic", &XLDataBarColor::setAutomatic, "set"_a = true)
        .def("set_indexed", &XLDataBarColor::setIndexed, "index"_a)
        .def("set_theme", &XLDataBarColor::setTheme, "theme"_a)
        .def("summary", &XLDataBarColor::summary);

    // Bind XLLine
    // NOTE: XLLine has no set_style(). Line style must be modified via
    // XLBorder.set_left/right/top/bottom/diagonal (which guarantee correct XML structure).
    // Color can be modified via the XLDataBarColor object returned by color().
    py::class_<XLLine>(m, "XLLine")
        .def(py::init<>())
        .def("style", &XLLine::style,
             "Get the line style as an XLLineStyle enum.")
        .def("color", &XLLine::color,
             "Get the line color as an XLDataBarColor object (supports in-place mutation).")
        .def("summary", &XLLine::summary)
        .def("__bool__", &XLLine::operator bool,
             "True if the line is set (i.e. has a non-None style).");

    // Bind XLBorder
    py::class_<XLBorder>(m, "XLBorder")
        .def(py::init<>())
        // --- Getters for line objects ---
        .def("left", &XLBorder::left)
        .def("right", &XLBorder::right)
        .def("top", &XLBorder::top)
        .def("bottom", &XLBorder::bottom)
        .def("diagonal", &XLBorder::diagonal)
        .def("vertical", &XLBorder::vertical,
             "Get the inner vertical line (used in merged cell borders).")
        .def("horizontal", &XLBorder::horizontal,
             "Get the inner horizontal line (used in merged cell borders).")
        // --- Boolean flag getters ---
        .def("diagonal_up", &XLBorder::diagonalUp,
             "True if the diagonal runs from bottom-left to top-right.")
        .def("diagonal_down", &XLBorder::diagonalDown,
             "True if the diagonal runs from top-left to bottom-right.")
        .def("outline", &XLBorder::outline,
             "True if the border outline attribute is set.")
        // --- Setters for line objects (return XLBorder& for chaining, exposed as None) ---
        .def("set_left", &XLBorder::setLeft, "lineStyle"_a, "lineColor"_a,
             "lineTint"_a = 0.0)
        .def("set_right", &XLBorder::setRight, "lineStyle"_a, "lineColor"_a,
             "lineTint"_a = 0.0)
        .def("set_top", &XLBorder::setTop, "lineStyle"_a, "lineColor"_a,
             "lineTint"_a = 0.0)
        .def("set_bottom", &XLBorder::setBottom, "lineStyle"_a, "lineColor"_a,
             "lineTint"_a = 0.0)
        .def("set_diagonal", &XLBorder::setDiagonal, "lineStyle"_a, "lineColor"_a,
             "lineTint"_a = 0.0)
        .def("set_vertical", &XLBorder::setVertical, "lineStyle"_a, "lineColor"_a,
             "lineTint"_a = 0.0,
             "Set the inner vertical line style (used in merged cell borders).")
        .def("set_horizontal", &XLBorder::setHorizontal, "lineStyle"_a, "lineColor"_a,
             "lineTint"_a = 0.0,
             "Set the inner horizontal line style (used in merged cell borders).")
        // --- Boolean flag setters (return XLBorder& for chaining, exposed as None) ---
        .def("set_diagonal_up", &XLBorder::setDiagonalUp, "set"_a = true,
             "Enable/disable diagonal from bottom-left to top-right.")
        .def("set_diagonal_down", &XLBorder::setDiagonalDown, "set"_a = true,
             "Enable/disable diagonal from top-left to bottom-right.")
        .def("set_outline", &XLBorder::setOutline, "set"_a = true,
             "Enable/disable the border outline attribute.")
        .def("summary", &XLBorder::summary);

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
        .def("set_wrap_text", &XLAlignment::setWrapText, "set"_a = true)
        .def("indent", &XLAlignment::indent)
        .def("set_indent", &XLAlignment::setIndent)
        .def("shrink_to_fit", &XLAlignment::shrinkToFit)
        .def("set_shrink_to_fit", &XLAlignment::setShrinkToFit, "set"_a = true);

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
        .def("set_apply_number_format", &XLCellFormat::setApplyNumberFormat, "set"_a = true)
        .def("alignment", &XLCellFormat::alignment, "createIfMissing"_a = false)
        .def("apply_font", &XLCellFormat::applyFont)
        .def("set_apply_font", &XLCellFormat::setApplyFont, "set"_a = true)
        .def("apply_fill", &XLCellFormat::applyFill)
        .def("set_apply_fill", &XLCellFormat::setApplyFill, "set"_a = true)
        .def("apply_border", &XLCellFormat::applyBorder)
        .def("set_apply_border", &XLCellFormat::setApplyBorder, "set"_a = true)
        .def("apply_alignment", &XLCellFormat::applyAlignment)
        .def("set_apply_alignment", &XLCellFormat::setApplyAlignment, "set"_a = true)
        .def("locked", &XLCellFormat::locked)
        .def("set_locked", &XLCellFormat::setLocked, "set"_a = true)
        .def("hidden", &XLCellFormat::hidden)
        .def("set_hidden", &XLCellFormat::setHidden, "set"_a = true)
        .def("apply_protection", &XLCellFormat::applyProtection)
        .def("set_apply_protection", &XLCellFormat::setApplyProtection, "set"_a = true);

    // Bind XLFonts
    py::class_<XLFonts>(m, "XLFonts")
        .def("count", &XLFonts::count)
        .def("font_by_index", &XLFonts::fontByIndex, py::keep_alive<0, 1>())
        .def("__getitem__", &XLFonts::operator[], py::keep_alive<0, 1>())
        .def("create", &XLFonts::create, "copyFrom"_a = XLFont{},
             "styleEntriesPrefix"_a = XLDefaultStyleEntriesPrefix);

    // Bind XLFills
    py::class_<XLFills>(m, "XLFills")
        .def("count", &XLFills::count)
        .def("fill_by_index", &XLFills::fillByIndex, py::keep_alive<0, 1>())
        .def("__getitem__", &XLFills::operator[], py::keep_alive<0, 1>())
        .def("create", &XLFills::create, "copyFrom"_a = XLFill{},
             "styleEntriesPrefix"_a = XLDefaultStyleEntriesPrefix);

    // Bind XLBorders
    py::class_<XLBorders>(m, "XLBorders")
        .def("count", &XLBorders::count)
        .def("border_by_index", &XLBorders::borderByIndex, py::keep_alive<0, 1>())
        .def("__getitem__", &XLBorders::operator[], py::keep_alive<0, 1>())
        .def("create", &XLBorders::create, "copyFrom"_a = XLBorder{},
             "styleEntriesPrefix"_a = XLDefaultStyleEntriesPrefix);

    // Bind XLCellFormats
    py::class_<XLCellFormats>(m, "XLCellFormats")
        .def("count", &XLCellFormats::count)
        .def("cell_format_by_index", &XLCellFormats::cellFormatByIndex, py::keep_alive<0, 1>())
        .def("__getitem__", &XLCellFormats::operator[], py::keep_alive<0, 1>())
        .def("create", &XLCellFormats::create, "copyFrom"_a = XLCellFormat{},
             "styleEntriesPrefix"_a = XLDefaultStyleEntriesPrefix);

    // Bind XLStyles
    py::class_<XLStyles>(m, "XLStyles")
        .def("fonts", &XLStyles::fonts, py::rv_policy::reference_internal)
        .def("fills", &XLStyles::fills, py::rv_policy::reference_internal)
        .def("borders", &XLStyles::borders, py::rv_policy::reference_internal)
        .def("cell_formats", &XLStyles::cellFormats, py::rv_policy::reference_internal)
        .def("number_formats", &XLStyles::numberFormats, py::rv_policy::reference_internal);

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
        .def("create", &XLNumberFormats::create, "copyFrom"_a = XLNumberFormat{},
             "styleEntriesPrefix"_a = XLDefaultStyleEntriesPrefix);

    // Bind XLDxf
    py::class_<XLDxf>(m, "XLDxf")
        .def(py::init<>())
        .def("empty", &XLDxf::empty)
        .def("font", &XLDxf::font)
        .def("num_fmt", &XLDxf::numFmt)
        .def("fill", &XLDxf::fill)
        .def("alignment", &XLDxf::alignment)
        .def("border", &XLDxf::border)
        .def("has_font", &XLDxf::hasFont)
        .def("summary", &XLDxf::summary);

    // Bind XLNumberFormatter
    py::class_<XLNumberFormatter>(m, "XLNumberFormatter")
        .def(py::init<const std::string&>(), "formatString"_a)
        .def("format", [](const XLNumberFormatter& self, py::object val) {
            CellData cd = CellData::from_python(val);
            return self.format(cd.to_xlcellvalue());
        }, "value"_a, "Formats the given Python value according to the format string");
}
