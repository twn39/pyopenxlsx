from ._openxlsx import (
    XLColor,
    XLFont,
    XLFill,
    XLBorder,
    XLAlignment,
    XLPatternType,
    XLLineStyle,
)


class Font(XLFont):
    def __init__(self, name="Arial", size=11, bold=False, italic=False, color=None):
        super().__init__()
        self.set_name(name)
        self.set_size(size)
        self.set_bold(bold)
        self.set_italic(italic)
        if color:
            if isinstance(color, str):
                self.set_color(XLColor(color))
            else:
                self.set_color(color)


class Fill(XLFill):
    def __init__(
        self, pattern_type=XLPatternType.Solid, color=None, background_color=None
    ):
        super().__init__()
        self.set_pattern_type(pattern_type)
        if color:
            if isinstance(color, str):
                self.set_color(XLColor(color))
            else:
                self.set_color(color)
        if background_color:
            if isinstance(background_color, str):
                self.set_background_color(XLColor(background_color))
            else:
                self.set_background_color(background_color)


class Alignment(XLAlignment):
    def __init__(self, horizontal=None, vertical=None, wrap_text=False):
        super().__init__()
        if horizontal:
            self.set_horizontal(horizontal)
        if vertical:
            self.set_vertical(vertical)
        self.set_wrap_text(wrap_text)


def is_date_format(c_format):
    """
    Returns True if the given format code or id implies a date/time.
    """
    if isinstance(c_format, int):
        # Standard date formats
        # 14-22: Date/Time
        # 27-36: Date/Time (mostly Asian)
        # 45-47: Time
        return (
            (14 <= c_format <= 22) or (27 <= c_format <= 36) or (45 <= c_format <= 47)
        )

    if isinstance(c_format, str):
        # Heuristic check for date characters
        # This is simplified. Reference: OpenXML standard or similar libraries.
        # Check for y, m, d, h, s in the string, but ignore colors [Red] etc.
        # Also need to handle escaped chars.
        # For now, a simple reliable check for common patterns.
        import re

        # Remove quoted text
        fmt = re.sub(r'"[^"]*"', "", c_format)
        # Remove known colors
        fmt = re.sub(r"\[[^\]]*\]", "", fmt)
        # Check for date time tokens
        return any(x in fmt for x in ["y", "m", "d", "h", "s", "Y", "M", "D", "H", "S"])

    return False


def _get_line_style(style):
    if isinstance(style, str):
        # Capitalize logic or dictionary?
        # Enums are like "Thin", "Thick", "Dashed"
        # Let's try to match case-insensitive
        for name, member in XLLineStyle.__members__.items():
            if name.lower() == style.lower():
                return member
        # Fallback to Thin if unknown or raise?
        # Let's raise or default. Defaulting to Thin might be confusing.
        # But raising might break simple assumptions.
        # Given "thick" -> "Thick" works.
        pass
    return style


class Side:
    def __init__(self, style=XLLineStyle.Thin, color=None):
        self.style = _get_line_style(style)
        if color:
            if isinstance(color, str):
                self.color = XLColor(color)
            else:
                self.color = color
        else:
            self.color = XLColor(0, 0, 0)


class Border(XLBorder):
    def __init__(self, left=None, right=None, top=None, bottom=None, diagonal=None):
        super().__init__()
        if left:
            if isinstance(left, Side):
                self.set_left(left.style, left.color)
            else:
                # Assume simple style string or enum if passed directly, default black color
                self.set_left(_get_line_style(left), XLColor(0, 0, 0))
        if right:
            if isinstance(right, Side):
                self.set_right(right.style, right.color)
            else:
                self.set_right(_get_line_style(right), XLColor(0, 0, 0))
        if top:
            if isinstance(top, Side):
                self.set_top(top.style, top.color)
            else:
                self.set_top(_get_line_style(top), XLColor(0, 0, 0))
        if bottom:
            if isinstance(bottom, Side):
                self.set_bottom(bottom.style, bottom.color)
            else:
                self.set_bottom(_get_line_style(bottom), XLColor(0, 0, 0))
        if diagonal:
            if isinstance(diagonal, Side):
                self.set_diagonal(diagonal.style, diagonal.color)
            else:
                self.set_diagonal(_get_line_style(diagonal), XLColor(0, 0, 0))


class Protection:
    def __init__(self, locked=True, hidden=False):
        self.locked = locked
        self.hidden = hidden


class Style:
    def __init__(
        self,
        font=None,
        fill=None,
        border=None,
        alignment=None,
        number_format=None,
        protection=None,
    ):
        self.font = font
        self.fill = fill
        self.border = border
        self.alignment = alignment
        self.number_format = number_format
        self.protection = protection
