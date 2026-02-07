from ._openxlsx import (
    XLColor,
    XLFont,
    XLFill,
    XLBorder,
    XLAlignment,
    XLPatternType,
    XLLineStyle,
    XLAlignmentStyle,
)


class Font:
    def __init__(self, name="Arial", size=11, bold=False, italic=False, color=None):
        self._name = name
        self._size = size
        self._bold = bold
        self._italic = italic
        if color:
            if isinstance(color, str):
                self._color = XLColor(color)
            else:
                self._color = color
        else:
            self._color = XLColor(0, 0, 0) # Default black

    def name(self): return self._name
    def set_name(self, value): self._name = value
    
    def size(self): return self._size
    def set_size(self, value): self._size = value
    
    def bold(self): return self._bold
    def set_bold(self, value): self._bold = value
    
    def italic(self): return self._italic
    def set_italic(self, value): self._italic = value
    
    def color(self): return self._color
    def set_color(self, value):
        if isinstance(value, str):
            self._color = XLColor(value)
        else:
            self._color = value


class Fill:
    def __init__(
        self, pattern_type=XLPatternType.Solid, color=None, background_color=None
    ):
        self._pattern_type = pattern_type
        self._color = None
        self._background_color = None

        if color:
            if isinstance(color, str):
                self._color = XLColor(color)
            else:
                self._color = color
        if background_color:
            if isinstance(background_color, str):
                self._background_color = XLColor(background_color)
            else:
                self._background_color = background_color

    def pattern_type(self): return self._pattern_type
    def set_pattern_type(self, value): self._pattern_type = value

    def color(self): return self._color
    def set_color(self, value):
        if isinstance(value, str):
            self._color = XLColor(value)
        else:
            self._color = value

    def background_color(self): return self._background_color
    def set_background_color(self, value):
        if isinstance(value, str):
            self._background_color = XLColor(value)
        else:
            self._background_color = value


class Alignment:
    def __init__(self, horizontal=None, vertical=None, wrap_text=False):
        self._horizontal = None
        self._vertical = None
        self._wrap_text = wrap_text

        if horizontal:
            if isinstance(horizontal, str):
                # Map string to enum
                for name, member in XLAlignmentStyle.__members__.items():
                    if name.lower() == horizontal.lower():
                        self._horizontal = member
                        break
            else:
                self._horizontal = horizontal
        
        if vertical:
            if isinstance(vertical, str):
                # Map string to enum
                for name, member in XLAlignmentStyle.__members__.items():
                    if name.lower() == vertical.lower():
                        self._vertical = member
                        break
            else:
                self._vertical = vertical

    def horizontal(self): return self._horizontal
    def set_horizontal(self, value): self._horizontal = value

    def vertical(self): return self._vertical
    def set_vertical(self, value): self._vertical = value

    def wrap_text(self): return self._wrap_text
    def set_wrap_text(self, value): self._wrap_text = value


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
        self._style = _get_line_style(style)
        if color:
            if isinstance(color, str):
                self._color = XLColor(color)
            else:
                self._color = color
        else:
            self._color = XLColor(0, 0, 0)

    def style(self): return self._style
    def color(self): return self._color

class Border:
    def __init__(
        self,
        left=None,
        right=None,
        top=None,
        bottom=None,
        diagonal=None,
        outline=None,
    ):
        self._left = Side(getattr(XLLineStyle, "None"), XLColor(0,0,0))
        self._right = Side(getattr(XLLineStyle, "None"), XLColor(0,0,0))
        self._top = Side(getattr(XLLineStyle, "None"), XLColor(0,0,0))
        self._bottom = Side(getattr(XLLineStyle, "None"), XLColor(0,0,0))
        self._diagonal = Side(getattr(XLLineStyle, "None"), XLColor(0,0,0))

        if outline:
            left = right = top = bottom = outline

        if left:
            if isinstance(left, Side): self._left = left
            else: self._left = Side(_get_line_style(left), XLColor(0, 0, 0))
        if right:
            if isinstance(right, Side): self._right = right
            else: self._right = Side(_get_line_style(right), XLColor(0, 0, 0))
        if top:
            if isinstance(top, Side): self._top = top
            else: self._top = Side(_get_line_style(top), XLColor(0, 0, 0))
        if bottom:
            if isinstance(bottom, Side): self._bottom = bottom
            else: self._bottom = Side(_get_line_style(bottom), XLColor(0, 0, 0))
        if diagonal:
            if isinstance(diagonal, Side): self._diagonal = diagonal
            else: self._diagonal = Side(_get_line_style(diagonal), XLColor(0, 0, 0))

    def left(self): return self._left
    def right(self): return self._right
    def top(self): return self._top
    def bottom(self): return self._bottom
    def diagonal(self): return self._diagonal

    def set_left(self, style, color): self._left = Side(style, color)
    def set_right(self, style, color): self._right = Side(style, color)
    def set_top(self, style, color): self._top = Side(style, color)
    def set_bottom(self, style, color): self._bottom = Side(style, color)
    def set_diagonal(self, style, color): self._diagonal = Side(style, color)


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
        self.style_index = 0  # To be set by workbook.add_style

    def __int__(self):
        return self.style_index
