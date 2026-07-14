"""Cell style registration helpers used by ``Workbook.add_style``.

Kept separate so Workbook remains a thin lifecycle façade while style
assembly (font/fill/border/number format clone into XF) stays testable
and independently reviewable.
"""

from __future__ import annotations

from typing import Any

from ._openxlsx import XLLineStyle, XLPatternType, XLUnderlineStyle
from .styles import Style


def register_cell_style(
    styles: Any,
    font=None,
    fill=None,
    border=None,
    alignment=None,
    number_format=None,
    protection=None,
) -> int:
    """Create a cell format on *styles* (``XLStyles``) and return its index.

    Accepts a ``Style`` object as *font* for openpyxl-style convenience
    (same as historical ``Workbook.add_style(Style(...))``).
    """
    style_obj = None
    if isinstance(font, Style):
        style_obj = font
        font = style_obj.font
        fill = style_obj.fill
        border = style_obj.border
        alignment = style_obj.alignment
        number_format = style_obj.number_format
        protection = style_obj.protection

    index = styles.cell_formats().create()
    xf = styles.cell_formats().cell_format_by_index(index)

    if font is not None:
        if isinstance(font, int):
            xf.set_font_index(font)
        else:
            fonts = styles.fonts()
            idx = fonts.create()
            target_font = fonts.font_by_index(idx)
            target_font.set_name(font.name())
            target_font.set_size(font.size())
            target_font.set_bold(font.bold())
            target_font.set_italic(font.italic())
            if hasattr(font, "underline"):
                u = font.underline()
                # Skip default "None" underline to avoid invalid OOXML.
                if u is not None and u != getattr(XLUnderlineStyle, "None"):
                    target_font.set_underline(u)
            if hasattr(font, "strikethrough") and font.strikethrough():
                target_font.set_strikethrough(True)
            if font.color():
                target_font.set_color(font.color())
            xf.set_font_index(idx)
        xf.set_apply_font(True)

    if fill is not None:
        if isinstance(fill, int):
            xf.set_fill_index(fill)
        else:
            fills = styles.fills()
            idx = fills.create()
            target_fill = fills.fill_by_index(idx)

            p_type = fill.pattern_type()
            if p_type != getattr(XLPatternType, "None"):
                target_fill.set_pattern_type(p_type)

            if fill.color():
                target_fill.set_color(fill.color())
            if fill.background_color():
                target_fill.set_background_color(fill.background_color())
            xf.set_fill_index(idx)
        xf.set_apply_fill(True)

    if border is not None:
        if isinstance(border, int):
            xf.set_border_index(border)
        else:
            borders = styles.borders()
            idx = borders.create()
            target_border = borders.border_by_index(idx)
            line_none = getattr(XLLineStyle, "None")

            left_side = border.left()
            if left_side and left_side.style() and left_side.style() != line_none:
                target_border.set_left(left_side.style(), left_side.color())

            r = border.right()
            if r and r.style() and r.style() != line_none:
                target_border.set_right(r.style(), r.color())

            t = border.top()
            if t and t.style() and t.style() != line_none:
                target_border.set_top(t.style(), t.color())

            b = border.bottom()
            if b and b.style() and b.style() != line_none:
                target_border.set_bottom(b.style(), b.color())

            d = border.diagonal()
            if d and d.style() and d.style() != line_none:
                target_border.set_diagonal(d.style(), d.color())

            xf.set_border_index(idx)
        xf.set_apply_border(True)

    if alignment:
        target_align = xf.alignment(True)
        if alignment.horizontal():
            target_align.set_horizontal(alignment.horizontal())
        if alignment.vertical():
            target_align.set_vertical(alignment.vertical())
        target_align.set_wrap_text(alignment.wrap_text())
        if hasattr(alignment, "indent"):
            target_align.set_indent(alignment.indent())
        if hasattr(alignment, "text_rotation"):
            target_align.set_rotation(alignment.text_rotation())
        if hasattr(alignment, "shrink_to_fit"):
            target_align.set_shrink_to_fit(alignment.shrink_to_fit())
        xf.set_apply_alignment(True)

    if number_format:
        if isinstance(number_format, int):
            xf.set_number_format_id(number_format)
        elif isinstance(number_format, str):
            nfs = styles.number_formats()
            found = False
            target_id = 0
            count = nfs.count()
            for i in range(count):
                nf = nfs.number_format_by_index(i)
                if nf.format_code() == number_format:
                    target_id = nf.number_format_id()
                    found = True
                    break

            if found:
                xf.set_number_format_id(target_id)
            else:
                max_id = 163
                for i in range(count):
                    nf = nfs.number_format_by_index(i)
                    if nf.number_format_id() > max_id:
                        max_id = nf.number_format_id()

                new_id = max_id + 1
                nfs.create()
                nf = nfs.number_format_by_index(nfs.count() - 1)
                nf.set_number_format_id(new_id)
                nf.set_format_code(number_format)
                xf.set_number_format_id(new_id)

        xf.set_apply_number_format(True)

    if protection:
        if hasattr(protection, "locked"):
            xf.set_locked(protection.locked)
        if hasattr(protection, "hidden"):
            xf.set_hidden(protection.hidden)
        xf.set_apply_protection(True)

    if style_obj is not None:
        style_obj.style_index = index

    return index
