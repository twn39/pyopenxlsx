import asyncio
import tempfile
import os
from weakref import WeakValueDictionary

from . import _openxlsx
from ._openxlsx import XLProperty, XLLineStyle, XLPatternType
from .worksheet import Worksheet
from .styles import Style



class DocumentProperties:
    """
    High-level wrapper for Excel document properties.
    Supports both Core and App/Extended properties.
    """

    _PROPERTY_MAP = {
        "title": XLProperty.Title,
        "subject": XLProperty.Subject,
        "creator": XLProperty.Creator,
        "keywords": XLProperty.Keywords,
        "description": XLProperty.Description,
        "last_modified_by": XLProperty.LastModifiedBy,
        "last_printed": XLProperty.LastPrinted,
        "created": XLProperty.CreationDate,
        "modified": XLProperty.ModificationDate,
        "category": XLProperty.Category,
        "application": XLProperty.Application,
        "doc_security": XLProperty.DocSecurity,
        "scale_crop": XLProperty.ScaleCrop,
        "manager": XLProperty.Manager,
        "company": XLProperty.Company,
        "links_up_to_date": XLProperty.LinksUpToDate,
        "shared_doc": XLProperty.SharedDoc,
        "hyperlink_base": XLProperty.HyperlinkBase,
        "hyperlinks_changed": XLProperty.HyperlinksChanged,
        "app_version": XLProperty.AppVersion,
    }

    def __init__(self, doc):
        self._doc = doc

    def __getitem__(self, key):
        if isinstance(key, XLProperty):
            return self._doc.property(key)

        prop = self._PROPERTY_MAP.get(key.lower().replace(" ", "_"))
        if prop is not None:
            return self._doc.property(prop)

        # Fallback to string-based lookup in AppProperties (most flexible)
        return self._doc.app_properties().property(key)

    def __setitem__(self, key, value):
        if isinstance(key, XLProperty):
            self._doc.set_property(key, str(value))
            return

        prop = self._PROPERTY_MAP.get(key.lower().replace(" ", "_"))
        if prop is not None:
            self._doc.set_property(prop, str(value))
        else:
            # Fallback to string-based set in AppProperties
            self._doc.app_properties().set_property(key, str(value))

    def __delitem__(self, key):
        if isinstance(key, XLProperty):
            self._doc.delete_property(key)
            return

        prop = self._PROPERTY_MAP.get(key.lower().replace(" ", "_"))
        if prop is not None:
            self._doc.delete_property(prop)
        else:
            self._doc.app_properties().delete_property(key)

    @property
    def title(self):
        return self[XLProperty.Title]

    @title.setter
    def title(self, value):
        self[XLProperty.Title] = value

    @property
    def creator(self):
        return self[XLProperty.Creator]

    @creator.setter
    def creator(self, value):
        self[XLProperty.Creator] = value

    @property
    def last_modified_by(self):
        return self[XLProperty.LastModifiedBy]

    @last_modified_by.setter
    def last_modified_by(self, value):
        self[XLProperty.LastModifiedBy] = value

    @property
    def subject(self):
        return self[XLProperty.Subject]

    @subject.setter
    def subject(self, value):
        self[XLProperty.Subject] = value

    @property
    def description(self):
        return self[XLProperty.Description]

    @description.setter
    def description(self, value):
        self[XLProperty.Description] = value

    @property
    def keywords(self):
        return self[XLProperty.Keywords]

    @keywords.setter
    def keywords(self, value):
        self[XLProperty.Keywords] = value

    @property
    def category(self):
        return self[XLProperty.Category]

    @category.setter
    def category(self, value):
        self[XLProperty.Category] = value

    @property
    def company(self):
        return self[XLProperty.Company]

    @company.setter
    def company(self, value):
        self[XLProperty.Company] = value


class Workbook:
    """
    Represents an Excel workbook.

    Uses WeakValueDictionary for worksheet caching to allow garbage collection
    of Worksheet objects when they are no longer referenced elsewhere.
    """

    def __init__(self, filename=None, force_overwrite=True):
        self._doc = _openxlsx.XLDocument()
        self._temp_file = None  # Track temp file for cleanup
        if filename:
            self._doc.open(str(filename))
            self._filename = str(filename)
        else:
            # Use a temporary file to avoid polluting the current directory
            # OpenXLSX's create() writes to disk immediately
            fd, temp_path = tempfile.mkstemp(suffix=".xlsx", prefix="pyopenxlsx_")
            os.close(fd)  # Close the file descriptor, XLDocument will manage the file
            self._temp_file = temp_path
            self._doc.create(temp_path, force_overwrite)
            self._filename = None
        self._wb = self._doc.workbook()
        # Use WeakValueDictionary to avoid keeping Worksheet objects alive indefinitely
        # Worksheets will be garbage collected when no external references remain
        self._sheets = WeakValueDictionary()
        self._styles = None
        self._date_format_cache = {}

    def save(self, filename=None, force_overwrite=True):
        if filename:
            self._doc.save_as(str(filename), force_overwrite)
        elif self._filename:
            self._doc.save()
        else:
            raise ValueError("No filename specified")

    async def save_async(self, filename=None, force_overwrite=True):
        await asyncio.to_thread(self.save, filename, force_overwrite)

    def close(self):
        self._doc.close()
        # Clean up temporary file if it was created
        if self._temp_file and os.path.exists(self._temp_file):
            try:
                os.unlink(self._temp_file)
            except OSError:
                pass  # Ignore errors during cleanup
            self._temp_file = None

    async def close_async(self):
        await asyncio.to_thread(self.close)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    async def __aenter__(self):
        return self

    async def __aexit__(self, exc_type, exc_val, exc_tb):
        await self.close_async()

    @property
    def styles(self):
        if self._styles is None:
            self._styles = self._doc.styles()
        return self._styles

    @property
    def properties(self):
        if not hasattr(self, "_properties"):
            self._properties = DocumentProperties(self._doc)
        return self._properties

    def add_style(
        self,
        font=None,
        fill=None,
        border=None,
        alignment=None,
        number_format=None,
        protection=None,
    ):
        style_obj = None
        if isinstance(font, Style):
            style_obj = font
            font = style_obj.font
            fill = style_obj.fill
            border = style_obj.border
            alignment = style_obj.alignment
            number_format = style_obj.number_format
            protection = style_obj.protection

        # Create a new cell format entry (default)
        index = self.styles.cell_formats().create()
        xf = self.styles.cell_formats().cell_format_by_index(index)

        if font is not None:
            if isinstance(font, int):
                xf.set_font_index(font)
            else:
                fonts = self.styles.fonts()
                idx = fonts.create()
                target_font = fonts.font_by_index(idx)
                target_font.set_name(font.name())
                target_font.set_size(font.size())
                target_font.set_bold(font.bold())
                target_font.set_italic(font.italic())
                # TODO: Handle underline, etc. if added to Font class
                if font.color():
                    target_font.set_color(font.color())
                xf.set_font_index(idx)
            xf.set_apply_font(True)

        if fill is not None:
            if isinstance(fill, int):
                xf.set_fill_index(fill)
            else:
                fills = self.styles.fills()
                idx = fills.create()
                target_fill = fills.fill_by_index(idx)
                
                # Check for None pattern
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
                borders = self.styles.borders()
                idx = borders.create()
                target_border = borders.border_by_index(idx)
                
                line_none = getattr(XLLineStyle, "None")
                
                l = border.left()
                if l and l.style() and l.style() != line_none: target_border.set_left(l.style(), l.color())
                
                r = border.right()
                if r and r.style() and r.style() != line_none: target_border.set_right(r.style(), r.color())
                
                t = border.top()
                if t and t.style() and t.style() != line_none: target_border.set_top(t.style(), t.color())
                
                b = border.bottom()
                if b and b.style() and b.style() != line_none: target_border.set_bottom(b.style(), b.color())
                
                d = border.diagonal()
                if d and d.style() and d.style() != line_none: target_border.set_diagonal(d.style(), d.color())

                xf.set_border_index(idx)
            xf.set_apply_border(True)

        if alignment:
            target_align = xf.alignment(True)
            if alignment.horizontal():
                target_align.set_horizontal(alignment.horizontal())
            if alignment.vertical():
                target_align.set_vertical(alignment.vertical())
            target_align.set_wrap_text(alignment.wrap_text())
            xf.set_apply_alignment(True)

        if number_format:
            if isinstance(number_format, int):
                # Assume it's a numberFormatId
                xf.set_number_format_id(number_format)
            elif isinstance(number_format, str):
                # Check if this format code already exists
                nfs = self.styles.number_formats()
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
                    # Create new custom format
                    max_id = 163
                    for i in range(count):
                        nf = nfs.number_format_by_index(i)
                        if nf.number_format_id() > max_id:
                            max_id = nf.number_format_id()

                    new_id = max_id + 1

                    # Create new empty number format entry
                    nfs.create()
                    # Retrieve it (assume appended)
                    nf = nfs.number_format_by_index(nfs.count() - 1)
                    nf.set_number_format_id(new_id)
                    nf.set_format_code(number_format)

                    xf.set_number_format_id(new_id)

            xf.set_apply_number_format(True)

        if protection:
            target_prot = xf
            if hasattr(protection, "locked"):
                target_prot.set_locked(protection.locked)
            if hasattr(protection, "hidden"):
                target_prot.set_hidden(protection.hidden)
            xf.set_apply_protection(True)

        if style_obj:
            style_obj.style_index = index

        return index

    async def add_style_async(
        self,
        font=None,
        fill=None,
        border=None,
        alignment=None,
        number_format=None,
        protection=None,
    ):
        return await asyncio.to_thread(
            self.add_style, font, fill, border, alignment, number_format, protection
        )

    @property
    def workbook(self):
        return self._wb

    @property
    def active(self):
        try:
            names = self.sheetnames
            for name in names:
                ws = self.workbook.worksheet(name)
                if ws.is_active():
                    return Worksheet(ws, self)
        except Exception:
            pass

        if self.sheetnames:
            return self[self.sheetnames[0]]
        return None

    @active.setter
    def active(self, ws):
        if not isinstance(ws, Worksheet):
            raise TypeError("Must be a Worksheet object")
        self.workbook.clear_active_tab()
        ws._sheet.set_active()

    def create_sheet(self, title=None, index=None):
        if title is None:
            i = 1
            while f"Sheet{i}" in self.sheetnames:
                i += 1
            title = f"Sheet{i}"
        self.workbook.add_worksheet(title)
        ws = self[title]
        if index is not None:
            ws._sheet.set_index(index + 1)
        return ws

    async def create_sheet_async(self, title=None, index=None):
        return await asyncio.to_thread(self.create_sheet, title, index)

    def remove(self, worksheet):
        self.workbook.delete_sheet(worksheet.title)

    async def remove_async(self, worksheet):
        await asyncio.to_thread(self.remove, worksheet)

    def copy_worksheet(self, from_worksheet):
        new_name = f"{from_worksheet.title} Copy"
        i = 1
        while new_name in self.sheetnames:
            new_name = f"{from_worksheet.title} Copy{i}"
            i += 1
        self.workbook.clone_sheet(from_worksheet.title, new_name)
        return self[new_name]

    async def copy_worksheet_async(self, from_worksheet):
        return await asyncio.to_thread(self.copy_worksheet, from_worksheet)

    @property
    def sheetnames(self):
        return list(self.workbook.worksheet_names())

    def __getitem__(self, key):
        if key in self._sheets:
            return self._sheets[key]
        if key in self.sheetnames:
            ws = Worksheet(self.workbook.worksheet(key), self)
            self._sheets[key] = ws
            return ws
        raise KeyError(f"Worksheet {key} does not exist")

    def __delitem__(self, key):
        if key in self.sheetnames:
            self.workbook.delete_sheet(key)
            if key in self._sheets:
                del self._sheets[key]
        else:
            raise KeyError(f"Worksheet {key} does not exist")

    def __iter__(self):
        for name in self.sheetnames:
            yield self[name]

    def __len__(self):
        return self.workbook.sheet_count()

    def __contains__(self, key):
        return key in self.sheetnames

    def get_embedded_images(self):
        """
        Get a list of all embedded images in the workbook.

        Returns:
            list[ImageInfo]: List of ImageInfo objects with name, path, and extension attributes.

        Example:
            >>> wb = Workbook("test.xlsx")
            >>> images = wb.get_embedded_images()
            >>> for img in images:
            ...     print(f"Found image: {img.name}")
        """
        return self._doc.get_embedded_images()

    def get_image_data(self, image_path_or_name):
        """
        Get the binary data of an embedded image.

        Args:
            image_path_or_name: Either the full path (e.g., 'xl/media/image1.png')
                               or just the filename (e.g., 'image1.png')

        Returns:
            bytes: The raw binary data of the image.

        Raises:
            RuntimeError: If the image is not found in the archive.

        Example:
            >>> wb = Workbook("test.xlsx")
            >>> images = wb.get_embedded_images()
            >>> if images:
            ...     data = wb.get_image_data(images[0].name)
            ...     with open("extracted_image.png", "wb") as f:
            ...         f.write(data)
        """
        return self._doc.get_image_data(image_path_or_name)

    def extract_images(self, output_dir):
        """
        Extract all embedded images to a directory.

        Args:
            output_dir: Directory path where images will be saved.
                       Will be created if it doesn't exist.

        Returns:
            list[str]: List of paths to the extracted image files.

        Example:
            >>> wb = Workbook("test.xlsx")
            >>> extracted = wb.extract_images("./images/")
            >>> print(f"Extracted {len(extracted)} images")
        """
        import os

        os.makedirs(output_dir, exist_ok=True)

        images = self.get_embedded_images()
        extracted_paths = []

        for img in images:
            data = self.get_image_data(img.path)
            output_path = os.path.join(output_dir, img.name)
            with open(output_path, "wb") as f:
                f.write(data)
            extracted_paths.append(output_path)

        return extracted_paths

    async def extract_images_async(self, output_dir):
        """Async version of extract_images."""
        return await asyncio.to_thread(self.extract_images, output_dir)

    def __del__(self):
        # Ensure temporary file is cleaned up even if close() was not called
        if (
            hasattr(self, "_temp_file")
            and self._temp_file
            and os.path.exists(self._temp_file)
        ):
            try:
                os.unlink(self._temp_file)
            except OSError:
                pass


def load_workbook(filename):
    return Workbook(filename)


async def load_workbook_async(filename):
    return await asyncio.to_thread(load_workbook, filename)
