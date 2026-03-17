from weakref import ref as weakref
from ._openxlsx import XLPageOrientation


class PageMargins:
    """
    Represents the page margins of a worksheet.
    Values are in inches.
    """

    __slots__ = ("_margins", "_worksheet_ref")

    def __init__(self, raw_margins, worksheet=None):
        self._margins = raw_margins
        self._worksheet_ref = weakref(worksheet) if worksheet else None

    @property
    def left(self):
        return self._margins.left()

    @left.setter
    def left(self, value):
        self._margins.set_left(float(value))

    @property
    def right(self):
        return self._margins.right()

    @right.setter
    def right(self, value):
        self._margins.set_right(float(value))

    @property
    def top(self):
        return self._margins.top()

    @top.setter
    def top(self, value):
        self._margins.set_top(float(value))

    @property
    def bottom(self):
        return self._margins.bottom()

    @bottom.setter
    def bottom(self, value):
        self._margins.set_bottom(float(value))

    @property
    def header(self):
        return self._margins.header()

    @header.setter
    def header(self, value):
        self._margins.set_header(float(value))

    @property
    def footer(self):
        return self._margins.footer()

    @footer.setter
    def footer(self, value):
        self._margins.set_footer(float(value))


class PrintOptions:
    """
    Represents the print options of a worksheet.
    """

    __slots__ = ("_options", "_worksheet_ref")

    def __init__(self, raw_options, worksheet=None):
        self._options = raw_options
        self._worksheet_ref = weakref(worksheet) if worksheet else None

    @property
    def grid_lines(self):
        """Whether grid lines are printed."""
        return self._options.grid_lines()

    @grid_lines.setter
    def grid_lines(self, value):
        self._options.set_grid_lines(bool(value))

    @property
    def headings(self):
        """Whether row and column headings are printed."""
        return self._options.headings()

    @headings.setter
    def headings(self, value):
        self._options.set_headings(bool(value))

    @property
    def horizontal_centered(self):
        """Whether content is horizontally centered on the page."""
        return self._options.horizontal_centered()

    @horizontal_centered.setter
    def horizontal_centered(self, value):
        self._options.set_horizontal_centered(bool(value))

    @property
    def vertical_centered(self):
        """Whether content is vertically centered on the page."""
        return self._options.vertical_centered()

    @vertical_centered.setter
    def vertical_centered(self, value):
        self._options.set_vertical_centered(bool(value))


class PageSetup:
    """
    Represents the page setup of a worksheet.
    """

    __slots__ = ("_setup", "_worksheet_ref")

    def __init__(self, raw_setup, worksheet=None):
        self._setup = raw_setup
        self._worksheet_ref = weakref(worksheet) if worksheet else None

    @property
    def paper_size(self):
        """The paper size (Excel internal paper size enumeration)."""
        return self._setup.paper_size()

    @paper_size.setter
    def paper_size(self, value):
        self._setup.set_paper_size(int(value))

    @property
    def orientation(self):
        """The page orientation (XLPageOrientation)."""
        return self._setup.orientation()

    @orientation.setter
    def orientation(self, value):
        self._setup.set_orientation(value)

    @property
    def scale(self):
        """The print scale (in percentage)."""
        return self._setup.scale()

    @scale.setter
    def scale(self, value):
        self._setup.set_scale(int(value))

    @property
    def fit_to_width(self):
        """The number of pages to fit to width."""
        return self._setup.fit_to_width()

    @fit_to_width.setter
    def fit_to_width(self, value):
        self._setup.set_fit_to_width(int(value))

    @property
    def fit_to_height(self):
        """The number of pages to fit to height."""
        return self._setup.fit_to_height()

    @fit_to_height.setter
    def fit_to_height(self, value):
        self._setup.set_fit_to_height(int(value))

    @property
    def black_and_white(self):
        """Whether to print in black and white."""
        return self._setup.black_and_white()

    @black_and_white.setter
    def black_and_white(self, value):
        self._setup.set_black_and_white(bool(value))
