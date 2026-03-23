from weakref import ref as weakref
from ._openxlsx import (
    XLDataValidationType,
    XLDataValidationOperator,
    XLDataValidationErrorStyle,
    XLDataValidationConfig,
)


class DataValidation:
    """
    Represents an Excel data validation rule.
    """

    __slots__ = ("_dv", "_worksheet_ref")

    def __init__(self, raw_dv, worksheet=None):
        self._dv = raw_dv
        self._worksheet_ref = weakref(worksheet) if worksheet else None

    @property
    def _worksheet(self):
        return self._worksheet_ref() if self._worksheet_ref else None

    @property
    def sqref(self):
        """Get the range (sqref) this validation applies to."""
        return self._dv.sqref()

    @sqref.setter
    def sqref(self, value):
        self._dv.set_sqref(str(value))

    @property
    def type(self):
        """Get the validation type."""
        return self._dv.type()

    @type.setter
    def type(self, value):
        self._dv.set_type(value)

    @property
    def operator(self):
        """Get the validation operator."""
        return self._dv.operator_()

    @operator.setter
    def operator(self, value):
        self._dv.set_operator(value)

    @property
    def allow_blank(self):
        """Whether blank values are allowed."""
        return self._dv.allow_blank()

    @allow_blank.setter
    def allow_blank(self, value):
        self._dv.set_allow_blank(bool(value))

    @property
    def show_drop_down(self):
        """Whether to show a drop-down list."""
        return self._dv.show_drop_down()

    @show_drop_down.setter
    def show_drop_down(self, value):
        self._dv.set_show_drop_down(bool(value))

    @property
    def show_input_message(self):
        """Whether to show the input message."""
        return self._dv.show_input_message()

    @show_input_message.setter
    def show_input_message(self, value):
        self._dv.set_show_input_message(bool(value))

    @property
    def show_error_message(self):
        """Whether to show the error message."""
        return self._dv.show_error_message()

    @show_error_message.setter
    def show_error_message(self, value):
        self._dv.set_show_error_message(bool(value))

    @property
    def formula1(self):
        """The first formula for the validation."""
        return self._dv.formula1()

    @formula1.setter
    def formula1(self, value):
        self._dv.set_formula1(str(value))

    @property
    def formula2(self):
        """The second formula for the validation."""
        return self._dv.formula2()

    @formula2.setter
    def formula2(self, value):
        self._dv.set_formula2(str(value))

    @property
    def error_style(self):
        """Get the error style."""
        return self._dv.error_style()

    @error_style.setter
    def error_style(self, value):
        # Allow setting by string or enum
        if isinstance(value, str):
            if value.lower() == "stop":
                self._dv.set_error(
                    self.error_title, self.error, XLDataValidationErrorStyle.Stop
                )
            elif value.lower() == "warning":
                self._dv.set_error(
                    self.error_title, self.error, XLDataValidationErrorStyle.Warning
                )
            elif value.lower() == "information":
                self._dv.set_error(
                    self.error_title, self.error, XLDataValidationErrorStyle.Information
                )
        else:
            self._dv.set_error(self.error_title, self.error, value)

    @property
    def ime_mode(self):
        """Get the IME mode."""
        return self._dv.ime_mode()

    @ime_mode.setter
    def ime_mode(self, value):
        self._dv.set_ime_mode(value)

    @property
    def prompt_title(self):
        """The title of the prompt message."""
        return self._dv.prompt_title()

    @property
    def prompt(self):
        """The prompt message."""
        return self._dv.prompt()

    @property
    def error_title(self):
        """The title of the error message."""
        return self._dv.error_title()

    @property
    def error(self):
        """The error message."""
        return self._dv.error()

    def set_prompt(self, title, message):
        """Set the input prompt title and message."""
        self._dv.set_prompt(str(title), str(message))

    def set_error(self, title, message, style="stop"):
        """
        Set the error message title, message and style.
        Style can be 'stop', 'warning', or 'information'.
        """
        s = XLDataValidationErrorStyle.Stop
        if style.lower() == "warning":
            s = XLDataValidationErrorStyle.Warning
        elif style.lower() == "information":
            s = XLDataValidationErrorStyle.Information
        self._dv.set_error(str(title), str(message), s)

    def add_cell(self, cell_ref):
        """Add a cell to the validation range."""
        self._dv.add_cell(str(cell_ref))

    def add_range(self, range_ref):
        """Add a range to the validation."""
        self._dv.add_range(str(range_ref))

    def set_list(self, items):
        """Set a list of allowed values."""
        self._dv.set_list([str(i) for i in items])

    def set_reference_drop_list(self, sheet_name, range_ref):
        """Set a drop-down list from a range reference on another sheet."""
        self._dv.set_reference_drop_list(str(sheet_name), str(range_ref))


class DataValidations:
    """
    Manages all data validation rules in a worksheet.
    """

    def __init__(self, raw_dvs, worksheet=None):
        self._dvs = raw_dvs
        self._worksheet_ref = weakref(worksheet) if worksheet else None

    @property
    def _worksheet(self):
        return self._worksheet_ref() if self._worksheet_ref else None

    def __len__(self):
        return self._dvs.count()

    def __iter__(self):
        ws = self._worksheet
        for raw_dv in self._dvs:
            yield DataValidation(raw_dv, ws)

    def __getitem__(self, index):
        if isinstance(index, int):
            return DataValidation(self._dvs.at(index), self._worksheet)
        elif isinstance(index, str):
            return DataValidation(self._dvs.at(index), self._worksheet)
        raise TypeError("Index must be an integer or a string (sqref)")

    def append(self):
        """Append a new empty data validation rule."""
        return DataValidation(self._dvs.append(), self._worksheet)

    def add_validation(
        self, sqref, type="none", operator="between", formula1="", formula2="", **kwargs
    ):
        """
        Convenience method to add a data validation rule.
        """
        config = XLDataValidationConfig()
        # Mapping for type
        types = {
            "none": getattr(XLDataValidationType, "None"),
            "custom": XLDataValidationType.Custom,
            "date": XLDataValidationType.Date,
            "decimal": XLDataValidationType.Decimal,
            "list": XLDataValidationType.List,
            "text_length": XLDataValidationType.TextLength,
            "time": XLDataValidationType.Time,
            "whole": XLDataValidationType.Whole,
        }
        config.type = types.get(type.lower(), getattr(XLDataValidationType, "None"))

        # Mapping for operator
        ops = {
            "between": XLDataValidationOperator.Between,
            "equal": XLDataValidationOperator.Equal,
            "greater_than": XLDataValidationOperator.GreaterThan,
            "greater_than_or_equal": XLDataValidationOperator.GreaterThanOrEqual,
            "less_than": XLDataValidationOperator.LessThan,
            "less_than_or_equal": XLDataValidationOperator.LessThanOrEqual,
            "not_between": XLDataValidationOperator.NotBetween,
            "not_equal": XLDataValidationOperator.NotEqual,
        }
        config.operator_ = ops.get(operator.lower(), XLDataValidationOperator.Between)  # type: ignore

        config.formula1 = str(formula1)
        config.formula2 = str(formula2)

        # Handle optional kwargs
        if "allow_blank" in kwargs:
            config.allow_blank = bool(kwargs["allow_blank"])
        if "show_drop_down" in kwargs:
            config.show_drop_down = bool(kwargs["show_drop_down"])
        if "show_input_message" in kwargs:
            config.show_input_message = bool(kwargs["show_input_message"])
        if "show_error_message" in kwargs:
            config.show_error_message = bool(kwargs["show_error_message"])
        if "prompt_title" in kwargs:
            config.prompt_title = str(kwargs["prompt_title"])
        if "prompt" in kwargs:
            config.prompt = str(kwargs["prompt"])
        if "error_title" in kwargs:
            config.error_title = str(kwargs["error_title"])
        if "error" in kwargs:
            config.error = str(kwargs["error"])

        return DataValidation(
            self._dvs.add_validation(config, str(sqref)), self._worksheet
        )

    def remove(self, index_or_sqref):
        """Remove a data validation rule by index or sqref."""
        if isinstance(index_or_sqref, int):
            self._dvs.remove(index_or_sqref)
        else:
            self._dvs.remove(str(index_or_sqref))

    def clear(self):
        """Clear all data validation rules."""
        self._dvs.clear()
