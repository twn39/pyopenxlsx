class Formula:
    """
    Wrapper for Excel formulas.
    Allows interacting with the formula assigned to a cell.
    """

    def __init__(self, raw_cell):
        # raw_cell is the XLCell binding instance
        self._cell = raw_cell

    def __str__(self):
        return self._cell.get_formula()

    def __repr__(self):
        return f"Formula('{self._cell.get_formula()}')"

    def __eq__(self, other):
        return self._cell.get_formula() == str(other)

    @property
    def text(self):
        """Get or set the formula string."""
        return self._cell.get_formula()

    @text.setter
    def text(self, value):
        self._cell.set_formula(str(value))

    def clear(self):
        """Clear the formula from the cell."""
        self._cell.clear_formula()
