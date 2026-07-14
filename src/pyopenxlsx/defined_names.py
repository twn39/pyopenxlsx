"""High-level façade for workbook defined names (named ranges)."""

from __future__ import annotations

from typing import Any, Iterator, List, Optional


class DefinedName:
    """Thin wrapper around a native ``XLDefinedName``.

    Forwards attribute access to the native object (``name()``,
    ``refers_to()``, …) so existing call sites keep working.
    """

    __slots__ = ("_dn",)

    def __init__(self, native: Any):
        self._dn = native

    @property
    def raw(self) -> Any:
        return self._dn

    def __getattr__(self, name: str) -> Any:
        return getattr(self._dn, name)

    def __repr__(self) -> str:
        try:
            return f"DefinedName({self._dn.name()!r} -> {self._dn.refers_to()!r})"
        except Exception:
            return "DefinedName(<invalid>)"


class DefinedNames:
    """Collection API for workbook named ranges.

    Wraps native ``XLDefinedNames`` while keeping ``append`` / ``get`` /
    ``exists`` / ``remove`` compatible with existing call sites. Prefer
    :meth:`define` for new code.
    """

    __slots__ = ("_names", "_workbook")

    def __init__(self, native: Any, workbook: Any = None):
        self._names = native
        self._workbook = workbook

    @property
    def raw(self) -> Any:
        return self._names

    def define(
        self,
        name: str,
        refers_to: str,
        *,
        local_sheet_id: Optional[int] = None,
        sheet: Any = None,
        hidden: bool = False,
        comment: str = "",
    ) -> DefinedName:
        """Define (or redefine) a named range.

        :param name: Defined name (no spaces recommended).
        :param refers_to: Formula/reference, e.g. ``\"Sheet1!$A$1:$B$10\"``.
        :param local_sheet_id: Optional 0-based sheet-local scope.
        :param sheet: Alternative to *local_sheet_id*; a high-level
            ``Worksheet`` or sheet title string.
        :param hidden: Hide the name from Excel's UI when True.
        :param comment: Optional comment on the defined name.
        """
        sheet_id = local_sheet_id
        if sheet is not None:
            sheet_id = self._resolve_sheet_id(sheet)

        # Replace existing name in the same scope for idempotent defines.
        if sheet_id is None:
            if self._names.exists(name):
                self._names.remove(name)
            self._names.append(name, refers_to)
            dn = self._names.get(name)
        else:
            if self._names.exists(name, local_sheet_id=sheet_id):
                self._names.remove(name, local_sheet_id=sheet_id)
            self._names.append(name, refers_to, sheet_id)
            dn = self._names.get(name, local_sheet_id=sheet_id)

        if hidden:
            dn.set_hidden(True)
        if comment:
            dn.set_comment(comment)
        return DefinedName(dn)

    def append(
        self,
        name: str,
        formula: str,
        local_sheet_id: Optional[int] = None,
    ) -> Any:
        """Native-compatible append (does not replace existing names)."""
        if local_sheet_id is None:
            return self._names.append(name, formula)
        return self._names.append(name, formula, local_sheet_id)

    def get(
        self, name: str, local_sheet_id: Optional[int] = None
    ) -> DefinedName:
        if local_sheet_id is None:
            return DefinedName(self._names.get(name))
        return DefinedName(self._names.get(name, local_sheet_id=local_sheet_id))

    def remove(self, name: str, local_sheet_id: Optional[int] = None) -> None:
        if local_sheet_id is None:
            self._names.remove(name)
        else:
            self._names.remove(name, local_sheet_id=local_sheet_id)

    def exists(self, name: str, local_sheet_id: Optional[int] = None) -> bool:
        if local_sheet_id is None:
            return self._names.exists(name)
        return self._names.exists(name, local_sheet_id=local_sheet_id)

    def count(self) -> int:
        return self._names.count()

    def all(self) -> List[DefinedName]:
        return [DefinedName(dn) for dn in self._names.all()]

    def __len__(self) -> int:
        return self.count()

    def __contains__(self, name: str) -> bool:
        return self.exists(name)

    def __getitem__(self, name: str) -> DefinedName:
        if not self.exists(name):
            raise KeyError(name)
        return self.get(name)

    def __iter__(self) -> Iterator[DefinedName]:
        for dn in self._names:
            yield DefinedName(dn)

    def _resolve_sheet_id(self, sheet: Any) -> int:
        if isinstance(sheet, int):
            return sheet
        if hasattr(sheet, "index"):
            return int(sheet.index)
        if isinstance(sheet, str) and self._workbook is not None:
            names = list(self._workbook.sheetnames)
            try:
                return names.index(sheet)
            except ValueError as exc:
                raise KeyError(f"Sheet not found: {sheet!r}") from exc
        raise TypeError(
            "sheet must be a Worksheet, sheet title, or 0-based sheet index"
        )

    def __getattr__(self, name: str) -> Any:
        return getattr(self._names, name)
