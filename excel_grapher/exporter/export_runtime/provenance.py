"""Coordinate-to-cell provenance for generated named-coordinate modules.

A regular series occupies a worksheet rectangle whose rows and columns
enumerate its axes. `row_cells`, `column_cells`, and `block_cells` describe
that rectangle once; every coordinate resolves to its authored cell on
demand, so provenance stays inspectable without listing each cell.
"""

from __future__ import annotations

from collections.abc import Callable, Iterator, Mapping, Sequence
from typing import Any, overload

from .tensor import (
    Axis,
    AxisError,
    AxisTemplate,
    Coordinate,
    CoordinateError,
    Domain,
    DomainTemplate,
)

AxisGroup = tuple[tuple[str, ...], Mapping[Any, Any]]
LayoutKeys = Axis | Sequence[str | int]


def _layout_keys(axis: LayoutKeys | None) -> tuple[str | int, ...] | None:
    if axis is None:
        return None
    return axis.keys if isinstance(axis, Axis) else tuple(axis)


def _quote_sheet(sheet: str) -> str:
    if " " in sheet or "-" in sheet or "'" in sheet:
        return "'" + sheet.replace("'", "''") + "'"
    return sheet


def column_letter(index: int) -> str:
    """Return the worksheet column letters for a 1-based column index."""
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters


def column_index(letters: str) -> int:
    """Return the 1-based column index for worksheet column letters."""
    index = 0
    for char in letters.upper():
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index


class RectangleCells(Mapping[Coordinate, str]):
    """Authored cells of one series laid out as a worksheet rectangle.

    Coordinates are the row key followed by the column key, or the reverse
    when `cols_first` is set; an absent axis contributes no key. Explicit
    `exceptions` override individual coordinates.
    """

    __slots__ = (
        "_col_axis",
        "_cols_first",
        "_exceptions",
        "_first_col",
        "_first_row",
        "_row_axis",
        "_sheet",
    )

    def __init__(
        self,
        sheet: str,
        first_row: int,
        first_col: int,
        row_axis: LayoutKeys | None,
        col_axis: LayoutKeys | None,
        *,
        cols_first: bool = False,
        exceptions: Mapping[Coordinate, str] | None = None,
    ) -> None:
        self._sheet = _quote_sheet(sheet)
        self._first_row = first_row
        self._first_col = first_col
        self._row_axis = _layout_keys(row_axis)
        self._col_axis = _layout_keys(col_axis)
        self._cols_first = cols_first
        self._exceptions = dict(exceptions or {})

    def _split(self, coordinate: Coordinate) -> tuple[Any, Any]:
        keys = list(coordinate)
        if self._cols_first:
            keys.reverse()
        row_key = keys.pop(0) if self._row_axis is not None else None
        col_key = keys.pop(0) if self._col_axis is not None else None
        if keys:
            raise KeyError(coordinate)
        return row_key, col_key

    def __getitem__(self, coordinate: Coordinate) -> str:
        if coordinate in self._exceptions:
            return self._exceptions[coordinate]
        row_key, col_key = self._split(coordinate)
        row = self._first_row
        col = self._first_col
        try:
            if self._row_axis is not None:
                row += self._row_axis.index(row_key)
            if self._col_axis is not None:
                col += self._col_axis.index(col_key)
        except ValueError:
            raise KeyError(coordinate) from None
        return f"{self._sheet}!{column_letter(col)}{row}"

    def __iter__(self) -> Iterator[Coordinate]:
        rows: tuple[Any, ...] = (None,) if self._row_axis is None else self._row_axis
        cols: tuple[Any, ...] = (None,) if self._col_axis is None else self._col_axis
        if self._cols_first:
            for col_key in cols:
                for row_key in rows:
                    yield self._coordinate(row_key, col_key)
            return
        for row_key in rows:
            for col_key in cols:
                yield self._coordinate(row_key, col_key)

    def _coordinate(self, row_key: Any, col_key: Any) -> Coordinate:
        parts = [
            key for axis, key in ((self._row_axis, row_key), (self._col_axis, col_key)) if axis
        ]
        if self._cols_first:
            parts.reverse()
        return tuple(parts)

    def __len__(self) -> int:
        rows = 1 if self._row_axis is None else len(self._row_axis)
        cols = 1 if self._col_axis is None else len(self._col_axis)
        return rows * cols

    def __repr__(self) -> str:
        return f"RectangleCells({self._sheet}, {dict(self)!r})"


class ProvenanceTemplate:
    """Authored cells that bind runtime axes before lookup."""

    __slots__ = ("_bind",)

    def __init__(self, bind: Callable[..., Mapping[Coordinate, str]]) -> None:
        self._bind = bind

    def bind(self, **labellers: Any) -> Mapping[Coordinate, str]:
        """Materialize authored cells over the bound runtime axes."""
        return self._bind(**labellers)


def _axis_from_labellers(axis: Axis | AxisTemplate, labellers: Mapping[str, Any]) -> Axis:
    if isinstance(axis, Axis):
        return axis
    labeller = labellers.get(axis.name)
    if labeller is None:
        raise AxisError(f"axis {axis.name!r}: missing labeller {axis.labeller!r}")
    return axis.bind(labeller.domain.axes[0])


@overload
def row_cells(sheet: str, row: int, first_column: str, axis: LayoutKeys) -> RectangleCells: ...


@overload
def row_cells(
    sheet: str, row: int, first_column: str, axis: AxisTemplate
) -> ProvenanceTemplate: ...


def row_cells(
    sheet: str, row: int, first_column: str, axis: LayoutKeys | AxisTemplate
) -> RectangleCells | ProvenanceTemplate:
    """Cells of a one-row series whose columns enumerate `axis`."""
    if isinstance(axis, AxisTemplate):

        def bind_row(**labellers: Any) -> RectangleCells:
            return row_cells(sheet, row, first_column, _axis_from_labellers(axis, labellers))

        return ProvenanceTemplate(bind_row)
    return RectangleCells(sheet, row, column_index(first_column), None, axis)


@overload
def column_cells(sheet: str, column: str, first_row: int, axis: LayoutKeys) -> RectangleCells: ...


@overload
def column_cells(
    sheet: str, column: str, first_row: int, axis: AxisTemplate
) -> ProvenanceTemplate: ...


def column_cells(
    sheet: str, column: str, first_row: int, axis: LayoutKeys | AxisTemplate
) -> RectangleCells | ProvenanceTemplate:
    """Cells of a one-column series whose rows enumerate `axis`."""
    if isinstance(axis, AxisTemplate):

        def bind_column(**labellers: Any) -> RectangleCells:
            return column_cells(sheet, column, first_row, _axis_from_labellers(axis, labellers))

        return ProvenanceTemplate(bind_column)
    return RectangleCells(sheet, first_row, column_index(column), axis, None)


@overload
def block_cells(
    sheet: str,
    first_row: int,
    first_column: str,
    row_axis: LayoutKeys,
    col_axis: LayoutKeys,
    *,
    cols_first: bool = False,
    exceptions: Mapping[Coordinate, str] | None = None,
) -> RectangleCells: ...


@overload
def block_cells(
    sheet: str,
    first_row: int,
    first_column: str,
    row_axis: Axis | AxisTemplate,
    col_axis: Axis | AxisTemplate,
    *,
    cols_first: bool = False,
    exceptions: Mapping[Coordinate, str] | None = None,
) -> ProvenanceTemplate: ...


def block_cells(
    sheet: str,
    first_row: int,
    first_column: str,
    row_axis: LayoutKeys | AxisTemplate,
    col_axis: LayoutKeys | AxisTemplate,
    *,
    cols_first: bool = False,
    exceptions: Mapping[Coordinate, str] | None = None,
) -> RectangleCells | ProvenanceTemplate:
    """Cells of a matrix series whose rows and columns enumerate two axes."""
    if isinstance(row_axis, AxisTemplate) or isinstance(col_axis, AxisTemplate):

        def bind_block(**labellers: Any) -> RectangleCells:
            rows: LayoutKeys = (
                _axis_from_labellers(row_axis, labellers)
                if isinstance(row_axis, AxisTemplate)
                else row_axis
            )
            cols: LayoutKeys = (
                _axis_from_labellers(col_axis, labellers)
                if isinstance(col_axis, AxisTemplate)
                else col_axis
            )
            return block_cells(
                sheet,
                first_row,
                first_column,
                rows,
                cols,
                cols_first=cols_first,
                exceptions=exceptions,
            )

        return ProvenanceTemplate(bind_block)
    return RectangleCells(
        sheet,
        first_row,
        column_index(first_column),
        row_axis,
        col_axis,
        cols_first=cols_first,
        exceptions=exceptions,
    )


class GridCells(Mapping[Coordinate, str]):
    """Authored cells of a series whose sheet, row, and column each follow a key group.

    Each worksheet position is either fixed or a mapping from the keys of a
    group of fields to that position. A group with one field maps bare keys;
    wider groups map key tuples. Coordinates enumerate `domain`, so a sparse
    domain lists only its authored cells.
    """

    __slots__ = ("_cols", "_domain", "_exceptions", "_rows", "_sheet")

    def __init__(
        self,
        sheet: str | AxisGroup,
        domain: Domain,
        *,
        rows: int | AxisGroup,
        cols: str | AxisGroup,
        exceptions: Mapping[Coordinate, str] | None = None,
    ) -> None:
        names = tuple(axis.name for axis in domain.axes)
        self._domain = domain
        self._sheet = self._group(sheet, names)
        self._rows = self._group(rows, names)
        self._cols = self._group(cols, names)
        self._exceptions = dict(exceptions or {})

    @staticmethod
    def _group(
        spec: str | int | AxisGroup, names: tuple[str, ...]
    ) -> tuple[tuple[int, ...], Mapping[Any, Any]]:
        if isinstance(spec, tuple):
            fields, mapping = spec
            return tuple(names.index(field) for field in fields), mapping
        return (), {(): spec}

    @staticmethod
    def _lookup(group: tuple[tuple[int, ...], Mapping[Any, Any]], coordinate: Coordinate) -> Any:
        positions, mapping = group
        if len(positions) == 1:
            return mapping[coordinate[positions[0]]]
        return mapping[tuple(coordinate[position] for position in positions)]

    def __getitem__(self, coordinate: Coordinate) -> str:
        if coordinate in self._exceptions:
            return self._exceptions[coordinate]
        try:
            self._domain.position(coordinate)
            sheet = self._lookup(self._sheet, coordinate)
            row = self._lookup(self._rows, coordinate)
            col = self._lookup(self._cols, coordinate)
        except (CoordinateError, KeyError):
            raise KeyError(coordinate) from None
        return f"{_quote_sheet(sheet)}!{col}{row}"

    def __iter__(self) -> Iterator[Coordinate]:
        return iter(self._domain)

    def __len__(self) -> int:
        return len(self._domain)

    def __repr__(self) -> str:
        return f"GridCells({dict(self)!r})"


@overload
def grid_cells(
    sheet: str | AxisGroup,
    domain: Domain,
    *,
    rows: int | AxisGroup,
    cols: str | AxisGroup,
    exceptions: Mapping[Coordinate, str] | None = None,
) -> GridCells: ...


@overload
def grid_cells(
    sheet: str | AxisGroup,
    domain: DomainTemplate,
    *,
    rows: int | AxisGroup,
    cols: str | AxisGroup,
    exceptions: Mapping[Coordinate, str] | None = None,
) -> ProvenanceTemplate: ...


def grid_cells(
    sheet: str | AxisGroup,
    domain: Domain | DomainTemplate,
    *,
    rows: int | AxisGroup,
    cols: str | AxisGroup,
    exceptions: Mapping[Coordinate, str] | None = None,
) -> GridCells | ProvenanceTemplate:
    """Cells of a series whose worksheet positions follow groups of its key fields."""
    if isinstance(domain, DomainTemplate):

        def bind_grid(**labellers: Any) -> GridCells:
            return grid_cells(
                sheet,
                domain.bind(**labellers),
                rows=rows,
                cols=cols,
                exceptions=exceptions,
            )

        return ProvenanceTemplate(bind_grid)
    return GridCells(sheet, domain, rows=rows, cols=cols, exceptions=exceptions)
