"""Positional 2D cell access over lazy ranges and nested-list arrays."""

from __future__ import annotations

from collections.abc import Iterator
from typing import TypeAlias, cast

from excel_grapher.core.grid.ranges import Range
from excel_grapher.core.types import CellValue, XlError

__all__ = ["Grid", "Scalar"]

Scalar: TypeAlias = float | int | str | bool | XlError | None


def _as_nested_rows_from_ndarray(value: object) -> list[list[CellValue]] | None:
    """Convert an ndarray-like value to nested lists without importing NumPy.

    Duck-types via ``ndim`` / ``tolist`` so the grid module stays import-light
    for standalone exports that must remain NumPy-free.
    """
    ndim = getattr(value, "ndim", None)
    tolist = getattr(value, "tolist", None)
    if not isinstance(ndim, int) or not callable(tolist):
        return None
    if ndim == 0:
        return None
    raw = tolist()
    if ndim == 1:
        return [[cast(CellValue, cell)] for cell in raw]
    return cast("list[list[CellValue]]", raw)


class Grid:
    """Positional raw-value access over a lazy `Range` or nested-list array."""

    __slots__ = ("nrows", "ncols", "_range", "_rows")

    def __init__(
        self,
        nrows: int,
        ncols: int,
        rng: Range | None,
        rows: list[list[CellValue]] | None,
    ) -> None:
        self.nrows = nrows
        self.ncols = ncols
        self._range = rng
        self._rows = rows

    @staticmethod
    def wrap(value: object) -> Grid | None:
        """Wrap a range/array value; return `None` for scalar values."""
        if isinstance(value, Range):
            nrows, ncols = value.shape
            return Grid(nrows, ncols, value, None)
        ndarray_rows = _as_nested_rows_from_ndarray(value)
        if ndarray_rows is not None:
            if not ndarray_rows:
                ndarray_rows = [[None]]
            return Grid(len(ndarray_rows), len(ndarray_rows[0]), None, ndarray_rows)
        if isinstance(value, (list, tuple)):
            rows = [
                list(row) if isinstance(row, (list, tuple)) else [row]
                for row in cast("list[CellValue]", value)
            ]
            if not rows:
                rows = [[None]]
            return Grid(len(rows), len(rows[0]), None, cast("list[list[CellValue]]", rows))
        return None

    def at(self, row0: int, col0: int) -> Scalar:
        """Return the raw value at a 0-based position (error sentinels included)."""
        if self._range is not None:
            return cast(Scalar, self._range.value_at(row0 + 1, col0 + 1))
        assert self._rows is not None
        return cast(Scalar, self._rows[row0][col0])

    def at_flat(self, index0: int) -> Scalar:
        """Return the raw value at a 0-based row-major flat index."""
        row0, col0 = divmod(index0, self.ncols)
        return self.at(row0, col0)

    @property
    def size(self) -> int:
        """Total cell count."""
        return self.nrows * self.ncols

    def iter_raw(self) -> Iterator[Scalar]:
        """Yield raw values (error sentinels included) in row-major order."""
        for row0 in range(self.nrows):
            for col0 in range(self.ncols):
                yield self.at(row0, col0)

    def row_slice(self, row0: int) -> Range | list[list[CellValue]]:
        """Return one row as a lazy view (`Range` input) or nested list."""
        if self._range is not None:
            return self._range.row(row0 + 1)
        assert self._rows is not None
        return [list(self._rows[row0])]

    def col_slice(self, col0: int) -> Range | list[list[CellValue]]:
        """Return one column as a lazy view (`Range` input) or nested list."""
        if self._range is not None:
            return self._range.column(col0 + 1)
        assert self._rows is not None
        return [[row[col0]] for row in self._rows]
