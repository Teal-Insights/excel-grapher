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


def _ndarray_grid_shape(value: object) -> tuple[int, int] | None:
    """Read a 1-D/2-D ndarray-like shape as ``(nrows, ncols)`` without converting it.

    Lets `Grid` hold the array itself instead of eagerly copying every cell into
    nested lists. 1-D buffers read as single-column grids, matching
    `_as_nested_rows_from_ndarray`. Returns `None` for anything else (including
    3-D buffers) so callers keep the nested-list path.
    """
    ndim = getattr(value, "ndim", None)
    if ndim not in (1, 2) or not callable(getattr(value, "tolist", None)):
        return None
    shape = getattr(value, "shape", None)
    if not isinstance(shape, tuple) or len(shape) != ndim:
        return None
    if not all(isinstance(extent, int) for extent in shape):
        return None
    if ndim == 1:
        return (shape[0], 1)
    return (shape[0], shape[1])


class Grid:
    """Positional raw-value access over a lazy `Range`, ndarray, or nested-list array."""

    __slots__ = ("nrows", "ncols", "_range", "_rows", "_array")

    def __init__(
        self,
        nrows: int,
        ncols: int,
        rng: Range | None,
        rows: list[list[CellValue]] | None,
        array: object = None,
    ) -> None:
        self.nrows = nrows
        self.ncols = ncols
        self._range = rng
        self._rows = rows
        self._array = array

    @staticmethod
    def wrap(value: object) -> Grid | None:
        """Wrap a range/array value; return `None` for scalar values.

        ndarray operands are kept as-is: consumers that want the buffer read
        `array`, and positional consumers pay for nested rows only on first use.
        """
        if isinstance(value, Range):
            nrows, ncols = value.shape
            return Grid(nrows, ncols, value, None)
        ndarray_shape = _ndarray_grid_shape(value)
        if ndarray_shape is not None:
            nrows, ncols = ndarray_shape
            if nrows == 0:
                return Grid(1, 1, None, [[None]])
            return Grid(nrows, ncols, None, None, value)
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

    @property
    def array(self) -> object:
        """The backing ndarray for ndarray operands, else `None`.

        Array consumers (vectorized operator fast paths) use this to skip the
        nested-list round trip; everything else goes through `at`.
        """
        return self._array

    def _nested_rows(self) -> list[list[CellValue]]:
        """Materialize (and cache) nested rows for positional access."""
        rows = self._rows
        if rows is None:
            rows = _as_nested_rows_from_ndarray(self._array)
            assert rows is not None
            self._rows = rows
        return rows

    def at(self, row0: int, col0: int) -> Scalar:
        """Return the raw value at a 0-based position (error sentinels included)."""
        if self._range is not None:
            return cast(Scalar, self._range.value_at(row0 + 1, col0 + 1))
        return cast(Scalar, self._nested_rows()[row0][col0])

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
        return [list(self._nested_rows()[row0])]

    def col_slice(self, col0: int) -> Range | list[list[CellValue]]:
        """Return one column as a lazy view (`Range` input) or nested list."""
        if self._range is not None:
            return self._range.column(col0 + 1)
        return [[row[col0]] for row in self._nested_rows()]

    def as_array(self) -> Range | list[list[CellValue]]:
        """Return the full grid as a lazy `Range` or nested-list copy."""
        if self._range is not None:
            return self._range
        return [list(row) for row in self._nested_rows()]
