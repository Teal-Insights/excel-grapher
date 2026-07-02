"""Export-owned Excel value model shared by exported runtime modules."""

from __future__ import annotations

from collections.abc import Iterator
from dataclasses import dataclass
from math import isfinite
from typing import TypeAlias, cast

from excel_grapher.core import XlError
from excel_grapher.core.types import XlErrorException

from .ranges import Range

__all__ = ["CellValue", "ExcelRange", "Grid", "Scalar", "flatten"]


@dataclass(frozen=True, slots=True)
class ExcelRange:
    """Rectangular worksheet reference geometry for exported code.

    Unlike the evaluator's `ExcelRange`, this variant carries geometry only;
    exported code resolves cell values through the lazy `Range` type instead
    of eager array materialization.
    """

    sheet: str
    start_row: int
    start_col: int
    end_row: int
    end_col: int

    @property
    def shape(self) -> tuple[int, int]:
        """The reference shape as `(rows, columns)`."""
        return (self.end_row - self.start_row + 1, self.end_col - self.start_col + 1)


Scalar: TypeAlias = float | int | str | bool | XlError | None
CellValue: TypeAlias = Scalar | ExcelRange | Range | list["CellValue"]


def as_scalar(value: CellValue) -> Scalar:
    """Collapse range/array values to `#VALUE!` for scalar coercion contexts."""
    if isinstance(value, (Range, ExcelRange, list, tuple)):
        return XlError.VALUE
    return value


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


def _convergence_delta(prev: CellValue, curr: CellValue) -> float:
    if isinstance(prev, (Range, list)) or isinstance(curr, (Range, list)):
        prev_rows = prev.rows_raw() if isinstance(prev, Range) else prev
        curr_rows = curr.rows_raw() if isinstance(curr, Range) else curr
        return 0.0 if prev_rows == curr_rows else float("inf")

    if isinstance(prev, bool) or isinstance(curr, bool):
        return 0.0 if prev == curr else float("inf")
    if isinstance(prev, (int, float)) and isinstance(curr, (int, float)):
        pf = float(prev)
        cf = float(curr)
        if isfinite(pf) and isfinite(cf):
            return abs(cf - pf)
    try:
        eq = prev == curr
    except Exception:
        return float("inf")
    if isinstance(eq, bool):
        return 0.0 if eq else float("inf")
    return float("inf")


def flatten(*args: CellValue) -> Iterator[Scalar]:
    """Yield scalar values from scalars, nested lists, and lazy ranges.

    Excel errors raise `XlErrorException` on encounter, mirroring the
    evaluator's argument error precheck for generic worksheet functions.
    Lookup scans keep skip semantics through `Grid`/`Range.value_at` instead.
    """
    for arg in args:
        if isinstance(arg, Range):
            for v in arg.iter_raw():
                if isinstance(v, XlError):
                    raise XlErrorException(v)
                yield cast(Scalar, v)
        elif isinstance(arg, (list, tuple)):
            yield from flatten(*arg)
        else:
            if isinstance(arg, XlError):
                raise XlErrorException(arg)
            yield cast(Scalar, arg)
