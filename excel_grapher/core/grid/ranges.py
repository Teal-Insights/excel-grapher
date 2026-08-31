"""Lazy rectangular ranges shared by evaluator and export runtimes."""

from __future__ import annotations

from collections.abc import Callable, Iterator
from dataclasses import dataclass, field

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import format_cell_key
from excel_grapher.core.types import FormulaValue, XlError, XlErrorException

__all__ = ["Range"]


@dataclass(frozen=True, slots=True)
class Range:
    """Rectangular lazy range with consumer-driven cell access.

    Coordinates passed to `cell`, `row`, `column`, and `view` are 1-based and
    relative to this range, matching Excel function arguments.
    """

    sheet: str
    start_row: int
    start_col: int
    end_row: int
    end_col: int
    # Resolvers may come from evaluation contexts with their own value
    # vocabulary; values are validated/coerced at consumption time.
    _resolver: Callable[[str], FormulaValue] = field(repr=False, compare=False)
    # Optional coordinate reader: `(row, col)` absolute 1-based. When set,
    # `cell` / `value_at` use it and do not construct NodeKey strings.
    _coord_resolver: Callable[[int, int], FormulaValue] | None = field(
        default=None, repr=False, compare=False
    )

    def __post_init__(self) -> None:
        """Validate the rectangular bounds."""
        if self.start_row < 1 or self.start_col < 1:
            raise ValueError("Range coordinates must be positive")
        if self.end_row < self.start_row or self.end_col < self.start_col:
            raise ValueError("Range end must be greater than or equal to start")

    @property
    def shape(self) -> tuple[int, int]:
        """The range shape as `(rows, columns)`."""
        return (self.end_row - self.start_row + 1, self.end_col - self.start_col + 1)

    def cell_addresses(self) -> Iterator[str]:
        """Yield row-major addresses without evaluating cells."""
        for row in range(self.start_row, self.end_row + 1):
            for col in range(self.start_col, self.end_col + 1):
                yield self._address(row, col)

    def cell(self, row: int, col: int) -> FormulaValue:
        """Return a single relative cell value without evaluating siblings.

        Args:
            row: 1-based row within the range.
            col: 1-based column within the range.

        Raises:
            IndexError: If `row` or `col` is outside the range.
            XlErrorException: If the resolved cell is an Excel error.
        """
        self._validate_relative_cell(row, col)
        value = self._resolve_at(self.start_row + row - 1, self.start_col + col - 1)
        return self._raise_if_error(value)

    def row(self, row: int) -> Range:
        """Return a lazy view for one relative row."""
        nrows, _ = self.shape
        if row < 1 or row > nrows:
            raise IndexError("Range row is out of bounds")
        absolute_row = self.start_row + row - 1
        return Range(
            self.sheet,
            absolute_row,
            self.start_col,
            absolute_row,
            self.end_col,
            self._resolver,
            _coord_resolver=self._coord_resolver,
        )

    def column(self, col: int) -> Range:
        """Return a lazy view for one relative column."""
        _, ncols = self.shape
        if col < 1 or col > ncols:
            raise IndexError("Range column is out of bounds")
        absolute_col = self.start_col + col - 1
        return Range(
            self.sheet,
            self.start_row,
            absolute_col,
            self.end_row,
            absolute_col,
            self._resolver,
            _coord_resolver=self._coord_resolver,
        )

    def view(
        self,
        row_start: int = 1,
        row_end: int | None = None,
        col_start: int = 1,
        col_end: int | None = None,
    ) -> Range:
        """Return a lazy rectangular subrange view using relative coordinates."""
        nrows, ncols = self.shape
        row_end = nrows if row_end is None else row_end
        col_end = ncols if col_end is None else col_end
        self._validate_relative_cell(row_start, col_start)
        self._validate_relative_cell(row_end, col_end)
        if row_end < row_start or col_end < col_start:
            raise ValueError("Range view end must be greater than or equal to start")
        return Range(
            self.sheet,
            self.start_row + row_start - 1,
            self.start_col + col_start - 1,
            self.start_row + row_end - 1,
            self.start_col + col_end - 1,
            self._resolver,
            _coord_resolver=self._coord_resolver,
        )

    def value_at(self, row: int, col: int) -> FormulaValue:
        """Return a single relative cell value with errors as sentinels.

        Unlike `cell`, Excel errors surface as `XlError` sentinel values (raised
        `XlErrorException`s from the resolver are caught and converted). Range
        consumers that implement Excel skip semantics (lookup scans, criteria
        matching) use this accessor; `cell`/iteration raise instead.
        """
        self._validate_relative_cell(row, col)
        try:
            return self._resolve_at(self.start_row + row - 1, self.start_col + col - 1)
        except XlErrorException as exc:
            return exc.code

    def iter_raw(self) -> Iterator[FormulaValue]:
        """Yield raw values (error sentinels included) in row-major order."""
        nrows, ncols = self.shape
        for row in range(1, nrows + 1):
            for col in range(1, ncols + 1):
                yield self.value_at(row, col)

    def rows_raw(self) -> list[list[FormulaValue]]:
        """Materialize the range as nested row lists of raw values."""
        nrows, ncols = self.shape
        return [[self.value_at(r, c) for c in range(1, ncols + 1)] for r in range(1, nrows + 1)]

    def iter_values(self) -> Iterator[FormulaValue]:
        """Yield values in deterministic row-major order."""
        nrows, ncols = self.shape
        for row in range(1, nrows + 1):
            for col in range(1, ncols + 1):
                yield self.cell(row, col)

    def __iter__(self) -> Iterator[FormulaValue]:
        """Yield values in deterministic row-major order."""
        return self.iter_values()

    def _resolve_at(self, row: int, col: int) -> FormulaValue:
        if self._coord_resolver is not None:
            return self._coord_resolver(row, col)
        return self._resolver(self._address(row, col))

    def _address(self, row: int, col: int) -> str:
        col_letter = fastpyxl.utils.cell.get_column_letter(col)
        return format_cell_key(self.sheet, col_letter, row)

    def _validate_relative_cell(self, row: int, col: int) -> None:
        nrows, ncols = self.shape
        if row < 1 or row > nrows or col < 1 or col > ncols:
            raise IndexError("Range cell is out of bounds")

    @staticmethod
    def _raise_if_error(value: FormulaValue) -> FormulaValue:
        if isinstance(value, XlError):
            raise XlErrorException(value)
        return value
