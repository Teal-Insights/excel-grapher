"""Representation-agnostic Excel value types and error enum."""

from __future__ import annotations

from collections.abc import Callable, Iterator
from dataclasses import dataclass
from enum import StrEnum
from typing import TypeAlias

import fastpyxl.utils.cell

from excel_grapher.core.address_keys import format_key


class XlError(StrEnum):
    VALUE = "#VALUE!"
    REF = "#REF!"
    DIV = "#DIV/0!"
    NA = "#N/A"
    NAME = "#NAME?"
    NUM = "#NUM!"
    NULL = "#NULL!"

    @classmethod
    def from_text(cls, value: str) -> XlError | None:
        upper = value.strip().upper()
        for err in cls:
            if err.value == upper:
                return err
        return None


class XlErrorException(Exception):
    """Exception form of an Excel error code.

    The exported runtime raises Excel errors as exceptions; the evaluator keeps
    `XlError` sentinel values and never raises this type.
    """

    code: XlError

    def __init__(self, code: XlError) -> None:
        """Initialize the exception with an Excel error code."""
        if not isinstance(code, XlError):
            raise TypeError(f"Expected XlError, got {type(code).__name__}")
        self.code = code
        super().__init__(code.value)


@dataclass(frozen=True, slots=True)
class ExcelRange:
    """Rectangular worksheet reference geometry for evaluator and export."""

    sheet: str
    start_row: int
    start_col: int
    end_row: int
    end_col: int

    @property
    def shape(self) -> tuple[int, int]:
        """The reference shape as `(rows, columns)`."""
        return (self.end_row - self.start_row + 1, self.end_col - self.start_col + 1)

    def cell_addresses(self) -> Iterator[str]:
        """Yield row-major sheet-qualified addresses without evaluating cells."""
        for r in range(self.start_row, self.end_row + 1):
            for c in range(self.start_col, self.end_col + 1):
                col = fastpyxl.utils.cell.get_column_letter(c)
                yield format_key(self.sheet, f"{col}{r}")


def resolve_excel_range(
    rng: ExcelRange,
    evaluate_fn: Callable[[str], CellValue],
) -> NestedGrid:
    """Eagerly materialize `rng` to a nested list via `evaluate_fn`.

    Prefer lazy `Range` consumers when only a subset of cells is needed.
    """
    values: list[CellValue] = [evaluate_fn(addr) for addr in rng.cell_addresses()]
    rows, cols = rng.shape
    grid: list[list[CellValue]] = []
    index = 0
    for _row in range(rows):
        row_values: list[CellValue] = []
        for _col in range(cols):
            row_values.append(values[index])
            index += 1
        grid.append(row_values)
    return grid


# Scalar values and references. Lazy `Range` / nested-list grids from
# `excel_grapher.core.grid` are used as function operands in the evaluator;
# omitted here to avoid a circular import with `core.grid.ranges`. NumPy
# object ndarrays are confined to fast-path materialization buffers, not this
# alias.
CellValue: TypeAlias = float | int | str | bool | XlError | ExcelRange | None

# Row-major nested-list grid of evaluated cells (formula / operator results).
NestedGrid: TypeAlias = list[list[CellValue]]

# Evaluator cell results and multi-cell operands (excludes lazy `Range`).
FormulaValue: TypeAlias = CellValue | NestedGrid
