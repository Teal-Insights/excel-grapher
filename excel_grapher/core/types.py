"""Representation-agnostic Excel value types and error enum."""

from __future__ import annotations

from collections.abc import Callable, Iterator
from dataclasses import dataclass
from enum import StrEnum
from typing import TypeAlias

import fastpyxl.utils.cell
import numpy as np

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
    sheet: str
    start_row: int
    start_col: int
    end_row: int
    end_col: int

    @property
    def shape(self) -> tuple[int, int]:
        return (self.end_row - self.start_row + 1, self.end_col - self.start_col + 1)

    def cell_addresses(self) -> Iterator[str]:
        for r in range(self.start_row, self.end_row + 1):
            for c in range(self.start_col, self.end_col + 1):
                col = fastpyxl.utils.cell.get_column_letter(c)
                yield format_key(self.sheet, f"{col}{r}")

    def resolve(self, evaluate_fn: Callable[[str], CellValue]) -> np.ndarray:
        values: list[CellValue] = [evaluate_fn(addr) for addr in self.cell_addresses()]
        rows, cols = self.shape
        return np.array(values, dtype=object).reshape((rows, cols))


# Scalar values, references, and object-dtype ndarrays of CellValue (e.g. OFFSET /
# SUMPRODUCT). Lazy `Range` values from `excel_grapher.core.grid` are also used as
# function operands in the evaluator; they are not listed here to avoid a circular
# import with `core.grid.ranges`.
CellValue: TypeAlias = float | int | str | bool | XlError | ExcelRange | np.ndarray | None
