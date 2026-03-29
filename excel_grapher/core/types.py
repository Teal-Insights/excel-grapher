"""Representation-agnostic Excel value types and error enum."""

from __future__ import annotations

from collections.abc import Callable, Iterator
from dataclasses import dataclass
from enum import StrEnum
from typing import TypeAlias

import fastpyxl.utils.cell
import numpy as np


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
                yield f"{self.sheet}!{col}{r}"

    def resolve(self, evaluate_fn: Callable[[str], CellValue]) -> np.ndarray:
        values: list[CellValue] = [evaluate_fn(addr) for addr in self.cell_addresses()]
        rows, cols = self.shape
        return np.array(values, dtype=object).reshape((rows, cols))


# Scalar values, references, and object-dtype ndarrays of CellValue (e.g. OFFSET result).
CellValue: TypeAlias = (
    float | int | str | bool | XlError | ExcelRange | np.ndarray | None
)
