from __future__ import annotations

from collections.abc import Callable, Iterable, Iterator
from dataclasses import dataclass
from enum import Enum
from typing import Any, TypeAlias

import numpy as np
import openpyxl.utils.cell


class XlError(str, Enum):
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
                col = openpyxl.utils.cell.get_column_letter(c)
                yield f"{self.sheet}!{col}{r}"

    def resolve(self, evaluate_fn: Callable[[str], CellValue]) -> np.ndarray:
        values: list[CellValue] = [evaluate_fn(addr) for addr in self.cell_addresses()]
        rows, cols = self.shape
        return np.array(values, dtype=object).reshape((rows, cols))


# Excel formulas can produce scalar values, references, and (in the generated-code path)
# 2D array values (e.g., OFFSET consumed by SUM). Arrays are modeled as object-dtype
# numpy ndarrays containing CellValue elements.
CellValue: TypeAlias = (
    float | int | str | bool | XlError | ExcelRange | np.ndarray | None
)


def to_native(value: Any) -> Any:
    if hasattr(value, "item"):
        return value.item()
    return value


def to_number(value: CellValue) -> float | XlError:
    if value is None:
        return 0.0
    if isinstance(value, XlError):
        return value
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        s = value.strip()
        if s == "":
            return 0.0
        try:
            return float(s)
        except ValueError:
            return XlError.VALUE
    if isinstance(value, ExcelRange):
        return XlError.VALUE
    return XlError.VALUE


def to_int(value: CellValue) -> int | XlError:
    """Coerce a CellValue to an integer using Excel-style numeric coercion.

    This is a helper for functions that conceptually operate on integer indices
    (e.g. CHOOSE/INDEX/MATCH), while still propagating Excel errors.
    """
    n = to_number(value)
    if isinstance(n, XlError):
        return n
    return int(n)


def _format_general_number(value: float) -> str:
    if value.is_integer():
        return str(int(value))
    return str(value)


def to_string(value: CellValue) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, XlError):
        return value.value
    if isinstance(value, (int, float)):
        return _format_general_number(float(value))
    if isinstance(value, ExcelRange):
        return XlError.VALUE.value
    return str(value)


def to_bool(value: CellValue) -> bool | XlError:
    if value is None:
        return False
    if isinstance(value, XlError):
        return value
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return float(value) != 0.0
    if isinstance(value, str):
        s = value.strip().upper()
        if s == "":
            return False
        if s == "TRUE":
            return True
        if s == "FALSE":
            return False
        return XlError.VALUE
    if isinstance(value, ExcelRange):
        return XlError.VALUE
    return XlError.VALUE


def excel_casefold(value: str) -> str:
    return value.casefold()


def flatten(*args: Any) -> Iterator[CellValue]:
    for arg in args:
        if isinstance(arg, np.ndarray):
            yield from (v for v in arg.flat)
            continue
        if isinstance(arg, (list, tuple)):
            yield from flatten(*arg)
            continue
        yield arg


def get_error(*args: Any) -> XlError | None:
    for v in flatten(*args):
        if isinstance(v, XlError):
            return v
    return None


def numeric_values(values: Iterable[CellValue]) -> tuple[list[float], XlError | None]:
    nums: list[float] = []
    for v in values:
        n = to_number(v)
        if isinstance(n, XlError):
            return ([], n)
        nums.append(float(n))
    return (nums, None)

