"""Shared logical worksheet function implementations."""

from __future__ import annotations

from collections.abc import Iterator
from typing import cast

from .coercions import to_bool
from .grid import Grid
from .types import CellValue, XlError

__all__ = ["logical_and", "logical_not", "logical_or"]


def _iter_logical_cells(arg: CellValue) -> Iterator[CellValue]:
    """Yield scalar cells from a range/array argument in row-major order."""
    grid = Grid.wrap(arg)
    if grid is not None:
        for cell in grid.iter_raw():
            yield cast(CellValue, cell)
        return
    yield arg


def _consume_logical_cells(
    args: tuple[CellValue, ...],
    *,
    short_circuit_true: bool,
) -> bool | XlError:
    """Walk logical cells across arguments with Excel AND/OR semantics."""
    found_logical = False
    for arg in args:
        for cell in _iter_logical_cells(arg):
            if cell is None:
                continue
            if isinstance(cell, XlError):
                return cell
            found_logical = True
            b = to_bool(cell)
            if isinstance(b, XlError):
                return b
            if short_circuit_true:
                if b:
                    return True
            elif not b:
                return False
    if not found_logical:
        return XlError.VALUE
    return not short_circuit_true


def logical_and(*args: CellValue) -> bool | XlError:
    """Return logical AND across scalar and range arguments."""
    return _consume_logical_cells(args, short_circuit_true=False)


def logical_or(*args: CellValue) -> bool | XlError:
    """Return logical OR across scalar and range arguments."""
    return _consume_logical_cells(args, short_circuit_true=True)


def logical_not(arg: CellValue) -> bool | XlError:
    """Return logical NOT of an argument."""
    b = to_bool(arg)
    if isinstance(b, XlError):
        return b
    return not b
