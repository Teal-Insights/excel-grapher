"""Read generated input leaves from a populated workbook of this vintage.

Generated packages copy this module so `Model.from_workbook` can bind
input-direction cells through authored provenance without an Excel process.
Constants stay on the codegen snapshot in `data`.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any

from .tensor import Coordinate, Domain, Tensor


def split_cell_address(address: str) -> tuple[str, str]:
    """Split a sheet-qualified A1 address into `(sheet, coord)`.

    Args:
        address: A cell such as `Inputs!A1` or `'My Sheet'!B2`.

    Returns:
        The unquoted sheet title and the A1 coordinate.

    Raises:
        ValueError: `address` is not sheet-qualified.
    """
    if address.startswith("'"):
        index = 1
        while index < len(address):
            if address[index] == "'":
                if index + 1 < len(address) and address[index + 1] == "'":
                    index += 2
                    continue
                break
            index += 1
        sheet = address[1:index].replace("''", "'")
        rest = address[index + 1 :]
        if rest.startswith("!"):
            rest = rest[1:]
        if not rest:
            raise ValueError(f"expected sheet-qualified address, received {address!r}")
        return sheet, rest
    sheet, sep, coord = address.partition("!")
    if not sep or not coord:
        raise ValueError(f"expected sheet-qualified address, received {address!r}")
    return sheet, coord


def _read_cell(book: Any, address: str) -> object:
    sheet, coord = split_cell_address(address)
    if sheet not in book.sheetnames:
        raise KeyError(f"input cell {address}: sheet {sheet!r} is missing")
    return book[sheet][coord].value


def _axis_probe(axis: Any) -> Tensor[Any]:
    return Tensor(Domain.product(axis), tuple(axis.keys))


def _snapshot_cells(binding: object) -> Mapping[Coordinate, str]:
    """Materialize authored cells, binding runtime templates to snapshot axes."""
    cells = getattr(binding, "cells", None)
    if cells is None:
        raise TypeError(f"{binding!r} has no cells provenance")
    bind = getattr(cells, "bind", None)
    if bind is None:
        return cells
    domain = getattr(binding, "domain", None)
    if domain is None:
        raise TypeError(f"{binding!r} has no domain for snapshot cell bind")
    return bind(**{axis.name: _axis_probe(axis) for axis in domain.axes})


def _scalar_address(cells: object) -> str:
    if isinstance(cells, str):
        return cells
    if isinstance(cells, Mapping) and len(cells) == 1:
        return str(next(iter(cells.values())))
    raise TypeError(f"expected scalar cell provenance, received {cells!r}")


def _read_series(book: Any, name: str, data_module: object) -> object:
    default = getattr(data_module, f"{name.upper()}_DEFAULT")
    with_records = getattr(default, "with_records", None)
    if with_records is None:
        cells = getattr(data_module, f"{name.upper()}_CELLS")
        return _read_cell(book, _scalar_address(cells))
    records = [
        (coord, _read_cell(book, address)) for coord, address in _snapshot_cells(default).items()
    ]
    return with_records(records)


def read_bound_inputs(
    workbook: Path | str,
    names: Sequence[str],
    data_module: object,
) -> dict[str, object]:
    """Read named input-direction cells from `workbook`.

    Only `names` are bound. Constant series stay on `data_module`. Extra sheets
    and unbound cells are ignored. This is not Excel F9 parity: formula engine
    sheets are not re-read.

    Args:
        workbook: Path to a populated workbook of this export vintage.
        names: Input series ids to bind.
        data_module: Generated `data` module exposing `*_DEFAULT` and `*_CELLS`.

    Returns:
        Mapping from series id to a scalar or tensor for `Model.__init__`.

    Raises:
        KeyError: A bound input address names a sheet that is not in `workbook`.
    """
    from fastpyxl import load_workbook

    path = Path(workbook)
    book = load_workbook(path, data_only=True, keep_vba=path.suffix.lower() == ".xlsm")
    try:
        return {name: _read_series(book, name, data_module) for name in names}
    finally:
        book.close()
