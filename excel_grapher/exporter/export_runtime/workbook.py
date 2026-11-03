"""Read generated input leaves from a populated workbook of this vintage.

Generated packages copy this module so `Model.from_workbook` can bind
input-direction cells through authored provenance without an Excel process.
Constants stay on the codegen snapshot in `data`.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any

from .tensor import (
    UNSET,
    AxisError,
    Coordinate,
    Domain,
    SchemaError,
    Series,
    Tensor,
    label_axis,
    relabel_series,
)


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


def _worksheet_cell(book: Any, address: str) -> Any:
    """Return the worksheet cell for a sheet-qualified address."""
    sheet, coord = split_cell_address(address)
    if sheet not in book.sheetnames:
        raise KeyError(f"input cell {address}: sheet {sheet!r} is missing")
    return book[sheet][coord]


def _read_cell(book: Any, address: str) -> object:
    return _worksheet_cell(book, address).value


def _is_formula_cell(cell: Any) -> bool:
    """Return True when `cell` still stores a formula rather than a literal."""
    if getattr(cell, "data_type", None) == "f":
        return True
    value = getattr(cell, "value", None)
    if isinstance(value, str):
        return value.startswith("=")
    return getattr(value, "text", None) is not None


def _read_override_value(cell: Any) -> object:
    """Return `UNSET` for a remaining formula and the literal otherwise."""
    if _is_formula_cell(cell):
        return UNSET
    return cell.value


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


def _coerce_label(value: object, key_type: type, address: str, axis_name: str) -> object:
    """Return a cached header as `key_type`, or fail when the cache is missing."""
    if value is None or (isinstance(value, str) and (not value.strip() or value.startswith("="))):
        raise SchemaError(f"axis {axis_name!r}: header cell {address} has no cached value")
    if key_type is int and isinstance(value, float) and value.is_integer():
        value = int(value)
    if type(value) is not key_type:
        raise SchemaError(
            f"axis {axis_name!r}: header cell {address} label {value!r} must be {key_type.__name__}"
        )
    return value


def _label_axis_from_cache(
    book: Any, axis_name: str, key_type: type[str] | type[int], addresses: tuple[str, ...]
) -> Any:
    labels = [
        _coerce_label(_read_cell(book, address), key_type, address, axis_name)
        for address in addresses
    ]
    try:
        return label_axis(axis_name, labels, key_type)
    except AxisError as exc:
        raise SchemaError(str(exc)) from exc


def _relabel_from_cache(book: Any, name: str, data_module: object, series: object) -> object:
    """Key `series` by cached formula headers, leaving literal headers in place."""
    if not isinstance(series, Series):
        return series
    spec = getattr(data_module, "LABEL_CELLS", {}).get(name)
    if not spec:
        return series
    axes = {axis.name: axis for axis in series.domain.axes}
    replacements = {}
    for axis_name, addresses in spec.items():
        axis = axes.get(axis_name)
        if axis is None:
            raise SchemaError(f"{name}: label cells name unknown axis {axis_name!r}")
        if len(addresses) != len(axis.keys):
            raise SchemaError(
                f"{name}: axis {axis_name!r} has {len(addresses)} header cells "
                f"for {len(axis.keys)} keys"
            )
        replacements[axis_name] = _label_axis_from_cache(
            book, axis.name, axis.key_type, tuple(addresses)
        )
    return relabel_series(series, **replacements)


def _read_series(book: Any, name: str, data_module: object) -> object:
    default = getattr(data_module, f"{name.upper()}_DEFAULT")
    with_records = getattr(default, "with_records", None)
    if with_records is None:
        cells = getattr(data_module, f"{name.upper()}_CELLS")
        return _read_cell(book, _scalar_address(cells))
    records = [
        (coord, _read_cell(book, address)) for coord, address in _snapshot_cells(default).items()
    ]
    return _relabel_from_cache(book, name, data_module, with_records(records))


def _read_override(book: Any, name: str, data_module: object, value_book: Any) -> object:
    """Read one override series from the cells that still hold formulas.

    A cell that still stores a formula contributes `UNSET`, so evaluation keeps
    the template formula instead of a stale calculated cache. A literal,
    including a blank, displaces that formula. When every cell is still a
    formula, the whole series is `UNSET`.
    """
    binding = getattr(data_module, name.upper(), None)
    with_records = getattr(binding, "with_records", None)
    if with_records is not None and getattr(binding, "cells", None) is not None:
        saw_literal = False
        records: list[tuple[Coordinate, object]] = []
        for coord, address in _snapshot_cells(binding).items():
            value = _read_override_value(_worksheet_cell(book, address))
            if value is UNSET:
                records.append((coord, UNSET))
            else:
                saw_literal = True
                records.append((coord, value))
        if not saw_literal:
            return UNSET
        return _relabel_from_cache(value_book, name, data_module, with_records(records))
    cells = getattr(data_module, f"{name.upper()}_CELLS")
    return _read_override_value(_worksheet_cell(book, _scalar_address(cells)))


def read_bound_inputs(
    workbook: Path | str,
    names: Sequence[str],
    data_module: object,
    override_names: Sequence[str] = (),
) -> dict[str, object]:
    """Read named input-direction cells from `workbook`.

    Only `names` are bound. Constant series stay on `data_module`. Extra sheets
    and unbound cells are ignored. Values are cached workbook cells. Generated
    constructors translate `input.value_map` needles to public keys before
    validation. This is not Excel F9 parity: formula engine sheets are not
    re-read.

    `override_names` are `input.mode: override` series. Those cells are read
    from the stored formula, not the cached calculated value. A remaining
    formula leaves that coordinate unset. A literal displaces the template
    formula.

    When a `column_header` or `row_label` cell is a formula, its cached value
    in `workbook` becomes that axis key. Provenance stays positional: the same
    column or row keeps the same observation. Literal headers keep the export
    snapshot keys. A missing cache, a repeated label, or a label of the wrong
    type fails closed. Formula engine sheets are not evaluated.

    Args:
        workbook: Path to a populated workbook of this export vintage.
        names: Input series ids to bind.
        data_module: Generated `data` module exposing `*_DEFAULT` and `*_CELLS`.
        override_names: Input ids whose cells may still be template formulas.

    Returns:
        Mapping from series id to a scalar, tensor, or `UNSET` for `Model.__init__`.

    Raises:
        KeyError: A bound input address names a sheet that is not in `workbook`.
    """
    from fastpyxl import load_workbook

    path = Path(workbook)
    keep_vba = path.suffix.lower() == ".xlsm"
    overrides = set(override_names)
    value_book = load_workbook(path, data_only=True, keep_vba=keep_vba)
    formula_book = None
    try:
        if overrides:
            formula_book = load_workbook(path, data_only=False, keep_vba=keep_vba)
        values: dict[str, object] = {}
        for name in names:
            if name in overrides:
                if formula_book is None:
                    raise RuntimeError(f"override input {name} requires the formula workbook")
                values[name] = _read_override(formula_book, name, data_module, value_book)
            else:
                values[name] = _read_series(value_book, name, data_module)
        return values
    finally:
        value_book.close()
        if formula_book is not None:
            formula_book.close()
