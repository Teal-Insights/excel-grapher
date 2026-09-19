"""Derive dynamic-ref cell domains from series bindings.

Spike for issue #188. A binding manifest already says, per series, which cells
are `constant` (frozen template values) and which are `input` cells with an
optional `domain` / `value_map`. Those are exactly the facts a `constraints.py`
table restates cell by cell, so the manifest can be compiled into the same
`CellTypeEnv` that `DynamicRefConfig.from_constraints` builds.

Rules:

* `constant: {}` series pin every bound cell to its cached workbook value
  (the `FromWorkbook` semantics of `DynamicRefConfig.from_constraints_and_workbook`).
* `input` series contribute `input.domain` (`enum` / `between` / `real_between`)
  to every bound cell. When only `input.value_map` is declared, the cell domain
  is the set of workbook needles (the map values), not the public keys.
* Series without a domain contribute nothing.
"""

from __future__ import annotations

from collections.abc import Mapping
from pathlib import Path
from typing import Annotated, Any, Literal, cast

import fastpyxl
from fastpyxl.utils.cell import coordinate_to_tuple

from excel_grapher.core.cell_types import (
    Between,
    CellKind,
    CellType,
    CellTypeEnv,
    EnumDomain,
    RealBetween,
    constraints_to_cell_type_env,
    normalize_cell_type_env_key,
)
from excel_grapher.grapher.dynamic_refs import DynamicRefConfig, DynamicRefLimits
from excel_grapher.series_bindings.normalize import has_constant_direction, has_input_direction
from excel_grapher.series_bindings.ranges import expand_bound_series_addresses

# `Literal[...]` subscripted with a runtime tuple is fine at runtime but not for type checkers.
_RuntimeLiteral: Any = cast(Any, Literal)


def _input_domain_annotation(series: Mapping[str, Any]) -> Any | None:
    """Return the typing annotation equivalent to a series' `input.domain`."""
    input_block = series.get("input")
    if not isinstance(input_block, Mapping):
        return None
    domain = input_block.get("domain")
    if isinstance(domain, Mapping):
        if "enum" in domain:
            return _RuntimeLiteral[tuple(domain["enum"])]
        if "between" in domain:
            bounds = domain["between"]
            return Annotated[int, Between(bounds.get("min"), bounds.get("max"))]
        if "real_between" in domain:
            bounds = domain["real_between"]
            return Annotated[float, RealBetween(bounds.get("min"), bounds.get("max"))]
        return None
    value_map = input_block.get("value_map")
    if isinstance(value_map, Mapping) and value_map:
        return _RuntimeLiteral[tuple(value_map.values())]
    return None


def _kind_from_value(value: object) -> CellKind:
    if isinstance(value, bool):
        return CellKind.BOOL
    if isinstance(value, (int, float)):
        return CellKind.NUMBER
    if isinstance(value, str):
        return CellKind.STRING
    return CellKind.ANY


def _split_key(key: str) -> tuple[str, str]:
    sheet, coord = key.split("!", 1)
    return sheet, coord


def _constant_cell_types(addresses: list[str], *, workbook: Path | str) -> dict[str, CellType]:
    """Pin each address to a singleton domain from its cached workbook value."""
    if not addresses:
        return {}
    items = sorted(
        ((key, *_split_key(key)) for key in addresses),
        key=lambda item: (item[1], coordinate_to_tuple(item[2])),
    )
    env: dict[str, CellType] = {}
    wb = fastpyxl.load_workbook(Path(workbook), data_only=True, read_only=True)
    try:
        for key, sheet, coord in items:
            if sheet not in wb.sheetnames:
                raise ValueError(f"Sheet {sheet!r} (constant series cell {key!r}) not in workbook")
            value = wb[sheet][coord].value
            if value is None:
                continue
            env[key] = CellType(
                kind=_kind_from_value(value), enum=EnumDomain(values=frozenset({value}))
            )
    finally:
        wb.close()
    return env


def cell_type_env_from_bindings(
    bindings: Mapping[str, Any], *, workbook: Path | str
) -> dict[str, CellType]:
    """Compile a binding manifest into a `CellTypeEnv`.

    Args:
        bindings: Loaded (merged) series binding manifest.
        workbook: Workbook path; read for `constant` series cached values and
            for `data_range` expansion.

    Returns:
        Normalized-address to `CellType` mapping, keyed like
        `constraints_to_cell_type_env` output.
    """
    schema: dict[str, Any] = {}
    constant_addresses: list[str] = []
    for series in bindings.get("series", []):
        if not isinstance(series, dict):
            continue
        addresses = [
            normalize_cell_type_env_key(addr)
            for addr in expand_bound_series_addresses(series, workbook=workbook)
        ]
        if has_constant_direction(series):
            constant_addresses.extend(addresses)
            continue
        if not has_input_direction(series):
            continue
        annotation = _input_domain_annotation(series)
        if annotation is None:
            continue
        for addr in addresses:
            schema[addr] = annotation

    env = constraints_to_cell_type_env(schema, {})
    env.update(_constant_cell_types(constant_addresses, workbook=workbook))
    return env


def dynamic_refs_from_bindings(
    bindings: Mapping[str, Any],
    *,
    workbook: Path | str,
    limits: DynamicRefLimits | None = None,
) -> DynamicRefConfig:
    """Build a `DynamicRefConfig` whose env is derived from series bindings."""
    env: CellTypeEnv = cell_type_env_from_bindings(bindings, workbook=workbook)
    return DynamicRefConfig(cell_type_env=env, limits=limits or DynamicRefLimits())
