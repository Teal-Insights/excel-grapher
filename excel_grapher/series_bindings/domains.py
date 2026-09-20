"""Compile series bindings into the `CellTypeEnv` dynamic-ref inference consumes.

A binding manifest already says, per series, which `input.domain` (or
`input.value_map`) applies to every bound cell. Series-level `relations`
name a partner series; each declaring cell compiles to `GreaterThanCell` or
`NotEqualCell` naming the partner cell at the same key.

Rules:

* `input` series contribute `input.domain` (`enum` / `between` / `real_between`)
  to every bound cell. When only `input.value_map` is declared, the cell domain
  is the set of workbook needles (the map values), not the public keys.
* `relations` expand per bound cell. Missing partner cells at a key fail closed.
* `constant` series are not pinned from workbook values. Series without a
  domain or relations contribute nothing.
"""

from __future__ import annotations

from collections.abc import Mapping
from pathlib import Path
from typing import Annotated, Any, Literal, cast, get_args, get_origin

from excel_grapher.core.cell_types import (
    Between,
    CellRelation,
    CellType,
    RealBetween,
    constraints_to_cell_type_env,
)
from excel_grapher.series_bindings.relations import (
    RELATION_TYPES,
    SeriesRelationError,
    canonical_measure_dtype,
    iter_series_relations,
    missing_partner_cell_issues,
    raise_if_relation_errors,
    relation_cell_indexes,
    relation_declaration_issues,
    series_by_id,
)

# `Literal[...]` subscripted with a runtime tuple is fine at runtime but not for type checkers.
_RuntimeLiteral: Any = cast(Any, Literal)

__all__ = [
    "SeriesRelationError",
    "cell_type_env_from_bindings",
]


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


def _python_type_for_series(series: Mapping[str, Any]) -> type:
    dtype = canonical_measure_dtype(series)
    if dtype in {"int", "integer"}:
        return int
    if dtype in {"string", "str"}:
        return str
    if dtype == "bool":
        return bool
    return float


def _with_relation_metadata(annotation: Any, relations: tuple[CellRelation, ...]) -> Any:
    if not relations:
        return annotation
    if get_origin(annotation) is Annotated:
        args = get_args(annotation)
        return Annotated[args[0], *args[1:], *relations]
    return Annotated[annotation, *relations]


def cell_type_env_from_bindings(
    bindings: Mapping[str, Any], *, workbook: Path | str
) -> dict[str, CellType]:
    """Compile a binding manifest into a `CellTypeEnv`.

    Args:
        bindings: Loaded (merged) series binding manifest.
        workbook: Workbook path; read for `data_range` expansion and key binds.

    Returns:
        Normalized-address to `CellType` mapping, keyed like
        `constraints_to_cell_type_env` output. Constant series are omitted
        unless they declare `relations`.

    Raises:
        SeriesRelationError: A relation partner is missing, incomparable,
            cyclic, reflexive, or has no cell at the declaring key.
    """
    raise_if_relation_errors(relation_declaration_issues(bindings))

    indexed = series_by_id(bindings)
    needed: set[str] = set()
    for series_id, series in indexed.items():
        relations = iter_series_relations(series)
        if _input_domain_annotation(series) is not None or relations:
            needed.add(series_id)
            needed.update(partner_id for _kind, partner_id in relations)

    indexes, issues = relation_cell_indexes(bindings, workbook=workbook, series_ids=needed)
    issues.extend(missing_partner_cell_issues(bindings, indexes))
    raise_if_relation_errors(issues)

    schema: dict[str, Any] = {}
    for series_id, series in indexed.items():
        base = _input_domain_annotation(series)
        relations = iter_series_relations(series)
        if base is None and not relations:
            continue
        declaring_index = indexes[series_id]
        for frozen, address in declaring_index.items():
            relation_meta: list[CellRelation] = []
            for kind, partner_id in relations:
                partner_addr = indexes[partner_id][frozen]
                relation_meta.append(RELATION_TYPES[kind](partner_addr))
            annotation = base if base is not None else _python_type_for_series(series)
            schema[address] = _with_relation_metadata(annotation, tuple(relation_meta))

    return constraints_to_cell_type_env(schema, {})
