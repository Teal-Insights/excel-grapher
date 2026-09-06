from __future__ import annotations

from collections.abc import Iterable, Mapping
from dataclasses import dataclass
from enum import StrEnum
from typing import Any, TypeAlias, get_args, get_origin

from fastpyxl.utils.cell import coordinate_from_string


class CellKind(StrEnum):
    NUMBER = "number"
    STRING = "string"
    BOOL = "bool"
    DATE = "date"
    ERROR = "error"
    ANY = "any"


@dataclass(frozen=True, slots=True)
class IntervalDomain:
    """Closed integer interval domain for a cell (discrete steps, enumerable for dynamic refs)."""

    min: int | None = None
    max: int | None = None


# Backwards-compatible alias
IntIntervalDomain = IntervalDomain


@dataclass(frozen=True, slots=True)
class RealIntervalDomain:
    """Closed real-valued interval metadata; not enumerable for dynamic-ref branching."""

    min: float | None = None
    max: float | None = None


@dataclass(frozen=True, slots=True)
class EnumDomain:
    """Finite enum domain for a cell."""

    values: frozenset[Any]


@dataclass(frozen=True, slots=True)
class GreaterThanCell:
    """Metadata marker: the annotated cell is always greater than another cell."""

    other: str


@dataclass(frozen=True, slots=True)
class NotEqualCell:
    """Metadata marker: the annotated cell is never equal to another cell."""

    other: str


CellRelation: TypeAlias = GreaterThanCell | NotEqualCell


@dataclass(frozen=True, slots=True)
class CellType:
    """Internal description of the allowed values for a single cell."""

    kind: CellKind
    interval: IntervalDomain | None = None
    real_interval: RealIntervalDomain | None = None
    enum: EnumDomain | None = None
    relations: tuple[CellRelation, ...] = ()


CellTypeEnv: TypeAlias = Mapping[str, CellType]


@dataclass(frozen=True, slots=True)
class Between:
    """Integer interval constraint for Annotated numeric types (discrete / enumerable)."""

    min: int | None = None
    max: int | None = None


@dataclass(frozen=True, slots=True)
class RealBetween:
    """Real-valued interval constraint for Annotated float types (not enumerable for dynamic refs)."""

    min: float | int | None = None
    max: float | int | None = None


def _cell_type_from_annotation(annotated_type: Any) -> CellType:
    """Build a `CellType` from a constraint annotation (Annotated / Literal / plain type)."""
    # Import here to avoid forcing Annotated / Literal into __all__ of core.
    from typing import Annotated, Literal

    base_type = annotated_type
    metadata: list[object] = []

    if get_origin(annotated_type) is Annotated:
        args = get_args(annotated_type)
        if not args:
            base_type = Any
        else:
            base_type = args[0]
            metadata = list(args[1:])

    int_domain, real_domain = _interval_domains_from_metadata(metadata)
    relations = _relations_from_metadata(metadata)

    origin = get_origin(base_type)
    enum_domain: EnumDomain | None = None
    if origin is Literal:
        literal_values = get_args(base_type)
        kind = _infer_kind_from_literal_values(literal_values)
        if int_domain is None and real_domain is None:
            enum_domain = EnumDomain(values=frozenset(literal_values))
    else:
        kind = _infer_kind_from_python_type(base_type)

    return CellType(
        kind=kind,
        interval=int_domain,
        real_interval=real_domain,
        enum=enum_domain,
        relations=relations,
    )


def constraints_to_cell_type_env(
    constraints_schema: Mapping[str, Any], constraints_data: Mapping[str, Any]
) -> dict[str, CellType]:
    r"""Derive a `CellTypeEnv` from a constraints schema and optional instance data.

    *constraints_schema* maps sheet-qualified addresses (e.g. `\"Sheet1!B1\"`) to
    type objects describing domains (`Annotated`, `Literal`, plain `int` / `str`, etc.).
    *constraints_data* may hold runtime values for validation elsewhere; this function
    only inspects type metadata.

    Env dict keys are `normalize_cell_type_env_key` of each schema key so they
    align with `format_key` addresses from the grapher after normalization.
    """
    env: dict[str, CellType] = {}
    for key, annotated_type in constraints_schema.items():
        env[normalize_cell_type_env_key(key)] = _cell_type_from_annotation(annotated_type)

    _ = constraints_data

    return env


def _as_real_bound(x: float | int | None) -> float | None:
    if x is None:
        return None
    return float(x)


def _interval_domains_from_metadata(
    metadata: list[object],
) -> tuple[IntervalDomain | None, RealIntervalDomain | None]:
    int_domain: IntervalDomain | None = None
    real_domain: RealIntervalDomain | None = None
    for meta in metadata:
        if isinstance(meta, Between):
            int_domain = IntervalDomain(min=meta.min, max=meta.max)
        elif isinstance(meta, RealBetween):
            real_domain = RealIntervalDomain(
                min=_as_real_bound(meta.min),
                max=_as_real_bound(meta.max),
            )
    return int_domain, real_domain


def _relations_from_metadata(metadata: list[object]) -> tuple[CellRelation, ...]:
    relations: list[CellRelation] = []
    for meta in metadata:
        if isinstance(meta, GreaterThanCell):
            relations.append(GreaterThanCell(normalize_cell_type_env_key(meta.other)))
        elif isinstance(meta, NotEqualCell):
            relations.append(NotEqualCell(normalize_cell_type_env_key(meta.other)))
    return tuple(relations)


def _infer_kind_from_literal_values(values: tuple[object, ...]) -> CellKind:
    # If all values share the same basic type, infer from that; otherwise fall back to ANY.
    if not values:
        return CellKind.ANY

    first_type = type(values[0])
    if all(isinstance(v, int) for v in values):
        return CellKind.NUMBER
    if all(isinstance(v, str) for v in values):
        return CellKind.STRING
    if all(isinstance(v, bool) for v in values):
        return CellKind.BOOL
    if all(isinstance(v, first_type) for v in values):
        # Treat other homogeneous literals (e.g. date objects) as ANY for now.
        return CellKind.ANY
    return CellKind.ANY


def _infer_kind_from_python_type(tp: Any) -> CellKind:
    if tp is int or tp is float:
        return CellKind.NUMBER
    if tp is bool:
        return CellKind.BOOL
    if tp is str:
        return CellKind.STRING
    # A richer implementation could handle dates, errors, etc.
    return CellKind.ANY


def normalize_cell_type_env_key(address: str) -> str:
    """Return the canonical key for `CellTypeEnv` / dynamic-ref constraint maps.

    Graph code uses `excel_grapher.grapher.parser.format_key`, which wraps
    sheet names in single quotes when Excel requires it. Constraint schema keys
    may use the same spelling. This
    function strips those delimiters and normalizes the cell coordinate (column
    letters uppercased) so env lookups match regardless of quoting or case.

    Not to be confused with `excel_grapher.core.address_keys.normalize_key`,
    which follows evaluator node-key quoting rules and can differ for sheets
    that contain spaces.
    """
    sheet_part, coord = address.split("!", 1)
    sheet = sheet_part.strip()
    if sheet.startswith("'") and sheet.endswith("'"):
        sheet = sheet[1:-1].replace("''", "'")

    col, row = coordinate_from_string(coord.strip().replace("$", ""))
    return f"{sheet}!{col.upper()}{row}"


def leaves_missing_cell_type_constraints(
    leaves: Iterable[str], cell_type_env: Mapping[str, CellType]
) -> set[str]:
    """Leaves whose normalized address has no entry in `cell_type_env`.

    Looks up normalized keys with `Mapping` membership so a large env is not
    copied into a `frozenset` on every INDEX/OFFSET formula.
    """
    return {addr for addr in leaves if normalize_cell_type_env_key(addr) not in cell_type_env}


_normalize_cell_address = normalize_cell_type_env_key
