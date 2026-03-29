from __future__ import annotations

from collections.abc import Iterable, Mapping
from dataclasses import dataclass
from enum import StrEnum
from typing import Any, TypeAlias, get_args, get_origin, get_type_hints

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


def constraints_to_cell_type_env(
    constraints_type: type[Any], constraints: Mapping[str, Any]
) -> dict[str, CellType]:
    """Derive a CellTypeEnv from a validated constraints object.

    The constraints_type is expected to be a TypedDict- or pydantic-style model
    whose annotations use Annotated / Literal to describe domains, and
    constraints is a validated instance (e.g. from TypeAdapter). The current
    implementation only inspects type metadata; it assumes the instance has
    already been validated elsewhere.

    Dict keys are :func:`normalize_cell_type_env_key` of each hint key so they
    align with ``format_key`` addresses from the grapher after normalization.
    """

    # Import here to avoid forcing Annotated / Literal into __all__ of core.
    from typing import Annotated, Literal

    hints = get_type_hints(constraints_type, include_extras=True)
    env: dict[str, CellType] = {}

    for key, annotated_type in hints.items():
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

        env[normalize_cell_type_env_key(key)] = CellType(
            kind=kind,
            interval=int_domain,
            real_interval=real_domain,
            enum=enum_domain,
            relations=relations,
        )

    # We currently ignore the concrete values in `constraints` and rely solely
    # on type metadata; this leaves room to validate presence/shape later.
    _ = constraints

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
    """Return the canonical key for :class:`CellTypeEnv` / dynamic-ref constraint maps.

    Graph code uses :func:`excel_grapher.grapher.parser.format_key`, which wraps
    sheet names in single quotes when Excel requires it. TypedDict and
    :func:`constraints_to_cell_type_env` may use the same spelling. This
    function strips those delimiters and normalizes the cell coordinate (column
    letters uppercased) so env lookups match regardless of quoting or case.

    Not to be confused with :func:`excel_grapher.evaluator.name_utils.normalize_address`,
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
    """Leaves whose normalized address has no entry in ``cell_type_env``."""
    env_keys = frozenset(cell_type_env.keys())
    return {addr for addr in leaves if normalize_cell_type_env_key(addr) not in env_keys}


_normalize_cell_address = normalize_cell_type_env_key

