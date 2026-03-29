from __future__ import annotations

from collections.abc import Mapping
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
    """Closed numeric interval domain for a cell."""

    min: int | float | None = None
    max: int | float | None = None


# Backwards-compatible alias
IntIntervalDomain = IntervalDomain


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
    enum: EnumDomain | None = None
    relations: tuple[CellRelation, ...] = ()


CellTypeEnv: TypeAlias = Mapping[str, CellType]


@dataclass(frozen=True, slots=True)
class Between:
    """Metadata marker for numeric interval constraints in Annotated types."""

    min: int | float | None = None
    max: int | float | None = None


def constraints_to_cell_type_env(
    constraints_type: type[Any], constraints: Mapping[str, Any]
) -> dict[str, CellType]:
    """Derive a CellTypeEnv from a validated constraints object.

    The constraints_type is expected to be a TypedDict- or pydantic-style model
    whose annotations use Annotated / Literal to describe domains, and
    constraints is a validated instance (e.g. from TypeAdapter). The current
    implementation only inspects type metadata; it assumes the instance has
    already been validated elsewhere.
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

        domain = _domain_from_metadata(metadata)
        relations = _relations_from_metadata(metadata)

        origin = get_origin(base_type)
        if origin is Literal:
            literal_values = get_args(base_type)
            kind = _infer_kind_from_literal_values(literal_values)
            if domain is None:
                domain = EnumDomain(values=frozenset(literal_values))
        else:
            kind = _infer_kind_from_python_type(base_type)

        env[key] = CellType(
            kind=kind,
            interval=_interval_from_domain(domain),
            enum=_enum_from_domain(domain),
            relations=relations,
        )

    # We currently ignore the concrete values in `constraints` and rely solely
    # on type metadata; this leaves room to validate presence/shape later.
    _ = constraints

    return env


def _domain_from_metadata(metadata: list[object]) -> IntervalDomain | EnumDomain | None:
    for meta in metadata:
        if isinstance(meta, Between):
            return IntervalDomain(min=meta.min, max=meta.max)
    return None


def _relations_from_metadata(metadata: list[object]) -> tuple[CellRelation, ...]:
    relations: list[CellRelation] = []
    for meta in metadata:
        if isinstance(meta, GreaterThanCell):
            relations.append(GreaterThanCell(_normalize_cell_address(meta.other)))
        elif isinstance(meta, NotEqualCell):
            relations.append(NotEqualCell(_normalize_cell_address(meta.other)))
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


def _interval_from_domain(domain: IntervalDomain | EnumDomain | None) -> IntervalDomain | None:
    if isinstance(domain, IntervalDomain):
        return domain
    return None


def _enum_from_domain(domain: IntervalDomain | EnumDomain | None) -> EnumDomain | None:
    if isinstance(domain, EnumDomain):
        return domain
    return None


def _normalize_cell_address(addr: str) -> str:
    sheet_part, coord = addr.split("!", 1)
    sheet = sheet_part.strip()
    if sheet.startswith("'") and sheet.endswith("'"):
        sheet = sheet[1:-1].replace("''", "'")

    col, row = coordinate_from_string(coord.strip().replace("$", ""))
    return f"{sheet}!{col}{row}"

