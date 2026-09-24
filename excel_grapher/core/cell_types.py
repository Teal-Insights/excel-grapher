from __future__ import annotations

import types
from collections.abc import Iterable, Mapping
from dataclasses import dataclass
from enum import StrEnum
from typing import Any, TypeAlias, Union, get_args, get_origin

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


class CellTypeEnvDict(dict[str, CellType]):
    """Mutable `CellTypeEnv` that stores keys after `normalize_cell_type_env_key`.

    Graph code passes `format_key` addresses (sheet quotes when Excel requires
    them). `in` / `[]` / `get` accept that spelling or the unquoted env form;
    writes always persist the normalized key so `lookup_cell_type` and
    `env.get(normalized)` agree.

    `keys()`, `items()`, and iteration yield only the stored (normalized)
    spelling. `quoted in env` can be true while `quoted in set(env)` is false.
    """

    def __init__(
        self,
        data: Mapping[str, CellType] | Iterable[tuple[str, CellType]] | None = None,
    ) -> None:
        super().__init__()
        if data:
            self.update(data)

    def __setitem__(self, key: str, value: CellType) -> None:
        super().__setitem__(normalize_cell_type_env_key(key), value)

    def __getitem__(self, key: str) -> CellType:
        found = super().get(key)
        if found is not None:
            return found
        return super().__getitem__(normalize_cell_type_env_key(key))

    def __delitem__(self, key: str) -> None:
        try:
            super().__delitem__(key)
        except KeyError:
            super().__delitem__(normalize_cell_type_env_key(key))

    def __contains__(self, key: object) -> bool:
        if not isinstance(key, str):
            return False
        if super().__contains__(key):
            return True
        try:
            return super().__contains__(normalize_cell_type_env_key(key))
        except (IndexError, ValueError):
            return False

    def get(self, key: object, default: Any = None) -> Any:
        if not isinstance(key, str):
            return default
        if super().__contains__(key):
            return super().__getitem__(key)
        try:
            return super().get(normalize_cell_type_env_key(key), default)
        except (IndexError, ValueError):
            return default

    def update(self, *args: Any, **kwargs: Any) -> None:
        if args:
            other = args[0]
            pairs = other.items() if isinstance(other, Mapping) else other
            for key, value in pairs:
                self[key] = value
        for key, value in kwargs.items():
            self[key] = value


def lookup_cell_type(env: Mapping[str, CellType], address: str) -> CellType | None:
    """Return the env entry for `address` after `normalize_cell_type_env_key`."""
    return env.get(normalize_cell_type_env_key(address))


def canonicalize_cell_type_env_keys(env: dict[str, CellType]) -> None:
    """Rewrite `env` keys in place to `normalize_cell_type_env_key` form."""
    stale = [(key, value) for key, value in env.items() if key != normalize_cell_type_env_key(key)]
    for key, value in stale:
        del env[key]
        env[normalize_cell_type_env_key(key)] = value


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


def is_union_domain(cell_type: CellType) -> bool:
    """Return whether `cell_type` is an enum unioned with one interval.

    Membership is the enum or the interval. A singleton enum on a union is
    not an extract-only pin.

    Args:
        cell_type: Cell constraint to inspect.

    Returns:
        True when `enum` is set together with `interval` or `real_interval`.
    """
    has_enum = cell_type.enum is not None
    has_interval = cell_type.interval is not None
    has_real = cell_type.real_interval is not None
    return has_enum and (has_interval or has_real)


def _cell_type_from_annotation(annotated_type: Any) -> CellType:
    """Build a `CellType` from a constraint annotation (Annotated / Literal / plain type).

    `Literal[...] | Annotated[..., Between|RealBetween]` stores both the enum
    and the interval. The two interval kinds cannot share one cell.
    """
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

    origin = get_origin(base_type)
    if origin is Union or origin is types.UnionType:
        return _cell_type_from_union(get_args(base_type), metadata)

    int_domain, real_domain = _interval_domains_from_metadata(metadata)
    relations = _relations_from_metadata(metadata)

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


def _cell_type_from_union(arms: tuple[Any, ...], metadata: list[object]) -> CellType:
    """Merge union arms into one `CellType`.

    Raises:
        ValueError: An arm is unconstrained while another is not, or the arms
            carry two different intervals, including `between` with
            `real_between`.
    """
    parts = [_cell_type_from_annotation(arm) for arm in arms]
    open_arms = [_arm_is_open(part) for part in parts]
    if any(open_arms) and not all(open_arms):
        raise ValueError(
            "union domain mixes a constrained arm with an unconstrained type; "
            "declare Between or RealBetween on the numeric arm"
        )
    relations = list(_relations_from_metadata(metadata))
    for part in parts:
        relations.extend(part.relations)
    if all(open_arms):
        kinds = {part.kind for part in parts}
        kind = kinds.pop() if len(kinds) == 1 else CellKind.ANY
        return CellType(kind=kind, relations=tuple(relations))

    int_domain, real_domain = _interval_domains_from_metadata(metadata)
    enums: list[object] = []
    kinds: list[CellKind] = []
    for part in parts:
        kinds.append(part.kind)
        if part.enum is not None:
            enums.extend(part.enum.values)
        int_domain = _merge_interval(int_domain, part.interval, "between")
        real_domain = _merge_interval(real_domain, part.real_interval, "real_between")
    if int_domain is not None and real_domain is not None:
        raise ValueError("union domain cannot combine between and real_between constraints")
    kind = kinds[0] if len(set(kinds)) == 1 else CellKind.ANY
    return CellType(
        kind=kind,
        interval=int_domain,
        real_interval=real_domain,
        enum=EnumDomain(values=frozenset(enums)) if enums else None,
        relations=tuple(relations),
    )


def _arm_is_open(cell_type: CellType) -> bool:
    return cell_type.enum is None and cell_type.interval is None and cell_type.real_interval is None


def _merge_interval(current: Any, new: Any, label: str) -> Any:
    if new is None:
        return current
    if current is not None and current != new:
        raise ValueError(f"union domain cannot combine multiple {label} constraints")
    return new


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
