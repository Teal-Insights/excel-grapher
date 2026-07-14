"""Metadata for Excel functions used in expression evaluation and domain inference.

Argument roles describe how each argument is used so that domain inference can
decide which cell references need a numeric domain:
- value: the argument is evaluated; cell refs require a domain when used in
  OFFSET/INDEX row/column expressions.
- ref_only: the implementation uses only the reference (e.g. row/column of the
  cell), not the cell's value; no domain is required for that ref.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal, TypeAlias

ArgRole: TypeAlias = Literal["value", "ref_only"]

# Sufficient for Excel varargs (SUM, SUMPRODUCT, AND, …) until Phase 3 drops
# eager materialization for full-scan consumers.
_ALL_ARGS: frozenset[int] = frozenset(range(32))


@dataclass(frozen=True, slots=True)
class ExcelFunctionMeta:
    """Argument roles for a function supported by the expression evaluator."""

    name: str
    arg_roles: tuple[ArgRole, ...]


FUNCTION_META: dict[str, ExcelFunctionMeta] = {
    "ROW": ExcelFunctionMeta("ROW", ("ref_only",)),
    "COLUMN": ExcelFunctionMeta("COLUMN", ("ref_only",)),
    "SUM": ExcelFunctionMeta("SUM", ()),
    "MIN": ExcelFunctionMeta("MIN", ()),
    "MAX": ExcelFunctionMeta("MAX", ()),
    "ABS": ExcelFunctionMeta("ABS", ("value",)),
    "EXP": ExcelFunctionMeta("EXP", ("value",)),
    "IF": ExcelFunctionMeta("IF", ("value", "value", "value")),
    "CONCAT": ExcelFunctionMeta("CONCAT", ()),
}

# Full-scan reductions that still eager-materialize to object-dtype ndarrays
# (evaluator bridge until Phase 3 teaches flatten / aggregates over Grid).
# Selective consumers (lookups, INDEX geometry) are *not* listed — they receive
# lazy `Range` by default in FormulaEvaluator._resolve_function_arg.
NUMPY_ARRAY_ARG_INDICES: dict[str, frozenset[int]] = {
    "SUM": _ALL_ARGS,
    "AVERAGE": _ALL_ARGS,
    "MIN": _ALL_ARGS,
    "MAX": _ALL_ARGS,
    "COUNT": _ALL_ARGS,
    "COUNTA": _ALL_ARGS,
    "COUNTIF": frozenset({0}),
    "AVERAGEIF": frozenset({0, 2}),
    "STDEV": _ALL_ARGS,
    "LARGE": frozenset({0}),
    "NPV": frozenset(range(1, 32)),
    "RANK": frozenset({1}),
    "SUMPRODUCT": _ALL_ARGS,
    "AND": _ALL_ARGS,
    "OR": _ALL_ARGS,
}

# Deprecated after Phase 1: range args default to lazy `Range`. Kept empty so
# call sites remain valid until the helper is removed in Phase 4.
LAZY_RANGE_ARG_INDICES: dict[str, frozenset[int]] = {}


def numpy_array_arg_indices(function_name: str) -> frozenset[int]:
    """Return argument indices that keep numpy arrays for ``function_name``."""
    return NUMPY_ARRAY_ARG_INDICES.get(function_name.upper(), frozenset())


def lazy_range_arg_indices(function_name: str) -> frozenset[int]:
    """Return argument indices that keep lazy `Range` values for ``function_name``.

    After Phase 1 this is always empty: multi-cell ranges default to lazy
    `Range` unless listed in `NUMPY_ARRAY_ARG_INDICES`.
    """
    return LAZY_RANGE_ARG_INDICES.get(function_name.upper(), frozenset())


def is_ref_only_arg(function_name: str, arg_index: int) -> bool:
    """True if this argument position is ref_only (no domain required for cell refs)."""
    meta = FUNCTION_META.get(function_name.upper())
    if meta is None or arg_index >= len(meta.arg_roles):
        return False
    return meta.arg_roles[arg_index] == "ref_only"
