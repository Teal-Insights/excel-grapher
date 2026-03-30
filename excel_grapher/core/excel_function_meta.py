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
    "IF": ExcelFunctionMeta("IF", ("value", "value", "value")),
    "CONCAT": ExcelFunctionMeta("CONCAT", ()),
}


def is_ref_only_arg(function_name: str, arg_index: int) -> bool:
    """True if this argument position is ref_only (no domain required for cell refs)."""
    meta = FUNCTION_META.get(function_name.upper())
    if meta is None or arg_index >= len(meta.arg_roles):
        return False
    return meta.arg_roles[arg_index] == "ref_only"
