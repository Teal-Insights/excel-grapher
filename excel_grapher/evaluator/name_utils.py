"""Utilities for converting Excel names to Python identifiers.

Address parsing/formatting/normalization live in
`excel_grapher.core.address_keys` as the canonical implementation.
"""

from __future__ import annotations

from excel_grapher.core.excel_function_names import EXCEL_FUNCTION_PREFIXES
from excel_grapher.core.excel_function_names import (
    normalize_excel_function_name as _normalize_excel_function_name,
)
from excel_grapher.evaluator.functions import FUNCTIONS

_REGISTERED_BUILTINS = frozenset(FUNCTIONS)


def normalize_excel_function_name(name: str) -> str:
    """Normalize using the evaluator built-in registry for `_XLUDF.` stripping."""
    return _normalize_excel_function_name(name, registered_builtins=_REGISTERED_BUILTINS)


__all__ = [
    "EXCEL_FUNCTION_PREFIXES",
    "normalize_excel_function_name",
]
