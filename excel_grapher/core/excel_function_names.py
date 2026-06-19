"""Helpers for Excel function naming in string-based formula utilities and codegen."""

from __future__ import annotations

# Excel on-disk compatibility prefixes; extend as needed (_XLPM., etc.).
EXCEL_FUNCTION_PREFIXES: tuple[str, ...] = ("_XLFN.", "_XLUDF.")


def normalize_excel_function_name(name: str) -> str:
    """Normalize a parsed Excel function name to its canonical built-in form.

    Strips known Excel compatibility prefixes so dispatch, codegen, and the
    evaluator registry use canonical names only (``XLOOKUP``, not
    ``_XLFN.XLOOKUP``).

    Args:
        name: Function name as it appears in a formula token (any casing).

    Returns:
        Canonical upper-case function name for dispatch and codegen.
    """
    upper = name.upper()
    for prefix in EXCEL_FUNCTION_PREFIXES:
        if upper.startswith(prefix):
            return upper.split(".", 1)[1]
    return upper


def excel_func_to_python_runtime_name(normalized_name: str) -> str:
    """Map a canonical Excel function name to the export-runtime Python callable."""
    result = normalized_name.upper().lower().replace(".", "_")
    return f"xl_{result}"


def excel_function_call_prefixes(function_name: str) -> tuple[str, ...]:
    """Return leading formula prefixes that call ``function_name`` at top level.

    Used by string-based formula helpers (e.g. ``split_top_level_function``).

    Args:
        function_name: Bare Excel function name (e.g. ``IFS``).

    Returns:
        Tuple of prefixes including ``FN(`` and compatibility-prefixed variants.
    """
    fn = function_name.upper()
    prefixes: list[str] = [f"{fn}("]
    for compat_prefix in EXCEL_FUNCTION_PREFIXES:
        prefixes.append(f"{compat_prefix}{fn}(")
    return tuple(prefixes)


__all__ = [
    "EXCEL_FUNCTION_PREFIXES",
    "excel_func_to_python_runtime_name",
    "excel_function_call_prefixes",
    "normalize_excel_function_name",
]
