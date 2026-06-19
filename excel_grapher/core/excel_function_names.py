"""Helpers for Excel function naming in string-based formula utilities and codegen."""

from __future__ import annotations

# Excel on-disk compatibility prefixes; extend as needed (_XLPM., etc.).
EXCEL_FUNCTION_PREFIXES: tuple[str, ...] = ("_XLFN.", "_XLUDF.")
_XLFN_PREFIX = "_XLFN."
_XLUDF_PREFIX = "_XLUDF."


def normalize_excel_function_name(
    name: str,
    *,
    registered_builtins: frozenset[str] | None = None,
) -> str:
    """Normalize a parsed Excel function name to its canonical built-in form.

    ``_XLFN.`` is always stripped (Excel future-function namespace). ``_XLUDF.``
    is stripped only when the suffix names a built-in in ``registered_builtins``,
    so genuine add-in UDFs keep their prefixed spelling.

    Args:
        name: Function name as it appears in a formula token (any casing).
        registered_builtins: Upper-case built-in names that may drop ``_XLUDF.``.
            When omitted, ``_XLUDF.`` prefixes are preserved.

    Returns:
        Canonical upper-case function name for dispatch and codegen.
    """
    upper = name.upper()
    if upper.startswith(_XLFN_PREFIX):
        return upper.split(".", 1)[1]
    if upper.startswith(_XLUDF_PREFIX):
        suffix = upper.split(".", 1)[1]
        if registered_builtins is not None and suffix in registered_builtins:
            return suffix
        return upper
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
