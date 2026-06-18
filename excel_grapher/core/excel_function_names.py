"""Normalize Excel function names from on-disk formula spellings."""

from __future__ import annotations

_XLFN_PREFIX = "_XLFN."
_XLUDF_PREFIX = "_XLUDF."

# Built-ins that Excel may store with a ``_xludf.`` prefix after round-trip.
# Unknown ``_xludf.*`` names are left unchanged (custom / add-in UDFs).
_XLUDF_STRIPPABLE_BUILTINS: frozenset[str] = frozenset(
    {
        "FALSE",
        "IFNA",
        "NUMBERVALUE",
        "TRUE",
        "XLOOKUP",
    }
)

# Export-runtime Python names that differ from ``xl_{normalized_lower}``.
_RUNTIME_PYTHON_NAMES: dict[str, str] = {
    "XLOOKUP": "xl__xlfn_xlookup",
}


def excel_func_to_python_runtime_name(normalized_name: str) -> str:
    """Map a canonical Excel function name to the export-runtime Python callable."""
    upper = normalized_name.upper()
    if upper in _RUNTIME_PYTHON_NAMES:
        return _RUNTIME_PYTHON_NAMES[upper]
    result = upper.lower().replace(".", "_")
    return f"xl_{result}"


def normalize_excel_function_name(name: str) -> str:
    """Normalize a parsed Excel function name to its canonical built-in form.

    Strips the documented ``_xlfn.`` future-function prefix unconditionally.
    Strips ``_xludf.`` only for known built-ins that excel-grapher implements.

    Args:
        name: Function name as it appears in a formula token (any casing).

    Returns:
        Canonical upper-case function name for dispatch and codegen.
    """
    upper = name.upper()
    if upper.startswith(_XLFN_PREFIX):
        return upper[len(_XLFN_PREFIX) :]
    if upper.startswith(_XLUDF_PREFIX):
        suffix = upper[len(_XLUDF_PREFIX) :]
        if suffix in _XLUDF_STRIPPABLE_BUILTINS:
            return suffix
    return upper


def excel_function_call_prefixes(function_name: str) -> tuple[str, ...]:
    """Return leading formula prefixes that call ``function_name`` at top level.

    Used by string-based formula helpers (e.g. ``split_top_level_function``).

    Args:
        function_name: Bare Excel function name (e.g. ``IFS``).

    Returns:
        Tuple of prefixes including ``FN(``, ``_XLFN.FN(``, and ``_XLUDF.FN(``
        when applicable.
    """
    fn = function_name.upper()
    prefixes: list[str] = [f"{fn}("]
    prefixes.append(f"{_XLFN_PREFIX}{fn}(")
    if fn in _XLUDF_STRIPPABLE_BUILTINS:
        prefixes.append(f"{_XLUDF_PREFIX}{fn}(")
    return tuple(prefixes)


__all__ = [
    "excel_func_to_python_runtime_name",
    "excel_function_call_prefixes",
    "normalize_excel_function_name",
]
