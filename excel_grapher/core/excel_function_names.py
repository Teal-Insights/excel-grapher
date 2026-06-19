"""Normalize Excel function names from on-disk formula spellings.

Excel compatibility prefixes
----------------------------

``_xlfn.``
    Future-function prefix stored by Excel for newer built-ins (e.g.
    ``=_xlfn.XLOOKUP(...)``). Stripped **unconditionally** by
    ``normalize_excel_function_name`` so dispatch, codegen, and the evaluator
    registry use canonical names only (``XLOOKUP``, not ``_XLFN.XLOOKUP``).

``_xludf.``
    Prefix Excel uses for some built-ins after round-trip, or for real add-in
    UDFs. Stripped **only** for names in ``_XLUDF_STRIPPABLE_BUILTINS``;
    unknown ``_xludf.*`` tokens are left unchanged.

All code paths that resolve Excel function names must call
``normalize_excel_function_name`` (or parse via ``formula_ast``, which does).
Do not register ``_XLFN.*`` aliases in the evaluator or hand-strip prefixes
elsewhere.
"""

from __future__ import annotations

_XLFN_PREFIX = "_XLFN."
_XLUDF_PREFIX = "_XLUDF."

# Built-ins that Excel may store with a ``_xludf.`` prefix after round-trip.
# Unknown ``_xludf.*`` names are left unchanged (custom / add-in UDFs).
#
# To support a new built-in after round-trip with ``_xludf.``:
# 1. Add the canonical upper-case name here.
# 2. Implement the function in the shared runtime and register under the
#    canonical name only.
# 3. Extend regression fixtures via ``rewrite_prefixed_workbook`` if needed.
_XLUDF_STRIPPABLE_BUILTINS: frozenset[str] = frozenset(
    {
        "FALSE",
        "IFNA",
        "NUMBERVALUE",
        "TRUE",
        "XLOOKUP",
    }
)


# Public alias for fixture helpers and documentation.
XLUDF_STRIPPABLE_BUILTINS = _XLUDF_STRIPPABLE_BUILTINS


def excel_func_to_python_runtime_name(normalized_name: str) -> str:
    """Map a canonical Excel function name to the export-runtime Python callable."""
    result = normalized_name.upper().lower().replace(".", "_")
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
    "XLUDF_STRIPPABLE_BUILTINS",
    "excel_func_to_python_runtime_name",
    "excel_function_call_prefixes",
    "normalize_excel_function_name",
]
