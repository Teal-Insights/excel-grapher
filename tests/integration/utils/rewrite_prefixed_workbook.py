"""Rewrite allowlisted built-in calls to ``_xlfn.`` or ``_xludf.`` spellings for fixtures."""

from __future__ import annotations

import re
import shutil
from pathlib import Path
from typing import Literal

from fastpyxl import load_workbook

# Built-ins commonly rewritten in prefix-regression workbook fixtures.
_REWRITABLE_BUILTINS: frozenset[str] = frozenset(
    {
        "FALSE",
        "IFNA",
        "NUMBERVALUE",
        "TRUE",
        "XLOOKUP",
    }
)

ExcelBuiltinPrefix = Literal["_xlfn", "_xludf"]

# Longest names first so shorter prefixes do not shadow longer function names.
_REWRITABLE_ORDER: tuple[str, ...] = tuple(sorted(_REWRITABLE_BUILTINS, key=len, reverse=True))


def _call_site_pattern(name: str) -> re.Pattern[str]:
    return re.compile(
        rf"([=(,])\s*(?:_xlfn\.|_xludf\.)?{re.escape(name)}\s*\(",
        re.IGNORECASE,
    )


def rewrite_formula_builtin_prefix(formula: str, prefix: ExcelBuiltinPrefix) -> str:
    """Return ``formula`` with allowlisted built-ins spelled as ``{prefix}.NAME``."""
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula
    result = formula
    for name in _REWRITABLE_ORDER:
        result = _call_site_pattern(name).sub(rf"\1{prefix}.{name}(", result)
    return result


def rewrite_formula_to_xludf(formula: str) -> str:
    """Return ``formula`` with allowlisted built-ins spelled as ``_xludf.NAME``."""
    return rewrite_formula_builtin_prefix(formula, "_xludf")


def rewrite_formula_to_xlfn(formula: str) -> str:
    """Return ``formula`` with allowlisted built-ins spelled as ``_xlfn.NAME``."""
    return rewrite_formula_builtin_prefix(formula, "_xlfn")


def write_prefixed_workbook_copy(
    source: Path,
    destination: Path,
    *,
    prefix: ExcelBuiltinPrefix,
    workbook_name: str | None = None,
) -> Path:
    """Copy ``source`` to ``destination`` with compatibility-prefix formula spellings."""
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination)
    wb = load_workbook(destination)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    cell.value = rewrite_formula_builtin_prefix(cell.value, prefix)
    wb.save(destination)
    if workbook_name is not None:
        bindings_dir = destination.with_suffix(".bindings")
        if bindings_dir.is_dir():
            for shard in bindings_dir.glob("*.bindings.yaml"):
                text = shard.read_text(encoding="utf-8")
                shard.write_text(
                    text.replace(
                        f"workbook: {source.name}",
                        f"workbook: {workbook_name}",
                    ),
                    encoding="utf-8",
                )
    return destination


def write_xludf_workbook_copy(
    source: Path,
    destination: Path,
    *,
    workbook_name: str | None = None,
) -> Path:
    """Copy ``source`` to ``destination`` with ``_xludf.`` formula spellings."""
    return write_prefixed_workbook_copy(
        source,
        destination,
        prefix="_xludf",
        workbook_name=workbook_name,
    )


def write_xlfn_workbook_copy(
    source: Path,
    destination: Path,
    *,
    workbook_name: str | None = None,
) -> Path:
    """Copy ``source`` to ``destination`` with ``_xlfn.`` formula spellings."""
    return write_prefixed_workbook_copy(
        source,
        destination,
        prefix="_xlfn",
        workbook_name=workbook_name,
    )


__all__ = [
    "ExcelBuiltinPrefix",
    "rewrite_formula_builtin_prefix",
    "rewrite_formula_to_xlfn",
    "rewrite_formula_to_xludf",
    "write_prefixed_workbook_copy",
    "write_xlfn_workbook_copy",
    "write_xludf_workbook_copy",
]
