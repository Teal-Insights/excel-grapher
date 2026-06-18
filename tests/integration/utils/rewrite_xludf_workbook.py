"""Rewrite allowlisted built-in calls to ``_xludf.`` spelling for regression fixtures."""

from __future__ import annotations

import re
import shutil
from pathlib import Path

from fastpyxl import load_workbook

# Longest names first so shorter prefixes do not shadow longer function names.
_REWRITABLE_ORDER: tuple[str, ...] = ("NUMBERVALUE", "XLOOKUP", "IFNA")


def _call_site_pattern(name: str) -> re.Pattern[str]:
    return re.compile(
        rf"([=(,])\s*(?:_xlfn\.)?{re.escape(name)}\s*\(",
        re.IGNORECASE,
    )


def rewrite_formula_to_xludf(formula: str) -> str:
    """Return ``formula`` with allowlisted built-ins spelled as ``_xludf.NAME``."""
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula
    result = formula
    for name in _REWRITABLE_ORDER:
        result = _call_site_pattern(name).sub(rf"\1_xludf.{name}(", result)
    return result


def write_xludf_workbook_copy(
    source: Path,
    destination: Path,
    *,
    workbook_name: str | None = None,
) -> Path:
    """Copy ``source`` to ``destination`` with ``_xludf.`` formula spellings."""
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination)
    wb = load_workbook(destination)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    cell.value = rewrite_formula_to_xludf(cell.value)
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


__all__ = ["rewrite_formula_to_xludf", "write_xludf_workbook_copy"]
