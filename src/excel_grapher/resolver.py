from __future__ import annotations

import re
from typing import Protocol

import openpyxl

from .parser import CellRef


class NameResolver(Protocol):
    def resolve(self, name: str) -> tuple[str, str] | None:
        """
        Return (sheet, A1) for a defined name, or None if unknown/unsupported.
        """


class DictNameResolver:
    def __init__(self, mapping: dict[str, tuple[str, str]]):
        self._mapping = mapping

    def resolve(self, name: str) -> tuple[str, str] | None:
        return self._mapping.get(name)


def build_named_range_map(wb: openpyxl.Workbook) -> dict[str, tuple[str, str]]:
    """
    Map defined names to single-cell references.

    Only includes simple definitions like Sheet1!$A$1 (optionally quoted sheet name).
    Skips ranges and complex formulas.
    """
    out: dict[str, tuple[str, str]] = {}
    for name, defn in wb.defined_names.items():
        attr_text = getattr(defn, "attr_text", None)
        if not isinstance(attr_text, str) or not attr_text:
            continue
        if "," in attr_text:
            continue
        if ":" in attr_text:
            continue
        if "(" in attr_text:
            continue
        if attr_text.startswith("{") or attr_text.startswith("#") or attr_text.startswith('"'):
            continue
        m = re.match(r"'?([^'!]+)'?!\$?([A-Z]{1,3})\$?(\d+)$", attr_text)
        if not m:
            continue
        sheet_name = m.group(1)
        col = m.group(2)
        row = m.group(3)
        out[str(name)] = (sheet_name, f"{col}{row}")
    return out


def qualify_cell_ref(ref: CellRef, current_sheet: str) -> tuple[str, str]:
    sheet = ref.sheet if ref.sheet is not None else current_sheet
    return sheet, f"{ref.column}{ref.row}"

