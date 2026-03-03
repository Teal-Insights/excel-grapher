from __future__ import annotations

import re
from dataclasses import dataclass
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


@dataclass(frozen=True)
class NamedRangeMaps:
    cell_map: dict[str, tuple[str, str]]
    range_map: dict[str, tuple[str, str, str]]


def build_named_range_map(wb: openpyxl.Workbook) -> NamedRangeMaps:
    """
    Map defined names to single-cell and range references.

    Only includes simple definitions like Sheet1!$A$1 or Sheet1!$A$1:$B$10
    (optionally quoted sheet name). Skips multi-area and complex formulas.
    """
    cell_map: dict[str, tuple[str, str]] = {}
    range_map: dict[str, tuple[str, str, str]] = {}
    for name, defn in wb.defined_names.items():
        attr_text = getattr(defn, "attr_text", None)
        if not isinstance(attr_text, str) or not attr_text:
            continue
        if "," in attr_text:
            continue
        if attr_text.startswith("{") or attr_text.startswith("#") or attr_text.startswith('"'):
            continue
        if ":" in attr_text:
            m = re.match(
                r"'?(?P<sheet>[^'!]+)'?!\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+):\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+)$",
                attr_text,
            )
            if not m:
                continue
            sheet_name = m.group("sheet")
            start = f"{m.group('c1')}{m.group('r1')}"
            end = f"{m.group('c2')}{m.group('r2')}"
            range_map[str(name)] = (sheet_name, start, end)
            continue

        m = re.match(r"'?([^'!]+)'?!\$?([A-Z]{1,3})\$?(\d+)$", attr_text)
        if not m:
            continue
        sheet_name = m.group(1)
        col = m.group(2)
        row = m.group(3)
        cell_map[str(name)] = (sheet_name, f"{col}{row}")
    return NamedRangeMaps(cell_map=cell_map, range_map=range_map)


def qualify_cell_ref(ref: CellRef, current_sheet: str) -> tuple[str, str]:
    sheet = ref.sheet if ref.sheet is not None else current_sheet
    return sheet, f"{ref.column}{ref.row}"

