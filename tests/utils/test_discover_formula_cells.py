"""discover_formula_cells_in_rows uses a single keep_formula_cache load."""

from __future__ import annotations

from pathlib import Path
from typing import Any
from unittest.mock import patch

import fastpyxl
import xlsxwriter

from tests.utils.discover_formula_cells import discover_formula_cells_in_rows


def test_discover_formula_cells_loads_workbook_once(tmp_path: Path) -> None:
    path = tmp_path / "scan.xlsx"
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("S")
    ws.write_number("A1", 1)
    ws.write_formula("B1", "=A1+1", None, 2)
    ws.write("C1", "text")
    wb.close()

    calls: list[dict[str, Any]] = []
    real_load = fastpyxl.load_workbook

    def wrapped(*args: Any, **kwargs: Any):
        calls.append(dict(kwargs))
        return real_load(*args, **kwargs)

    with patch("tests.utils.discover_formula_cells.fastpyxl.load_workbook", side_effect=wrapped):
        targets = discover_formula_cells_in_rows(path, "S", [1])

    assert targets == ["S!B1"]
    assert len(calls) == 1
    assert calls[0].get("keep_formula_cache") is True
