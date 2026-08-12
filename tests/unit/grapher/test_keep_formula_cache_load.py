"""create_dependency_graph should load formulas + caches in one pass."""

from __future__ import annotations

from pathlib import Path
from typing import Any
from unittest.mock import patch

import fastpyxl
import xlsxwriter

from excel_grapher import create_dependency_graph


def _write_formula_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("S")
    ws.write_number("A1", 10)
    ws.write_formula("B1", "=A1*2", None, 20)
    wb.close()


def test_create_dependency_graph_loads_workbook_once_with_formula_cache(
    tmp_path: Path,
) -> None:
    path = tmp_path / "dual.xlsx"
    _write_formula_workbook(path)

    calls: list[dict[str, Any]] = []
    real_load = fastpyxl.load_workbook

    def wrapped(*args: Any, **kwargs: Any):
        calls.append(dict(kwargs))
        return real_load(*args, **kwargs)

    with patch("excel_grapher.grapher.builder.fastpyxl.load_workbook", side_effect=wrapped):
        graph = create_dependency_graph(path, ["S!B1"], load_values=True)

    assert len(calls) == 1
    assert calls[0].get("keep_formula_cache") is True
    assert calls[0].get("data_only") in (False, None)
    node = graph.get_node("S!B1")
    assert node is not None
    assert node.value == 20
    leaf = graph.get_node("S!A1")
    assert leaf is not None
    assert leaf.value == 10


def test_create_dependency_graph_skips_formula_cache_when_values_disabled(
    tmp_path: Path,
) -> None:
    path = tmp_path / "formulas_only.xlsx"
    _write_formula_workbook(path)

    calls: list[dict[str, Any]] = []
    real_load = fastpyxl.load_workbook

    def wrapped(*args: Any, **kwargs: Any):
        calls.append(dict(kwargs))
        return real_load(*args, **kwargs)

    with patch("excel_grapher.grapher.builder.fastpyxl.load_workbook", side_effect=wrapped):
        graph = create_dependency_graph(path, ["S!B1"], load_values=False)

    assert len(calls) == 1
    assert calls[0].get("keep_formula_cache") in (False, None)
    node = graph.get_node("S!B1")
    assert node is not None
    assert node.value is None
