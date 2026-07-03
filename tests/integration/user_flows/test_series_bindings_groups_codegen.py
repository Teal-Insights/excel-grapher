"""Integration: grouped series bindings drive export sequencing and manifest."""

from __future__ import annotations

import json
import re
from copy import deepcopy
from pathlib import Path
from typing import Any

import xlsxwriter

from excel_grapher.exporter import CodeGenerator
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import expand_data_range, validate_bindings_document


def _write_grouped_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number("A1", 1.0)
    ws.write_number("B1", 2.0)
    ws.write_number("C1", 3.0)
    wb.close()


GROUPED_BINDINGS: dict[str, Any] = {
    "schema_version": "1.6.0",
    "series": [
        {
            "id": "paris",
            "sheet": "Sheet1",
            "data_range": "Sheet1!C1",
            "layout": "scalar",
            "input": {"setter": {"name": "set_paris"}},
            "groups": [{"path": ["Climate scenarios", "Paris"], "order": 2}],
            "structure": {
                "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell", "read": "float"}},
                "dimensions": [],
            },
            "key": [],
        },
        {
            "id": "baseline",
            "sheet": "Sheet1",
            "data_range": "Sheet1!A1",
            "layout": "scalar",
            "input": {"setter": {"name": "set_baseline"}},
            "groups": [{"path": ["Baseline setup"], "order": 1}],
            "structure": {
                "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell", "read": "float"}},
                "dimensions": [],
            },
            "key": [],
        },
        {
            "id": "moderate",
            "sheet": "Sheet1",
            "data_range": "Sheet1!B1",
            "layout": "scalar",
            "input": {"setter": {"name": "set_moderate"}},
            "groups": [{"path": ["Climate scenarios", "Moderate"], "order": 1}],
            "structure": {
                "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell", "read": "float"}},
                "dimensions": [],
            },
            "key": [],
        },
    ],
}


def test_grouped_export_sequences_setters_and_emits_manifest(tmp_path: Path) -> None:
    workbook = tmp_path / "grouped.xlsx"
    _write_grouped_workbook(workbook)
    bindings = validate_bindings_document(deepcopy(GROUPED_BINDINGS))
    targets = expand_data_range("Sheet1!A1:C1", workbook=workbook)
    graph = create_dependency_graph(workbook, targets, load_values=True)

    with CodeGenerator(graph) as gen:
        modules = gen.generate_modules(
            targets,
            series_bindings=bindings,
            bindings_workbook=workbook,
        )

    assert "api_manifest.json" in modules
    manifest = json.loads(modules["api_manifest.json"])
    assert manifest["flat"]["setters"] == [
        "set_baseline",
        "set_moderate",
        "set_paris",
    ]

    api_py = modules["api.py"]
    setter_defs = [
        match.group(1) for match in re.finditer(r"^def (set_\w+)\(", api_py, flags=re.MULTILINE)
    ]
    assert setter_defs == ["set_baseline", "set_moderate", "set_paris"]

    init_py = modules["__init__.py"]
    all_match = re.search(r"__all__ = (\[.*?\])", init_py, flags=re.DOTALL)
    assert all_match is not None
    all_exports = eval(all_match.group(1))
    setter_exports = [name for name in all_exports if name.startswith("set_")]
    assert setter_exports == ["set_baseline", "set_moderate", "set_paris"]

    list_setters_match = re.search(
        r"def list_setters\(\).*?return (\[.*?\])",
        api_py,
        flags=re.DOTALL,
    )
    assert list_setters_match is not None
    assert eval(list_setters_match.group(1)) == [
        "set_baseline",
        "set_moderate",
        "set_paris",
    ]

    climate = manifest["group_tree"]["Climate scenarios"]
    assert climate["Moderate"]["setters"] == ["set_moderate"]
    assert climate["Paris"]["setters"] == ["set_paris"]
