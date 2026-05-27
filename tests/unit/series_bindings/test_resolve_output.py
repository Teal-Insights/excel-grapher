"""Unit tests for output-direction series binding resolution."""

from __future__ import annotations

from pathlib import Path

import xlsxwriter

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import (
    expand_data_range,
    load_series_bindings,
    resolve_series_binding,
)

FIXTURES = Path(__file__).resolve().parents[2] / "fixtures" / "series_bindings"


def _write_formula_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    ws.write_number("B2", 10.0)
    ws.write_formula("C2", "=B2*2")
    wb.close()


def test_output_resolve_includes_formula_cells_not_only_leaves(tmp_path: Path) -> None:
    wb_path = tmp_path / "formula.xlsx"
    _write_formula_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Sheet1!C2"], load_values=True)
    assert "Sheet1!C2" in graph
    node = graph.get_node("Sheet1!C2")
    assert node is not None
    assert not node.is_leaf

    series = {
        "id": "scaled_output",
        "sheet": "Sheet1",
        "data_range": "Sheet1!C2",
        "layout": "scalar",
        "output": {"compute": {"name": "compute_scaled_output"}},
        "structure": {
            "measure": {"concept": "OBS_VALUE", "bind": {"kind": "data_cell"}},
            "dimensions": [
                {
                    "concept": "LABEL",
                    "role": "key",
                    "scope": "series",
                    "bind": {"kind": "constant", "value": "scaled"},
                }
            ],
        },
        "key": ["LABEL"],
    }

    resolved = resolve_series_binding(graph, wb_path, series, direction="output")
    assert resolved["ok"] is True
    assert len(resolved["leaves"]) == 1
    assert resolved["leaves"][0]["address"] == "Sheet1!C2"
    assert resolved["leaves"][0]["record"]["LABEL"] == "scaled"
    assert "OBS_VALUE" in resolved["leaves"][0]["record"]


def test_output_record_includes_non_key_dimensions(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    from tests.unit.series_bindings.test_resolve import _write_borvelia_workbook

    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(
        wb_path,
        expand_data_range("Inputs!F5:J5"),
        load_values=True,
    )
    bindings = load_series_bindings(
        FIXTURES / "shard_borvelia_output.yaml",
    )
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series, direction="output")
    leaf = next(item for item in resolved["leaves"] if item["key"]["TIME_PERIOD"] == 3)
    assert leaf["record"]["FREQUENCY"] == "A"
    assert leaf["record"]["TIME_PERIOD"] == 3
    assert leaf["record"]["REF_AREA"] == "Borvelia"


def test_output_partial_overlap_warns(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    from tests.unit.series_bindings.test_resolve import _write_borvelia_workbook

    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Inputs!H5"], load_values=True)
    bindings = load_series_bindings(FIXTURES / "shard_borvelia_output.yaml")
    series = dict(bindings["series"][0])
    resolved = resolve_series_binding(graph, wb_path, series, direction="output")
    assert len(resolved["leaves"]) == 1
    assert any(i["code"] == "partial_graph_overlap" for i in resolved["issues"])


def test_output_partial_overlap_warning_can_be_disabled(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    from tests.unit.series_bindings.test_resolve import _write_borvelia_workbook

    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, ["Inputs!H5"], load_values=True)
    bindings = load_series_bindings(FIXTURES / "shard_borvelia_output.yaml")
    series = dict(bindings["series"][0])
    series["validation"] = {"warn_on_partial_overlap": False}
    resolved = resolve_series_binding(graph, wb_path, series, direction="output")
    assert not any(i["code"] == "partial_graph_overlap" for i in resolved["issues"])
