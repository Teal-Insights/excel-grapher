"""Issue 677 — inverted-tree `input.value_map` on scalar selector inputs."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    series_entry,
    write_workbook,
)

_SELECTOR_MAP = {"High": "High ", "Medium": "Medium", "Low": "Low "}


def _selector_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "selector.xlsx",
        {
            "Inputs": {"A1": "High "},
            "Engine": {
                "A1": (
                    '=IF(Inputs!A1="High ",10,IF(Inputs!A1="Medium",5,IF(Inputs!A1="Low ",1,0)))'
                )
            },
            "Outputs": {"A1": "=Engine!A1"},
        },
    )


def _selector_bindings(*, value_map: dict[str, str] | None = _SELECTOR_MAP) -> dict:
    selector = series_entry(
        "selector",
        "Inputs!A1",
        layout="scalar",
        direction="input",
        dtype="string",
        value_map=value_map,
    )
    return bindings_document(
        selector,
        series_entry("chosen", "Engine!A1", layout="scalar", direction="internal", dtype="float"),
        series_entry("out", "Outputs!A1", layout="scalar", direction="output", dtype="float"),
        schema_version="1.15.0",
    )


def test_clean_key_matches_evaluator_with_workbook_needle(tmp_path: Path) -> None:
    workbook = _selector_workbook(tmp_path)
    document = _selector_bindings()
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="value_map_if")
    _catalog, _deps, graph = inverted_graph_parts(workbook, document)
    expected = FormulaEvaluator(graph).evaluate(["Outputs!A1"])["Outputs!A1"]
    assert pkg.compute_out(selector="High") == (expected,)
    assert pkg.compute_out(selector="High") == (10,)


def test_unmapped_value_raises_domain_error_naming_clean_keys(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_selector_workbook(tmp_path), _selector_bindings()),
        tmp_path,
        name="value_map_reject",
    )
    with pytest.raises(ValueError, match=r"selector out of domain: 'Nope'"):
        pkg.compute_out(selector="Nope")
    with pytest.raises(ValueError, match=r"'High'"):
        pkg.compute_out(selector="High ")


def test_map_runs_on_orchestrator_after_domain_not_in_internals(tmp_path: Path) -> None:
    modules = generate_inverted(_selector_workbook(tmp_path), _selector_bindings())
    api = modules["api.py"]
    assert "apply_input_value_map(selector" in api
    assert "require_input_domain(selector" in api
    assert api.index("require_input_domain(selector") < api.index("apply_input_value_map(selector")
    assert "apply_input_value_map" not in modules["internals.py"]
    assert '"High "' in modules["internals.py"] or "'High '" in modules["internals.py"]


def test_shared_runner_maps_once_in_evaluation_body(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "shared.xlsx",
        {
            "Inputs": {"A1": "High "},
            "Engine": {
                "A1": '=IF(Inputs!A1="High ",10,0)',
                "B1": '=IF(Inputs!A1="High ",20,0)',
            },
            "Outputs": {"A1": "=Engine!A1", "B1": "=Engine!B1"},
        },
    )
    document = bindings_document(
        series_entry(
            "selector",
            "Inputs!A1",
            layout="scalar",
            direction="input",
            dtype="string",
            value_map=_SELECTOR_MAP,
        ),
        series_entry("engine_a", "Engine!A1", layout="scalar", direction="internal", dtype="float"),
        series_entry("engine_b", "Engine!B1", layout="scalar", direction="internal", dtype="float"),
        series_entry("out_a", "Outputs!A1", layout="scalar", direction="output", dtype="float"),
        series_entry("out_b", "Outputs!B1", layout="scalar", direction="output", dtype="float"),
        schema_version="1.15.0",
    )
    modules = generate_inverted(workbook, document)
    api = modules["api.py"]
    assert "def _run_0" in api
    assert api.count("apply_input_value_map(selector") == 1
    pkg = load_package(modules, tmp_path, name="value_map_shared")
    assert pkg.compute_out_a(selector="High") == (10,)
    assert pkg.compute_out_b(selector="High") == (20,)
    with pytest.raises(ValueError, match=r"selector out of domain"):
        pkg.compute_out_a(selector="Nope")
