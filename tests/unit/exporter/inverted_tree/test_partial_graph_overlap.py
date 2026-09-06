"""Issue 693 — inverted-tree emit must keep partial_graph_overlap series.

Off-graph cells inside an internal/output `data_range` are a validator
warning, not an export error. Filtering happens in `build_catalog` before
the key domain is resolved (not by shrinking `formula_series()`). Input and
constant series stay unfiltered. Helpers reject sequence arguments that are
not dense over the producer's `__domain__`.

`layout: matrix` keeps hole cells in the catalog (issue 696) so the
rectangle, stride, and `__domain__` stay intact. Helpers emit `None` or a
cached literal at those positions.

`layout: series` internals keep on-graph leaves in `data_range` (issue 707)
so a consumer range that names a literal leaf still has a catalog owner.
Off-graph blanks, literals, and off-closure formulas stay stripped (#693).
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.catalog import build_catalog
from excel_grapher.exporter.inverted_tree.emit import generate_inverted_tree_modules
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import validate_bindings_document, validate_series_bindings
from excel_grapher.series_bindings.normalize import has_output_direction
from excel_grapher.series_bindings.ranges import expand_data_range, series_data_ranges
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    load_package,
    series_entry,
    write_workbook,
)


def _output_targets(document: dict[str, Any], workbook: Path) -> list[str]:
    """Return `data_range` cells of output series only (not all bound cells)."""
    bindings = validate_bindings_document(document)
    targets: list[str] = []
    for series in bindings["series"]:
        if not has_output_direction(series):
            continue
        for data_range in series_data_ranges(series):
            targets.extend(expand_data_range(data_range, workbook=workbook))
    return sorted(set(targets))


def _output_graph(workbook: Path, document: dict[str, Any]):
    return create_dependency_graph(
        workbook,
        _output_targets(document, workbook),
        load_values=True,
        capture_dependency_provenance=True,
    )


def _emit_from_outputs(workbook: Path, document: dict[str, Any]) -> dict[str, str]:
    bindings = validate_bindings_document(document)
    graph = _output_graph(workbook, document)
    return generate_inverted_tree_modules(
        graph,
        series_bindings=bindings,
        bindings_workbook=workbook,
    )


def _mcve_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "partial_overlap_mcve.xlsx",
        {
            "Inputs": {"A1": 10},
            "Engine": {
                "B1": "=Inputs!A1+1",
                "C1": "=Inputs!A1+2",
            },
            "Outputs": {"B1": "=Engine!B1"},
        },
    )


def _mcve_document() -> dict[str, Any]:
    return bindings_document(
        series_entry("rate", "Inputs!A1", layout="scalar", direction="input"),
        series_entry(
            "engine_row",
            "Engine!B1:C1",
            layout="scalar",
            direction="internal",
        ),
        series_entry("result", "Outputs!B1", layout="scalar", direction="output"),
    )


def _year_workbook(tmp_path: Path) -> Path:
    """Five-year model; Engine!D2 (2023) is an unused sibling formula."""
    return write_workbook(
        tmp_path / "partial_overlap_years.xlsx",
        {
            "Inputs": {
                "B1": 2021,
                "C1": 2022,
                "D1": 2023,
                "E1": 2024,
                "F1": 2025,
                "B2": 1.0,
                "C2": 2.0,
                "D2": 3.0,
                "E2": 4.0,
                "F2": 5.0,
            },
            "Engine": {
                "B1": 2021,
                "C1": 2022,
                "D1": 2023,
                "E1": 2024,
                "F1": 2025,
                "B2": "=Inputs!B2+1",
                "C2": "=Inputs!C2+1",
                "D2": "=Inputs!D2+99",
                "E2": "=Inputs!E2+1",
                "F2": "=Inputs!F2+1",
            },
            "Outputs": {
                "B1": 2021,
                "C1": 2022,
                "D1": 2023,
                "E1": 2024,
                "F1": 2025,
                "B2": "=Engine!B2",
                "C2": "=Engine!C2",
                "D2": "=Inputs!D2",
                "E2": "=Engine!E2",
                "F2": "=Engine!F2",
            },
        },
    )


def _year_document() -> dict[str, Any]:
    return bindings_document(
        series_entry(
            "rate",
            "Inputs!B2:F2",
            layout="series",
            direction="input",
            header_row=1,
        ),
        series_entry(
            "engine_row",
            "Engine!B2:F2",
            layout="series",
            direction="internal",
            header_row=1,
        ),
        series_entry(
            "result",
            "Outputs!B2:F2",
            layout="series",
            direction="output",
            header_row=1,
        ),
    )


def test_validate_ok_with_partial_graph_overlap(tmp_path: Path) -> None:
    workbook = _mcve_workbook(tmp_path)
    document = _mcve_document()
    bindings = validate_bindings_document(document)
    graph = _output_graph(workbook, document)
    report = validate_series_bindings(graph, bindings, workbook=workbook)
    assert report["ok"] is True
    assert any(i["code"] == "partial_graph_overlap" for i in report["issues"])


def test_catalog_filters_off_graph_formula_cells(tmp_path: Path) -> None:
    workbook = _mcve_workbook(tmp_path)
    document = _mcve_document()
    bindings = validate_bindings_document(document)
    graph = _output_graph(workbook, document)
    with pytest.warns(UserWarning, match="not graph formula cells"):
        catalog = build_catalog(bindings, workbook=workbook, graph=graph)
    assert catalog.get("engine_row").cells == ("Engine!B1",)
    assert catalog.get("rate").cells == ("Inputs!A1",)
    assert catalog.get("result").cells == ("Outputs!B1",)


def test_catalog_does_not_filter_input_or_constant(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "unfiltered_leaves.xlsx",
        {
            "Inputs": {"A1": 2021, "B1": 2022, "A2": 10, "B2": 20},
            "Const": {"A1": 2021, "B1": 2022, "A2": 1, "B2": 2},
            "Engine": {"A1": "=Inputs!A2+Const!A2"},
            "Outputs": {"A1": "=Engine!A1"},
        },
    )
    document = bindings_document(
        series_entry("rate", "Inputs!A2:B2", layout="series", direction="input", header_row=1),
        series_entry("base", "Const!A2:B2", layout="series", direction="constant", header_row=1),
        series_entry("engine", "Engine!A1", layout="scalar", direction="internal"),
        series_entry("result", "Outputs!A1", layout="scalar", direction="output"),
    )
    bindings = validate_bindings_document(document)
    graph = _output_graph(workbook, document)
    catalog = build_catalog(bindings, workbook=workbook, graph=graph)
    assert catalog.get("rate").cells == ("Inputs!A2", "Inputs!B2")
    assert catalog.get("base").cells == ("Const!A2", "Const!B2")


def test_catalog_warns_when_filtering_collapses_to_scalar(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "collapse_to_scalar.xlsx",
        {
            "Inputs": {"A1": 10},
            "Engine": {
                "B1": 2021,
                "C1": 2022,
                "B2": "=Inputs!A1+1",
                "C2": "=Inputs!A1+2",
            },
            "Outputs": {"B2": "=Engine!B2"},
        },
    )
    document = bindings_document(
        series_entry("rate", "Inputs!A1", layout="scalar", direction="input"),
        series_entry(
            "engine_row",
            "Engine!B2:C2",
            layout="series",
            direction="internal",
            header_row=1,
        ),
        series_entry("result", "Outputs!B2", layout="scalar", direction="output"),
    )
    bindings = validate_bindings_document(document)
    graph = _output_graph(workbook, document)
    with pytest.warns(UserWarning, match="single cell"):
        catalog = build_catalog(bindings, workbook=workbook, graph=graph)
    series = catalog.get("engine_row")
    assert series.layout == "series"
    assert series.cells == ("Engine!B2",)
    assert series.is_scalar


def test_mcve_emit_succeeds(tmp_path: Path) -> None:
    workbook = _mcve_workbook(tmp_path)
    with pytest.warns(UserWarning):
        modules = _emit_from_outputs(workbook, _mcve_document())
    pkg = load_package(modules, tmp_path, name="partial_mcve")
    assert pkg.compute_result(rate=10.0) == pytest.approx((11.0,))


def test_year_keyed_interior_hole_exports(tmp_path: Path) -> None:
    workbook = _year_workbook(tmp_path)
    document = _year_document()
    bindings = validate_bindings_document(document)
    graph = _output_graph(workbook, document)
    with pytest.warns(UserWarning, match="not graph formula cells"):
        catalog = build_catalog(bindings, workbook=workbook, graph=graph)
    engine = catalog.get("engine_row")
    result = catalog.get("result")
    rate = catalog.get("rate")
    assert engine.cells == ("Engine!B2", "Engine!C2", "Engine!E2", "Engine!F2")
    assert [point["TIME_PERIOD"] for point in engine.domain] == [2021, 2022, 2024, 2025]
    assert len(result.cells) == 5
    assert len(rate.cells) == 5

    with pytest.warns(UserWarning):
        modules = _emit_from_outputs(workbook, document)
    pkg = load_package(modules, tmp_path, name="partial_years")
    assert pkg.compute_result.__domain__ == (2021, 2022, 2023, 2024, 2025)
    assert pkg.internals.engine_row.__domain__ == (2021, 2022, 2024, 2025)
    got = pkg.compute_result(rate=(1.0, 2.0, 3.0, 4.0, 5.0))
    assert got == pytest.approx((2.0, 3.0, 3.0, 5.0, 6.0))


def test_helper_rejects_public_length_on_holed_series(tmp_path: Path) -> None:
    workbook = _year_workbook(tmp_path)
    with pytest.warns(UserWarning):
        modules = _emit_from_outputs(workbook, _year_document())
    pkg = load_package(modules, tmp_path, name="partial_guards")
    rate5 = (1.0, 2.0, 3.0, 4.0, 5.0)
    engine4 = pkg.internals.engine_row.__domain__
    assert len(engine4) == 4
    with pytest.raises(ValueError, match="expected length"):
        pkg.internals.engine_row(rate5)
    engine5 = (2.0, 3.0, 102.0, 5.0, 6.0)
    with pytest.raises(ValueError, match="expected length"):
        pkg.internals.result(rate5, engine5)
    with pytest.raises(ValueError, match="expected length"):
        pkg.internals.result(rate5, (2.0, 3.0, 5.0))


def test_helper_docstring_states_dense_domain_contract(tmp_path: Path) -> None:
    workbook = _year_workbook(tmp_path)
    with pytest.warns(UserWarning):
        modules = _emit_from_outputs(workbook, _year_document())
    internals = modules["internals.py"]
    assert "dense over" in internals
    assert "__domain__" in internals
    assert "shorter than the public domain" in internals


def test_emit_refuses_non_leaf_input_overlap(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "non_leaf_input.xlsx",
        {
            "Inputs": {"A1": 10},
            "Engine": {"B1": "=Inputs!A1+1"},
            "Outputs": {"B1": "=Engine!B1"},
        },
    )
    document = bindings_document(
        series_entry("rate", "Engine!B1", layout="scalar", direction="input"),
        series_entry("result", "Outputs!B1", layout="scalar", direction="output"),
    )
    bindings = validate_bindings_document(document)
    graph = _output_graph(workbook, document)
    report = validate_series_bindings(graph, bindings, workbook=workbook)
    assert report["ok"] is False
    assert any(i["code"] == "non_leaf_input_overlap" for i in report["issues"])
    with pytest.raises(InvertedTreeExportError, match="non_leaf_input_overlap"):
        generate_inverted_tree_modules(
            graph,
            series_bindings=bindings,
            bindings_workbook=workbook,
        )


def _matrix_table(data_range: str = "Profile!B2:C3") -> dict[str, Any]:
    return {
        "id": "profile_table",
        "sheet": data_range.split("!", 1)[0],
        "data_range": data_range,
        "layout": "matrix",
        "internal": {},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [
                {
                    "id": "COUNTRY",
                    "concept": "COUNTRY",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "row_label", "label_column": "A", "read": "string"},
                },
                {
                    "id": "TIME_PERIOD",
                    "concept": "TIME_PERIOD",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
                },
            ],
        },
        "key": ["COUNTRY", "TIME_PERIOD"],
    }


def _matrix_document(*outputs: dict[str, Any]) -> dict[str, Any]:
    return bindings_document(_matrix_table(), *outputs)


def _matrix_profile_cells(**overrides: object) -> dict[str, object]:
    cells: dict[str, object] = {
        "B1": 2020,
        "C1": 2021,
        "A2": "France",
        "B2": "=1",
        "C2": "=2",
        "A3": "Kenya",
        "B3": "=3",
        "C3": "=4",
    }
    cells.update(overrides)
    return cells


def test_matrix_holes_are_retained(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "matrix_hole.xlsx",
        {
            "Profile": _matrix_profile_cells(),
            "Outputs": {"A1": "=Profile!B2"},
        },
    )
    document = _matrix_document(
        series_entry("output_cell", "Outputs!A1", layout="scalar", direction="output"),
    )
    bindings = validate_bindings_document(document)
    graph = _output_graph(workbook, document)
    with pytest.warns(UserWarning, match="not graph formula cells"):
        catalog = build_catalog(bindings, workbook=workbook, graph=graph)
    series = catalog.get("profile_table")
    assert series.cells == ("Profile!B2", "Profile!C2", "Profile!B3", "Profile!C3")
    assert series.rect == ("Profile", 2, 2, 3, 3)
    assert series.block_width == 2
    assert len(series.domain) == 4
    assert series.hole_indices == (1, 2, 3)
    assert [hole.kind for hole in series.holes] == [
        "off_closure",
        "off_closure",
        "off_closure",
    ]


def test_matrix_interior_blank_emits_none_and_keeps_stride(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "matrix_blank.xlsx",
        {
            "Profile": _matrix_profile_cells(C2=None),
            "Outputs": {"A1": "=INDEX(Profile!B2:C3,2,1)"},
        },
    )
    document = _matrix_document(
        series_entry("output_cell", "Outputs!A1", layout="scalar", direction="output"),
    )
    bindings = validate_bindings_document(document)
    graph = _output_graph(workbook, document)
    with pytest.warns(UserWarning, match="not graph formula cells"):
        catalog = build_catalog(bindings, workbook=workbook, graph=graph)
    series = catalog.get("profile_table")
    assert series.cells == ("Profile!B2", "Profile!C2", "Profile!B3", "Profile!C3")
    assert series.rect is not None
    assert series.block_width == 2
    assert series.hole_indices == (1,)
    assert series.holes[0].kind == "blank"
    assert series.holes[0].address == "Profile!C2"

    with pytest.warns(UserWarning):
        modules = _emit_from_outputs(workbook, document)
    internals = modules["internals.py"]
    assert "Profile!C2" in internals
    assert "blank" in internals
    assert "float | str | None" in internals
    pkg = load_package(modules, tmp_path, name="matrix_blank")
    helper = pkg.internals.profile_table
    assert helper.__domain__ == (
        ("France", 2020),
        ("France", 2021),
        ("Kenya", 2020),
        ("Kenya", 2021),
    )
    assert len(helper.__domain__) == 4
    assert helper.__holes__ == (1,)
    got = helper()
    assert got[0] == pytest.approx(1.0)
    assert got[1] is None
    assert got[2] == pytest.approx(3.0)
    assert got[3] == pytest.approx(4.0)
    records = pkg.runtime.as_records(helper, got)
    filled = [record for index, record in enumerate(records) if index not in helper.__holes__]
    assert [record["OBS_VALUE"] for record in filled] == pytest.approx([1.0, 3.0, 4.0])
    expected = FormulaEvaluator(graph).evaluate(["Outputs!A1"])
    assert pkg.compute_output_cell() == pytest.approx((expected["Outputs!A1"],))


def test_matrix_off_closure_formula_is_named_in_docstring(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "matrix_off_closure.xlsx",
        {
            "Profile": _matrix_profile_cells(),
            "Outputs": {"A1": "=Profile!B2"},
        },
    )
    document = _matrix_document(
        series_entry("output_cell", "Outputs!A1", layout="scalar", direction="output"),
    )
    with pytest.warns(UserWarning):
        modules = _emit_from_outputs(workbook, document)
    internals = modules["internals.py"]
    assert "Profile!C2" in internals
    assert "not computed" in internals
    pkg = load_package(modules, tmp_path, name="matrix_off_closure")
    got = pkg.internals.profile_table()
    assert got[0] == pytest.approx(1.0)
    assert got[1] is None
    assert got[2] is None
    assert got[3] is None
    assert pkg.internals.profile_table.__holes__ == (1, 2, 3)


def test_matrix_graph_leaf_literal_matches_evaluator(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "matrix_literal.xlsx",
        {
            "Profile": _matrix_profile_cells(C2=99.0),
            "Outputs": {
                "A1": "=Profile!B2",
                "A2": "=Profile!C2+1",
            },
        },
    )
    document = _matrix_document(
        series_entry("on_graph", "Outputs!A1", layout="scalar", direction="output"),
        series_entry("from_literal", "Outputs!A2", layout="scalar", direction="output"),
    )
    bindings = validate_bindings_document(document)
    graph = _output_graph(workbook, document)
    with pytest.warns(UserWarning, match="not graph formula cells"):
        catalog = build_catalog(bindings, workbook=workbook, graph=graph)
    series = catalog.get("profile_table")
    assert series.cells == ("Profile!B2", "Profile!C2", "Profile!B3", "Profile!C3")
    assert series.hole_indices == (1, 2, 3)
    assert series.holes[0].kind == "graph_leaf"
    assert series.holes[0].address == "Profile!C2"

    with pytest.warns(UserWarning):
        modules = _emit_from_outputs(workbook, document)
    assert "99.0" in modules["internals.py"]
    pkg = load_package(modules, tmp_path, name="matrix_literal")
    got = pkg.internals.profile_table()
    assert got[1] == pytest.approx(99.0)
    expected = FormulaEvaluator(graph).evaluate(["Outputs!A1", "Outputs!A2"])
    assert pkg.compute_on_graph() == pytest.approx((expected["Outputs!A1"],))
    assert pkg.compute_from_literal() == pytest.approx((expected["Outputs!A2"],))


def test_matrix_graph_leaf_without_cached_value_raises(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "matrix_missing_cache.xlsx",
        {
            "Profile": _matrix_profile_cells(C2=99.0),
            "Outputs": {
                "A1": "=Profile!B2",
                "A2": "=Profile!C2+1",
            },
        },
    )
    document = _matrix_document(
        series_entry("on_graph", "Outputs!A1", layout="scalar", direction="output"),
        series_entry("from_literal", "Outputs!A2", layout="scalar", direction="output"),
    )
    bindings = validate_bindings_document(document)
    graph = _output_graph(workbook, document)
    leaf = graph.get_node("Profile!C2")
    assert leaf is not None
    graph.set_node_value("Profile!C2", None)
    with pytest.raises(InvertedTreeExportError, match="cached value"):
        build_catalog(bindings, workbook=workbook, graph=graph)


def test_matrix_unreferenced_literal_is_embedded(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "matrix_unref_literal.xlsx",
        {
            "Profile": _matrix_profile_cells(C2=42.0),
            "Outputs": {"A1": "=Profile!B2"},
        },
    )
    document = _matrix_document(
        series_entry("output_cell", "Outputs!A1", layout="scalar", direction="output"),
    )
    bindings = validate_bindings_document(document)
    graph = _output_graph(workbook, document)
    with pytest.warns(UserWarning, match="not graph formula cells"):
        catalog = build_catalog(bindings, workbook=workbook, graph=graph)
    series = catalog.get("profile_table")
    assert series.holes[0].kind == "literal"
    assert series.holes[0].address == "Profile!C2"
    with pytest.warns(UserWarning):
        modules = _emit_from_outputs(workbook, document)
    pkg = load_package(modules, tmp_path, name="matrix_unref_literal")
    got = pkg.internals.profile_table()
    assert got[1] == pytest.approx(42.0)


def _series_leaf_workbook(tmp_path: Path, *, extra_off_graph: bool = False) -> tuple[Path, str]:
    engine: dict[str, object] = {
        "B1": 2021,
        "C1": 2022,
        "D1": 2023,
        "B2": "=Inputs!B2+1",
        "C2": 0,
        "D2": "=Inputs!D2+3",
    }
    data_range = "Engine!B2:D2"
    if extra_off_graph:
        engine["E1"] = 2024
        engine["E2"] = 99
        data_range = "Engine!B2:E2"
    return write_workbook(
        tmp_path / ("series_leaf_extra.xlsx" if extra_off_graph else "series_graph_leaf.xlsx"),
        {
            "Inputs": {
                "B1": 2021,
                "C1": 2022,
                "D1": 2023,
                "B2": 1.0,
                "C2": 2.0,
                "D2": 3.0,
            },
            "Engine": engine,
            "Outputs": {"A1": "=SUM(Engine!B2:D2)"},
        },
    ), data_range


def _series_leaf_document(data_range: str) -> dict[str, Any]:
    return bindings_document(
        series_entry(
            "rate",
            "Inputs!B2:D2",
            layout="series",
            direction="input",
            header_row=1,
        ),
        series_entry(
            "engine_row",
            data_range,
            layout="series",
            direction="internal",
            header_row=1,
        ),
        series_entry("result", "Outputs!A1", layout="scalar", direction="output"),
    )


def test_series_graph_leaf_is_owned_and_sum_emits(tmp_path: Path) -> None:
    workbook, data_range = _series_leaf_workbook(tmp_path)
    document = _series_leaf_document(data_range)
    bindings = validate_bindings_document(document)
    graph = _output_graph(workbook, document)
    catalog = build_catalog(bindings, workbook=workbook, graph=graph)
    series = catalog.get("engine_row")
    assert series.cells == ("Engine!B2", "Engine!C2", "Engine!D2")
    assert catalog.series_id_for("Engine!C2") == "engine_row"
    hole = series.hole_at(1)
    assert hole is not None
    assert hole.kind == "graph_leaf"
    assert hole.address == "Engine!C2"

    report = validate_series_bindings(graph, bindings, workbook=workbook)
    assert report["ok"] is True
    engine_issues = [i for i in report["issues"] if i.get("series_id") == "engine_row"]
    assert any(i["code"] == "leaf_in_formula_series" for i in engine_issues)
    assert not any(i["code"] == "partial_graph_overlap" for i in engine_issues)

    modules = _emit_from_outputs(workbook, document)
    pkg = load_package(modules, tmp_path, name="series_graph_leaf")
    expected = FormulaEvaluator(graph).evaluate(["Outputs!A1"])
    assert pkg.compute_result(rate=(1.0, 2.0, 3.0)) == pytest.approx((expected["Outputs!A1"],))
    assert pkg.internals.engine_row(rate=(1.0, 2.0, 3.0))[1] == pytest.approx(0.0)


def test_series_graph_leaf_strips_off_graph_and_keeps_leaf(tmp_path: Path) -> None:
    workbook, data_range = _series_leaf_workbook(tmp_path, extra_off_graph=True)
    document = _series_leaf_document(data_range)
    bindings = validate_bindings_document(document)
    graph = _output_graph(workbook, document)
    with pytest.warns(UserWarning, match="not graph formula cells"):
        catalog = build_catalog(bindings, workbook=workbook, graph=graph)
    series = catalog.get("engine_row")
    assert series.cells == ("Engine!B2", "Engine!C2", "Engine!D2")
    assert catalog.series_id_for("Engine!E2") is None
    assert series.hole_at(1) is not None
    assert series.hole_at(1).kind == "graph_leaf"

    report = validate_series_bindings(graph, bindings, workbook=workbook)
    engine_codes = {i["code"] for i in report["issues"] if i.get("series_id") == "engine_row"}
    assert "leaf_in_formula_series" in engine_codes
    assert "partial_graph_overlap" in engine_codes

    with pytest.warns(UserWarning):
        modules = _emit_from_outputs(workbook, document)
    pkg = load_package(modules, tmp_path, name="series_leaf_extra")
    expected = FormulaEvaluator(graph).evaluate(["Outputs!A1"])
    assert pkg.compute_result(rate=(1.0, 2.0, 3.0)) == pytest.approx((expected["Outputs!A1"],))


def test_series_graph_leaf_without_cached_value_raises(tmp_path: Path) -> None:
    workbook, data_range = _series_leaf_workbook(tmp_path)
    document = _series_leaf_document(data_range)
    bindings = validate_bindings_document(document)
    graph = _output_graph(workbook, document)
    leaf = graph.get_node("Engine!C2")
    assert leaf is not None
    graph.set_node_value("Engine!C2", None)
    with pytest.raises(InvertedTreeExportError, match="cached value"):
        build_catalog(bindings, workbook=workbook, graph=graph)
