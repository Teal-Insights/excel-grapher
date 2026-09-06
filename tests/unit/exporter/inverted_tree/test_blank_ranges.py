"""Inverted-tree emit honors `blank_ranges` named in formulas (#700, #703).

Declared structural blanks are omitted from the graph and from bindings
coverage. `FormulaEvaluator` resolves them as empty. Range walks and
single-cell `CellRef` sites in inverted-tree emit must do the same instead
of fail-closing on unowned catalog cells.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import (
    load_series_bindings,
    validate_bindings_document,
    validate_series_bindings,
)
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    call_compute,
    generate_inverted,
    input_kwargs,
    inverted_graph_parts,
    load_package,
    series_entry,
    write_workbook,
)

_BLANK = ("Lookup!A1:C3",)


def _mcve_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "blank_range_vlookup.xlsx",
        {
            "Inputs": {"A1": 1},
            "Lookup": {},
            "Engine": {"B1": "=VLOOKUP(Inputs!A1,Lookup!A1:C3,3,FALSE)"},
            "Outputs": {"B1": "=Engine!B1"},
        },
    )


def _mcve_bindings() -> dict[str, Any]:
    return bindings_document(
        series_entry("key", "Inputs!A1", layout="scalar", direction="input"),
        series_entry("lookup_result", "Engine!B1", layout="scalar", direction="internal"),
        series_entry("result", "Outputs!B1", layout="scalar", direction="output"),
    )


def _if_vlookup_workbook(tmp_path: Path) -> Path:
    """LIC-DSF shape: one VLOOKUP table is bound, the other is structural blank."""
    return write_workbook(
        tmp_path / "blank_range_if_vlookup.xlsx",
        {
            "Inputs": {"A1": 1, "B1": 0},
            "Data": {"A1": 1, "B1": 10, "C1": 100, "A2": 2, "B2": 20, "C2": 200},
            "Lookup": {},
            "Engine": {
                "B1": (
                    "=IF(Inputs!B1=1,"
                    "VLOOKUP(Inputs!A1,Lookup!A1:C3,3,FALSE),"
                    "VLOOKUP(Inputs!A1,Data!A1:C2,3,FALSE))"
                )
            },
            "Outputs": {"B1": "=Engine!B1"},
        },
    )


def _if_vlookup_bindings() -> dict[str, Any]:
    return bindings_document(
        series_entry("key", "Inputs!A1", layout="scalar", direction="input"),
        series_entry("flag", "Inputs!B1", layout="scalar", direction="input", dtype="int"),
        series_entry("table_a1", "Data!A1", layout="scalar", direction="constant"),
        series_entry("table_b1", "Data!B1", layout="scalar", direction="constant"),
        series_entry("table_c1", "Data!C1", layout="scalar", direction="constant"),
        series_entry("table_a2", "Data!A2", layout="scalar", direction="constant"),
        series_entry("table_b2", "Data!B2", layout="scalar", direction="constant"),
        series_entry("table_c2", "Data!C2", layout="scalar", direction="constant"),
        series_entry("lookup_result", "Engine!B1", layout="scalar", direction="internal"),
        series_entry("result", "Outputs!B1", layout="scalar", direction="output"),
    )


def _sum_partial_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "blank_range_sum.xlsx",
        {
            "Inputs": {"A1": 2024, "B1": 2025, "A2": 1.5, "B2": 2.5},
            "Outputs": {"Z1": "=SUM(Inputs!A2:B4)"},
        },
    )


def _sum_partial_bindings() -> dict[str, Any]:
    return bindings_document(
        series_entry("src", "Inputs!A2:B2", layout="series", direction="input", header_row=1),
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output"),
    )


def _scalar(value: object) -> object:
    if isinstance(value, tuple):
        assert len(value) == 1
        return value[0]
    return value


def test_blank_vlookup_table_is_not_a_graph_node(tmp_path: Path) -> None:
    workbook = _mcve_workbook(tmp_path)
    bindings = validate_bindings_document(_mcve_bindings())
    graph = create_dependency_graph(
        workbook,
        ["Outputs!B1"],
        load_values=True,
        blank_ranges=_BLANK,
    )
    assert set(graph.leaf_keys()) | set(graph.formula_keys()) == {
        "Inputs!A1",
        "Engine!B1",
        "Outputs!B1",
    }
    assert graph.get_node("Lookup!A1") is None
    report = validate_series_bindings(graph, bindings, workbook=workbook)
    assert report["ok"] is True


def test_blank_vlookup_range_fail_closes_without_blank_ranges(tmp_path: Path) -> None:
    workbook = _mcve_workbook(tmp_path)
    with pytest.raises(InvertedTreeExportError, match="is not a bound series"):
        generate_inverted(workbook, _mcve_bindings())


def test_generate_inverted_tree_modules_accepts_blank_vlookup_table(tmp_path: Path) -> None:
    workbook = _mcve_workbook(tmp_path)
    modules = generate_inverted(workbook, _mcve_bindings(), blank_ranges=_BLANK)
    assert "xl_vlookup(" in modules["internals.py"]
    assert "None, None, None" in modules["internals.py"]


def test_generate_modules_forwards_blank_ranges_to_inverted_tree(tmp_path: Path) -> None:
    workbook = _mcve_workbook(tmp_path)
    bindings = validate_bindings_document(_mcve_bindings())
    graph = create_dependency_graph(
        workbook,
        ["Outputs!B1"],
        load_values=True,
        blank_ranges=_BLANK,
    )
    with CodeGenerator(graph) as generator:
        modules = generator.generate_modules(
            series_bindings=bindings,
            bindings_workbook=workbook,
            blank_ranges=_BLANK,
        )
    assert "xl_vlookup(" in modules["internals.py"]


def test_blank_vlookup_package_matches_evaluator(tmp_path: Path) -> None:
    workbook = _mcve_workbook(tmp_path)
    catalog, _deps, graph = inverted_graph_parts(workbook, _mcve_bindings(), blank_ranges=_BLANK)
    pkg = load_package(
        generate_inverted(workbook, _mcve_bindings(), blank_ranges=_BLANK),
        tmp_path,
        name="blank_vlookup",
    )
    with FormulaEvaluator(graph, blank_ranges=_BLANK) as ev:
        expected = ev.evaluate(["Outputs!B1"])["Outputs!B1"]
    got = call_compute(pkg, catalog.output_series()[0].series_id, input_kwargs(catalog, graph))
    assert _scalar(got) == expected


def test_if_vlookup_uses_bound_table_when_blank_branch_is_unused(tmp_path: Path) -> None:
    workbook = _if_vlookup_workbook(tmp_path)
    catalog, _deps, graph = inverted_graph_parts(
        workbook, _if_vlookup_bindings(), blank_ranges=_BLANK
    )
    pkg = load_package(
        generate_inverted(workbook, _if_vlookup_bindings(), blank_ranges=_BLANK),
        tmp_path,
        name="if_vlookup",
    )
    with FormulaEvaluator(graph, blank_ranges=_BLANK) as ev:
        expected = ev.evaluate(["Outputs!B1"])["Outputs!B1"]
    got = call_compute(pkg, catalog.output_series()[0].series_id, input_kwargs(catalog, graph))
    assert _scalar(got) == pytest.approx(expected)
    assert _scalar(got) == pytest.approx(100)


def test_sum_drops_blank_interior_from_ownership_check(tmp_path: Path) -> None:
    workbook = _sum_partial_workbook(tmp_path)
    blank = ("Inputs!A3:B4",)
    modules = generate_inverted(workbook, _sum_partial_bindings(), blank_ranges=blank)
    assert "xl_sum(" in modules["internals.py"]
    pkg = load_package(modules, tmp_path, name="blank_sum")
    catalog, _deps, graph = inverted_graph_parts(
        workbook, _sum_partial_bindings(), blank_ranges=blank
    )
    with FormulaEvaluator(graph, blank_ranges=blank) as ev:
        expected = ev.evaluate(["Outputs!Z1"])["Outputs!Z1"]
    got = call_compute(pkg, catalog.output_series()[0].series_id, input_kwargs(catalog, graph))
    assert _scalar(got) == pytest.approx(expected)


def test_issue_mcve_generate_does_not_raise(tmp_path: Path) -> None:
    """Reproduce the self-contained MCVE from issue 700."""
    import yaml
    from fastpyxl import Workbook

    from excel_grapher.exporter.inverted_tree.emit import generate_inverted_tree_modules

    root = tmp_path / "mcve_blank_range_vlookup"
    root.mkdir()
    workbook_path = root / "workbook.xlsx"
    bindings_path = root / "bindings"
    bindings_path.mkdir()

    wb = Workbook()
    default = wb.active
    wb.remove(default)
    inputs = wb.create_sheet("Inputs")
    wb.create_sheet("Lookup")
    engine = wb.create_sheet("Engine")
    outputs = wb.create_sheet("Outputs")
    inputs["A1"] = 1
    engine["B1"] = "=VLOOKUP(Inputs!A1,Lookup!A1:C3,3,FALSE)"
    outputs["B1"] = "=Engine!B1"
    wb.save(workbook_path)

    scalar = {
        "layout": "scalar",
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [],
        },
        "key": [],
    }
    (bindings_path / "inputs.bindings.yaml").write_text(
        yaml.safe_dump(
            {
                "schema_version": "1.14.0",
                "workbook": "workbook.xlsx",
                "concept_scheme": {
                    "id": "mcve",
                    "concepts": [{"id": "OBS_VALUE", "name": "Observation", "dtype": "number"}],
                },
                "series": [
                    {
                        "id": "key",
                        "sheet": "Inputs",
                        "data_range": "Inputs!A1",
                        "input": {
                            "setter": {
                                "name": "set_key",
                                "record_contract": "records",
                                "strict": True,
                            }
                        },
                        **scalar,
                    }
                ],
            },
            sort_keys=False,
        ),
        encoding="utf-8",
    )
    (bindings_path / "outputs.bindings.yaml").write_text(
        yaml.safe_dump(
            {
                "schema_version": "1.14.0",
                "workbook": "workbook.xlsx",
                "series": [
                    {
                        "id": "result",
                        "sheet": "Outputs",
                        "data_range": "Outputs!B1",
                        "output": {
                            "compute": {
                                "name": "compute_result",
                                "record_contract": "records",
                            }
                        },
                        **scalar,
                    }
                ],
            },
            sort_keys=False,
        ),
        encoding="utf-8",
    )
    (bindings_path / "internals.bindings.yaml").write_text(
        yaml.safe_dump(
            {
                "schema_version": "1.14.0",
                "workbook": "workbook.xlsx",
                "series": [
                    {
                        "id": "lookup_result",
                        "sheet": "Engine",
                        "data_range": "Engine!B1",
                        "internal": {},
                        **scalar,
                    }
                ],
            },
            sort_keys=False,
        ),
        encoding="utf-8",
    )

    blank = ("Lookup!A1:C3",)
    graph = create_dependency_graph(
        workbook_path,
        ["Outputs!B1"],
        load_values=True,
        blank_ranges=blank,
    )
    assert set(graph.leaf_keys()) | set(graph.formula_keys()) == {
        "Inputs!A1",
        "Engine!B1",
        "Outputs!B1",
    }
    assert graph.get_node("Lookup!A1") is None

    bindings = load_series_bindings(bindings_path)
    report = validate_series_bindings(graph, bindings, workbook=workbook_path)
    assert report["ok"] is True

    generate_inverted_tree_modules(
        graph,
        series_bindings=bindings,
        bindings_workbook=workbook_path,
        blank_ranges=blank,
    )


_BLANK_CELL = ("Lookup!A1",)


def _cellref_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "blank_cellref.xlsx",
        {
            "Inputs": {"A1": 1},
            "Lookup": {},
            "Engine": {"B1": "=IF(ISNUMBER(Lookup!A1),Lookup!A1,0)+Inputs!A1"},
            "Outputs": {"B1": "=Engine!B1"},
        },
    )


def _cellref_bindings() -> dict[str, Any]:
    return bindings_document(
        series_entry("key", "Inputs!A1", layout="scalar", direction="input"),
        series_entry("lookup_result", "Engine!B1", layout="scalar", direction="internal"),
        series_entry("result", "Outputs!B1", layout="scalar", direction="output"),
    )


def test_blank_cellref_is_not_a_graph_node(tmp_path: Path) -> None:
    workbook = _cellref_workbook(tmp_path)
    bindings = validate_bindings_document(_cellref_bindings())
    graph = create_dependency_graph(
        workbook,
        ["Outputs!B1"],
        load_values=True,
        blank_ranges=_BLANK_CELL,
    )
    assert set(graph.leaf_keys()) | set(graph.formula_keys()) == {
        "Inputs!A1",
        "Engine!B1",
        "Outputs!B1",
    }
    assert graph.get_node("Lookup!A1") is None
    report = validate_series_bindings(graph, bindings, workbook=workbook)
    assert report["ok"] is True


def test_blank_cellref_fail_closes_without_blank_ranges(tmp_path: Path) -> None:
    workbook = _cellref_workbook(tmp_path)
    with pytest.raises(InvertedTreeExportError, match="cell Lookup!A1 is not in any bound series"):
        generate_inverted(workbook, _cellref_bindings())


def test_generate_inverted_tree_modules_accepts_blank_cellref(tmp_path: Path) -> None:
    workbook = _cellref_workbook(tmp_path)
    modules = generate_inverted(workbook, _cellref_bindings(), blank_ranges=_BLANK_CELL)
    internals = modules["internals.py"]
    assert "Lookup!A1" not in internals
    assert "None" in internals
    assert "xl_isnumber(" in internals


def test_generate_modules_forwards_blank_ranges_for_cellref(tmp_path: Path) -> None:
    workbook = _cellref_workbook(tmp_path)
    bindings = validate_bindings_document(_cellref_bindings())
    graph = create_dependency_graph(
        workbook,
        ["Outputs!B1"],
        load_values=True,
        blank_ranges=_BLANK_CELL,
    )
    with CodeGenerator(graph) as generator:
        modules = generator.generate_modules(
            series_bindings=bindings,
            bindings_workbook=workbook,
            blank_ranges=_BLANK_CELL,
        )
    assert "None" in modules["internals.py"]


def test_blank_cellref_package_matches_evaluator(tmp_path: Path) -> None:
    workbook = _cellref_workbook(tmp_path)
    catalog, _deps, graph = inverted_graph_parts(
        workbook, _cellref_bindings(), blank_ranges=_BLANK_CELL
    )
    pkg = load_package(
        generate_inverted(workbook, _cellref_bindings(), blank_ranges=_BLANK_CELL),
        tmp_path,
        name="blank_cellref",
    )
    with FormulaEvaluator(graph, blank_ranges=_BLANK_CELL) as ev:
        expected = ev.evaluate(["Outputs!B1"])["Outputs!B1"]
    got = call_compute(pkg, catalog.output_series()[0].series_id, input_kwargs(catalog, graph))
    assert _scalar(got) == expected
    assert _scalar(got) == pytest.approx(1)


def test_issue_703_mcve_generate_does_not_raise(tmp_path: Path) -> None:
    """Reproduce the self-contained MCVE from issue 703."""
    import yaml
    from fastpyxl import Workbook

    from excel_grapher.exporter.inverted_tree.emit import generate_inverted_tree_modules

    root = tmp_path / "mcve_blank_cellref"
    root.mkdir()
    workbook_path = root / "workbook.xlsx"
    bindings_path = root / "bindings"
    bindings_path.mkdir()

    wb = Workbook()
    default = wb.active
    wb.remove(default)
    inputs = wb.create_sheet("Inputs")
    wb.create_sheet("Lookup")
    engine = wb.create_sheet("Engine")
    outputs = wb.create_sheet("Outputs")
    inputs["A1"] = 1
    engine["B1"] = "=IF(ISNUMBER(Lookup!A1),Lookup!A1,0)+Inputs!A1"
    outputs["B1"] = "=Engine!B1"
    wb.save(workbook_path)

    scalar = {
        "layout": "scalar",
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [],
        },
        "key": [],
    }
    (bindings_path / "inputs.bindings.yaml").write_text(
        yaml.safe_dump(
            {
                "schema_version": "1.14.0",
                "workbook": "workbook.xlsx",
                "concept_scheme": {
                    "id": "mcve",
                    "concepts": [{"id": "OBS_VALUE", "name": "Observation", "dtype": "number"}],
                },
                "series": [
                    {
                        "id": "key",
                        "sheet": "Inputs",
                        "data_range": "Inputs!A1",
                        "input": {
                            "setter": {
                                "name": "set_key",
                                "record_contract": "records",
                                "strict": True,
                            }
                        },
                        **scalar,
                    }
                ],
            },
            sort_keys=False,
        ),
        encoding="utf-8",
    )
    (bindings_path / "outputs.bindings.yaml").write_text(
        yaml.safe_dump(
            {
                "schema_version": "1.14.0",
                "workbook": "workbook.xlsx",
                "series": [
                    {
                        "id": "result",
                        "sheet": "Outputs",
                        "data_range": "Outputs!B1",
                        "output": {
                            "compute": {
                                "name": "compute_result",
                                "record_contract": "records",
                            }
                        },
                        **scalar,
                    }
                ],
            },
            sort_keys=False,
        ),
        encoding="utf-8",
    )
    (bindings_path / "internals.bindings.yaml").write_text(
        yaml.safe_dump(
            {
                "schema_version": "1.14.0",
                "workbook": "workbook.xlsx",
                "series": [
                    {
                        "id": "lookup_result",
                        "sheet": "Engine",
                        "data_range": "Engine!B1",
                        "internal": {},
                        **scalar,
                    }
                ],
            },
            sort_keys=False,
        ),
        encoding="utf-8",
    )

    blank = ("Lookup!A1",)
    graph = create_dependency_graph(
        workbook_path,
        ["Outputs!B1"],
        load_values=True,
        blank_ranges=blank,
    )
    assert set(graph.leaf_keys()) | set(graph.formula_keys()) == {
        "Inputs!A1",
        "Engine!B1",
        "Outputs!B1",
    }
    assert graph.get_node("Lookup!A1") is None

    bindings = load_series_bindings(bindings_path)
    report = validate_series_bindings(graph, bindings, workbook=workbook_path)
    assert report["ok"] is True

    generate_inverted_tree_modules(
        graph,
        series_bindings=bindings,
        bindings_workbook=workbook_path,
        blank_ranges=blank,
    )
