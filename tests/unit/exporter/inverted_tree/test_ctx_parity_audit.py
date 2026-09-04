"""Issue 662 probes: ctx features that inverted-tree may lack.

Each test is a shape through inverted-tree (and FormulaEvaluator when the
shape exports) so audit rows marked `?` are evidence, not guesses.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree import InvertedTreeExportError
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    series_entry,
    write_workbook,
)


def _scalar(value: object) -> object:
    if isinstance(value, tuple):
        assert len(value) == 1
        return value[0]
    return value


def _copy_bindings(
    input_range: str,
    output_range: str,
    *,
    dtype: str = "float",
    layout: str = "scalar",
    header_row: int | None = None,
) -> dict:
    return bindings_document(
        series_entry(
            "src",
            input_range,
            layout=layout,
            direction="input",
            dtype=dtype,
            header_row=header_row,
        ),
        series_entry(
            "out",
            output_range,
            layout=layout,
            direction="output",
            dtype=dtype,
            header_row=header_row,
        ),
    )


def test_bool_dtype_round_trips(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "bool.xlsx",
        {
            "Inputs": {"A1": True},
            "Outputs": {"A1": "=Inputs!A1"},
        },
    )
    modules = generate_inverted(workbook, _copy_bindings("Inputs!A1", "Outputs!A1", dtype="bool"))
    pkg = load_package(modules, tmp_path, name="audit_bool")
    assert pkg.compute_out(src=False) == (False,)
    catalog, _deps, graph = inverted_graph_parts(
        workbook, _copy_bindings("Inputs!A1", "Outputs!A1")
    )
    expected = FormulaEvaluator(graph).evaluate(["Outputs!A1"])["Outputs!A1"]
    assert _scalar(pkg.compute_out(src=True)) == expected


def test_datetime_input_emits_data_literals(tmp_path: Path) -> None:
    stamp = datetime(2020, 1, 15, 12, 30)
    workbook = write_workbook(
        tmp_path / "dt.xlsx",
        {
            "Inputs": {"A1": stamp},
            "Outputs": {"A1": "=Inputs!A1"},
        },
    )
    document = _copy_bindings("Inputs!A1", "Outputs!A1", dtype="datetime")
    modules = generate_inverted(workbook, document)
    assert "from datetime import datetime" in modules["data.py"]
    pkg = load_package(modules, tmp_path, name="audit_dt")
    assert stamp == pkg.data.SRC_DEFAULT
    assert pkg.compute_out(src=stamp) == (stamp,)


def test_list_data_range_and_sheet_name_keys_match_evaluator(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "shards.xlsx",
        {
            "Baseline": {"B1": 2024, "C1": 2025, "B2": 1.0, "C2": 2.0},
            "Shock": {"B1": 2024, "C1": 2025, "B2": 3.0, "C2": 4.0},
            "Outputs": {"Z1": "=Baseline!B2"},
        },
    )
    growth = series_entry(
        "growth",
        "Baseline!B2:C2",
        layout="series",
        direction="input",
        header_row=1,
    )
    growth["sheet"] = ["Baseline", "Shock"]
    growth["data_range"] = ["Baseline!B2:C2", "Shock!B2:C2"]
    growth["key"] = ["SCENARIO", "TIME_PERIOD"]
    growth["structure"]["dimensions"].insert(
        0,
        {
            "concept": "SCENARIO",
            "role": "key",
            "scope": "cell",
            "bind": {"kind": "sheet_name"},
        },
    )
    document = bindings_document(
        growth,
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output"),
        schema_version="1.14.0",
    )
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    assert catalog.get("growth").cells == (
        "Baseline!B2",
        "Baseline!C2",
        "Shock!B2",
        "Shock!C2",
    )
    scenarios = [point["SCENARIO"] for point in catalog.get("growth").domain]
    assert scenarios == ["Baseline", "Baseline", "Shock", "Shock"]
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="audit_shards")
    expected = FormulaEvaluator(graph).evaluate(["Outputs!Z1"])["Outputs!Z1"]
    assert _scalar(pkg.compute_out(growth=(1.0, 2.0, 3.0, 4.0))) == expected


def test_value_map_key_domain_is_resolved(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "vmap.xlsx",
        {
            "Inputs": {"A1": 2024, "A2": 10.0, "A3": 20.0},
            "Outputs": {"Z1": "=Inputs!A2+Inputs!A3"},
        },
    )
    src = series_entry(
        "src",
        "Inputs!A2:A3",
        layout="series",
        direction="input",
        header_row=1,
    )
    src["key"] = ["SCENARIO", "TIME_PERIOD"]
    src["structure"]["dimensions"].insert(
        0,
        {
            "concept": "SCENARIO",
            "role": "key",
            "scope": "cell",
            "bind": {"kind": "value_map", "values": {"Base": 2, "Alt": 3}},
        },
    )
    document = bindings_document(
        src,
        series_entry("out", "Outputs!Z1", layout="scalar", direction="output"),
        schema_version="1.14.0",
    )
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    assert [point["SCENARIO"] for point in catalog.get("src").domain] == ["Base", "Alt"]
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="audit_vmap")
    expected = FormulaEvaluator(graph).evaluate(["Outputs!Z1"])["Outputs!Z1"]
    assert _scalar(pkg.compute_out(src=(10.0, 20.0))) == expected


def test_named_range_data_range_exports(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "named.xlsx",
        {
            "Inputs": {"B1": 2024, "C1": 2025, "B2": 1.5, "C2": 2.5},
            "Outputs": {"A1": "=Inputs!B2", "B1": "=Inputs!C2", "A10": 1, "B10": 2},
        },
        defined_names={"GROWTH": "Inputs!$B$2:$C$2"},
    )
    src = series_entry(
        "src",
        "GROWTH",
        layout="series",
        direction="input",
        header_row=1,
    )
    src["sheet"] = "Inputs"
    document = bindings_document(
        src,
        series_entry(
            "out",
            "Outputs!A1:B1",
            layout="series",
            direction="output",
            header_row=10,
        ),
    )
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    assert catalog.get("src").cells == ("Inputs!B2", "Inputs!C2")
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="audit_named")
    expected = FormulaEvaluator(graph).evaluate(["Outputs!A1", "Outputs!B1"])
    assert pkg.compute_out(src=(1.5, 2.5)) == (expected["Outputs!A1"], expected["Outputs!B1"])


def test_named_range_formula_expands_to_bound_cell(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "named_formula.xlsx",
        {
            "Inputs": {"A1": 4.0},
            "Outputs": {"A1": "=RATE"},
        },
        defined_names={"RATE": "Inputs!$A$1"},
    )
    document = _copy_bindings("Inputs!A1", "Outputs!A1")
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="audit_named_f")
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    expected = FormulaEvaluator(graph).evaluate(["Outputs!A1"])["Outputs!A1"]
    assert _scalar(pkg.compute_out(src=4.0)) == expected


@pytest.mark.parametrize(
    ("formula", "extra_cells", "match"),
    [
        ('=INDIRECT("Inputs!A1")', {}, r"no inverted-tree runtime helper|INDIRECT"),
        (
            "=SUMPRODUCT(Inputs!A1:A2,Inputs!B1:B2)",
            {"A2": 2.0, "B1": 3.0, "B2": 4.0},
            r"bare range|no inverted-tree runtime helper|SUMPRODUCT",
        ),
        (
            "=IFERROR(1/Inputs!A1,0)",
            {},
            r"no inverted-tree runtime helper|IFERROR",
        ),
        (
            "=SUM(Inputs!A1:A2)",
            {"A2": 2.0},
            r"bare range|no inverted-tree runtime helper|SUM",
        ),
        (
            "=SUM(IF(Inputs!A1:A2>0,Inputs!A1:A2))",
            {"A2": 2.0},
            r"bare range|no inverted-tree runtime helper|unsupported",
        ),
        (
            "=SUM(Inputs!A:A)",
            {},
            r"unsupported AST node|WholeColumn|no inverted-tree runtime helper|bare range",
        ),
    ],
)
def test_ctx_library_shapes_fail_closed(
    tmp_path: Path,
    formula: str,
    extra_cells: dict[str, object],
    match: str,
) -> None:
    inputs = {"A1": 1.0, **extra_cells}
    workbook = write_workbook(
        tmp_path / "fail_closed.xlsx",
        {
            "Inputs": inputs,
            "Outputs": {"A1": formula},
        },
    )
    src_range = (
        "Inputs!A1:B2"
        if "B1" in extra_cells
        else ("Inputs!A1:A2" if "A2" in extra_cells else "Inputs!A1")
    )
    layout = "series" if ":" in src_range else "scalar"
    document = bindings_document(
        series_entry(
            "src",
            src_range,
            layout=layout,
            direction="input",
            header_row=10 if layout == "series" else None,
        ),
        series_entry("out", "Outputs!A1", layout="scalar", direction="output"),
    )
    if layout == "series":
        workbook = write_workbook(
            tmp_path / "fail_closed.xlsx",
            {
                "Inputs": {**inputs, "A10": 1, "B10": 2},
                "Outputs": {"A1": formula},
            },
        )
    with pytest.raises(InvertedTreeExportError, match=match):
        generate_inverted(workbook, document)


def test_cross_sheet_range_fails_closed(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "cross.xlsx",
        {
            "Inputs": {"A1": 1.0},
            "Other": {"A1": 2.0},
            "Outputs": {"A1": "=SUM(Inputs!A1:Other!A1)"},
        },
    )
    document = bindings_document(
        series_entry("left", "Inputs!A1", layout="scalar", direction="input"),
        series_entry("right", "Other!A1", layout="scalar", direction="input"),
        series_entry("out", "Outputs!A1", layout="scalar", direction="output"),
    )
    with pytest.raises(InvertedTreeExportError, match=r"cross-sheet range|unsupported|not a bound"):
        generate_inverted(workbook, document)


def test_input_domain_rejects_out_of_range_argument(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "domain.xlsx",
        {
            "Inputs": {"A1": 0},
            "Outputs": {"A1": "=Inputs!A1"},
        },
    )
    document = bindings_document(
        series_entry(
            "flag",
            "Inputs!A1",
            layout="scalar",
            direction="input",
            dtype="int",
            domain={"enum": [0, 1]},
        ),
        series_entry("out", "Outputs!A1", layout="scalar", direction="output", dtype="int"),
    )
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="audit_domain")
    assert pkg.compute_out(flag=0) == (0,)
    with pytest.raises(ValueError, match=r"flag out of domain"):
        pkg.compute_out(flag=2)
