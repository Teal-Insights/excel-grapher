"""Adjacent identical inverted-tree branches coalesce at emission (#785)."""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from excel_grapher.exporter.inverted_tree.ast_emit import (
    _cannot_raise_xl_error,
    _coalesce_adjacent_bodies,
)
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    series_entry,
    write_workbook,
)


def _series_bindings(data_range: str) -> dict:
    return bindings_document(
        series_entry(
            "result",
            data_range,
            layout="series",
            direction="output",
            label_column="A",
        ),
    )


def _vertical_workbook(tmp_path: Path, name: str, column_b: dict[int, object]) -> Path:
    cells: dict[str, object] = {}
    for row, value in column_b.items():
        cells[f"A{row}"] = 2000 + row
        if value is not None:
            cells[f"B{row}"] = value
    return write_workbook(tmp_path / name, {"Data": cells})


def _instance_compute_source(internals: str, series_id: str = "result") -> str:
    header = f"    def {series_id}_compute("
    start = internals.index(header)
    lines = internals[start:].splitlines()
    body = [lines[0]]
    for line in lines[1:]:
        if line == "":
            break
        if line.startswith("    ") and not line.startswith("        "):
            break
        body.append(line)
    return "\n".join(body)


def _helper_body_source(internals: str, series_id: str = "result") -> str:
    match = re.search(rf"^def {series_id}\(", internals, re.MULTILINE)
    assert match is not None, f"helper {series_id} not found"
    lines = internals[match.start() :].splitlines()
    body = [lines[0]]
    for line in lines[1:]:
        if line.startswith("def ") or line.startswith("@publish"):
            break
        body.append(line)
    return "\n".join(body)


def _flow_lines(source: str) -> list[str]:
    keys = ("if i <", "elif i <", "else:", "return ", "try:", "except XlError")
    return [line for line in source.splitlines() if any(key in line for key in keys)]


def test_coalesce_adjacent_bodies_merges_identical_runs() -> None:
    runs = [("", 0, 1), ("", 1, 2), ("a", 2, 3), ("a", 3, 5), ("b", 5, 6)]
    bodies = ["None", "None", "x[i]", "x[i]", "y[i]"]
    assert _coalesce_adjacent_bodies(runs, bodies) == [
        (0, 2, "None", ()),
        (2, 5, "x[i]", ()),
        (5, 6, "y[i]", ()),
    ]


def test_cannot_raise_xl_error_accepts_literals_only() -> None:
    assert _cannot_raise_xl_error("None")
    assert _cannot_raise_xl_error("as_measure(1.0)")
    assert _cannot_raise_xl_error("as_measure(1, 'int')")
    assert _cannot_raise_xl_error("True")
    assert not _cannot_raise_xl_error("as_measure(xl_add(x[i], 1))")
    assert not _cannot_raise_xl_error("xl_raise('#REF!')")


@pytest.mark.parametrize("size", [10, 40])
def test_demand_driven_blank_run_control_flow_is_independent_of_length(
    tmp_path: Path, size: int
) -> None:
    column = {1: "=1", size: f"=B{size - 1}+1"}
    for row in range(2, size):
        column[row] = None
    workbook = _vertical_workbook(tmp_path, f"holes_{size}.xlsx", column)
    document = _series_bindings(f"Data!B1:B{size}")
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    series = catalog.get("result")
    assert series.hole_indices == tuple(range(1, size - 1))
    assert graph.get_node(f"Data!B{size}") is not None

    modules = generate_inverted(workbook, document, force_rung=3)
    compute = _instance_compute_source(modules["_kernels.py"])
    assert compute.count("return None") == 1
    assert compute.count("if i <") == 2
    assert compute.count("except XlError") <= 1
    none_line = next(line for line in compute.splitlines() if line.strip() == "return None")
    assert "try:" not in none_line
    pkg = load_package(modules, tmp_path, name=f"holes_{size}")
    got = pkg.compute_result()
    assert got[2001] == pytest.approx(1.0)
    assert got[2002] is None
    assert (2002,) in pkg.data.RESULT_DOMAIN
    assert got[2000 + size - 1] is None
    assert len(got.domain) == size
    assert len(pkg.data.RESULT_DOMAIN) == size


def test_demand_driven_blank_run_source_size_does_not_grow_with_holes(tmp_path: Path) -> None:
    lengths: list[int] = []
    for size in (10, 40):
        column = {1: "=1", size: f"=B{size - 1}+1"}
        for row in range(2, size):
            column[row] = None
        workbook = _vertical_workbook(tmp_path, f"holes_size_{size}.xlsx", column)
        modules = generate_inverted(workbook, _series_bindings(f"Data!B1:B{size}"), force_rung=3)
        compute = _instance_compute_source(modules["_kernels.py"])
        lengths.append(len(_flow_lines(compute)))
    assert lengths[0] == lengths[1]


def test_distinct_neighboring_literals_stay_separate(tmp_path: Path) -> None:
    workbook = _vertical_workbook(
        tmp_path,
        "distinct_literals.xlsx",
        {1: "=1", 2: "=2", 3: "=1", 4: "=3"},
    )
    modules = generate_inverted(workbook, _series_bindings("Data!B1:B4"), force_rung=3)
    compute = _instance_compute_source(modules["_kernels.py"])
    assert compute.count("if i <") == 3
    assert "if i < 2:" in compute
    pkg = load_package(modules, tmp_path, name="distinct_lits")
    assert [pkg.compute_result()[year] for year in (2001, 2002, 2003, 2004)] == pytest.approx(
        (1.0, 2.0, 1.0, 3.0)
    )


def test_adjacent_identical_nonliterals_merge_in_demand_dispatch(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "same_nonliteral.xlsx",
        {
            "Inputs": {
                "A1": 2001,
                "A2": 2002,
                "A3": 2003,
                "A4": 2004,
                "B1": 10.0,
                "B2": 20.0,
                "B3": 30.0,
                "B4": 40.0,
            },
            "Data": {
                "A1": 2001,
                "A2": 2002,
                "A3": 2003,
                "A4": 2004,
                "B1": "=Inputs!B1+TRUE",
                "B2": "=Inputs!B2+TRUE()",
                "B3": "=Inputs!B3+FALSE",
                "B4": "=Inputs!B4+1",
            },
        },
    )
    document = bindings_document(
        series_entry(
            "values",
            "Inputs!B1:B4",
            layout="series",
            direction="input",
            label_column="A",
        ),
        series_entry(
            "result",
            "Data!B1:B4",
            layout="series",
            direction="output",
            label_column="A",
        ),
    )
    modules = generate_inverted(workbook, document, force_rung=3)
    compute = _instance_compute_source(modules["_kernels.py"])
    assert "if i < 1:" not in compute
    assert "if i < 2:" in compute
    pkg = load_package(modules, tmp_path, name="same_nonlit_r3")
    got = pkg.compute_result(
        values=pkg.data.Values.from_nested(
            domain=pkg.data.VALUES_DOMAIN, values=(10.0, 20.0, 30.0, 40.0)
        )
    )
    assert [value for _, value in got.items()] == pytest.approx((11.0, 21.0, 30.0, 41.0))
    assert got[2001] == pytest.approx(11.0)
    assert got[2002] == pytest.approx(21.0)
    assert got[2003] == pytest.approx(30.0)


def test_adjacent_identical_nonliterals_merge_in_append_loop(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "same_nonliteral_append.xlsx",
        {
            "Inputs": {
                "A1": 10.0,
                "B1": 20.0,
                "C1": 30.0,
                "D1": 40.0,
                "A10": 1,
                "B10": 2,
                "C10": 3,
                "D10": 4,
            },
            "Engine": {
                "A1": 1,
                "B1": 2,
                "C1": 3,
                "D1": 4,
                "A2": "=Inputs!A1+TRUE",
                "B2": "=Inputs!B1+TRUE()",
                "C2": "=Inputs!C1+FALSE",
                "D2": "=Inputs!D1+1",
            },
        },
    )
    document = bindings_document(
        series_entry(
            "values",
            "Inputs!A1:D1",
            layout="series",
            direction="input",
            header_row=10,
        ),
        series_entry(
            "result",
            "Engine!A2:D2",
            layout="series",
            direction="output",
            header_row=1,
        ),
    )
    modules = generate_inverted(workbook, document)
    helper = _helper_body_source(modules["_kernels.py"])
    assert "elif i < 1:" not in helper
    assert "if i < 1:" not in helper
    assert "i < 2" in helper
    assert helper.count("except XlError") == 1
    pkg = load_package(modules, tmp_path, name="same_nonlit_r0")
    got = pkg.compute_result(
        values=pkg.data.Values.from_nested(
            domain=pkg.data.VALUES_DOMAIN, values=(10.0, 20.0, 30.0, 40.0)
        )
    )
    assert [value for _, value in got.items()] == pytest.approx((11.0, 21.0, 30.0, 41.0))


def test_append_loop_coalesces_interior_blanks(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "append_blanks.xlsx",
        {
            "Inputs": {
                "A1": 2.0,
                "B1": 3.0,
                "C1": 4.0,
                "D1": 5.0,
                "A10": 1,
                "B10": 2,
                "C10": 3,
                "D10": 4,
            },
            "Engine": {
                "A1": 1,
                "B1": 2,
                "C1": 3,
                "D1": 4,
                "A2": "=Inputs!A1*2",
                "D2": "=Inputs!D1*2",
            },
        },
    )
    document = bindings_document(
        series_entry(
            "values",
            "Inputs!A1:D1",
            layout="series",
            direction="input",
            header_row=10,
        ),
        series_entry(
            "result",
            "Engine!A2:D2",
            layout="series",
            direction="output",
            header_row=1,
        ),
    )
    catalog, _deps, _graph = inverted_graph_parts(workbook, document)
    assert catalog.get("result").hole_indices == (1, 2)
    modules = generate_inverted(workbook, document)
    helper = _helper_body_source(modules["_kernels.py"])
    assert helper.count("out.append(None)") == 1
    assert "i < 1" in helper
    assert "i < 3" in helper
    assert "i < 2" not in helper
    pkg = load_package(modules, tmp_path, name="append_blanks")
    got = pkg.compute_result(
        values=pkg.data.Values.from_nested(
            domain=pkg.data.VALUES_DOMAIN, values=(2.0, 3.0, 4.0, 5.0)
        )
    )
    assert got[1] == pytest.approx(4.0)
    assert got[2] is None
    assert (2,) in pkg.data.RESULT_DOMAIN
    assert got[3] is None
    assert (3,) in pkg.data.RESULT_DOMAIN
    assert got[4] == pytest.approx(10.0)


def test_raising_expressions_still_catch_xl_error(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "div_error.xlsx",
        {
            "Inputs": {
                "A1": 2001,
                "A2": 2002,
                "A3": 2003,
                "A4": 2004,
                "B1": 0.0,
                "B2": 2.0,
                "B3": 0.0,
                "B4": 4.0,
            },
            "Data": {
                "A1": 2001,
                "A2": 2002,
                "A3": 2003,
                "A4": 2004,
                "B1": "=1/Inputs!B1",
                "B2": "=1/Inputs!B2",
                "B3": "=1/Inputs!B3",
                "B4": "=1",
            },
        },
    )
    document = bindings_document(
        series_entry(
            "denoms",
            "Inputs!B1:B4",
            layout="series",
            direction="input",
            label_column="A",
        ),
        series_entry(
            "result",
            "Data!B1:B4",
            layout="series",
            direction="output",
            label_column="A",
        ),
    )
    modules = generate_inverted(workbook, document, force_rung=3)
    compute = _instance_compute_source(modules["_kernels.py"])
    assert "except XlError" in compute
    pkg = load_package(modules, tmp_path, name="div_err_r3")
    got = pkg.compute_result(
        denoms=pkg.data.Denoms.from_nested(
            domain=pkg.data.DENOMS_DOMAIN, values=(0.0, 2.0, 0.0, 4.0)
        )
    )
    assert got[2001] == "#DIV/0!"
    assert got[2002] == pytest.approx(0.5)
    assert got[2003] == "#DIV/0!"
    assert got[2004] == pytest.approx(1.0)

    modules_r0 = generate_inverted(workbook, document)
    helper = _helper_body_source(modules_r0["_kernels.py"])
    assert helper.count("except XlError") == 1
    pkg_r0 = load_package(modules_r0, tmp_path, name="div_err_r0")
    got_r0 = pkg_r0.compute_result(
        denoms=pkg_r0.data.Denoms.from_nested(
            domain=pkg_r0.data.DENOMS_DOMAIN, values=(0.0, 2.0, 0.0, 4.0)
        )
    )
    assert got_r0[2001] == "#DIV/0!"
    assert got_r0[2002] == pytest.approx(0.5)
    assert got_r0[2003] == "#DIV/0!"
    assert got_r0[2004] == pytest.approx(1.0)


def test_literal_demand_dispatch_omits_error_wrapper(tmp_path: Path) -> None:
    workbook = _vertical_workbook(
        tmp_path,
        "all_literals.xlsx",
        {1: "=1", 2: None, 3: None, 4: "=2"},
    )
    modules = generate_inverted(workbook, _series_bindings("Data!B1:B4"), force_rung=3)
    compute = _instance_compute_source(modules["_kernels.py"])
    assert "try:" not in compute
    assert "except XlError" not in compute
    pkg = load_package(modules, tmp_path, name="all_lits")
    assert pkg.compute_result()[2001] == pytest.approx(1.0)
    assert pkg.compute_result()[2002] is None
    assert pkg.compute_result()[2003] is None
    assert pkg.compute_result()[2004] == pytest.approx(2.0)
