"""Reuse lookup tables and identical eager lookups in generated internals (#796)."""

from __future__ import annotations

import ast
import timeit
import tracemalloc
from pathlib import Path
from typing import Any

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    call_compute,
    generate_inverted,
    input_kwargs,
    inverted_graph_parts,
    load_package,
    write_workbook,
)


def _country_dim() -> dict[str, Any]:
    return {
        "id": "COUNTRY",
        "concept": "COUNTRY",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "row_label", "label_column": "A", "read": "string"},
    }


def _measure() -> dict[str, Any]:
    return {
        "concept": "OBS_VALUE",
        "dtype": "float",
        "bind": {"kind": "data_cell", "read": "float"},
    }


def _series_entry(
    name: str,
    area: str,
    direction: str,
) -> dict[str, Any]:
    entry: dict[str, Any] = {
        "id": name,
        "sheet": "Data",
        "data_range": f"Data!{area}",
        "layout": "series",
        "key": ["COUNTRY"],
        "structure": {
            "dimensions": [_country_dim()],
            "measure": _measure(),
        },
    }
    if direction == "output":
        entry["output"] = {"compute": {"name": f"compute_{name}"}}
    else:
        entry["input"] = {"setter": {"name": f"set_{name}"}}
    return entry


def _lookup_workbook(tmp_path: Path, size: int, *, formula: str | None = None) -> Path:
    cells: dict[str, object] = {"B1": 2020, "C1": 2021}
    for row in range(2, size + 2):
        cells[f"A{row}"] = f"Country {row}"
        cells[f"B{row}"] = float(row)
        cells[f"C{row}"] = float(row * 2)
        cells[f"D{row}"] = float(row)
        table = f"$B$2:$C${size + 1}"
        cells[f"E{row}"] = (
            formula.format(row=row, table=table)
            if formula
            else (f"=VLOOKUP(D{row},{table},2,FALSE)+VLOOKUP(D{row},{table},2,FALSE)")
        )
    return write_workbook(tmp_path / f"lookup_reuse_{size}.xlsx", {"Data": cells})


def _lookup_bindings(size: int) -> dict[str, Any]:
    return bindings_document(
        _series_entry("left_values", f"B2:B{size + 1}", "input"),
        _series_entry("right_values", f"C2:C{size + 1}", "input"),
        _series_entry("selectors", f"D2:D{size + 1}", "input"),
        _series_entry("result", f"E2:E{size + 1}", "output"),
    )


def _helper_ast(internals: str, name: str = "result") -> ast.FunctionDef:
    module = ast.parse(internals)
    for node in module.body:
        if isinstance(node, ast.FunctionDef) and node.name == name:
            return node
    raise AssertionError(f"helper {name!r} not found")


def _table_assigns_outside_loops(fn: ast.FunctionDef) -> list[str]:
    """Return `_TABLE_*` assignment sources that are not inside a `for`/`while`."""
    loop_nodes = tuple(node for node in ast.walk(fn) if isinstance(node, (ast.For, ast.While)))
    names: list[str] = []
    for stmt in fn.body:
        if isinstance(stmt, ast.AnnAssign) and isinstance(stmt.target, ast.Name):
            continue
        if not isinstance(stmt, ast.Assign):
            continue
        for target in stmt.targets:
            if isinstance(target, ast.Name) and target.id.startswith("_TABLE_"):
                nested = any(stmt in ast.walk(loop) for loop in loop_nodes)
                assert not nested, f"{target.id} is assigned inside a loop"
                names.append(target.id)
    return names


def _expected_hits(size: int) -> tuple[float, ...]:
    return tuple(float(row * 2 + row * 2) for row in range(2, size + 2))


def test_repeated_vlookup_table_is_zipped_and_hoisted_once(tmp_path: Path) -> None:
    size = 10
    workbook = _lookup_workbook(tmp_path, size)
    internals = generate_inverted(workbook, _lookup_bindings(size))["internals.py"]
    assert internals.count("xl_vlookup(") == 1
    assert "left_values[0]" not in internals
    assert "tuple(zip(" in internals
    helper = _helper_ast(internals)
    tables = _table_assigns_outside_loops(helper)
    assert tables == ["_TABLE_0"]
    for_nodes = [node for node in ast.walk(helper) if isinstance(node, ast.For)]
    assert for_nodes
    loop_src = ast.unparse(for_nodes[0])
    assert "tuple(zip(" not in loop_src
    assert loop_src.count("xl_vlookup(") == 1


def test_lookup_table_source_does_not_copy_per_row(tmp_path: Path) -> None:
    internals_sizes: list[int] = []
    compile_times: list[float] = []
    for size in (10, 40):
        workbook = _lookup_workbook(tmp_path, size)
        internals = generate_inverted(workbook, _lookup_bindings(size))["internals.py"]
        internals_sizes.append(len(internals.encode()))
        compile_times.append(
            timeit.timeit(lambda src=internals: compile(src, "<internals>", "exec"), number=20)
        )
        assert internals.count("xl_vlookup(") == 1
        assert internals.count("tuple(zip(") == 1
        assert internals.count("left_values[0]") == 0
    assert internals_sizes[-1] / internals_sizes[0] < 1.5
    assert compile_times[-1] / compile_times[0] < 2.5


def test_repeated_vlookup_matches_evaluator_and_tracks_inputs(tmp_path: Path) -> None:
    size = 8
    workbook = _lookup_workbook(tmp_path, size)
    document = _lookup_bindings(size)
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="lookup_reuse")
    cells = [f"Data!E{row}" for row in range(2, size + 2)]
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, cells, load_values=True)
    ).evaluate(cells)
    kwargs = input_kwargs(catalog, graph)
    got = call_compute(pkg, "result", kwargs)
    assert got == pytest.approx(tuple(expected[cell] for cell in cells))
    assert got == pytest.approx(_expected_hits(size))

    swapped = dict(kwargs)
    swapped["right_values"] = tuple(float(row * 3) for row in range(2, size + 2))
    got_swapped = call_compute(pkg, "result", swapped)
    assert got_swapped == pytest.approx(tuple(float(row * 3 * 2) for row in range(2, size + 2)))
    assert got_swapped != pytest.approx(got)


def test_lookup_miss_and_error_propagate(tmp_path: Path) -> None:
    size = 4
    workbook = _lookup_workbook(tmp_path, size)
    document = _lookup_bindings(size)
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="lookup_miss")
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    kwargs = input_kwargs(catalog, graph)
    selectors = kwargs["selectors"]
    assert isinstance(selectors, tuple)
    missing = dict(kwargs)
    missing["selectors"] = (99.0, *selectors[1:])
    got_miss = call_compute(pkg, "result", missing)
    assert isinstance(got_miss, tuple)
    assert got_miss[0] == "#N/A"
    assert got_miss[1:] == pytest.approx(_expected_hits(size)[1:])

    errors = dict(kwargs)
    errors["selectors"] = ("#DIV/0!", *selectors[1:])
    got_err = call_compute(pkg, "result", errors)
    assert isinstance(got_err, tuple)
    assert got_err[0] == "#DIV/0!"
    assert got_err[1:] == pytest.approx(_expected_hits(size)[1:])


def test_if_unused_lookup_branch_is_not_evaluated(tmp_path: Path) -> None:
    size = 4
    formula = "=IF(TRUE,VLOOKUP(D{row},{table},2,FALSE),VLOOKUP(99,{table},2,FALSE))"
    workbook = _lookup_workbook(tmp_path, size, formula=formula)
    document = _lookup_bindings(size)
    internals = generate_inverted(workbook, document)["internals.py"]
    assert internals.count("xl_vlookup(") == 2
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="lookup_if")
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    got = call_compute(pkg, "result", input_kwargs(catalog, graph))
    assert got == pytest.approx(tuple(float(row * 2) for row in range(2, size + 2)))


def test_iferror_miss_does_not_escape_the_handler(tmp_path: Path) -> None:
    size = 4
    formula = "=IFERROR(VLOOKUP(99,{table},2,FALSE),0)"
    workbook = _lookup_workbook(tmp_path, size, formula=formula)
    document = _lookup_bindings(size)
    internals = generate_inverted(workbook, document)["internals.py"]
    assert "lambda:" in internals
    assert "_LOOKUP_" not in internals
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="lookup_iferror")
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    got = call_compute(pkg, "result", input_kwargs(catalog, graph))
    assert got == pytest.approx((0.0,) * size)


@pytest.mark.parametrize("force_rung", [None, 3])
def test_demand_driven_lookup_reuses_table(tmp_path: Path, force_rung: int | None) -> None:
    size = 6
    workbook = _lookup_workbook(tmp_path, size)
    document = _lookup_bindings(size)
    modules = generate_inverted(workbook, document, force_rung=force_rung)
    internals = modules["internals.py"]
    assert internals.count("tuple(zip(") == 1
    helper = _helper_ast(internals)
    assert _table_assigns_outside_loops(helper) == ["_TABLE_0"]
    pkg = load_package(modules, tmp_path, name=f"lookup_rung_{force_rung}")
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    got = call_compute(pkg, "result", input_kwargs(catalog, graph))
    assert got == pytest.approx(_expected_hits(size))


def test_lookup_table_allocation_is_once_per_call(tmp_path: Path) -> None:
    size = 40
    workbook = _lookup_workbook(tmp_path, size)
    document = _lookup_bindings(size)
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="lookup_alloc")
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    kwargs = input_kwargs(catalog, graph)
    call_compute(pkg, "result", kwargs)
    tracemalloc.start()
    got = call_compute(pkg, "result", kwargs)
    _current, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    assert got == pytest.approx(_expected_hits(size))
    # One 40-row pair table plus the output tuple; far below a per-row copy.
    assert peak < 80_000
