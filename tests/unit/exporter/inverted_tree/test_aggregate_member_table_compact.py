"""Issue 795 — compact repeated and patterned aggregate member tables."""

from __future__ import annotations

import timeit
from pathlib import Path
from typing import Any, Literal

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.ast_emit import KeyedReadIntern
from excel_grapher.exporter.inverted_tree.domains import DomainEmitPlan
from excel_grapher.exporter.inverted_tree.schedule import slot_table_to_source
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    named_input_kwargs,
    series_entry,
    write_workbook,
)

_REPEAT_PADDING = ((),) * 8 + ((0, 1), (1, 2))
_RANGE_FAMILY = tuple(tuple(range(k, 9 + k, 3)) for k in range(3))
_IRREGULAR_FAMILY = (
    tuple(range(0, 9, 3)),
    tuple(range(1, 10, 3)),
    tuple(range(2, 11, 3)),
    (9, 10),
    tuple(range(4, 13, 3)),
    tuple(range(5, 14, 3)),
    tuple(range(6, 15, 3)),
)


def _materialize(source: str) -> tuple[tuple[int, ...], ...]:
    namespace = {"range": range, "tuple": tuple}
    table = eval(source, namespace, namespace)  # noqa: S307
    return tuple(tuple(row) for row in table)


def test_slot_table_repeats_identical_immutable_rows() -> None:
    source = slot_table_to_source(_REPEAT_PADDING)
    assert "(), (), ()" not in source
    assert "((),) * 8" in source
    assert _materialize(source) == _REPEAT_PADDING


def test_slot_table_emits_range_family_comprehension() -> None:
    source = slot_table_to_source(_RANGE_FAMILY)
    assert "for k in range(3)" in source
    assert "range(0, 9, 3), range(1, 10, 3)" not in source
    assert _materialize(source) == _RANGE_FAMILY


def test_slot_table_keeps_irregular_exception_in_range_family() -> None:
    source = slot_table_to_source(_IRREGULAR_FAMILY)
    assert "range(9, 11)" in source
    assert "for k in range(3)" in source
    assert _materialize(source) == _IRREGULAR_FAMILY


def test_slot_table_keeps_short_literal_display() -> None:
    rows = ((0, 1), (1, 2))
    source = slot_table_to_source(rows)
    assert source == "(range(0, 2), range(1, 3))"
    assert _materialize(source) == rows


def test_keyed_read_intern_reuses_constructed_member_tables() -> None:
    plan = DomainEmitPlan(
        field_domains={},
        interned=(),
        series_expr={},
        series_key={},
        scc_expr={},
        scc_key={},
    )
    intern = KeyedReadIntern(plan)
    compact = intern.member_slots(_REPEAT_PADDING)
    assert compact == intern.member_slots(_REPEAT_PADDING)
    assert compact == "_MEMBERS_0"
    assert intern.member_slots(((0, 1), (1, 2))) == "(range(0, 2), range(1, 3))"
    lines = intern.emit_lines()
    assignment = next(line for line in lines if line.startswith("_MEMBERS_0 = "))
    assert "((),) * 8" in assignment
    assert _materialize(assignment.split(" = ", 1)[1]) == _REPEAT_PADDING


def _holes_workbook(tmp_path: Path, size: int) -> Path:
    cells: dict[str, object] = {}
    for row in range(1, size + 1):
        cells[f"A{row}"] = 2000 + row
    for row in range(1, size - 1):
        cells[f"B{row}"] = "=1"
    cells[f"B{size - 1}"] = "=SUM(B1:B2)"
    cells[f"B{size}"] = "=SUM(B2:B3)"
    return write_workbook(tmp_path / f"holes_{size}.xlsx", {"Data": cells})


def _holes_bindings(size: int) -> dict[str, Any]:
    return bindings_document(
        series_entry(
            "result",
            f"Data!B1:B{size}",
            layout="series",
            direction="output",
            label_column="A",
            compute_name="compute_result",
        )
    )


def _expected_holes(size: int) -> tuple[float, ...]:
    return tuple(1.0 for _ in range(size - 2)) + (2.0, 2.0)


@pytest.mark.parametrize("force_rung", [None, 3])
def test_statement_local_member_table_drops_host_padding(
    tmp_path: Path, force_rung: Literal[3] | None
) -> None:
    size = 10
    workbook = _holes_workbook(tmp_path, size)
    modules = generate_inverted(workbook, _holes_bindings(size), force_rung=force_rung)
    internals = modules["_kernels.py"]
    assert "(), (), ()" not in internals
    assert "((),) *" not in internals
    assert "range(0, 2)" in internals
    assert "range(1, 3)" in internals
    assert "[i - 8]" in internals
    pkg = load_package(modules, tmp_path, name=f"holes_local_{force_rung}")
    assert [pkg.compute_result()[2000 + row] for row in range(1, size + 1)] == pytest.approx(
        _expected_holes(size)
    )


def test_padded_member_tables_scale_with_source_and_import_cost(tmp_path: Path) -> None:
    sizes = (10, 100, 1000)
    internals_sizes: list[int] = []
    compile_times: list[float] = []
    for size in sizes:
        workbook = _holes_workbook(tmp_path, size)
        modules = generate_inverted(workbook, _holes_bindings(size), force_rung=3)
        internals = modules["_kernels.py"]
        internals_sizes.append(len(internals.encode()))
        compile_times.append(
            timeit.timeit(lambda src=internals: compile(src, "<internals>", "exec"), number=20)
        )
        pkg = load_package(modules, tmp_path, name=f"holes_scale_{size}")
        assert [pkg.compute_result()[2000 + row] for row in range(1, size + 1)] == pytest.approx(
            _expected_holes(size)
        )
        assert "(), (), ()" not in internals
    assert internals_sizes[-1] / internals_sizes[0] < 2.5
    assert compile_times[-1] / compile_times[0] < 3.0


def test_statement_local_member_table_matches_evaluator(tmp_path: Path) -> None:
    size = 12
    workbook = _holes_workbook(tmp_path, size)
    document = _holes_bindings(size)
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    pkg = load_package(
        generate_inverted(workbook, document, force_rung=3),
        tmp_path,
        name="holes_parity",
    )
    cells = [f"Data!B{row}" for row in range(1, size + 1)]
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, cells, load_values=True)
    ).evaluate(cells)
    got = pkg.compute_result(**named_input_kwargs(pkg, catalog, graph))
    assert [value for _, value in got.items()] == pytest.approx(
        tuple(expected[cell] for cell in cells)
    )
    assert [value for _, value in got.items()] == pytest.approx(_expected_holes(size))


def _dense_columns_workbook(tmp_path: Path, columns: int, *, punch: int | None) -> Path:
    cells: dict[str, object] = {}
    for col_offset in range(columns):
        col = chr(ord("B") + col_offset)
        cells[f"{col}1"] = 2020 + col_offset
        cells[f"{col}2"] = float(col_offset + 1)
        cells[f"{col}3"] = float((col_offset + 1) * 10)
        cells[f"{col}4"] = float((col_offset + 1) * 100)
        if punch == col_offset:
            cells[f"{col}6"] = f"=SUM({col}2:{col}3)"
        else:
            cells[f"{col}6"] = f"=SUM({col}2:{col}4)"
    cells["A2"] = "a"
    cells["A3"] = "b"
    cells["A4"] = "c"
    return write_workbook(tmp_path / f"dense_cols_{columns}_{punch}.xlsx", {"Engine": cells})


def _grid_entry(data_range: str) -> dict[str, Any]:
    return {
        "id": "vintages",
        "sheet": "Engine",
        "data_range": data_range,
        "layout": "series",
        "constant": {},
        "key": ["COUNTRY", "TIME_PERIOD"],
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [
                {
                    "id": "TIME_PERIOD",
                    "concept": "TIME_PERIOD",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
                },
                {
                    "id": "COUNTRY",
                    "concept": "COUNTRY",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "row_label", "label_column": "A", "read": "string"},
                },
            ],
        },
    }


def _dense_bindings(columns: int) -> dict[str, Any]:
    last = chr(ord("B") + columns - 1)
    return bindings_document(
        _grid_entry(f"Engine!B2:{last}4"),
        series_entry(
            "totals",
            f"Engine!B6:{last}6",
            layout="series",
            direction="output",
            header_row=1,
            compute_name="compute_totals",
        ),
    )


def _dense_expected(columns: int, *, punch: int | None) -> tuple[float, ...]:
    values: list[float] = []
    for col_offset in range(columns):
        base = float(col_offset + 1)
        total = base + base * 10
        if punch != col_offset:
            total += base * 100
        values.append(total)
    return tuple(values)


def test_regular_member_rows_use_range_comprehension(tmp_path: Path) -> None:
    columns = 11
    workbook = _dense_columns_workbook(tmp_path, columns, punch=None)
    modules = generate_inverted(workbook, _dense_bindings(columns), force_rung=3)
    internals = modules["_kernels.py"]
    assert "for k in range(11)" in internals
    assert "range(0, 9, 3), range(1, 10, 3)" not in internals
    pkg = load_package(modules, tmp_path, name="dense_family")
    assert [pkg.compute_totals()[2020 + col] for col in range(columns)] == pytest.approx(
        _dense_expected(columns, punch=None)
    )


def test_irregular_exception_in_member_table_keeps_value_parity(tmp_path: Path) -> None:
    columns = 5
    punch = 2
    workbook = _dense_columns_workbook(tmp_path, columns, punch=punch)
    document = _dense_bindings(columns)
    modules = generate_inverted(workbook, document, force_rung=3)
    internals = modules["_kernels.py"]
    pkg = load_package(modules, tmp_path, name="dense_irregular")
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    cells = [f"Engine!{chr(ord('B') + col)}6" for col in range(columns)]
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, cells, load_values=True)
    ).evaluate(cells)
    got = pkg.compute_totals(**named_input_kwargs(pkg, catalog, graph))
    assert [value for _, value in got.items()] == pytest.approx(
        tuple(expected[cell] for cell in cells)
    )
    assert [value for _, value in got.items()] == pytest.approx(
        _dense_expected(columns, punch=punch)
    )
    assert "(), (), ()" not in internals
