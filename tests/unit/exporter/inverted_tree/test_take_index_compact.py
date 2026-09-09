"""Compact and intern catalog gather mappings for `take` (#798)."""

from __future__ import annotations

import ast
import timeit
from pathlib import Path
from typing import Any

import pytest
from fastpyxl.utils.cell import get_column_letter

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.ast_emit import KeyedReadIntern
from excel_grapher.exporter.inverted_tree.domains import DomainEmitPlan
from excel_grapher.exporter.inverted_tree.schedule import (
    IndexSourceIntern,
    indices_to_source,
)
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

_REPEAT_N = 10
_REPEAT_K = 21
_BLOCK_STRIDE = 27
_BLOCK_LEN = 21


def _eval_index_source(source: str) -> tuple[int, ...]:
    return tuple(eval(source, {"range": range}))


def _flat_tuple_source(indices: tuple[int, ...]) -> str:
    if not indices:
        return "()"
    if len(indices) == 1:
        return f"({indices[0]},)"
    return f"({', '.join(str(i) for i in indices)})"


def _repeat_indices(n: int = _REPEAT_N, k: int = _REPEAT_K) -> tuple[int, ...]:
    return tuple(i for i in range(n) for _ in range(k))


def _block_indices(
    n: int = _REPEAT_N, block: int = _BLOCK_LEN, stride: int = _BLOCK_STRIDE
) -> tuple[int, ...]:
    return tuple(stride * i + j for i in range(n) for j in range(block))


def test_index_source_intern_reuses_compact_mappings() -> None:
    intern = IndexSourceIntern(prefix="_TAKE_", inline_max_chars=80)
    repeated = _repeat_indices()
    blocks = _block_indices()
    assert intern.expr(repeated) == "_TAKE_0"
    assert intern.expr(repeated) == "_TAKE_0"
    assert intern.expr(blocks) == "_TAKE_1"
    assert intern.expr((22, 47, 71)) == "(22, 47, 71)"
    assert intern.expr(range(0, 6, 2)) == "range(0, 6, 2)"
    lines = intern.emit_lines()
    assert lines[0] == "_TAKE_0 = tuple(i for i in range(0, 10) for _ in range(21))"
    assert lines[1] == "_TAKE_1 = tuple(27 * i + j for i in range(10) for j in range(21))"
    assert _eval_index_source(lines[0].split(" = ", 1)[1]) == repeated
    assert _eval_index_source(lines[1].split(" = ", 1)[1]) == blocks


def test_keyed_read_intern_emits_compact_slot_tables() -> None:
    plan = DomainEmitPlan(
        field_domains={},
        interned=(),
        series_expr={},
        series_key={},
        scc_expr={},
        scc_key={},
    )
    intern = KeyedReadIntern(plan)
    repeated = _repeat_indices()
    assert intern.slots(repeated) == "_SLOTS_0"
    assert intern.slots(repeated) == "_SLOTS_0"
    lines = intern.emit_lines()
    assert "_SLOTS_0 = tuple(i for i in range(0, 10) for _ in range(21))" in lines


def test_compact_index_source_cost(capsys: pytest.CaptureFixture[str]) -> None:
    cases = {
        "repeat": _repeat_indices(),
        "blocks": _block_indices(),
        "strided": tuple(range(0, 200, 2)),
        "irregular": (0, 2, 5, 5, 1),
    }
    for name, indices in cases.items():
        source = indices_to_source(indices)
        flat = _flat_tuple_source(indices)
        compile_time = timeit.timeit(lambda src=source: compile(src, "<idx>", "eval"), number=200)
        eval_time = timeit.timeit(lambda src=source: _eval_index_source(src), number=200)
        print(
            f"{name}: n={len(indices)} source={len(source)} flat={len(flat)} "
            f"compile={compile_time:.6f}s eval={eval_time:.6f}s {source[:80]}"
        )
        assert _eval_index_source(source) == indices
        if name != "irregular":
            assert len(source) < len(flat) / 4
    captured = capsys.readouterr()
    assert "repeat:" in captured.out
    assert "blocks:" in captured.out


def _align_workbook(
    tmp_path: Path,
    *,
    producer_n: int,
    host_map: tuple[int, ...],
    name: str,
) -> Path:
    """Each host cell `i` reads `growth[host_map[i]] + interest[host_map[i]]`."""
    inputs: dict[str, object] = {}
    for index in range(producer_n):
        col = get_column_letter(index + 2)
        inputs[f"{col}1"] = 2020 + index
        inputs[f"{col}2"] = float(index)
        inputs[f"{col}3"] = float(index + 100)
    outputs: dict[str, object] = {}
    for host_i, prod_i in enumerate(host_map):
        dest = get_column_letter(host_i + 1)
        src = get_column_letter(prod_i + 2)
        outputs[f"{dest}1"] = f"C{host_i}"
        outputs[f"{dest}2"] = f"=Inputs!{src}2+Inputs!{src}3"
    return write_workbook(tmp_path / f"{name}.xlsx", {"Inputs": inputs, "Out": outputs})


def _align_bindings(producer_n: int, host_n: int) -> dict[str, Any]:
    producer_last = get_column_letter(producer_n + 1)
    host_last = get_column_letter(host_n)
    return bindings_document(
        series_entry(
            "growth",
            f"Inputs!B2:{producer_last}2",
            layout="series",
            direction="input",
            header_row=1,
        ),
        series_entry(
            "interest",
            f"Inputs!B3:{producer_last}3",
            layout="series",
            direction="input",
            header_row=1,
        ),
        series_entry(
            "result",
            f"Out!A2:{host_last}2",
            layout="series",
            direction="output",
            header_row=1,
            key_concept="COUNTRY",
            key_read="string",
            compute_name="compute_result",
        ),
    )


def _function_bodies_source(source: str) -> str:
    module = ast.parse(source)
    chunks: list[str] = []
    lines = source.splitlines()
    for node in module.body:
        if isinstance(node, ast.FunctionDef):
            chunks.append("\n".join(lines[node.lineno - 1 : node.end_lineno]))
    return "\n".join(chunks)


def test_api_interns_repeated_value_take_mappings(tmp_path: Path) -> None:
    n, copies = 8, 6
    host_map = tuple(i for i in range(n) for _ in range(copies))
    workbook = _align_workbook(tmp_path, producer_n=n, host_map=host_map, name="repeat_values")
    document = _align_bindings(n, len(host_map))
    modules = generate_inverted(workbook, document)
    api = modules["_kernel.py"]
    compact = indices_to_source(host_map)
    assert compact == "tuple(i for i in range(0, 8) for _ in range(6))"
    assert f"_TAKE_0 = {compact}" in api
    assert api.count("_TAKE_0") >= 3
    assert _flat_tuple_source(host_map) not in api
    bodies = _function_bodies_source(api)
    assert "for _ in range(6)" not in bodies

    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    pkg = load_package(modules, tmp_path, name="repeat_values")
    cells = [f"Out!{get_column_letter(i + 1)}2" for i in range(len(host_map))]
    expected_values = FormulaEvaluator(
        create_dependency_graph(workbook, cells, load_values=True)
    ).evaluate(cells)
    got = pkg.compute_result(**named_input_kwargs(pkg, catalog, graph))
    assert tuple(value for _, value in got.items()) == pytest.approx(
        tuple(expected_values[cell] for cell in cells)
    )


def test_api_interns_identical_compact_take_mappings(tmp_path: Path) -> None:
    n, copies = 8, 6
    host_map = tuple(j for _ in range(copies) for j in range(n))
    workbook = _align_workbook(tmp_path, producer_n=n, host_map=host_map, name="repeat_take")
    document = _align_bindings(n, len(host_map))
    modules = generate_inverted(workbook, document)
    api = modules["_kernel.py"]
    compact = indices_to_source(host_map)
    assert compact == "tuple(range(0, 8)) * 6"
    assert f"_TAKE_0 = {compact}" in api
    assert api.count("_TAKE_0") >= 3
    assert api.count(compact) == 1
    assert _flat_tuple_source(host_map) not in api
    bodies = _function_bodies_source(api)
    assert "tuple(range(0, 8)) * 6" not in bodies
    assert "tuple(i for i in" not in bodies

    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    pkg = load_package(modules, tmp_path, name="repeat_take")
    cells = [f"Out!{get_column_letter(i + 1)}2" for i in range(len(host_map))]
    expected_values = FormulaEvaluator(
        create_dependency_graph(workbook, cells, load_values=True)
    ).evaluate(cells)
    got = pkg.compute_result(**named_input_kwargs(pkg, catalog, graph))
    assert tuple(value for _, value in got.items()) == pytest.approx(
        tuple(expected_values[cell] for cell in cells)
    )


def test_api_compacts_strided_block_take_mappings(tmp_path: Path) -> None:
    block, groups, stride = 3, 6, 4
    host_map = tuple(stride * i + j for i in range(groups) for j in range(block))
    producer_n = stride * (groups - 1) + block
    workbook = _align_workbook(
        tmp_path, producer_n=producer_n, host_map=host_map, name="block_take"
    )
    document = _align_bindings(producer_n, len(host_map))
    modules = generate_inverted(workbook, document)
    api = modules["_kernel.py"]
    compact = indices_to_source(host_map)
    assert compact == "tuple(4 * i + j for i in range(6) for j in range(3))"
    assert _eval_index_source(compact) == host_map
    assert compact in api or f"_TAKE_0 = {compact}" in api
    assert _flat_tuple_source(host_map) not in api
    bodies = _function_bodies_source(api)
    assert "for i in range(6) for j in range(3)" not in bodies

    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    pkg = load_package(modules, tmp_path, name="block_take")
    cells = [f"Out!{get_column_letter(i + 1)}2" for i in range(len(host_map))]
    expected_values = FormulaEvaluator(
        create_dependency_graph(workbook, cells, load_values=True)
    ).evaluate(cells)
    got = pkg.compute_result(**named_input_kwargs(pkg, catalog, graph))
    assert tuple(value for _, value in got.items()) == pytest.approx(
        tuple(expected_values[cell] for cell in cells)
    )


def test_api_keeps_irregular_gather_order(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "irregular_take.xlsx",
        {
            "Inputs": {
                "B1": 2020,
                "C1": 2021,
                "D1": 2022,
                "E1": 2023,
                "B2": 10.0,
                "C2": 20.0,
                "D2": 30.0,
                "E2": 40.0,
            },
            "Engine": {
                "B1": 2020,
                "C1": 2021,
                "D1": 2022,
                "E1": 2023,
                "B2": "=Inputs!B2+1",
                "C2": "=Inputs!C2+1",
                "D2": "=Inputs!D2+1",
                "E2": "=Inputs!E2+1",
            },
            "Out": {
                "A1": 2020,
                "B1": 2022,
                "C1": 2021,
                "A2": "=Engine!B2",
                "B2": "=Engine!D2",
                "C2": "=Engine!C2",
            },
        },
    )
    document = bindings_document(
        series_entry("values", "Inputs!B2:E2", layout="series", direction="input", header_row=1),
        series_entry(
            "engine_plus",
            "Engine!B2:E2",
            layout="series",
            direction="internal",
            header_row=1,
        ),
        series_entry(
            "result",
            "Out!A2:C2",
            layout="series",
            direction="output",
            header_row=1,
            compute_name="compute_result",
        ),
    )
    modules = generate_inverted(workbook, document)
    api = modules["_kernel.py"]
    assert "take(engine_plus, (0, 2, 1))" in api
    assert "take(engine_plus, (0, 1, 2))" not in api
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    pkg = load_package(modules, tmp_path, name="irregular_take")
    cells = ["Out!A2", "Out!B2", "Out!C2"]
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, cells, load_values=True)
    ).evaluate(cells)
    got = pkg.compute_result(**named_input_kwargs(pkg, catalog, graph))
    assert dict(got.items()) == pytest.approx(
        {
            coordinate: expected[cell]
            for coordinate, cell in catalog.get("result").coordinate_cells.items()
        }
    )
    assert tuple(value for _, value in got.items()) == pytest.approx((11.0, 31.0, 21.0))


def test_generated_api_import_and_eval_cost(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    n, copies = 10, 8
    host_map = tuple(j for _ in range(copies) for j in range(n))
    workbook = _align_workbook(tmp_path, producer_n=n, host_map=host_map, name="take_cost")
    modules = generate_inverted(workbook, _align_bindings(n, len(host_map)))
    api = modules["_kernel.py"]
    compile_time = timeit.timeit(lambda src=api: compile(src, "<api>", "exec"), number=20)
    interned = [line for line in api.splitlines() if line.startswith("_TAKE_")]
    print(
        f"api_bytes={len(api.encode())} interned={len(interned)} "
        f"compile={compile_time:.6f}s {interned[:2]}"
    )
    assert interned
    assert compile_time < 1.0
    captured = capsys.readouterr()
    assert "api_bytes=" in captured.out
