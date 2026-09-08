"""Issue 786 — intern keyed domains instead of repeating `tuple.index` literals."""

from __future__ import annotations

import timeit
from pathlib import Path
from typing import Any

import pytest

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.ast_emit import KeyedReadIntern
from excel_grapher.exporter.inverted_tree.domains import DomainEmitPlan
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
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


def _time_dim() -> dict[str, Any]:
    return {
        "id": "TIME_PERIOD",
        "concept": "TIME_PERIOD",
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
    }


def _measure() -> dict[str, Any]:
    return {
        "concept": "OBS_VALUE",
        "dtype": "float",
        "bind": {"kind": "data_cell", "read": "float"},
    }


def _domain_workbook(tmp_path: Path, size: int) -> Path:
    cells: dict[str, object] = {"B1": 2020, "C1": 2021}
    for row in range(2, size + 2):
        cells[f"A{row}"] = f"Country {row}"
        cells[f"B{row}"] = float(row)
        cells[f"C{row}"] = float(row * 2)
        cells[f"E{row}"] = f"=B{row}+C{row}"
    return write_workbook(tmp_path / f"keyed_domain_{size}.xlsx", {"Data": cells})


def _values_entry(size: int) -> dict[str, Any]:
    return {
        "id": "values",
        "sheet": "Data",
        "data_range": f"Data!B2:C{size + 1}",
        "layout": "matrix",
        "input": {"setter": {"name": "set_values"}},
        "key": ["COUNTRY", "TIME_PERIOD"],
        "structure": {
            "measure": _measure(),
            "dimensions": [_country_dim(), _time_dim()],
        },
    }


def _result_entry(size: int) -> dict[str, Any]:
    return {
        "id": "result",
        "sheet": "Data",
        "data_range": f"Data!E2:E{size + 1}",
        "layout": "series",
        "output": {"compute": {"name": "compute_result"}},
        "key": ["COUNTRY"],
        "structure": {
            "measure": _measure(),
            "dimensions": [_country_dim()],
        },
    }


def _domain_bindings(size: int) -> dict[str, Any]:
    return bindings_document(_values_entry(size), _result_entry(size))


def _producer_domain(size: int) -> tuple[tuple[str, int], ...]:
    return tuple((f"Country {row}", year) for row in range(2, size + 2) for year in (2020, 2021))


def _host_countries(size: int) -> tuple[str, ...]:
    return tuple(f"Country {row}" for row in range(2, size + 2))


def test_repeated_keyed_reads_do_not_embed_producer_domain(tmp_path: Path) -> None:
    size = 20
    workbook = _domain_workbook(tmp_path, size)
    document = _domain_bindings(size)
    catalog, deps, _graph = inverted_graph_parts(workbook, document)
    assert "values" in deps["result"].keyed_ids
    internals = generate_inverted(workbook, document)["internals.py"]
    domain = _producer_domain(size)
    assert internals.count(repr(domain)) <= 1
    assert internals.count(".index(") == 0
    countries = _host_countries(size)
    assert internals.count(repr(countries)) <= 1


def test_keyed_domain_source_does_not_copy_per_read(tmp_path: Path) -> None:
    sizes = (10, 20, 40)
    internals_sizes: list[int] = []
    compile_times: list[float] = []
    for size in sizes:
        workbook = _domain_workbook(tmp_path, size)
        internals = generate_inverted(workbook, _domain_bindings(size))["internals.py"]
        internals_sizes.append(len(internals.encode()))
        compile_times.append(
            timeit.timeit(lambda src=internals: compile(src, "<internals>", "exec"), number=20)
        )
        assert internals.count(repr(_producer_domain(size))) <= 1
        assert internals.count(".index(") == 0
    # Two-read formulas must not pay the producer domain once per read.
    # Affine / interned maps grow far slower than embedding the cartesian key
    # tuple at each site (the pre-fix 10→40 growth was ~2.8×).
    assert internals_sizes[-1] / internals_sizes[0] < 2.0
    assert compile_times[-1] / compile_times[0] < 3.0


def test_repeated_keyed_reads_match_evaluator(tmp_path: Path) -> None:
    size = 8
    workbook = _domain_workbook(tmp_path, size)
    document = _domain_bindings(size)
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="keyed_dedup")
    cells = [f"Data!E{row}" for row in range(2, size + 2)]
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, cells, load_values=True)
    ).evaluate(cells)
    got = pkg.compute_result(**input_kwargs(catalog, graph))
    assert got == pytest.approx(tuple(expected[cell] for cell in cells))
    assert got == pytest.approx(tuple(float(row + row * 2) for row in range(2, size + 2)))


def _shuffled_workbook(tmp_path: Path) -> Path:
    """Producer rows C2, C4, C3; host walks C2, C3, C4 (not affine)."""
    return write_workbook(
        tmp_path / "keyed_shuffled.xlsx",
        {
            "Data": {
                "B1": 2020,
                "C1": 2021,
                "A2": "Country 2",
                "B2": 2.0,
                "C2": 4.0,
                "A3": "Country 4",
                "B3": 4.0,
                "C3": 8.0,
                "A4": "Country 3",
                "B4": 3.0,
                "C4": 6.0,
            },
            "Out": {
                "A2": "Country 2",
                "E2": "=Data!B2+Data!C2",
                "A3": "Country 3",
                "E3": "=Data!B4+Data!C4",
                "A4": "Country 4",
                "E4": "=Data!B3+Data!C3",
            },
        },
    )


def _shuffled_bindings() -> dict[str, Any]:
    values = _values_entry(3)
    values["data_range"] = "Data!B2:C4"
    result = {
        "id": "result",
        "sheet": "Out",
        "data_range": "Out!E2:E4",
        "layout": "series",
        "output": {"compute": {"name": "compute_result"}},
        "key": ["COUNTRY"],
        "structure": {
            "measure": _measure(),
            "dimensions": [_country_dim()],
        },
    }
    return bindings_document(values, result)


def test_non_affine_keyed_reads_intern_slot_table_once(tmp_path: Path) -> None:
    workbook = _shuffled_workbook(tmp_path)
    document = _shuffled_bindings()
    catalog, deps, graph = inverted_graph_parts(workbook, document)
    assert "values" in deps["result"].keyed_ids
    modules = generate_inverted(workbook, document)
    internals = modules["internals.py"]
    domain = (
        ("Country 2", 2020),
        ("Country 2", 2021),
        ("Country 4", 2020),
        ("Country 4", 2021),
        ("Country 3", 2020),
        ("Country 3", 2021),
    )
    assert internals.count(repr(domain)) <= 1
    assert internals.count(".index(") == 0
    assert "4 * i" in internals
    assert "values[2]" in internals
    pkg = load_package(modules, tmp_path, name="keyed_shuffled")
    cells = ["Out!E2", "Out!E3", "Out!E4"]
    expected = FormulaEvaluator(
        create_dependency_graph(workbook, cells, load_values=True)
    ).evaluate(cells)
    got = pkg.compute_result(**input_kwargs(catalog, graph))
    assert got == pytest.approx(tuple(expected[cell] for cell in cells))
    assert got == pytest.approx((6.0, 9.0, 12.0))


def test_keyed_read_intern_reuses_sequences_and_maps() -> None:
    plan = DomainEmitPlan(
        field_domains={"COUNTRY": ("A", "B")},
        interned=(),
        series_expr={"values": "data.COUNTRY_DOMAIN"},
        series_key={"values": ("COUNTRY",)},
        scc_expr={},
        scc_key={},
    )
    intern = KeyedReadIntern(plan)
    assert intern.sequence(("A", "B"), field="COUNTRY") == "data.COUNTRY_DOMAIN"
    assert intern.sequence(("X", "Y")) == "_KEYS_0"
    assert intern.sequence(("X", "Y")) == "_KEYS_0"
    assert intern.slot_map("values") == "_INDEX_values"
    assert intern.slot_map("values") == "_INDEX_values"
    assert intern.slots((0, 4, 2)) == "_SLOTS_0"
    assert intern.slots((0, 4, 2)) == "_SLOTS_0"
    lines = intern.emit_lines()
    assert "_KEYS_0 = ('X', 'Y')" in lines
    assert "_SLOTS_0 = (0, 4, 2)" in lines
    assert any(line.startswith("_INDEX_values = ") for line in lines)
    assert intern.uses_data
