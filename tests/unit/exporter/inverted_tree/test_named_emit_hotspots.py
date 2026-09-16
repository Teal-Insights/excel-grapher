"""Named emit must not re-lower identical range tables per host cell (#867).

Uniform formula families replay `_emit_range_table` once per statement, not
once per catalog member. Range-vs-blank tests are geometric, so unrelated
blank rectangles stay off the per-cell hot path. Lockstep producer slots
are cached per `(host, producer)`.
"""

from __future__ import annotations

from collections.abc import Sequence
from pathlib import Path
from typing import Any

import pytest
from fastpyxl.utils.cell import get_column_letter

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree import ast_emit as ast_emit_mod
from excel_grapher.exporter.inverted_tree.ast_emit import (
    EmitContext,
    _lockstep_producer_slots,
)
from excel_grapher.exporter.inverted_tree.catalog import BoundSeries, KeyPoint, Statement
from excel_grapher.exporter.inverted_tree.deps import DependenceEdge, SeriesDeps
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    call_compute,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    make_catalog,
    named_input_kwargs,
    series_entry,
    write_workbook,
)


def _matrix_series(
    series_id: str,
    area: str,
    direction: str,
    *,
    compute: str | None = None,
) -> dict[str, Any]:
    dims = [
        {
            "id": "INSTRUMENT",
            "concept": "INSTRUMENT",
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
    ]
    block: dict[str, Any] = {
        "id": series_id,
        "sheet": "Data",
        "data_range": f"Data!{area}",
        "layout": "matrix",
        "key": ["INSTRUMENT", "TIME_PERIOD"],
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": dims,
        },
    }
    if direction == "output":
        block["output"] = {"compute": {"name": compute or f"compute_{series_id}"}}
    else:
        block["constant"] = {}
    return block


def _lockstep_workbook(
    tmp_path: Path,
    instruments: int,
    years: int,
    *,
    blank_rects: int = 0,
    hole: str | None = "Data!C3",
) -> tuple[Path, tuple[str, ...], dict[str, Any]]:
    last_row = instruments + 1
    last_col = get_column_letter(1 + years)
    out_start_col = 2 + years
    out_last_col = get_column_letter(1 + 2 * years)
    data: dict[str, object] = {}
    other: dict[str, object] = {}
    for year_index in range(years):
        data[f"{get_column_letter(2 + year_index)}1"] = 2020 + year_index
        data[f"{get_column_letter(out_start_col + year_index)}1"] = 2020 + year_index
    for row in range(2, last_row + 1):
        data[f"A{row}"] = f"COM{row - 1}"
        for year_index in range(years):
            data[f"{get_column_letter(2 + year_index)}{row}"] = float(row + year_index)
            start = get_column_letter(2)
            data[f"{get_column_letter(out_start_col + year_index)}{row}"] = (
                f"=SUM((${start}$2:${last_col}${last_row})*(${start}$2:${last_col}${last_row}))"
            )
    for index in range(1, blank_rects + 1):
        other[f"A{index}"] = None
    workbook = write_workbook(
        tmp_path / f"lockstep_{instruments}_{years}_{blank_rects}.xlsx",
        {"Data": data, "Other": other},
    )
    blanks: list[str] = []
    if hole is not None:
        blanks.append(hole)
    blanks.extend(f"Other!A{index}" for index in range(1, blank_rects + 1))
    document = bindings_document(
        _matrix_series("values", f"B2:{last_col}{last_row}", "constant"),
        _matrix_series(
            "result",
            f"{get_column_letter(out_start_col)}2:{out_last_col}{last_row}",
            "output",
            compute="compute_result",
        ),
        schema_version="1.16.0",
    )
    document["concept_scheme"]["concepts"].append({"id": "INSTRUMENT", "dtype": "string"})
    return workbook, tuple(blanks), document


def _count_calls(monkeypatch: pytest.MonkeyPatch, target: Any, name: str) -> dict[str, int]:
    counts = {"n": 0}
    original = getattr(target, name)

    def counting(*args: object, **kwargs: object) -> object:
        counts["n"] += 1
        return original(*args, **kwargs)

    monkeypatch.setattr(target, name, counting)
    return counts


def test_uniform_sum_emit_range_table_once_per_statement(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    counts = _count_calls(monkeypatch, ast_emit_mod, "_emit_range_table")
    small_book, small_blanks, small_doc = _lockstep_workbook(
        tmp_path, 3, 4, blank_rects=40, hole=None
    )
    generate_inverted(small_book, small_doc, blank_ranges=small_blanks)
    small = counts["n"]
    counts["n"] = 0
    large_book, large_blanks, large_doc = _lockstep_workbook(
        tmp_path, 6, 4, blank_rects=40, hole=None
    )
    generate_inverted(large_book, large_doc, blank_ranges=large_blanks)
    large = counts["n"]
    assert small > 0
    assert large == small


def test_unrelated_blanks_do_not_enter_named_view_hot_path(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    counts = _count_calls(monkeypatch, ast_emit_mod, "address_in_blank_ranges")
    book, blanks, document = _lockstep_workbook(tmp_path, 4, 4, blank_rects=80, hole=None)
    generate_inverted(book, document, blank_ranges=blanks)
    assert counts["n"] == 0


def test_hole_does_not_replay_range_table_per_host(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    counts = _count_calls(monkeypatch, ast_emit_mod, "_emit_range_table")
    book, blanks, document = _lockstep_workbook(tmp_path, 4, 4, blank_rects=80, hole="Data!C3")
    generate_inverted(book, document, blank_ranges=blanks)
    assert counts["n"] == 6


def test_lockstep_export_matches_evaluator_with_hole_and_unrelated_blanks(
    tmp_path: Path,
) -> None:
    workbook, blanks, document = _lockstep_workbook(tmp_path, 3, 3, blank_rects=12, hole="Data!C3")
    catalog, _deps, graph = inverted_graph_parts(workbook, document, blank_ranges=blanks)
    pkg = load_package(
        generate_inverted(workbook, document, blank_ranges=blanks),
        tmp_path,
        name="lockstep_hotspot",
    )
    kwargs = named_input_kwargs(pkg, catalog, graph)
    for series in catalog.constant_series():
        kwargs[series.series_id] = getattr(pkg.data, series.series_id.upper())
    got = call_compute(pkg, "result", kwargs)
    result_series = catalog.get("result")
    with FormulaEvaluator(graph, blank_ranges=blanks) as evaluator:
        expected = evaluator.evaluate(list(result_series.cells))
    for cell, point in zip(result_series.cells, result_series.domain, strict=True):
        key = tuple(point[field] for field in result_series.key_fields)
        want = expected[cell]
        value = got[key]
        if isinstance(want, float) and isinstance(value, float):
            assert value == pytest.approx(want), cell
        else:
            assert value == want, f"{cell}: {value!r} != {want!r}"
    assert "xl_sum(" in generate_inverted(workbook, document, blank_ranges=blanks)["internals.py"]


def test_identity_string_keys_skip_lockstep_follow(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    workbook = write_workbook(
        tmp_path / "identity_keys.xlsx",
        {
            "Data": {
                "A1": "COM1",
                "A2": "COM2",
                "A3": "COM3",
                "B1": 1.0,
                "B2": 2.0,
                "B3": 3.0,
                "C1": "=B1",
                "C2": "=B2",
                "C3": "=B3",
            }
        },
    )
    document = bindings_document(
        series_entry(
            "values",
            "Data!B1:B3",
            layout="series",
            direction="input",
            label_column="A",
            key_concept="INSTRUMENT",
            key_read="string",
        ),
        series_entry(
            "result",
            "Data!C1:C3",
            layout="series",
            direction="output",
            label_column="A",
            key_concept="INSTRUMENT",
            key_read="string",
        ),
        schema_version="1.16.0",
    )
    document["concept_scheme"]["concepts"].append({"id": "INSTRUMENT", "dtype": "string"})
    counts = _count_calls(monkeypatch, ast_emit_mod, "_string_follow_expr")
    generate_inverted(workbook, document)
    assert counts["n"] == 0


class _CountingEdges:
    def __init__(self, items: Sequence[DependenceEdge]) -> None:
        self._items = tuple(items)
        self.scans = 0

    def __iter__(self):
        self.scans += 1
        return iter(self._items)


def _lockstep_pair(n: int) -> tuple[EmitContext, BoundSeries]:
    host_cells = tuple(f"Out!B{row}" for row in range(2, n + 2))
    producer_cells = tuple(f"In!B{row}" for row in range(2, n + 2))
    domain = tuple(KeyPoint((("INSTRUMENT", f"COM{i}"),)) for i in range(1, n + 1))
    host = BoundSeries(
        series_id="result",
        layout="series",
        direction="output",
        cells=host_cells,
        key_fields=("INSTRUMENT",),
        dtype="float",
        compute_name="compute_result",
        raw={},
        domain=domain,
        statements=(Statement("result", "result", None, 0, n, host_cells, domain),),
    )
    producer = BoundSeries(
        series_id="values",
        layout="series",
        direction="input",
        cells=producer_cells,
        key_fields=("INSTRUMENT",),
        dtype="float",
        compute_name=None,
        raw={},
        domain=domain,
        statements=(Statement("values", "values", None, 0, n, producer_cells, domain),),
    )
    catalog = make_catalog(
        series={"result": host, "values": producer},
        order=("values", "result"),
        address_to_id={
            **{cell: "result" for cell in host_cells},
            **{cell: "values" for cell in producer_cells},
        },
    )
    edges = _CountingEdges(
        DependenceEdge(
            consumer_id="result",
            producer_id="values",
            consumer_cell=host_cells[index],
            producer_cell=producer_cells[index],
            distance=0,
            access="identity",
        )
        for index in range(n)
    )
    deps = SeriesDeps(
        host_id="result",
        param_ids=("values",),
        is_scan=False,
        seed_id=None,
        aligned_ids=frozenset({"values"}),
        lookup_ids=frozenset(),
        index_maps={},
        affine_maps={"values": (1, 0)},
        edges=edges,  # type: ignore[arg-type]
    )
    ctx = EmitContext(
        host=host,
        catalog=catalog,
        deps=deps,
        host_index=0,
        host_cell=host_cells[0],
        coordinate_vars={"INSTRUMENT": "instrument"},
    )
    return ctx, producer


def test_lockstep_producer_slots_scans_edges_once() -> None:
    ctx, producer = _lockstep_pair(40)
    first = _lockstep_producer_slots(ctx, producer)
    second = _lockstep_producer_slots(ctx, producer)
    assert first == {index: index for index in range(40)}
    assert second == first
    edges = ctx.deps.edges
    assert isinstance(edges, _CountingEdges)
    assert edges.scans == 1
