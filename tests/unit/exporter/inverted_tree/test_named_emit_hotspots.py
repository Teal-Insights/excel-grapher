"""Named emit must not re-lower identical range tables per host cell (#867).

Uniform formula families replay `_emit_range_table` once per statement, not
once per catalog member. Range-vs-blank tests are geometric, so unrelated
blank rectangles stay off the per-cell hot path. Lockstep producer slots
are cached per `(host, producer)`. A mixed identity/remap whose outlier is
not a first/interior/last sample still lowers through the lockstep dict,
including when the producer is a 2D matrix column (#873).
"""

from __future__ import annotations

import re
from collections.abc import Sequence
from dataclasses import replace
from pathlib import Path
from typing import Any

import pytest
from fastpyxl.utils.cell import get_column_letter

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree import ast_emit as ast_emit_mod
from excel_grapher.exporter.inverted_tree.ast_emit import (
    EmitContext,
    _lockstep_producer_slots,
    _lockstep_string_map,
)
from excel_grapher.exporter.inverted_tree.catalog import BoundSeries, KeyPoint, Statement
from excel_grapher.exporter.inverted_tree.deps import DependenceEdge, SeriesDeps
from tests.unit.exporter.inverted_tree.helpers import (
    assert_package_matches_evaluator,
    bindings_document,
    call_compute,
    generate_inverted,
    inverted_graph_parts,
    invoke_public_compute,
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
    small_book, small_blanks, small_doc = _lockstep_workbook(
        tmp_path, 4, 4, blank_rects=80, hole="Data!C3"
    )
    generate_inverted(small_book, small_doc, blank_ranges=small_blanks)
    small = counts["n"]
    counts["n"] = 0
    large_book, large_blanks, large_doc = _lockstep_workbook(
        tmp_path, 8, 4, blank_rects=80, hole="Data!C3"
    )
    generate_inverted(large_book, large_doc, blank_ranges=large_blanks)
    large = counts["n"]
    assert small > 0
    assert large == small
    assert small <= 6


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


def test_off_sample_lockstep_remap_still_uses_the_dict(tmp_path: Path) -> None:
    """Identity samples must not hide a remap at an unsampled catalog index.

    Seven lockstep rows sample indices 0, 3, and 6. Those three host keys
    match the producer. The remap at index 1 is not a sample; replicating
    `instrument` would read a missing producer key.
    """
    cells: dict[str, object] = {}
    for index in range(7):
        row = index + 1
        cells[f"A{row}"] = f"COM{index + 1}"
        cells[f"B{row}"] = float(index + 1)
        cells[f"D{row}"] = "ALIAS" if index == 1 else f"COM{index + 1}"
        cells[f"E{row}"] = f"=B{row}"
    workbook = write_workbook(tmp_path / "off_sample_remap.xlsx", {"Data": cells})
    document = bindings_document(
        series_entry(
            "values",
            "Data!B1:B7",
            layout="series",
            direction="input",
            label_column="A",
            key_concept="INSTRUMENT",
            key_read="string",
        ),
        series_entry(
            "result",
            "Data!E1:E7",
            layout="series",
            direction="output",
            label_column="D",
            key_concept="INSTRUMENT",
            key_read="string",
        ),
        schema_version="1.16.0",
    )
    document["concept_scheme"]["concepts"].append({"id": "INSTRUMENT", "dtype": "string"})
    internals = generate_inverted(workbook, document)["internals.py"]
    assert "INSTRUMENT_TO_VALUES[instrument]" in internals
    assert "values[INSTRUMENT_TO_VALUES[instrument]]" in internals
    assert "ALIAS" in internals
    assert not re.search(r"if instrument ==", internals)
    pkg = assert_package_matches_evaluator(workbook, document, tmp_path, "off_sample_remap")
    got = invoke_public_compute(pkg, pkg.compute_result, dict(values=pkg.data.VALUES_DEFAULT))
    assert got["ALIAS"] == 2.0
    assert got["COM7"] == 7.0


def _aliased_terms_matrix_document() -> dict[str, Any]:
    """Host copies one matrix column; one interior instrument label is aliased (#873)."""
    instruments = {
        "Alpha": 2,
        "Beta": 3,
        "Gamma": 4,
        "Delta": 5,
        "Epsilon": 6,
        "Zeta": 7,
    }
    host_instruments = {
        "Alpha": 2,
        "Beta": 3,
        "Gamma": 4,
        "Delta": 5,
        "Epsilon alias": 6,
        "Zeta": 7,
    }
    document = bindings_document(
        {
            "id": "terms",
            "sheet": "Terms",
            "data_range": "Terms!B2:C7",
            "layout": "matrix",
            "input": {},
            "structure": {
                "measure": {
                    "concept": "OBS_VALUE",
                    "dtype": "float",
                    "bind": {"kind": "data_cell", "read": "float"},
                },
                "dimensions": [
                    {
                        "id": "INSTRUMENT",
                        "concept": "INSTRUMENT",
                        "role": "key",
                        "scope": "cell",
                        "bind": {"kind": "value_map", "values": instruments},
                    },
                    {
                        "id": "INDICATOR",
                        "concept": "INDICATOR",
                        "role": "key",
                        "scope": "cell",
                        "bind": {"kind": "column_header", "header_row": 1, "read": "string"},
                    },
                ],
            },
            "key": ["INSTRUMENT", "INDICATOR"],
        },
        {
            "id": "pv_interest",
            "sheet": "PV",
            "data_range": "PV!B2:B7",
            "layout": "series",
            "output": {"compute": {"name": "compute_pv_interest"}},
            "structure": {
                "measure": {
                    "concept": "OBS_VALUE",
                    "dtype": "float",
                    "bind": {"kind": "data_cell", "read": "float"},
                },
                "dimensions": [
                    {
                        "id": "INSTRUMENT",
                        "concept": "INSTRUMENT",
                        "role": "key",
                        "scope": "cell",
                        "bind": {"kind": "value_map", "values": host_instruments},
                    }
                ],
            },
            "key": ["INSTRUMENT"],
        },
        schema_version="1.16.0",
    )
    document["concept_scheme"]["concepts"].extend(
        [
            {"id": "INSTRUMENT", "dtype": "string"},
            {"id": "INDICATOR", "dtype": "string"},
        ]
    )
    return document


def test_off_sample_2d_lockstep_remap_uses_producer_keys(tmp_path: Path) -> None:
    """A 2D producer column copy must remap an unsampled aliased host key (#873).

    Six host members sample indices 0, 3, and 5. Those three names match the
    terms matrix. The alias at index 4 is not a sample; replicating
    `terms[instrument, 'Interest rate']` looks up a key the producer does not
    own. Lockstep remaps must key on the shared `INSTRUMENT` axis, not require
    slope 1 on the flattened (instrument × indicator) catalog index.
    """
    instruments = ("Alpha", "Beta", "Gamma", "Delta", "Epsilon", "Zeta")
    host_keys = ("Alpha", "Beta", "Gamma", "Delta", "Epsilon alias", "Zeta")
    terms: dict[str, object] = {"B1": "Interest rate", "C1": "Grace period"}
    pv: dict[str, object] = {}
    for row, (producer_name, host_name) in enumerate(
        zip(instruments, host_keys, strict=True), start=2
    ):
        terms[f"A{row}"] = producer_name
        terms[f"B{row}"] = row / 100.0
        terms[f"C{row}"] = float(row)
        pv[f"A{row}"] = host_name
        pv[f"B{row}"] = f"=Terms!B{row}"
    workbook = write_workbook(tmp_path / "aliased_terms.xlsx", {"Terms": terms, "PV": pv})
    document = _aliased_terms_matrix_document()
    internals = generate_inverted(workbook, document)["internals.py"]
    assert "INSTRUMENT_TO_TERMS[instrument]" in internals
    assert "terms[INSTRUMENT_TO_TERMS[instrument], 'Interest rate']" in internals
    assert "'Epsilon alias': 'Epsilon'" in internals
    assert "terms[instrument, 'Interest rate']" not in internals
    assert "terms['Epsilon alias'" not in internals
    pkg = assert_package_matches_evaluator(workbook, document, tmp_path, "aliased_terms")
    got = invoke_public_compute(pkg, pkg.compute_pv_interest, dict(terms=pkg.data.TERMS_DEFAULT))
    assert got["Epsilon alias"] == pytest.approx(0.06)
    assert got["Zeta"] == pytest.approx(0.07)


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


def _lockstep_column_pair(
    n: int, *, indicators: int = 2, rotate: bool = False
) -> tuple[EmitContext, BoundSeries]:
    """Host copies one column of an `n × indicators` producer matrix."""
    instruments = tuple(f"COM{i}" for i in range(1, n + 1))
    indicator_names = tuple(f"IND{i}" for i in range(indicators))
    host_cells = tuple(f"Out!B{row}" for row in range(2, n + 2))
    host_domain = tuple(KeyPoint((("INSTRUMENT", name),)) for name in instruments)
    producer_cells: list[str] = []
    producer_domain: list[KeyPoint] = []
    for row, name in enumerate(instruments, start=2):
        for offset, indicator in enumerate(indicator_names):
            producer_cells.append(f"In!{get_column_letter(2 + offset)}{row}")
            producer_domain.append(
                KeyPoint((("INSTRUMENT", name), ("INDICATOR", indicator))),
            )
    producer_cells_t = tuple(producer_cells)
    producer_domain_t = tuple(producer_domain)
    host = BoundSeries(
        series_id="result",
        layout="series",
        direction="output",
        cells=host_cells,
        key_fields=("INSTRUMENT",),
        dtype="float",
        compute_name="compute_result",
        raw={},
        domain=host_domain,
        statements=(Statement("result", "result", None, 0, n, host_cells, host_domain),),
    )
    producer = BoundSeries(
        series_id="terms",
        layout="matrix",
        direction="input",
        cells=producer_cells_t,
        key_fields=("INSTRUMENT", "INDICATOR"),
        dtype="float",
        compute_name=None,
        raw={},
        domain=producer_domain_t,
        statements=(
            Statement(
                "terms",
                "terms",
                None,
                0,
                len(producer_cells_t),
                producer_cells_t,
                producer_domain_t,
            ),
        ),
    )
    catalog = make_catalog(
        series={"result": host, "terms": producer},
        order=("terms", "result"),
        address_to_id={
            **{cell: "result" for cell in host_cells},
            **{cell: "terms" for cell in producer_cells_t},
        },
    )
    edges = _CountingEdges(
        DependenceEdge(
            consumer_id="result",
            producer_id="terms",
            consumer_cell=host_cells[index],
            producer_cell=producer_cells_t[((index + 1) % n if rotate else index) * indicators],
            distance=0,
            access="identity",
        )
        for index in range(n)
    )
    deps = SeriesDeps(
        host_id="result",
        param_ids=("terms",),
        is_scan=False,
        seed_id=None,
        aligned_ids=frozenset(),
        lookup_ids=frozenset(),
        index_maps={},
        affine_maps={},
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


@pytest.mark.parametrize("indicators", [2, 3])
def test_lockstep_producer_slots_accepts_shared_axis_column_stride(indicators: int) -> None:
    """A copied matrix column is lockstep on `INSTRUMENT`, not flat slope 1."""
    ctx, producer = _lockstep_column_pair(6, indicators=indicators)
    slots = _lockstep_producer_slots(ctx, producer)
    assert slots == {index: index * indicators for index in range(6)}


def test_lockstep_producer_slots_rejects_shared_axis_permutation() -> None:
    """A rotated column walk is not lockstep even when each host reads one slot."""
    ctx, producer = _lockstep_column_pair(4, indicators=2, rotate=True)
    assert _lockstep_producer_slots(ctx, producer) is None


class _CountingSlots(dict[int, int]):
    """Slot map that counts walks of its keys."""

    def __init__(self, *args: object, **kwargs: object) -> None:
        super().__init__(*args, **kwargs)  # type: ignore[arg-type]
        self.walks = 0

    def __iter__(self):
        self.walks += 1
        return super().__iter__()


def test_lockstep_string_map_caches_identity_none() -> None:
    """An identity pairing is `None`, and a repeat lookup does not walk slots."""
    ctx, producer = _lockstep_pair(8)
    slots = _CountingSlots({index: index for index in range(8)})
    assert _lockstep_string_map(ctx, producer, "INSTRUMENT", slots, "INSTRUMENT") is None
    assert _lockstep_string_map(ctx, producer, "INSTRUMENT", slots, "INSTRUMENT") is None
    assert slots.walks == 1


def test_lockstep_string_map_caches_remap_per_field() -> None:
    """A real remap is reused, and a different field does not share that entry."""
    ctx, producer = _lockstep_pair(3)
    mapped = replace(
        producer,
        domain=tuple(KeyPoint((("INSTRUMENT", label),)) for label in ("Alpha", "Beta", "Gamma")),
    )
    slots = _CountingSlots({index: index for index in range(3)})
    assert _lockstep_string_map(ctx, mapped, "MISSING", slots, "INSTRUMENT") is None
    found = _lockstep_string_map(ctx, mapped, "INSTRUMENT", slots, "INSTRUMENT")
    again = _lockstep_string_map(ctx, mapped, "INSTRUMENT", slots, "INSTRUMENT")
    other = _lockstep_string_map(ctx, mapped, "INSTRUMENT", slots, "OTHER")
    assert found == {"COM1": "Alpha", "COM2": "Beta", "COM3": "Gamma"}
    assert again == found
    assert other is None
    assert slots.walks == 3


def test_lockstep_string_map_cache_is_per_slot_map() -> None:
    """A different slot map for the same fields is computed on its own."""
    ctx, producer = _lockstep_pair(4)
    swapped = _CountingSlots({0: 1, 1: 0, 2: 2, 3: 3})
    identity = _CountingSlots({index: index for index in range(4)})
    found = _lockstep_string_map(ctx, producer, "INSTRUMENT", swapped, "INSTRUMENT")
    again = _lockstep_string_map(ctx, producer, "INSTRUMENT", identity, "INSTRUMENT")
    assert found == {"COM1": "COM2", "COM2": "COM1", "COM3": "COM3", "COM4": "COM4"}
    assert again is None
    assert swapped.walks == 1
    assert identity.walks == 1
