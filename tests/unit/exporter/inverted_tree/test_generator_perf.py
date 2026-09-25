"""Generator-side inverted-tree hot spots (#618, #636, #653, #683, #769)."""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import replace
from pathlib import Path
from typing import Literal

import pytest
from fastpyxl.utils.cell import get_column_letter

from excel_grapher.core import address_keys as address_keys_mod
from excel_grapher.core.address_keys import normalize_key
from excel_grapher.exporter.inverted_tree import access as access_mod
from excel_grapher.exporter.inverted_tree import catalog as catalog_mod
from excel_grapher.exporter.inverted_tree import deps as deps_mod
from excel_grapher.exporter.inverted_tree.access import overlapping_schedule_peer
from excel_grapher.exporter.inverted_tree.catalog import (
    BoundSeries,
    KeyPoint,
    SeriesCatalog,
    Statement,
    build_catalog,
    covering_series,
    covering_series_of_column,
    covering_series_of_range,
    schedule_coord,
)
from excel_grapher.exporter.inverted_tree.deps import DependenceEdge, iter_range_addresses
from excel_grapher.series_bindings import resolve as resolve_mod
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    make_catalog,
    series_entry,
    write_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a1_leaf_closure import (
    _a1_bindings,
    _a1_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a11_zipper import (
    _zipper_bindings,
    _zipper_workbook,
)


def _constant_series_bindings(n: int) -> dict:
    """Minimal catalog bindings: one constant series, no declared key.

    Schema validation requires a non-empty `key` on `layout: series`.
    `build_catalog` accepts `key: []` and treats expansion order as the
    schedule, which is the large-catalog case in #636.
    """
    return {
        "series": [
            {
                "id": "consts",
                "data_range": f"Sheet1!A1:A{n}",
                "layout": "series",
                "constant": {},
                "key": [],
                "structure": {"measure": {"dtype": "float"}},
            }
        ]
    }


def _count_normalize_key_calls(monkeypatch: pytest.MonkeyPatch) -> dict[str, int]:
    calls = {"n": 0}
    original = address_keys_mod.normalize_key

    def counting(key: str) -> str:
        calls["n"] += 1
        return original(key)

    monkeypatch.setattr(address_keys_mod, "normalize_key", counting)
    return calls


def test_generate_modules_walks_each_series_ast_once(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    walks: dict[str, int] = {}
    original = deps_mod.collect_series_edges

    def counting(
        series: BoundSeries,
        *,
        catalog: SeriesCatalog,
        graph: object,
        blank_rects: object = None,
    ) -> list[DependenceEdge]:
        walks[series.series_id] = walks.get(series.series_id, 0) + 1
        return original(series, catalog=catalog, graph=graph, blank_rects=blank_rects)

    monkeypatch.setattr(deps_mod, "collect_series_edges", counting)
    generate_inverted(_a1_workbook(tmp_path), _a1_bindings())
    generate_inverted(_zipper_workbook(tmp_path), _zipper_bindings())
    assert walks, "expected collect_series_edges to run during generate_modules"
    assert all(count == 1 for count in walks.values()), walks


def test_build_catalog_normalize_key_scales_linearly(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    workbook = write_workbook(tmp_path / "const.xlsx", {"Sheet1": {"A1": 1}})
    n_small, n_large = 1_000, 4_000
    calls = _count_normalize_key_calls(monkeypatch)
    build_catalog(_constant_series_bindings(n_small), workbook=workbook)
    small = calls["n"]
    calls["n"] = 0
    build_catalog(_constant_series_bindings(n_large), workbook=workbook)
    large = calls["n"]
    assert small == n_small, small
    assert large == n_large, large
    assert large == small * (n_large // n_small)


def test_schedule_coord_does_not_renormalize_catalog_cells(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    workbook = write_workbook(tmp_path / "const.xlsx", {"Sheet1": {"A1": 1}})
    catalog = build_catalog(_constant_series_bindings(n=8), workbook=workbook)
    calls = _count_normalize_key_calls(monkeypatch)
    for cell in catalog.get("consts").cells:
        schedule_coord(cell, catalog)
        assert catalog.series_for(cell) is catalog.get("consts")
        assert catalog.get("consts").index_of(cell) is not None
    assert calls["n"] == 0, calls["n"]


def _synthetic_series(
    series_id: str,
    cells: tuple[str, ...],
    years: Sequence[int],
    *,
    direction: str = "internal",
    peeled_seed: bool = False,
) -> BoundSeries:
    domain = tuple(KeyPoint((("TIME_PERIOD", year),)) for year in years)
    n = len(cells)
    if peeled_seed and n >= 2:
        statements = (
            Statement(f"{series_id}__0", series_id, None, 0, 1, cells[:1], domain[:1]),
            Statement(f"{series_id}__1", series_id, None, 1, n, cells[1:], domain[1:]),
        )
    else:
        statements = (Statement(series_id, series_id, None, 0, n, cells, domain),)
    return BoundSeries(
        series_id=series_id,
        layout="series",
        direction=direction,
        cells=cells,
        key_fields=("TIME_PERIOD",),
        dtype="float",
        compute_name=None,
        raw={},
        domain=domain,
        statements=statements,
    )


def _synthetic_zipper(n: int) -> tuple[SeriesCatalog, tuple[DependenceEdge, ...]]:
    """2-series zipper: `debt_t = debt_{t-1} + adj_t`, `adj_t = debt_{t-1} * r`.

    Debt is a peeled-seed two-statement series.
    """
    debt_cells = tuple(f"Engine!A{i}" for i in range(1, n + 1))
    adj_cells = tuple(f"Engine!B{i}" for i in range(2, n + 1))
    debt = _synthetic_series("debt", debt_cells, range(n), direction="output", peeled_seed=True)
    adj = _synthetic_series("adjustment", adj_cells, range(1, n), direction="internal")
    catalog = make_catalog(
        series={"debt": debt, "adjustment": adj},
        order=("debt", "adjustment"),
        address_to_id={
            **{cell: "debt" for cell in debt_cells},
            **{cell: "adjustment" for cell in adj_cells},
        },
    )
    edges: list[DependenceEdge] = []
    for t in range(1, n):
        debt_cell = debt_cells[t]
        prev_debt = debt_cells[t - 1]
        adj_cell = adj_cells[t - 1]
        edges.append(
            DependenceEdge(
                consumer_id="debt",
                producer_id="debt",
                consumer_cell=debt_cell,
                producer_cell=prev_debt,
                distance=1,
                access="shift",
            )
        )
        edges.append(
            DependenceEdge(
                consumer_id="debt",
                producer_id="adjustment",
                consumer_cell=debt_cell,
                producer_cell=adj_cell,
                distance=0,
                access="identity",
            )
        )
        edges.append(
            DependenceEdge(
                consumer_id="adjustment",
                producer_id="debt",
                consumer_cell=adj_cell,
                producer_cell=prev_debt,
                distance=1,
                access="shift",
            )
        )
    return catalog, tuple(edges)


def _count_load_workbook(monkeypatch: pytest.MonkeyPatch) -> dict[str, int]:
    counts = {"n": 0}
    original = resolve_mod.fastpyxl.load_workbook

    def counting(*args: object, **kwargs: object) -> object:
        counts["n"] += 1
        return original(*args, **kwargs)

    monkeypatch.setattr(resolve_mod.fastpyxl, "load_workbook", counting)
    return counts


def test_build_catalog_loads_workbook_once(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    cells = {"A1": 2009, "B1": 2010, "C1": 2011}
    for row in range(2, 8):
        for col in "ABC":
            cells[f"{col}{row}"] = float(row)
    workbook = write_workbook(tmp_path / "keyed.xlsx", {"Inputs": cells})
    document = bindings_document(
        *(
            series_entry(
                f"row_{row}",
                f"Inputs!A{row}:C{row}",
                layout="series",
                direction="input",
                header_row=1,
            )
            for row in range(2, 8)
        )
    )
    loads = _count_load_workbook(monkeypatch)
    catalog = build_catalog(document, workbook=workbook)
    assert len(catalog.order) == 6
    assert all(point["TIME_PERIOD"] == 2009 for point in catalog.get("row_2").domain[:1])
    assert loads["n"] == 1, loads


def _count_schedule_coord(monkeypatch: pytest.MonkeyPatch) -> dict[str, int]:
    counts = {"n": 0}
    original = access_mod.schedule_coord

    def counting(address: object, catalog: SeriesCatalog) -> int:
        counts["n"] += 1
        return original(address, catalog)

    monkeypatch.setattr(access_mod, "schedule_coord", counting)
    return counts


def test_overlapping_schedule_peer_uses_cached_coords(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    catalog, _edges = _synthetic_zipper(200)
    other = _synthetic_series("other", ("Engine!Z1", "Engine!Z2"), (1000, 1001))
    catalog = make_catalog(
        series={**catalog.series, "other": other},
        order=(*catalog.order, "other"),
        address_to_id={**catalog.address_to_id, "Engine!Z1": "other", "Engine!Z2": "other"},
    )
    host = catalog.get("debt")
    producer = catalog.get("adjustment")
    assert overlapping_schedule_peer(host, producer, catalog)
    assert not overlapping_schedule_peer(host, other, catalog)
    calls = _count_schedule_coord(monkeypatch)
    for _ in range(1_000):
        assert overlapping_schedule_peer(host, producer, catalog)
        assert not overlapping_schedule_peer(host, other, catalog)
    assert calls["n"] == 0, calls


def _block_cells(rows: int, cols: int, *, sheet: str = "Sheet1") -> tuple[str, ...]:
    return tuple(
        f"{sheet}!{get_column_letter(col)}{row}"
        for row in range(1, rows + 1)
        for col in range(1, cols + 1)
    )


def _block_series(series_id: str, rows: int, cols: int) -> BoundSeries:
    cells = _block_cells(rows, cols)
    domain = tuple(KeyPoint(()) for _ in cells)
    return BoundSeries(
        series_id=series_id,
        layout="matrix",
        direction="constant",
        cells=cells,
        key_fields=(),
        dtype="float",
        compute_name=None,
        raw={},
        domain=domain,
        statements=(Statement(series_id, series_id, None, 0, len(cells), cells, domain),),
    )


def _block_catalog(rows: int, cols: int) -> SeriesCatalog:
    block = _block_series("block", rows, cols)
    return make_catalog(
        series={"block": block},
        order=("block",),
        address_to_id={cell: "block" for cell in block.cells},
    )


def _count_covering_address_ops(monkeypatch: pytest.MonkeyPatch) -> dict[str, int]:
    counts = {"parse": 0, "format": 0}
    original_parse = catalog_mod.parse_cell_coords
    original_format = catalog_mod.format_cell_key

    def counting_parse(address: str) -> object:
        counts["parse"] += 1
        return original_parse(address)

    def counting_format(sheet: str, column: str, row: int) -> str:
        counts["format"] += 1
        return original_format(sheet, column, row)

    monkeypatch.setattr(catalog_mod, "parse_cell_coords", counting_parse)
    monkeypatch.setattr(catalog_mod, "format_cell_key", counting_format)
    return counts


def test_covering_series_range_parses_are_o1(monkeypatch: pytest.MonkeyPatch) -> None:
    catalog = _block_catalog(200, 30)
    start, end = "Sheet1!A1", "Sheet1!AD200"
    assert covering_series(catalog, iter_range_addresses(start, end)) is catalog.get("block")
    counts = _count_covering_address_ops(monkeypatch)
    assert covering_series_of_range(catalog, start, end) is catalog.get("block")
    full = dict(counts)
    counts["parse"] = counts["format"] = 0
    assert covering_series_of_range(catalog, "Sheet1!B2", "Sheet1!C10") is catalog.get("block")
    sub = dict(counts)
    counts["parse"] = counts["format"] = 0
    assert covering_series_of_range(catalog, "Sheet1!A1", "Sheet1!AE200") is None
    overhang = dict(counts)
    assert full["parse"] == sub["parse"] == overhang["parse"]
    assert full["parse"] <= 4, full
    assert full["format"] <= 2, full
    assert sub["format"] == full["format"]
    assert overhang["format"] == full["format"]


def test_covering_series_column_parses_are_o1(monkeypatch: pytest.MonkeyPatch) -> None:
    catalog = _block_catalog(200, 30)
    start, end = "Sheet1!A1", "Sheet1!AH200"
    counts = _count_covering_address_ops(monkeypatch)
    assert covering_series_of_column(catalog, start, end, 1) is catalog.get("block")
    first = dict(counts)
    counts["parse"] = counts["format"] = 0
    assert covering_series_of_column(catalog, start, end, 30) is catalog.get("block")
    last_in_block = dict(counts)
    counts["parse"] = counts["format"] = 0
    assert covering_series_of_column(catalog, start, end, 31) is None
    outside = dict(counts)
    assert first["parse"] == last_in_block["parse"] == outside["parse"]
    assert first["parse"] <= 4, first
    assert first["format"] <= 2, first
    assert last_in_block["format"] == first["format"]
    assert outside["format"] == first["format"]


@pytest.mark.parametrize("field, expected", [("VINTAGE", "row"), ("UNKNOWN", None)])
def test_key_axis_is_inferred_once_per_series(
    monkeypatch: pytest.MonkeyPatch, field: str, expected: Literal["row"] | None
) -> None:
    """Key-axis analysis must scale with series size, not reference count (#769)."""
    series = BoundSeries(
        series_id="vintages",
        layout="series",
        direction="internal",
        cells=tuple(normalize_key(f"Engine!A{row}") for row in range(1, 101)),
        key_fields=("VINTAGE",),
        dtype="float",
        compute_name=None,
        raw={},
        domain=tuple(KeyPoint((("VINTAGE", row),)) for row in range(1, 101)),
        statements=(),
    )
    original = deps_mod._infer_key_field_axis
    calls = 0

    def counted(series: BoundSeries, field: str) -> Literal["sheet", "row", "col"] | None:
        nonlocal calls
        calls += 1
        return original(series, field)

    monkeypatch.setattr(deps_mod, "_infer_key_field_axis", counted)
    for _ in range(20):
        assert deps_mod._key_field_axis(series, field) == expected
    assert calls == 1
    other = replace(series, domain=tuple(KeyPoint((("VINTAGE", 1),)) for _ in series.cells))
    assert deps_mod._key_field_axis(other, field) is None
    assert calls == 2


def _matrix_bound_series(n_inst: int, n_year: int) -> BoundSeries:
    """Rectangular INSTRUMENT × TIME_PERIOD series (issue 830)."""
    years = tuple(range(2024, 2024 + n_year))
    instruments = tuple(f"L{i}" for i in range(n_inst))
    domain = tuple(
        KeyPoint(items=(("INSTRUMENT", inst), ("TIME_PERIOD", year)))
        for inst in instruments
        for year in years
    )
    cells = tuple(f"Debt!A{i + 1}" for i in range(len(domain)))
    return BoundSeries(
        series_id="ext_debt_new_disbursements_external",
        layout="matrix",
        direction="internal",
        cells=cells,
        key_fields=("INSTRUMENT", "TIME_PERIOD"),
        dtype="float",
        compute_name=None,
        raw={},
        domain=domain,
        statements=(),
        graph_cells=frozenset(cells),
        key_types=("string", "int"),
    )


def _count_domain_explicit(monkeypatch: pytest.MonkeyPatch) -> dict[str, int]:
    counts = {"n": 0}
    original = catalog_mod.Domain.explicit

    def counting(
        *,
        axes: object,
        coordinates: object,
    ) -> object:
        counts["n"] += 1
        return original(axes=axes, coordinates=coordinates)

    monkeypatch.setattr(catalog_mod.Domain, "explicit", staticmethod(counting))
    return counts


def test_bound_series_domain_properties_are_cached(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """tensor_domain / coordinate_cells / required_coordinates build once (#830)."""
    series = _matrix_bound_series(20, 40)
    counts = _count_domain_explicit(monkeypatch)
    domain = series.tensor_domain
    cells = series.coordinate_cells
    required = series.required_coordinates
    assert counts["n"] == 1, counts
    assert series.tensor_domain is domain
    assert series.coordinate_cells is cells
    assert series.required_coordinates is required
    _ = tuple(coord for coord in domain if coord in series.required_coordinates)
    assert counts["n"] == 1, counts
    assert len(required) == len(series.domain)


def test_bound_series_domain_cache_resets_on_replace(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """`dataclasses.replace` must not share the previous instance's domain cache."""
    series = _matrix_bound_series(4, 5)
    counts = _count_domain_explicit(monkeypatch)
    original_domain = series.tensor_domain
    original_cells = series.coordinate_cells
    original_required = series.required_coordinates
    assert counts["n"] == 1, counts
    replaced = replace(series, statements=series.statements)
    assert replaced.tensor_domain == original_domain
    assert replaced.tensor_domain is not original_domain
    assert replaced.coordinate_cells == original_cells
    assert replaced.coordinate_cells is not original_cells
    assert replaced.required_coordinates == original_required
    assert replaced.required_coordinates is not original_required
    assert counts["n"] == 2, counts


def test_emit_named_data_binds_required_coordinates_once(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Membership tests must not re-enter `required_coordinates` per coordinate (#830)."""
    from excel_grapher.exporter.export_runtime.tensor import Axis
    from excel_grapher.exporter.inverted_tree.named_axes import NamedAxes
    from excel_grapher.exporter.inverted_tree.named_emit import emit_named_data

    n_inst, n_year = 20, 40
    series = _matrix_bound_series(n_inst, n_year)
    n = len(series.domain)
    catalog = make_catalog(
        series={series.series_id: series},
        order=(series.series_id,),
        address_to_id={cell: series.series_id for cell in series.cells},
    )
    named_axes = NamedAxes.plan(
        (
            Axis("INSTRUMENT", tuple(f"L{i}" for i in range(n_inst)), str),
            Axis("TIME_PERIOD", tuple(range(2024, 2024 + n_year)), int),
        )
    )
    original = BoundSeries.required_coordinates.fget
    assert original is not None
    accesses = {"n": 0}

    def counting(self: BoundSeries) -> frozenset:
        accesses["n"] += 1
        return original(self)

    monkeypatch.setattr(BoundSeries, "required_coordinates", property(counting))
    workbook = write_workbook(tmp_path / "issue_830.xlsx", {"Debt": {"A1": 0}})
    data = emit_named_data(catalog, workbook, named_axes, {})
    assert "define_series(" in data
    assert accesses["n"] < n, accesses
    assert accesses["n"] <= 2, accesses


def _index_blank_workbook(
    tmp_path: Path,
    *,
    n_formulas: int,
    n_rows: int,
    n_blanks: int,
    unbound: str | None = None,
) -> tuple[Path, tuple[str, ...], dict]:
    """Absolute INDEX/MATCH copies over a table punched by blank rows.

    Mirrors the #999 MCVE: a non-literal INDEX column is not one covering
    series, so plan and emit walk the same absolute rectangle per formula.
    """
    n_cols = 4
    last = n_rows + 1
    end_col = get_column_letter(n_cols)
    table: dict[str, object] = {}
    for row in range(2, last + 1):
        table[f"A{row}"] = f"k{row}"
        for col in range(2, n_cols + 1):
            table[f"{get_column_letter(col)}{row}"] = float(row + col)
    unbound_column = ""
    unbound_row = 0
    if unbound is not None:
        unbound_cell = unbound.split("!", 1)[1]
        unbound_column = "".join(ch for ch in unbound_cell if ch.isalpha())
        unbound_row = int("".join(ch for ch in unbound_cell if ch.isdigit()))
        table[unbound_cell] = 99.0
    engine: dict[str, object] = {}
    outputs: dict[str, object] = {}
    formula = (
        f"=INDEX(Table!$A$2:${end_col}${last},"
        f"MATCH(Inputs!$B$1,Table!$A$2:$A${last},0),Inputs!$C$1)"
    )
    for index in range(n_formulas):
        row = index + 2
        engine[f"A{row}"] = f"f{index}"
        engine[f"B{row}"] = formula
        outputs[f"A{row}"] = f"f{index}"
        outputs[f"B{row}"] = f"=Engine!B{row}"
    blanks = tuple(f"Table!A{row}:{end_col}{row}" for row in range(2, 2 + n_blanks))
    series = [
        series_entry("lookup_key", "Inputs!B1", layout="scalar", direction="input", dtype="string"),
        series_entry("col_index", "Inputs!C1", layout="scalar", direction="input", dtype="int"),
        series_entry(
            "labels",
            f"Table!A2:A{last}",
            layout="series",
            direction="input",
            dtype="string",
            label_column="A",
            key_concept="COUNTRY",
            key_read="string",
        ),
    ]
    for col in range(2, n_cols + 1):
        letter = get_column_letter(col)
        if unbound_column == letter:
            continue
        series.append(
            series_entry(
                f"col_{col}",
                f"Table!{letter}2:{letter}{last}",
                layout="series",
                direction="input",
                label_column="A",
                key_concept="COUNTRY",
                key_read="string",
            )
        )
    if unbound_column:
        # Leave the valued cell outside every series. Neighbor cells in that
        # column stay bound so the window is not an entirely empty table.
        above = unbound_row - 1
        below = unbound_row + 1
        if above >= 2:
            series.append(
                series_entry(
                    "unbound_above",
                    f"Table!{unbound_column}2:{unbound_column}{above}",
                    layout="series",
                    direction="input",
                    label_column="A",
                    key_concept="COUNTRY",
                    key_read="string",
                )
            )
        if below <= last:
            series.append(
                series_entry(
                    "unbound_below",
                    f"Table!{unbound_column}{below}:{unbound_column}{last}",
                    layout="series",
                    direction="input",
                    label_column="A",
                    key_concept="COUNTRY",
                    key_read="string",
                )
            )
    series.append(
        series_entry(
            "resolved",
            f"Engine!B2:B{n_formulas + 1}",
            layout="series",
            direction="internal",
            label_column="A",
            key_concept="COUNTRY",
            key_read="string",
        )
    )
    series.append(
        series_entry(
            "published",
            f"Outputs!B2:B{n_formulas + 1}",
            layout="series",
            direction="output",
            label_column="A",
            key_concept="COUNTRY",
            key_read="string",
        )
    )
    document = bindings_document(*series)
    workbook = write_workbook(
        tmp_path / f"index_blanks_{n_formulas}_{n_rows}_{n_blanks}.xlsx",
        {
            "Inputs": {"B1": "k2", "C1": 2},
            "Table": table,
            "Engine": engine,
            "Outputs": outputs,
        },
    )
    return workbook, blanks, document


def _prepare_index_blank_collect(
    tmp_path: Path, *, n_formulas: int, n_rows: int, n_blanks: int
) -> tuple[object, object, tuple]:
    from excel_grapher.exporter.inverted_tree.catalog import build_catalog
    from excel_grapher.grapher import create_dependency_graph
    from excel_grapher.grapher.blank_ranges import normalize_blank_range_specs
    from excel_grapher.series_bindings import validate_bindings_document
    from excel_grapher.series_bindings.workflow import all_series_targets

    workbook, blanks, document = _index_blank_workbook(
        tmp_path, n_formulas=n_formulas, n_rows=n_rows, n_blanks=n_blanks
    )
    bindings = validate_bindings_document(document)
    graph = create_dependency_graph(
        workbook,
        all_series_targets(bindings, workbook=workbook),
        load_values=True,
        use_cached_dynamic_refs=True,
        capture_dependency_provenance=True,
        blank_ranges=blanks,
    )
    catalog = build_catalog(bindings, workbook=workbook, graph=graph, blank_ranges=blanks)
    return catalog, graph, normalize_blank_range_specs(blanks)


def test_absolute_index_blank_checks_do_not_scale_with_copies(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Identical absolute ranges are resolved once, from corners (#999)."""
    from excel_grapher.core import address_keys as address_keys_mod
    from excel_grapher.exporter.inverted_tree.catalog import SeriesCatalog
    from excel_grapher.exporter.inverted_tree.deps import collect_catalog_edges
    from excel_grapher.grapher import blank_ranges as blank_ranges_mod
    from excel_grapher.grapher.graph import DependencyGraph

    n_rows = 30
    n_blanks = 10
    area = 4 * n_rows
    small_parts = _prepare_index_blank_collect(
        tmp_path, n_formulas=2, n_rows=n_rows, n_blanks=n_blanks
    )
    large_parts = _prepare_index_blank_collect(
        tmp_path, n_formulas=6, n_rows=n_rows, n_blanks=n_blanks
    )
    counts = {"parse": 0, "blank": 0}
    original_parse = address_keys_mod.parse_cell_coords
    original_blank = blank_ranges_mod.address_in_blank_ranges

    def counting_parse(address: str) -> tuple[str, int, int]:
        counts["parse"] += 1
        return original_parse(address)

    def counting_blank(address: str, rects: object) -> bool:
        counts["blank"] += 1
        return original_blank(address, rects)

    monkeypatch.setattr(address_keys_mod, "parse_cell_coords", counting_parse)
    monkeypatch.setattr(blank_ranges_mod, "parse_cell_coords", counting_parse)
    monkeypatch.setattr(deps_mod, "parse_cell_coords", counting_parse)
    monkeypatch.setattr(catalog_mod, "parse_cell_coords", counting_parse)
    monkeypatch.setattr(blank_ranges_mod, "address_in_blank_ranges", counting_blank)
    monkeypatch.setattr(deps_mod, "address_in_blank_ranges", counting_blank)

    def collect(parts: tuple[object, object, tuple]) -> None:
        catalog, graph, rects = parts
        assert isinstance(catalog, SeriesCatalog)
        assert isinstance(graph, DependencyGraph)
        edges = collect_catalog_edges(catalog, graph, blank_rects=rects)
        assert any(edge.producer_id == "col_2" for edge in edges.edges)

    collect(small_parts)
    small_parse, small_blank = counts["parse"], counts["blank"]
    counts["parse"] = 0
    counts["blank"] = 0
    collect(large_parts)
    assert counts["parse"] - small_parse < area, (small_parse, counts["parse"], area)
    assert counts["blank"] - small_blank < area, (small_blank, counts["blank"], area)


def test_unbound_nonblank_in_blank_window_still_fail_closes(tmp_path: Path) -> None:
    """A valued unbound cell stays missing when neighboring rows are blank."""
    from excel_grapher.exporter.inverted_tree.catalog import build_catalog
    from excel_grapher.exporter.inverted_tree.deps import collect_catalog_edges
    from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
    from excel_grapher.grapher import create_dependency_graph
    from excel_grapher.grapher.blank_ranges import normalize_blank_range_specs
    from excel_grapher.series_bindings import validate_bindings_document
    from excel_grapher.series_bindings.workflow import all_series_targets

    workbook, blanks, document = _index_blank_workbook(
        tmp_path,
        n_formulas=2,
        n_rows=6,
        n_blanks=1,
        unbound="Table!C4",
    )
    bindings = validate_bindings_document(document)
    graph = create_dependency_graph(
        workbook,
        all_series_targets(bindings, workbook=workbook),
        load_values=True,
        use_cached_dynamic_refs=True,
        capture_dependency_provenance=True,
        blank_ranges=blanks,
    )
    catalog = build_catalog(bindings, workbook=workbook, graph=graph, blank_ranges=blanks)
    with pytest.raises(InvertedTreeExportError, match="Table!C4"):
        collect_catalog_edges(catalog, graph, blank_rects=normalize_blank_range_specs(blanks))
