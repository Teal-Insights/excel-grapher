"""Catalog statements: per-cell key domains and auto shape-partition."""

from __future__ import annotations

import time
from pathlib import Path

import pytest

from excel_grapher.core.formula_ast import parse_formula_text
from excel_grapher.exporter.inverted_tree.catalog import (
    BoundSeries,
    KeyPoint,
    SeriesCatalog,
    Statement,
    build_catalog,
    partition_catalog,
)
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import make_cell_node
from excel_grapher.series_bindings import validate_bindings_document
from excel_grapher.series_bindings.resolve import UnknownBindKindError, resolve_key_domain
from excel_grapher.series_bindings.workflow import all_series_targets
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    make_catalog,
    series_entry,
    write_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a8_matrix import (
    _profile_bindings,
    _profile_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a12_formula_shape import (
    _a12_bindings,
    _a12_workbook,
)


def test_time_period_domain_resolves_header_values(tmp_path: Path) -> None:
    workbook = _a12_workbook(tmp_path)
    catalog = build_catalog(validate_bindings_document(_a12_bindings()), workbook=workbook)
    series = catalog.get("path")
    assert [point["TIME_PERIOD"] for point in series.domain] == [2009, 2010, 2011]
    assert len(series.domain) == len(series.cells)


def test_matrix_domain_is_country_by_year_key_points(tmp_path: Path) -> None:
    workbook = _profile_workbook(tmp_path)
    catalog = build_catalog(validate_bindings_document(_profile_bindings()), workbook=workbook)
    series = catalog.get("profile_table")
    assert [point.as_mapping() for point in series.domain] == [
        {"COUNTRY": "France", "TIME_PERIOD": 2020},
        {"COUNTRY": "France", "TIME_PERIOD": 2021},
        {"COUNTRY": "Kenya", "TIME_PERIOD": 2020},
        {"COUNTRY": "Kenya", "TIME_PERIOD": 2021},
    ]


def test_mixed_formulas_partition_into_one_statement_per_shape_run(tmp_path: Path) -> None:
    workbook = _a12_workbook(tmp_path)
    bindings = validate_bindings_document(_a12_bindings())
    graph = create_dependency_graph(
        workbook,
        all_series_targets(bindings, workbook=workbook),
        load_values=True,
    )
    catalog = build_catalog(bindings, workbook=workbook, graph=graph)
    series = catalog.get("path")
    assert [stmt.shape_key for stmt in series.statements] == [
        series.statements[0].shape_key,
        series.statements[1].shape_key,
        series.statements[2].shape_key,
    ]
    assert len({stmt.shape_key for stmt in series.statements}) == 3
    assert [(stmt.start, stmt.stop) for stmt in series.statements] == [(0, 1), (1, 2), (2, 3)]
    assert [stmt.statement_id for stmt in series.statements] == [
        "path__0",
        "path__1",
        "path__2",
    ]
    assert series.statements[0].domain[0]["TIME_PERIOD"] == 2009


def test_uniform_formula_series_is_one_statement(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "uniform.xlsx",
        {
            "Engine": {
                "A1": 2009,
                "B1": 2010,
                "A2": "=1",
                "B2": "=1",
            },
        },
    )
    document = bindings_document(
        series_entry(
            "path",
            "Engine!A2:B2",
            layout="series",
            direction="output",
            header_row=1,
        ),
    )
    bindings = validate_bindings_document(document)
    graph = create_dependency_graph(
        workbook,
        all_series_targets(bindings, workbook=workbook),
        load_values=True,
    )
    catalog = build_catalog(bindings, workbook=workbook, graph=graph)
    series = catalog.get("path")
    assert len(series.statements) == 1
    assert series.statements[0].statement_id == "path"
    assert series.statements[0].shape_key is not None
    assert series.statements[0].start == 0
    assert series.statements[0].stop == 2


def test_typo_bind_kind_fails_closed_instead_of_empty_domain(tmp_path: Path) -> None:
    workbook = _a12_workbook(tmp_path)
    document = _a12_bindings()
    document["series"][0]["structure"]["dimensions"][0]["bind"]["kind"] = "colum_header"
    entry = document["series"][0]
    with pytest.raises(ValueError, match="key field") as exc:
        resolve_key_domain(workbook, entry, ("Engine!A2", "Engine!B2", "Engine!C2"))
    assert isinstance(exc.value.__cause__, UnknownBindKindError)
    with pytest.raises(InvertedTreeExportError, match="key field"):
        build_catalog(document, workbook=workbook)


def test_partition_catalog_large_constant_series_performance() -> None:
    n = 200_000
    cells = tuple(f"Sheet1!A{i}" for i in range(1, n + 1))
    domain = tuple(KeyPoint((("idx", i),)) for i in range(n))
    series = BoundSeries(
        series_id="large_const",
        layout="series",
        direction="constant",
        cells=cells,
        key_fields=("idx",),
        dtype="float",
        compute_name=None,
        raw={},
        domain=domain,
        statements=(
            Statement(
                statement_id="large_const",
                series_id="large_const",
                shape_key=None,
                start=0,
                stop=n,
                cells=cells,
                domain=domain,
            ),
        ),
    )
    catalog = make_catalog(
        series={"large_const": series},
        order=("large_const",),
        address_to_id={cell: "large_const" for cell in cells},
    )
    graph = DependencyGraph()
    start_time = time.perf_counter()
    partitioned = partition_catalog(catalog, graph)
    elapsed = time.perf_counter() - start_time
    assert elapsed < 1.0
    res = partitioned.get("large_const")
    assert len(res.statements) == 1
    assert res.statements[0].statement_id == "large_const"
    assert res.statements[0].start == 0
    assert res.statements[0].stop == n


def _make_formula_catalog_and_graph(
    n: int,
) -> tuple[SeriesCatalog, DependencyGraph]:
    graph = DependencyGraph()
    src_cells = tuple(f"Sheet1!A{i}" for i in range(1, n + 1))
    dst_cells = tuple(f"Sheet1!B{i}" for i in range(1, n + 1))
    for i in range(1, n + 1):
        graph.add_node(make_cell_node("Sheet1", "A", i, value=1.0, is_leaf=True))
        addr = f"Sheet1!B{i}"
        f_text = f"=A{i}+1"
        ast = parse_formula_text(f_text, anchor=addr)
        graph.add_node(make_cell_node("Sheet1", "B", i, formula=f_text, formula_ast=ast))
    src_domain = tuple(KeyPoint((("idx", i),)) for i in range(n))
    dst_domain = tuple(KeyPoint((("idx", i),)) for i in range(n))
    src_series = BoundSeries(
        series_id="src",
        layout="series",
        direction="input",
        cells=src_cells,
        key_fields=("idx",),
        dtype="float",
        compute_name=None,
        raw={},
        domain=src_domain,
        statements=(Statement("src", "src", None, 0, n, src_cells, src_domain),),
    )
    dst_series = BoundSeries(
        series_id="dst",
        layout="series",
        direction="output",
        cells=dst_cells,
        key_fields=("idx",),
        dtype="float",
        compute_name=None,
        raw={},
        domain=dst_domain,
        statements=(Statement("dst", "dst", None, 0, n, dst_cells, dst_domain),),
    )
    addr_to_id = {c: "src" for c in src_cells}
    addr_to_id.update({c: "dst" for c in dst_cells})
    catalog = make_catalog(
        series={"src": src_series, "dst": dst_series},
        order=("src", "dst"),
        address_to_id=addr_to_id,
    )
    return catalog, graph


def test_partition_catalog_uniform_formula_is_linear_in_size() -> None:
    cat_small, g_small = _make_formula_catalog_and_graph(2_500)
    cat_large, g_large = _make_formula_catalog_and_graph(10_000)

    # Warmup
    partition_catalog(cat_small, g_small)

    t0 = time.perf_counter()
    res_small = partition_catalog(cat_small, g_small)
    t_small = time.perf_counter() - t0

    t0 = time.perf_counter()
    res_large = partition_catalog(cat_large, g_large)
    t_large = time.perf_counter() - t0

    assert len(res_small.get("dst").statements) == 1
    assert len(res_large.get("dst").statements) == 1
    # 4x increase in data should be roughly 4x-5x in runtime (< 7x tolerance for timer noise)
    if t_small > 0.001:
        ratio = t_large / t_small
        assert ratio < 7.0


def test_partition_catalog_splits_on_shape_and_producer_changes() -> None:
    graph = DependencyGraph()
    for i in range(1, 10):
        graph.add_node(make_cell_node("Sheet1", "A", i, value=1.0, is_leaf=True))
        graph.add_node(make_cell_node("Sheet1", "B", i, value=2.0, is_leaf=True))

    # dst:
    # 1: =A1+1 (shape1, prod 'src_a')
    # 2: =A2+1 (shape1, prod 'src_a') -> run [0, 2)
    # 3: =A3*2 (shape2, prod 'src_a') -> run [2, 3) (shape change)
    # 4: =B4*2 (shape2, prod 'src_b') -> run [3, 4) (producer change)
    # 5: =B5*2 (shape2, prod 'src_b') -> run [3, 5)
    formulas = ["=A1+1", "=A2+1", "=A3*2", "=B4*2", "=B5*2"]
    dst_cells = tuple(f"Sheet1!C{i}" for i in range(1, 6))
    for i, formula in enumerate(formulas, start=1):
        addr = f"Sheet1!C{i}"
        ast = parse_formula_text(formula, anchor=addr)
        graph.add_node(make_cell_node("Sheet1", "C", i, formula=formula, formula_ast=ast))

    src_a_cells = tuple(f"Sheet1!A{i}" for i in range(1, 6))
    src_b_cells = tuple(f"Sheet1!B{i}" for i in range(1, 6))
    domain = tuple(KeyPoint((("idx", i),)) for i in range(5))

    src_a = BoundSeries(
        series_id="src_a",
        layout="series",
        direction="input",
        cells=src_a_cells,
        key_fields=("idx",),
        dtype="float",
        compute_name=None,
        raw={},
        domain=domain,
        statements=(Statement("src_a", "src_a", None, 0, 5, src_a_cells, domain),),
    )
    src_b = BoundSeries(
        series_id="src_b",
        layout="series",
        direction="input",
        cells=src_b_cells,
        key_fields=("idx",),
        dtype="float",
        compute_name=None,
        raw={},
        domain=domain,
        statements=(Statement("src_b", "src_b", None, 0, 5, src_b_cells, domain),),
    )
    dst = BoundSeries(
        series_id="dst",
        layout="series",
        direction="output",
        cells=dst_cells,
        key_fields=("idx",),
        dtype="float",
        compute_name=None,
        raw={},
        domain=domain,
        statements=(Statement("dst", "dst", None, 0, 5, dst_cells, domain),),
    )
    addr_to_id = {c: "src_a" for c in src_a_cells}
    addr_to_id.update({c: "src_b" for c in src_b_cells})
    addr_to_id.update({c: "dst" for c in dst_cells})
    catalog = make_catalog(
        series={"src_a": src_a, "src_b": src_b, "dst": dst},
        order=("src_a", "src_b", "dst"),
        address_to_id=addr_to_id,
    )

    partitioned = partition_catalog(catalog, graph)
    statements = partitioned.get("dst").statements
    assert len(statements) == 3
    assert [(s.start, s.stop) for s in statements] == [(0, 2), (2, 3), (3, 5)]
    assert [s.statement_id for s in statements] == ["dst__0", "dst__2", "dst__3"]


def test_partition_catalog_handles_unbound_cell_refs() -> None:
    graph = DependencyGraph()
    # dst references Z1 which is not in any bound series
    formulas = ["=Z1+1", "=Z1+1"]
    dst_cells = tuple(f"Sheet1!C{i}" for i in range(1, 3))
    for i, formula in enumerate(formulas, start=1):
        addr = f"Sheet1!C{i}"
        ast = parse_formula_text(formula, anchor=addr)
        graph.add_node(make_cell_node("Sheet1", "C", i, formula=formula, formula_ast=ast))

    domain = tuple(KeyPoint((("idx", i),)) for i in range(2))
    dst = BoundSeries(
        series_id="dst",
        layout="series",
        direction="output",
        cells=dst_cells,
        key_fields=("idx",),
        dtype="float",
        compute_name=None,
        raw={},
        domain=domain,
        statements=(Statement("dst", "dst", None, 0, 2, dst_cells, domain),),
    )
    catalog = make_catalog(
        series={"dst": dst},
        order=("dst",),
        address_to_id={c: "dst" for c in dst_cells},
    )

    partitioned = partition_catalog(catalog, graph)
    statements = partitioned.get("dst").statements
    # Both share shape and have ('?', None) access pair -> grouped in 1 statement
    assert len(statements) == 1
    assert statements[0].statement_id == "dst"
    assert statements[0].start == 0
    assert statements[0].stop == 2


def test_partition_catalog_empty_and_single_cell_series() -> None:
    graph = DependencyGraph()
    empty_series = BoundSeries(
        series_id="empty",
        layout="series",
        direction="output",
        cells=(),
        key_fields=(),
        dtype="float",
        compute_name=None,
        raw={},
        domain=(),
        statements=(Statement("empty", "empty", None, 0, 0, (), ()),),
    )
    single_cell = ("Sheet1!A1",)
    single_domain = (KeyPoint((("idx", 0),)),)
    graph.add_node(
        make_cell_node(
            "Sheet1", "A", 1, formula="=1", formula_ast=parse_formula_text("=1", anchor="Sheet1!A1")
        )
    )
    single_series = BoundSeries(
        series_id="single",
        layout="scalar",
        direction="output",
        cells=single_cell,
        key_fields=("idx",),
        dtype="float",
        compute_name=None,
        raw={},
        domain=single_domain,
        statements=(Statement("single", "single", None, 0, 1, single_cell, single_domain),),
    )
    catalog = make_catalog(
        series={"empty": empty_series, "single": single_series},
        order=("empty", "single"),
        address_to_id={"Sheet1!A1": "single"},
    )
    partitioned = partition_catalog(catalog, graph)
    assert len(partitioned.get("empty").statements) == 1
    assert len(partitioned.get("single").statements) == 1
    assert partitioned.get("single").statements[0].statement_id == "single"


def test_bound_series_index_of() -> None:
    cells = ("Sheet1!A1", "Sheet1!A2", "Sheet1!A3")
    domain = tuple(KeyPoint((("idx", i),)) for i in range(3))
    series = BoundSeries(
        series_id="s",
        layout="series",
        direction="output",
        cells=cells,
        key_fields=("idx",),
        dtype="float",
        compute_name=None,
        raw={},
        domain=domain,
        statements=(Statement("s", "s", None, 0, 3, cells, domain),),
    )
    assert series.index_of("Sheet1!A1") == 0
    assert series.index_of("Sheet1!$A$2") == 1
    assert series.index_of("Sheet1!a3") == 2
    assert series.index_of("Sheet1!A4") is None
