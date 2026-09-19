"""nnz-aligned intern-id sidecars for guards and provenance (#920)."""

from __future__ import annotations

import copy
import pickle
from pathlib import Path
from typing import cast

import pytest

from excel_grapher import create_dependency_graph
from excel_grapher.grapher.cache import dependency_graph_from_json, dependency_graph_to_json
from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.guard import CellRef, Compare, GuardExpr, Literal
from excel_grapher.grapher.node import make_cell_node
from scripts.measure_graph_memory import (
    DEFAULT_TARGETS,
    DEFAULT_WORKBOOK,
    IF_HEAVY_TARGETS,
    IF_HEAVY_WORKBOOK,
    measure_graph_memory,
)


def _leaf(column: str, row: int, value: object = 1):
    return make_cell_node("Sheet1", column, row, value=value, is_leaf=True)


def _formula(column: str, row: int):
    return make_cell_node("Sheet1", column, row, normalized_formula="=1", is_leaf=False)


def _guard(key: str = "Sheet1!A1") -> Compare:
    return Compare(CellRef(key), ">", Literal(0))


def _prov(start: int, end: int) -> EdgeProvenance:
    return EdgeProvenance(
        causes=DependencyCause.direct_ref,
        direct_sites_normalized=((start, end),),
    )


def _meta_graph() -> DependencyGraph:
    graph = DependencyGraph()
    graph.add_node(_leaf("A", 1))
    graph.add_node(_formula("B", 1))
    graph.add_node(_formula("C", 1))
    shared = _guard()
    graph.add_edge("Sheet1!B1", "Sheet1!A1", guard=shared, provenance=_prov(1, 11))
    graph.add_edge("Sheet1!C1", "Sheet1!A1", guard=shared, provenance=_prov(1, 11))
    graph.add_edge("Sheet1!C1", "Sheet1!B1")
    return graph


def test_rebuild_drops_edge_key_maps_and_keeps_public_attrs() -> None:
    graph = _meta_graph()
    assert graph._guards
    assert graph._edge_provenance

    graph.rebuild_adjacency()

    assert graph._staging is False
    assert graph._guards == {}
    assert graph._edge_provenance == {}
    assert len(graph._guard_id) == graph.edge_count()
    assert len(graph._prov_id) == graph.edge_count()
    assert graph._guard_id.typecode == "I"
    assert graph._prov_id.typecode == "I"
    assert graph.is_guarded("Sheet1!B1", "Sheet1!A1")
    assert graph.is_guarded("Sheet1!C1", "Sheet1!A1")
    assert not graph.is_guarded("Sheet1!C1", "Sheet1!B1")
    assert graph.get_edge_guard("Sheet1!B1", "Sheet1!A1") is graph.get_edge_guard(
        "Sheet1!C1", "Sheet1!A1"
    )
    assert graph.get_edge_attrs("Sheet1!C1", "Sheet1!B1") == graph.get_edge_attrs(
        "Sheet1!C1", "Sheet1!B1"
    )
    assert graph.get_edge_attrs("Sheet1!C1", "Sheet1!B1").guard is None
    assert graph.get_edge_attrs("Sheet1!C1", "Sheet1!B1").provenance is None
    assert graph.get_edge_attrs("Sheet1!B1", "Sheet1!A1").provenance == _prov(1, 11)


def test_guard_id_is_aligned_with_csr_nnz_not_csc() -> None:
    graph = _meta_graph()
    graph.rebuild_adjacency()
    keys = graph._csr_keys
    for i, src in enumerate(keys):
        for k in range(graph._row_ptr[i], graph._row_ptr[i + 1]):
            dst = keys[graph._col_idx[k]]
            assert (graph._guard_id[k] != 0) == graph.is_guarded(src, dst)
            gid = graph._guard_id[k]
            expected = graph.get_edge_guard(src, dst)
            if gid == 0:
                assert expected is None
            else:
                assert graph._guard_exprs[gid] is expected
            pid = graph._prov_id[k]
            prov = graph.get_edge_attrs(src, dst).provenance
            if pid == 0:
                assert prov is None
            else:
                assert graph._provenances[pid] == prov


def test_shared_guards_share_one_intern_id() -> None:
    graph = _meta_graph()
    graph.rebuild_adjacency()
    used = {gid for gid in graph._guard_id if gid}
    assert used == {1}
    assert sum(expr is not None for expr in graph._guard_exprs) == 1
    equal_prov = {pid for pid in graph._prov_id if pid and graph._provenances[pid] == _prov(1, 11)}
    assert len(equal_prov) == 1


def test_is_guarded_does_not_hydrate_interned_ast_when_compact() -> None:
    graph = _meta_graph()
    graph.rebuild_adjacency()

    class Boom:
        def __getitem__(self, index: object) -> object:
            raise AssertionError("is_guarded must not hydrate interned guard ASTs")

    graph._guard_exprs = cast(list[GuardExpr | None], Boom())
    assert graph.is_guarded("Sheet1!B1", "Sheet1!A1")
    assert graph.is_guarded("Sheet1!C1", "Sheet1!A1")
    assert not graph.is_guarded("Sheet1!C1", "Sheet1!B1")
    assert not graph.is_guarded("Sheet1!A1", "Sheet1!B1")


def test_dangling_guarded_endpoint_is_dropped_on_rebuild() -> None:
    graph = DependencyGraph()
    graph.add_node(_leaf("A", 1))
    graph.add_node(_formula("B", 1))
    graph.add_edge("Sheet1!B1", "Sheet1!A1", guard=_guard(), provenance=_prov(1, 11))
    graph.add_edge("Sheet1!B1", "Sheet1!Z9", guard=_guard("Sheet1!Z9"))
    assert graph.is_guarded("Sheet1!B1", "Sheet1!Z9")

    graph.rebuild_adjacency()

    assert graph.get_dependencies("Sheet1!B1") == frozenset({"Sheet1!A1"})
    assert not graph.is_guarded("Sheet1!B1", "Sheet1!Z9")
    assert graph.get_edge_guard("Sheet1!B1", "Sheet1!Z9") is None
    assert graph.is_guarded("Sheet1!B1", "Sheet1!A1")
    assert graph._guards == {}


def test_add_edge_after_rebuild_rehydrates_metadata_maps() -> None:
    graph = _meta_graph()
    graph.rebuild_adjacency()
    graph.add_node(_formula("D", 1))
    extra = _guard("Sheet1!B1")
    graph.add_edge("Sheet1!C1", "Sheet1!D1", guard=extra, provenance=_prov(2, 12))
    assert graph._staging is True
    assert graph.is_guarded("Sheet1!B1", "Sheet1!A1")
    assert graph.is_guarded("Sheet1!C1", "Sheet1!D1")
    assert graph.get_edge_attrs("Sheet1!C1", "Sheet1!D1").provenance == _prov(2, 12)
    graph.rebuild_adjacency()
    assert graph._staging is False
    assert graph._guards == {}
    assert graph.is_guarded("Sheet1!C1", "Sheet1!D1")
    assert graph.get_edge_guard("Sheet1!C1", "Sheet1!D1") == extra


def test_compact_rebuild_preserves_metadata_without_reaccumulating_maps() -> None:
    graph = _meta_graph()
    graph.rebuild_adjacency()
    graph.rebuild_adjacency()
    assert graph._staging is False
    assert graph._guards == {}
    assert graph._edge_provenance == {}
    assert graph.is_guarded("Sheet1!B1", "Sheet1!A1")
    assert graph.get_edge_attrs("Sheet1!C1", "Sheet1!A1").provenance == _prov(1, 11)


def test_copy_for_projection_clones_meta_arrays() -> None:
    graph = _meta_graph()
    graph.rebuild_adjacency()
    cloned = graph._copy_for_projection()
    assert cloned._staging is False
    assert cloned._guards == {}
    assert cloned.is_guarded("Sheet1!B1", "Sheet1!A1")
    assert cloned._guard_id is not graph._guard_id
    assert cloned._prov_id is not graph._prov_id
    assert cloned._guard_exprs is not graph._guard_exprs
    assert cloned._provenances is not graph._provenances
    cloned._guard_id[0] = 0 if cloned._guard_id[0] != 0 else 1
    assert cloned._guard_id[0] != graph._guard_id[0]


def test_deepcopy_json_and_pickle_round_trip_compact_meta() -> None:
    graph = _meta_graph()
    graph.rebuild_adjacency()
    pickled: DependencyGraph = pickle.loads(pickle.dumps(graph))
    deep = copy.deepcopy(graph)
    json_restored = dependency_graph_from_json(dependency_graph_to_json(graph))
    for restored in (pickled, deep, json_restored):
        assert restored._staging is False
        assert restored._guards == {}
        assert restored._edge_provenance == {}
        for src in graph:
            for dst in graph.get_dependencies(src):
                assert restored.is_guarded(src, dst) == graph.is_guarded(src, dst)
                assert restored.get_edge_guard(src, dst) == graph.get_edge_guard(src, dst)
                assert (
                    restored.get_edge_attrs(src, dst).provenance
                    == graph.get_edge_attrs(src, dst).provenance
                )


def test_cycle_report_uses_compact_guards_without_rehydrating_maps() -> None:
    graph = DependencyGraph()
    graph.add_node(_formula("A", 1))
    graph.add_node(_formula("B", 1))
    graph.add_edge("Sheet1!A1", "Sheet1!B1", guard=_guard("Sheet1!B1"))
    graph.add_edge("Sheet1!B1", "Sheet1!A1", guard=_guard("Sheet1!A1"))
    graph.rebuild_adjacency()
    report = graph.cycle_report()
    assert report.has_may_cycles
    assert graph._edges == {}
    assert graph._guards == {}
    assert graph._staging is False


def test_compact_metadata_drops_edge_key_scaffold() -> None:
    graph = DependencyGraph()
    shared = _guard()
    n = 80
    for row in range(1, n + 1):
        graph.add_node(_leaf("A", row, value=float(row)))
        graph.add_node(_formula("B", row))
        graph.add_edge(
            f"Sheet1!B{row}",
            f"Sheet1!A{row}",
            guard=shared,
            provenance=_prov(row, row + 10),
        )
    staging = measure_graph_memory(graph)
    graph.rebuild_adjacency()
    compact = measure_graph_memory(graph)
    assert graph._guards == {}
    assert graph._edge_provenance == {}
    assert compact.component("guards").scaffolding_bytes < (
        staging.component("guards").scaffolding_bytes / 4
    )
    assert compact.component("guards").exclusive_bytes < (
        staging.component("guards").exclusive_bytes / 2
    )
    assert (
        compact.component("provenance").exclusive_bytes
        < staging.component("provenance").exclusive_bytes
    )
    assert (
        compact.component("provenance").scaffolding_bytes
        < staging.component("provenance").scaffolding_bytes
    )


def test_v6_pickle_blob_compacts_edge_key_lists_on_load(tmp_path: Path) -> None:
    """Old CSR pickles stored EdgeKey lists; load must install intern-id arrays."""
    from excel_grapher.grapher.graph import _collect_graph_keys
    from excel_grapher.grapher.graph_pickle import _GRAPH_BLOB_MAGIC, load_graph

    graph = _meta_graph()
    graph.rebuild_adjacency()
    keys_sorted = _collect_graph_keys(graph)
    idx = {k: i for i, k in enumerate(keys_sorted)}
    node_keys = list(graph._csr_keys)
    nodes = [graph._nodes[k] for k in node_keys]
    path = tmp_path / "v6-edge-meta.pkl"
    with path.open("wb") as handle:
        handle.write(_GRAPH_BLOB_MAGIC)
        handle.write((6).to_bytes(4, "little"))
        pickle.dump(
            {
                "keys": keys_sorted,
                "node_keys": node_keys,
                "nodes": nodes,
                "_hooks": [],
                "leaf_classification": None,
                "sheet_order": None,
                "named_ranges": None,
                "named_range_ranges": None,
            },
            handle,
            protocol=pickle.HIGHEST_PROTOCOL,
        )
        pickle.dump(
            {
                "row_ptr": graph._row_ptr,
                "col_idx": graph._col_idx,
                "col_ptr": graph._col_ptr,
                "row_idx": graph._row_idx,
                "_guards": [
                    (
                        idx["Sheet1!B1"],
                        idx["Sheet1!A1"],
                        graph.get_edge_guard("Sheet1!B1", "Sheet1!A1"),
                    ),
                    (
                        idx["Sheet1!C1"],
                        idx["Sheet1!A1"],
                        graph.get_edge_guard("Sheet1!C1", "Sheet1!A1"),
                    ),
                ],
                "_edge_provenance": [
                    (idx["Sheet1!B1"], idx["Sheet1!A1"], _prov(1, 11)),
                    (idx["Sheet1!C1"], idx["Sheet1!A1"], _prov(1, 11)),
                ],
            },
            handle,
            protocol=pickle.HIGHEST_PROTOCOL,
        )
    restored = load_graph(path)
    assert restored._staging is False
    assert restored._guards == {}
    assert restored.is_guarded("Sheet1!B1", "Sheet1!A1")
    assert restored.get_edge_attrs("Sheet1!C1", "Sheet1!A1").provenance == _prov(1, 11)
    assert restored._guard_id.typecode == "I"


def _assert_extract_is_compact(graph: DependencyGraph) -> None:
    assert graph._staging is False
    assert graph._guards == {}
    assert graph._edge_provenance == {}
    assert graph._edges == {}
    assert graph._reverse_edges == {}


@pytest.mark.skipif(not DEFAULT_WORKBOOK.is_file(), reason="taco_patterns.xlsx fixture missing")
def test_taco_extract_rebuild_drops_edge_key_maps() -> None:
    graph = create_dependency_graph(
        DEFAULT_WORKBOOK,
        DEFAULT_TARGETS,
        load_values=False,
        capture_dependency_provenance=True,
    )
    _assert_extract_is_compact(graph)
    report = measure_graph_memory(graph)
    # Character-span intern table collapses; leftover is a few KiB, not a map.
    assert report.identity_distinct_provenances < report.edge_count
    assert report.component("provenance").exclusive_bytes < 8_192


@pytest.mark.skipif(not IF_HEAVY_WORKBOOK.is_file(), reason="if_guards.xlsx fixture missing")
def test_if_guards_extract_rebuild_drops_edge_key_maps() -> None:
    graph = create_dependency_graph(
        IF_HEAVY_WORKBOOK,
        IF_HEAVY_TARGETS,
        load_values=False,
        capture_dependency_provenance=True,
    )
    _assert_extract_is_compact(graph)
    report = measure_graph_memory(graph)
    assert report.guarded_edge_count > 0
    assert report.identity_distinct_guards == report.guarded_edge_count
    assert report.identity_distinct_provenances == 1
    assert report.component("provenance").exclusive_bytes < 8_192
    # Guard leftover is interned GuardExpr payload, not EdgeKey maps.
    assert report.component("guards").scaffolding_bytes < report.component("guards").exclusive_bytes
