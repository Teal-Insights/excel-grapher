"""CSR+CSC adjacency sidecar spike (#908)."""

from __future__ import annotations

import array
import sys
import tracemalloc

from excel_grapher.grapher.csr_adjacency import (
    LANDING_API_BREAKS,
    CsrCscAdjacency,
    build_csr_csc,
    from_graph,
)
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import make_cell_node
from scripts.measure_graph_memory import deep_size, measure_graph_memory


def _leaf(column: str, row: int, value: object = 1):
    return make_cell_node("Sheet1", column, row, value=value, is_leaf=True)


def _formula(column: str, row: int):
    return make_cell_node("Sheet1", column, row, normalized_formula="=1", is_leaf=False)


def _add(graph: DependencyGraph, *nodes) -> None:
    for node in nodes:
        graph.add_node(node)


def _resolved_deps(graph: DependencyGraph, key: str) -> frozenset[str]:
    return frozenset(dep for dep in graph.get_dependencies(key) if dep in graph._nodes)


def _resolved_dependents(graph: DependencyGraph, key: str) -> frozenset[str]:
    return frozenset(dep for dep in graph.get_dependents(key) if dep in graph._nodes)


def _chain_graph(n: int, extra_degree: int = 3) -> DependencyGraph:
    """Build n formula cells that each read a leaf plus earlier formulas."""
    graph = DependencyGraph()
    for row in range(1, n + 1):
        graph.add_node(_leaf("A", row, value=float(row)))
        graph.add_node(_formula("B", row))
        graph.add_edge(f"Sheet1!B{row}", f"Sheet1!A{row}")
        for back in range(1, extra_degree + 1):
            src = row - back
            if src >= 1:
                graph.add_edge(f"Sheet1!B{row}", f"Sheet1!B{src}")
    return graph


def _array_bytes(csr: CsrCscAdjacency) -> int:
    return sum(sys.getsizeof(part) for part in (csr.row_ptr, csr.col_idx, csr.col_ptr, csr.row_idx))


def _index_table_bytes(csr: CsrCscAdjacency) -> int:
    """Exclusive-ish cost of `index` + `keys` without charging shared NodeKey strings."""
    string_ids = {id(key) for key in csr.keys}
    total = sys.getsizeof(csr.keys) + sys.getsizeof(csr.index)
    for key, value in csr.index.items():
        if id(key) not in string_ids:
            total += sys.getsizeof(key)
        total += sys.getsizeof(value)
    return total


# ---- correctness ----------------------------------------------------------


def test_empty_graph_has_zero_nnz() -> None:
    csr = from_graph(DependencyGraph())
    assert csr.n == 0
    assert csr.nnz == 0
    assert csr.dropped_endpoints == 0
    assert list(csr.row_ptr) == [0]
    assert list(csr.col_ptr) == [0]
    assert csr.stats.built_second_python_edge_index is False


def test_isolated_nodes_have_empty_rows_and_columns() -> None:
    graph = DependencyGraph()
    _add(graph, _leaf("A", 1), _leaf("A", 2))
    csr = from_graph(graph)
    assert csr.n == 2
    assert csr.nnz == 0
    assert csr.dependencies("Sheet1!A1") == frozenset()
    assert csr.dependents("Sheet1!A2") == frozenset()
    assert list(csr.row_ptr) == [0, 0, 0]
    assert list(csr.col_ptr) == [0, 0, 0]


def test_sidecar_matches_resolved_forward_and_reverse() -> None:
    graph = DependencyGraph()
    _add(graph, _leaf("A", 1), _formula("B", 1), _formula("C", 1))
    graph.add_edge("Sheet1!B1", "Sheet1!A1")
    graph.add_edge("Sheet1!C1", "Sheet1!B1")
    graph.add_edge("Sheet1!C1", "Sheet1!A1")

    csr = from_graph(graph)
    for key in graph:
        assert csr.dependencies(key) == _resolved_deps(graph, key)
        assert csr.dependents(key) == _resolved_dependents(graph, key)
    assert csr.nnz == 3
    assert csr.dropped_endpoints == 0


def test_dangling_endpoints_are_dropped_not_indexed() -> None:
    graph = DependencyGraph()
    _add(graph, _leaf("A", 1), _formula("B", 1))
    graph.add_edge("Sheet1!B1", "Sheet1!A1")
    graph.add_edge("Sheet1!B1", "Sheet1!Z9")

    csr = from_graph(graph)
    assert "Sheet1!Z9" not in csr.index
    assert csr.dropped_endpoints == 1
    assert csr.dependencies("Sheet1!B1") == frozenset({"Sheet1!A1"})
    assert csr.dependents("Sheet1!A1") == frozenset({"Sheet1!B1"})
    assert graph.get_dependencies("Sheet1!B1") == frozenset({"Sheet1!A1", "Sheet1!Z9"})


def test_csc_is_transpose_of_csr() -> None:
    graph = _chain_graph(12, extra_degree=4)
    csr = from_graph(graph)
    forward: set[tuple[int, int]] = set()
    for i in range(csr.n):
        for k in range(csr.row_ptr[i], csr.row_ptr[i + 1]):
            forward.add((i, csr.col_idx[k]))
    reverse: set[tuple[int, int]] = set()
    for j in range(csr.n):
        for k in range(csr.col_ptr[j], csr.col_ptr[j + 1]):
            reverse.add((csr.row_idx[k], j))
    assert forward == reverse
    assert len(forward) == csr.nnz


def test_unknown_key_reads_as_empty() -> None:
    csr = from_graph(DependencyGraph())
    assert csr.dependencies("Sheet1!A1") == frozenset()
    assert csr.dependents("Sheet1!A1") == frozenset()


def test_compact_builder_does_not_need_a_reverse_map() -> None:
    keys = ("a", "b", "c")
    neighbors = {"c": ("b", "a"), "b": ("a",), "a": ()}
    csr = build_csr_csc(keys, lambda key: neighbors[key])
    assert csr.dependencies("c") == frozenset({"a", "b"})
    assert csr.dependents("a") == frozenset({"b", "c"})
    assert csr.stats.held_input_adjacency is False
    assert csr.stats.built_second_python_edge_index is False


def test_from_graph_does_not_copy_neighbor_sets() -> None:
    graph = _chain_graph(8)
    before = {key: id(neighbors) for key, neighbors in graph._edges.items()}
    from_graph(graph)
    after = {key: id(neighbors) for key, neighbors in graph._edges.items()}
    assert before == after
    assert graph._reverse_edges  # still the production reverse index


def test_build_stats_scratch_is_on_nodes_not_nnz() -> None:
    graph = _chain_graph(40, extra_degree=5)
    csr = from_graph(graph)
    assert csr.nnz > csr.n
    assert csr.stats.scratch_uint32 == csr.n
    assert csr.stats.held_input_adjacency is True
    assert csr.stats.built_second_python_edge_index is False


def test_landing_api_breaks_are_documented() -> None:
    assert len(LANDING_API_BREAKS) >= 8
    joined = " ".join(LANDING_API_BREAKS)
    assert "add_edge" in joined
    assert "evaluation_order" in joined
    assert "_copy_for_projection" in joined


def test_arrays_are_uint32() -> None:
    csr = from_graph(_chain_graph(3))
    for part in (csr.row_ptr, csr.col_idx, csr.col_ptr, csr.row_idx):
        assert isinstance(part, array.array)
        assert part.typecode == "I"


def test_visit_checksums_match_manual_sums() -> None:
    csr = from_graph(_chain_graph(15, extra_degree=3))
    assert csr.visit_forward_ids() == sum(csr.col_idx)
    assert csr.visit_reverse_ids() == sum(csr.row_idx)


def test_walk_visit_counts_match_resolved_edges() -> None:
    graph = _chain_graph(25, extra_degree=4)
    csr = from_graph(graph)
    dict_arcs = sum(len(_resolved_deps(graph, key)) for key in graph)
    assert dict_arcs == csr.nnz
    api_arcs = sum(len(csr.dependencies(key)) for key in graph)
    assert api_arcs == csr.nnz


# ---- memory ---------------------------------------------------------------


def test_csr_arrays_are_far_smaller_than_dict_adjacency() -> None:
    graph = _chain_graph(250, extra_degree=6)
    report = measure_graph_memory(graph)
    adj_exclusive = (
        report.component("edges_forward").exclusive_bytes
        + report.component("edges_reverse").exclusive_bytes
    )
    csr = from_graph(graph)
    arrays = _array_bytes(csr)
    index_table = _index_table_bytes(csr)
    assert arrays < adj_exclusive / 8
    # The NodeKey -> row-id dict is the extra tax #472's 6 MiB figure omitted.
    assert index_table > 0
    assert arrays + index_table < adj_exclusive / 2


def test_csr_reuses_node_map_key_objects() -> None:
    graph = _chain_graph(20)
    csr = from_graph(graph)
    for stored, sidecar in zip(graph._nodes, csr.keys, strict=True):
        assert stored is sidecar
    assert deep_size(csr.row_ptr, csr.col_idx, csr.col_ptr, csr.row_idx) == _array_bytes(csr)


def test_build_peak_is_not_a_second_python_edge_index() -> None:
    graph = _chain_graph(400, extra_degree=6)
    report = measure_graph_memory(graph)
    adj_exclusive = (
        report.component("edges_forward").exclusive_bytes
        + report.component("edges_reverse").exclusive_bytes
    )
    tracemalloc.start()
    csr = from_graph(graph)
    _current, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    # A COO list of Python tuples would land near adjacency exclusive; compact
    # uint32 fill + O(n) scratch must stay well below that.
    assert peak < adj_exclusive
    assert _array_bytes(csr) < peak
    assert csr.stats.built_second_python_edge_index is False
