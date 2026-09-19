"""Production CSR+CSC adjacency storage on `DependencyGraph` (#915)."""

from __future__ import annotations

import array
import copy
import pickle
import sys
import tracemalloc

from excel_grapher.grapher.cache import dependency_graph_from_json, dependency_graph_to_json
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


def _csr_array_bytes(graph: DependencyGraph) -> int:
    return sum(
        sys.getsizeof(part)
        for part in (graph._row_ptr, graph._col_idx, graph._col_ptr, graph._row_idx)
    )


def test_add_edge_is_visible_before_rebuild() -> None:
    graph = DependencyGraph()
    _add(graph, _leaf("A", 1), _formula("B", 1))
    graph.add_edge("Sheet1!B1", "Sheet1!A1")
    assert graph.get_dependencies("Sheet1!B1") == frozenset({"Sheet1!A1"})
    assert graph.get_dependents("Sheet1!A1") == frozenset({"Sheet1!B1"})
    assert graph._staging is True


def test_rebuild_drops_dict_maps_and_keeps_public_neighbors() -> None:
    graph = DependencyGraph()
    _add(graph, _leaf("A", 1), _formula("B", 1), _formula("C", 1))
    graph.add_edge("Sheet1!B1", "Sheet1!A1")
    graph.add_edge("Sheet1!C1", "Sheet1!B1")
    graph.add_edge("Sheet1!C1", "Sheet1!A1")

    graph.rebuild_adjacency()

    assert graph._staging is False
    assert graph._edges == {}
    assert graph._reverse_edges == {}
    assert graph.edge_count() == 3
    for key in graph:
        assert isinstance(graph.get_dependencies(key), frozenset)
    assert graph.get_dependencies("Sheet1!C1") == frozenset({"Sheet1!B1", "Sheet1!A1"})
    assert graph.get_dependents("Sheet1!A1") == frozenset({"Sheet1!B1", "Sheet1!C1"})
    assert graph.get_dependencies("Sheet1!A1") == frozenset()
    assert graph.get_dependents("Sheet1!C1") == frozenset()


def test_isolated_nodes_have_empty_csr_rows() -> None:
    graph = DependencyGraph()
    _add(graph, _leaf("A", 1), _leaf("A", 2))
    graph.rebuild_adjacency()
    assert graph.edge_count() == 0
    assert list(graph._row_ptr) == [0, 0, 0]
    assert list(graph._col_ptr) == [0, 0, 0]
    assert graph.get_dependencies("Sheet1!A1") == frozenset()
    assert graph.get_dependents("Sheet1!A2") == frozenset()


def test_dangling_endpoints_are_dropped_on_rebuild() -> None:
    graph = DependencyGraph()
    _add(graph, _leaf("A", 1), _formula("B", 1))
    graph.add_edge("Sheet1!B1", "Sheet1!A1")
    graph.add_edge("Sheet1!B1", "Sheet1!Z9")
    assert graph.get_dependencies("Sheet1!B1") == frozenset({"Sheet1!A1", "Sheet1!Z9"})

    graph.rebuild_adjacency()

    assert "Sheet1!Z9" not in graph._node_index
    assert graph.get_dependencies("Sheet1!B1") == frozenset({"Sheet1!A1"})
    assert graph.get_dependents("Sheet1!A1") == frozenset({"Sheet1!B1"})
    assert graph.edge_count() == 1


def test_csr_arrays_are_uint32() -> None:
    graph = _chain_graph(3)
    graph.rebuild_adjacency()
    for part in (graph._row_ptr, graph._col_idx, graph._col_ptr, graph._row_idx):
        assert isinstance(part, array.array)
        assert part.typecode == "I"


def test_csc_is_transpose_of_csr() -> None:
    graph = _chain_graph(12, extra_degree=4)
    graph.rebuild_adjacency()
    forward: set[tuple[int, int]] = set()
    for i in range(len(graph._csr_keys)):
        for k in range(graph._row_ptr[i], graph._row_ptr[i + 1]):
            forward.add((i, graph._col_idx[k]))
    reverse: set[tuple[int, int]] = set()
    for j in range(len(graph._csr_keys)):
        for k in range(graph._col_ptr[j], graph._col_ptr[j + 1]):
            reverse.add((graph._row_idx[k], j))
    assert forward == reverse
    assert len(forward) == graph.edge_count()


def test_node_index_reuses_node_map_key_objects() -> None:
    graph = _chain_graph(20)
    graph.rebuild_adjacency()
    for stored, csr_key in zip(graph._nodes, graph._csr_keys, strict=True):
        assert stored is csr_key
        assert graph._node_index[stored] == graph._node_index[csr_key]
        assert stored is next(k for k in graph._node_index if k == stored)


def test_roots_use_csc_degrees() -> None:
    graph = DependencyGraph()
    _add(graph, _leaf("A", 1), _formula("B", 1), _formula("C", 1))
    graph.add_edge("Sheet1!B1", "Sheet1!A1")
    graph.add_edge("Sheet1!C1", "Sheet1!B1")
    graph.rebuild_adjacency()
    assert set(graph.roots()) == {"Sheet1!C1"}


def test_cycle_report_does_not_reaccumulate_dict_adjacency() -> None:
    graph = DependencyGraph()
    _add(graph, _formula("A", 1), _formula("B", 1))
    graph.add_edge("Sheet1!A1", "Sheet1!B1")
    graph.add_edge("Sheet1!B1", "Sheet1!A1")
    graph.rebuild_adjacency()
    report = graph.cycle_report()
    assert report.has_must_cycles
    assert graph._edges == {}
    assert graph._reverse_edges == {}
    assert graph._staging is False


def test_evaluation_order_walks_csr_without_dict_maps() -> None:
    graph = DependencyGraph(sheet_order=["Sheet1"])
    _add(graph, _leaf("A", 1), _formula("B", 1), _formula("C", 1))
    graph.add_edge("Sheet1!B1", "Sheet1!A1")
    graph.add_edge("Sheet1!C1", "Sheet1!B1")
    graph.rebuild_adjacency()
    order = graph.evaluation_order()
    assert order == ["Sheet1!A1", "Sheet1!B1", "Sheet1!C1"]
    assert graph._edges == {}
    assert graph._staging is False


def test_copy_for_projection_clones_csr_arrays() -> None:
    graph = _chain_graph(8)
    graph.rebuild_adjacency()
    cloned = graph._copy_for_projection()
    assert cloned._staging is False
    assert cloned._edges == {}
    assert cloned.get_dependencies("Sheet1!B3") == graph.get_dependencies("Sheet1!B3")
    assert cloned._row_ptr is not graph._row_ptr
    assert list(cloned._row_ptr) == list(graph._row_ptr)
    cloned._col_idx[0] = 0 if cloned._col_idx[0] != 0 else 1
    assert cloned._col_idx[0] != graph._col_idx[0]


def test_deepcopy_and_pickle_round_trip_csr() -> None:
    graph = _chain_graph(6)
    graph.rebuild_adjacency()
    pickled: DependencyGraph = pickle.loads(pickle.dumps(graph))
    deep = copy.deepcopy(graph)
    json_restored = dependency_graph_from_json(dependency_graph_to_json(graph))
    for restored in (pickled, deep, json_restored):
        assert restored._staging is False
        assert restored._edges == {}
        assert restored.edge_count() == graph.edge_count()
        for key in graph:
            assert restored.get_dependencies(key) == graph.get_dependencies(key)
            assert restored.get_dependents(key) == graph.get_dependents(key)


def test_add_edge_after_rebuild_rehydrates_then_can_rebuild() -> None:
    graph = DependencyGraph()
    _add(graph, _leaf("A", 1), _formula("B", 1), _formula("C", 1))
    graph.add_edge("Sheet1!B1", "Sheet1!A1")
    graph.rebuild_adjacency()
    graph.add_edge("Sheet1!C1", "Sheet1!B1")
    assert graph._staging is True
    assert graph.get_dependencies("Sheet1!B1") == frozenset({"Sheet1!A1"})
    assert graph.get_dependencies("Sheet1!C1") == frozenset({"Sheet1!B1"})
    graph.rebuild_adjacency()
    assert graph._staging is False
    assert graph.get_dependents("Sheet1!B1") == frozenset({"Sheet1!C1"})


def test_rebuild_peak_is_not_a_second_python_edge_index() -> None:
    graph = _chain_graph(400, extra_degree=6)
    report = measure_graph_memory(graph)
    adj_exclusive = (
        report.component("edges_forward").exclusive_bytes
        + report.component("edges_reverse").exclusive_bytes
    )
    tracemalloc.start()
    graph.rebuild_adjacency()
    _current, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    assert peak < adj_exclusive
    assert _csr_array_bytes(graph) < peak
    compact = measure_graph_memory(graph)
    csr = compact.component("adjacency_csr")
    index = compact.component("node_index")
    assert csr.exclusive_bytes + index.exclusive_bytes < adj_exclusive / 2
    assert graph._edges == {}
    assert deep_size(graph._row_ptr, graph._col_idx, graph._col_ptr, graph._row_idx) == (
        _csr_array_bytes(graph)
    )
