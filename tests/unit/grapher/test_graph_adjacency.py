"""Sparse adjacency maps: omit empty neighbor sets (#907)."""

from __future__ import annotations

import copy
import pickle

from excel_grapher.grapher.cache import dependency_graph_from_json, dependency_graph_to_json
from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import make_cell_node
from excel_grapher.grapher.subgraph import select_path_induced_subgraph


def _leaf(sheet: str, column: str, row: int, value: object = 1):
    return make_cell_node(sheet, column, row, value=value, is_leaf=True)


def _formula(sheet: str, column: str, row: int, formula: str):
    return make_cell_node(
        sheet,
        column,
        row,
        formula=formula,
        normalized_formula=formula,
        is_leaf=False,
    )


def _empty_adjacency_set_ids(graph: DependencyGraph) -> set[int]:
    ids: set[int] = set()
    for mapping in (graph._edges, graph._reverse_edges):
        for neighbors in mapping.values():
            if not neighbors:
                ids.add(id(neighbors))
    return ids


def _direct_edge(graph: DependencyGraph, src: str, dst: str) -> None:
    graph.add_edge(
        src,
        dst,
        provenance=EdgeProvenance(causes=DependencyCause.direct_ref),
    )


def test_isolated_nodes_do_not_own_distinct_empty_adjacency_sets() -> None:
    """An edgeless graph of N nodes must not allocate N×2 empty neighbor sets."""
    n = 16
    graph = DependencyGraph()
    for i in range(1, n + 1):
        graph.add_node(_leaf("Sheet1", "A", i, value=i))

    assert graph._edges == {}
    assert graph._reverse_edges == {}
    assert _empty_adjacency_set_ids(graph) == set()
    for i in range(1, n + 1):
        key = f"Sheet1!A{i}"
        assert graph.get_dependencies(key) == frozenset()
        assert graph.get_dependents(key) == frozenset()


def test_leaf_and_sink_omit_empty_direction() -> None:
    """Leaves omit outgoing keys; sinks omit reverse keys."""
    graph = DependencyGraph()
    graph.add_node(_leaf("Sheet1", "A", 1))
    graph.add_node(_formula("Sheet1", "B", 1, "=Sheet1!A1"))
    _direct_edge(graph, "Sheet1!B1", "Sheet1!A1")

    assert "Sheet1!A1" not in graph._edges
    assert "Sheet1!B1" not in graph._reverse_edges
    assert graph._edges["Sheet1!B1"] == {"Sheet1!A1"}
    assert graph._reverse_edges["Sheet1!A1"] == {"Sheet1!B1"}
    assert _empty_adjacency_set_ids(graph) == set()
    assert graph.get_dependencies("Sheet1!A1") == frozenset()
    assert graph.get_dependents("Sheet1!B1") == frozenset()


def test_remove_last_edge_drops_empty_adjacency_entries() -> None:
    graph = DependencyGraph()
    graph.add_node(_leaf("Sheet1", "A", 1))
    graph.add_node(_leaf("Sheet1", "C", 1))
    graph.add_node(_formula("Sheet1", "B", 1, "=Sheet1!A1+Sheet1!C1"))
    _direct_edge(graph, "Sheet1!B1", "Sheet1!A1")
    _direct_edge(graph, "Sheet1!B1", "Sheet1!C1")

    graph._remove_edge("Sheet1!B1", "Sheet1!A1")
    assert "Sheet1!A1" not in graph._reverse_edges
    assert graph._edges["Sheet1!B1"] == {"Sheet1!C1"}

    graph._remove_edge("Sheet1!B1", "Sheet1!C1")
    assert "Sheet1!B1" not in graph._edges
    assert "Sheet1!C1" not in graph._reverse_edges
    assert graph._edges == {}
    assert graph._reverse_edges == {}
    assert graph.get_dependencies("Sheet1!B1") == frozenset()
    assert graph.get_dependents("Sheet1!A1") == frozenset()


def test_remove_missing_edge_does_not_insert_empty_sets() -> None:
    graph = DependencyGraph()
    graph.add_node(_leaf("Sheet1", "A", 1))
    graph._remove_edge("Sheet1!A1", "Sheet1!Z9")
    assert graph._edges == {}
    assert graph._reverse_edges == {}
    assert _empty_adjacency_set_ids(graph) == set()


def test_move_isolated_node_does_not_insert_empty_sets() -> None:
    graph = DependencyGraph()
    graph.add_node(_leaf("Sheet1", "A", 1))
    graph.move_node("Sheet1!A1", "Sheet1!B2")
    assert "Sheet1!B2" in graph
    assert graph._edges == {}
    assert graph._reverse_edges == {}
    assert _empty_adjacency_set_ids(graph) == set()


def test_move_node_keeps_sparse_leaf_and_sink_keys() -> None:
    graph = DependencyGraph()
    graph.add_node(_leaf("Sheet1", "A", 1))
    graph.add_node(_formula("Sheet1", "B", 1, "=A1"))
    _direct_edge(graph, "Sheet1!B1", "Sheet1!A1")

    graph.move_node("Sheet1!A1", "Sheet1!C3")

    assert "Sheet1!C3" not in graph._edges
    assert "Sheet1!B1" not in graph._reverse_edges
    assert graph._edges["Sheet1!B1"] == {"Sheet1!C3"}
    assert graph._reverse_edges["Sheet1!C3"] == {"Sheet1!B1"}
    assert _empty_adjacency_set_ids(graph) == set()


def test_replace_formula_clearing_outgoing_drops_empty_forward_set() -> None:
    graph = DependencyGraph()
    graph.add_node(_leaf("Sheet1", "B", 1))
    graph.add_node(_formula("Sheet1", "A", 1, "=Sheet1!B1"))
    graph.add_node(_formula("Sheet1", "D", 1, "=Sheet1!A1"))
    _direct_edge(graph, "Sheet1!A1", "Sheet1!B1")
    _direct_edge(graph, "Sheet1!D1", "Sheet1!A1")

    graph.replace_node_formula("Sheet1!A1", None, None)

    assert "Sheet1!A1" not in graph._edges
    assert graph._reverse_edges["Sheet1!A1"] == {"Sheet1!D1"}
    assert _empty_adjacency_set_ids(graph) == set()
    assert graph.get_dependencies("Sheet1!A1") == frozenset()
    assert graph.get_dependents("Sheet1!A1") == frozenset({"Sheet1!D1"})


def test_copy_for_projection_does_not_reaccumulate_empty_sets() -> None:
    graph = DependencyGraph()
    graph.add_node(_leaf("Sheet1", "A", 1))
    graph.add_node(_formula("Sheet1", "B", 1, "=Sheet1!A1"))
    _direct_edge(graph, "Sheet1!B1", "Sheet1!A1")

    cloned = graph._copy_for_projection()
    assert cloned._edges == {"Sheet1!B1": {"Sheet1!A1"}}
    assert cloned._reverse_edges == {"Sheet1!A1": {"Sheet1!B1"}}
    assert cloned._edges["Sheet1!B1"] is not graph._edges["Sheet1!B1"]
    assert _empty_adjacency_set_ids(cloned) == set()

    isolated = DependencyGraph()
    isolated.add_node(_leaf("Sheet1", "A", 1))
    cloned_isolated = isolated._copy_for_projection()
    assert cloned_isolated._edges == {}
    assert cloned_isolated._reverse_edges == {}


def test_copy_for_projection_drops_legacy_empty_neighbor_sets() -> None:
    graph = DependencyGraph()
    graph.add_node(_leaf("Sheet1", "A", 1))
    graph._edges["Sheet1!A1"] = set()
    graph._reverse_edges["Sheet1!A1"] = set()

    cloned = graph._copy_for_projection()
    assert cloned._edges == {}
    assert cloned._reverse_edges == {}
    assert _empty_adjacency_set_ids(cloned) == set()


def test_induced_subgraph_omits_empty_adjacency() -> None:
    graph = DependencyGraph()
    graph.add_node(_leaf("Sheet1", "A", 1))
    graph.add_node(_formula("Sheet1", "B", 1, "=Sheet1!A1"))
    graph.add_node(_leaf("Sheet1", "Z", 1))
    _direct_edge(graph, "Sheet1!B1", "Sheet1!A1")

    sub = select_path_induced_subgraph(graph, source_keys=["Sheet1!B1"], target_keys=["Sheet1!A1"])
    assert "Sheet1!Z1" not in sub
    assert sub._edges == {"Sheet1!B1": {"Sheet1!A1"}}
    assert "Sheet1!A1" not in sub._edges
    assert "Sheet1!B1" not in sub._reverse_edges
    assert _empty_adjacency_set_ids(sub) == set()


def test_json_and_pickle_round_trip_do_not_restore_empty_sets() -> None:
    graph = DependencyGraph()
    for i in range(1, 5):
        graph.add_node(_leaf("Sheet1", "A", i, value=i))
    graph.add_node(_formula("Sheet1", "B", 1, "=Sheet1!A1"))
    _direct_edge(graph, "Sheet1!B1", "Sheet1!A1")

    pickled: DependencyGraph = pickle.loads(pickle.dumps(graph))
    json_restored = dependency_graph_from_json(dependency_graph_to_json(graph))
    deep = copy.deepcopy(graph)

    for restored in (pickled, json_restored, deep):
        assert restored._edges == {"Sheet1!B1": {"Sheet1!A1"}}
        assert restored._reverse_edges == {"Sheet1!A1": {"Sheet1!B1"}}
        assert "Sheet1!A2" not in restored._edges
        assert "Sheet1!A2" not in restored._reverse_edges
        assert _empty_adjacency_set_ids(restored) == set()
        for i in range(2, 5):
            key = f"Sheet1!A{i}"
            assert restored.get_dependencies(key) == frozenset()
            assert restored.get_dependents(key) == frozenset()
