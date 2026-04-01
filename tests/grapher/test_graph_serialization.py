"""Tests for DependencyGraph pickle serialization performance (issue #87).

Covers:
  - Round-trip correctness: nodes, edges, guards, extra attrs survive pickle
  - Compact storage: guards stored directly, not wrapped in per-edge dicts
  - String interning: NodeKey references share identity after deserialization
  - Unpickle builds one key-index map for all guards (not O(guards × nodes) rebuilds)
"""

from __future__ import annotations

import pickle

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.guard import CellRef, Compare, Literal
from excel_grapher.grapher.node import Node


def _make_test_graph() -> DependencyGraph:
    """Build a small graph with a mix of guarded/unguarded edges and extra attrs."""
    g = DependencyGraph()

    g.add_node(Node("Sheet1", "A", 1, None, None, 1, True))
    g.add_node(Node("Sheet1", "B", 1, None, None, 2, True))
    g.add_node(Node("Sheet1", "C", 1, None, None, 0, True))
    g.add_node(
        Node("Sheet1", "D", 1, "=IF(C1,A1,B1)", "=IF(Sheet1!C1,Sheet1!A1,Sheet1!B1)", None, False)
    )

    guard_true = Compare(CellRef("Sheet1!C1"), "=", Literal(True))
    guard_false = Compare(CellRef("Sheet1!C1"), "=", Literal(False))

    g.add_edge("Sheet1!D1", "Sheet1!C1")  # unguarded
    g.add_edge("Sheet1!D1", "Sheet1!A1", guard=guard_true)
    g.add_edge("Sheet1!D1", "Sheet1!B1", guard=guard_false)

    return g


# -------------------------------------------------------------------
# Round-trip correctness
# -------------------------------------------------------------------


def test_pickle_round_trip_preserves_nodes() -> None:
    original = _make_test_graph()
    restored: DependencyGraph = pickle.loads(pickle.dumps(original))

    assert len(restored) == len(original)
    for key in original:
        node_orig = original.get_node(key)
        node_rest = restored.get_node(key)
        assert node_orig is not None
        assert node_rest is not None
        assert node_rest.sheet == node_orig.sheet
        assert node_rest.column == node_orig.column
        assert node_rest.row == node_orig.row
        assert node_rest.formula == node_orig.formula
        assert node_rest.normalized_formula == node_orig.normalized_formula
        assert node_rest.value == node_orig.value
        assert node_rest.is_leaf == node_orig.is_leaf


def test_pickle_round_trip_preserves_edges() -> None:
    original = _make_test_graph()
    restored: DependencyGraph = pickle.loads(pickle.dumps(original))

    for key in original:
        assert restored.dependencies(key) == original.dependencies(key)
        assert restored.dependents(key) == original.dependents(key)


def test_pickle_round_trip_preserves_guards() -> None:
    original = _make_test_graph()
    restored: DependencyGraph = pickle.loads(pickle.dumps(original))

    # Unguarded edge
    assert restored.edge_guard("Sheet1!D1", "Sheet1!C1") is None

    # Guarded edges
    guard_a = restored.edge_guard("Sheet1!D1", "Sheet1!A1")
    assert guard_a is not None
    assert isinstance(guard_a, Compare)
    assert guard_a.op == "="

    guard_b = restored.edge_guard("Sheet1!D1", "Sheet1!B1")
    assert guard_b is not None
    assert isinstance(guard_b, Compare)


def test_pickle_round_trip_many_guarded_edges() -> None:
    """Regression: unpickle must not rebuild a full key→index map per guarded edge."""
    g = DependencyGraph()
    g.add_node(Node("Sheet1", "A", 1, None, None, 1, True))
    g.add_node(
        Node("Sheet1", "D", 1, "=1", "=1", None, False),
    )
    guard = Compare(CellRef("Sheet1!A1"), "=", Literal(True))
    n_extra = 800
    for i in range(2, 2 + n_extra):
        g.add_node(Node("Sheet1", "B", i, None, None, 1, True))
        g.add_edge("Sheet1!D1", f"Sheet1!B{i}", guard=guard)

    blob = pickle.dumps(g)
    restored: DependencyGraph = pickle.loads(blob)
    assert len(restored._guards) == n_extra
    for i in range(2, 2 + n_extra):
        g_edge = restored.edge_guard("Sheet1!D1", f"Sheet1!B{i}")
        assert g_edge is not None
        assert isinstance(g_edge, Compare)


def test_pickle_round_trip_preserves_extra_attrs() -> None:
    """Extra (non-guard) edge attributes must survive round-trip."""
    from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance

    g = DependencyGraph()
    g.add_node(Node("Sheet1", "A", 1, None, None, 1, True))
    g.add_node(Node("Sheet1", "B", 1, "=A1", "=Sheet1!A1", None, False))

    prov = EdgeProvenance(
        causes=frozenset({DependencyCause.direct_ref}),
        direct_sites_formula=((1, 3),),
        direct_sites_normalized=((1, 11),),
    )
    g.add_edge("Sheet1!B1", "Sheet1!A1", provenance=prov)

    restored: DependencyGraph = pickle.loads(pickle.dumps(g))
    attrs = restored.edge_attrs("Sheet1!B1", "Sheet1!A1")
    assert "provenance" in attrs
    assert attrs["provenance"].causes == frozenset({DependencyCause.direct_ref})


def test_pickle_round_trip_preserves_leaf_classification() -> None:
    g = _make_test_graph()
    g.leaf_classification = {"Sheet1!A1": "input", "Sheet1!B1": "constant"}

    restored: DependencyGraph = pickle.loads(pickle.dumps(g))
    assert restored.leaf_classification == {"Sheet1!A1": "input", "Sheet1!B1": "constant"}


# -------------------------------------------------------------------
# Compact storage: no per-edge wrapper dicts in pickle stream
# -------------------------------------------------------------------


def test_pickle_does_not_contain_per_edge_wrapper_dicts() -> None:
    """The serialized state must not wrap each edge's guard in a {'guard': ...} dict.

    After optimization, guards are stored directly in a _guards dict keyed by
    (from_key, to_key), eliminating millions of small wrapper dicts.
    """
    g = _make_test_graph()
    state = g.__getstate__()

    # The state should NOT contain '_edge_attrs' (old format with wrapper dicts)
    assert "_edge_attrs" not in state, (
        "Serialized state still contains _edge_attrs; expected compact representation"
    )
    # The state should contain '_guards' (compact representation)
    assert "_guards" in state


# -------------------------------------------------------------------
# String interning: NodeKeys share identity after deserialization
# -------------------------------------------------------------------


def test_deserialized_nodekeys_share_identity() -> None:
    """After deserialization, the same NodeKey string should be the same object
    (identity, not just equality) across _nodes, _edges, and _guards.

    This reduces memory and speeds up dict operations.
    """
    g = _make_test_graph()
    restored: DependencyGraph = pickle.loads(pickle.dumps(g))

    node_key_ids = {k: id(k) for k in restored._nodes}

    # Keys in _edges should be the same objects as keys in _nodes
    for k in restored._edges:
        if k in node_key_ids:
            assert id(k) == node_key_ids[k], (
                f"_edges key {k!r} is a different object than _nodes key"
            )

    # Keys in _reverse_edges should be the same objects
    for k in restored._reverse_edges:
        if k in node_key_ids:
            assert id(k) == node_key_ids[k], (
                f"_reverse_edges key {k!r} is a different object than _nodes key"
            )

    # Strings inside edge sets should also be interned
    for deps in restored._edges.values():
        for dep in deps:
            if dep in node_key_ids:
                assert id(dep) == node_key_ids[dep], (
                    f"Edge dep {dep!r} is a different object than _nodes key"
                )
