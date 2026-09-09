"""Tests for incremental computation and cache invalidation."""

from collections.abc import Iterator

import pytest

from excel_grapher import DependencyGraph, Node
from excel_grapher.core.address_keys import parse_address
from excel_grapher.evaluator.evaluator import FormulaEvaluator
from excel_grapher.grapher.node import NodeKey


def _make_node(address: str, formula: str | None, value: object) -> Node:
    """Helper to create a Node from a sheet-qualified address."""
    sheet, coord = parse_address(address)
    col = "".join(c for c in coord if c.isalpha())
    row = int("".join(c for c in coord if c.isdigit()))
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=formula,
        value=value,
        is_leaf=formula is None,
    )


def _make_graph(*nodes: Node) -> DependencyGraph:
    """Helper to create a DependencyGraph from nodes."""
    graph = DependencyGraph()
    for node in nodes:
        graph.add_node(node)
    return graph


# --- durable leaf update via graph.set_node_value ---


def test_graph_set_node_value_updates_node() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!B1", "=S!A1*2", None),
    )
    graph.set_node_value("S!A1", 20)
    node = graph.get_node("S!A1")
    assert node is not None
    assert node.value == 20


def test_reevaluation_after_graph_set_node_value_uses_new_value() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!B1", "=S!A1*2", None),
    )
    graph.add_edge("S!B1", "S!A1")

    with FormulaEvaluator(graph) as ev:
        result1 = ev.evaluate(["S!B1"])
        assert result1["S!B1"] == 20.0

        graph.set_node_value("S!A1", 5)
        result2 = ev.evaluate(["S!B1"])
        assert result2["S!B1"] == 10.0


def test_graph_set_node_value_dependents_re_evaluate() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!B1", "=S!A1*2", None),
        _make_node("S!C1", "=S!B1+1", None),
    )
    graph.add_edge("S!B1", "S!A1")
    graph.add_edge("S!C1", "S!B1")

    with FormulaEvaluator(graph) as ev:
        first = ev.evaluate(["S!C1"])
        assert first["S!C1"] == 21.0

        graph.set_node_value("S!A1", 20)
        second = ev.evaluate(["S!C1"])
        assert second["S!C1"] == 41.0


def test_graph_set_node_value_does_not_affect_unrelated_cells() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!A2", None, 5),
        _make_node("S!B1", "=S!A1*2", None),
        _make_node("S!B2", "=S!A2*3", None),
    )
    graph.add_edge("S!B1", "S!A1")
    graph.add_edge("S!B2", "S!A2")

    with FormulaEvaluator(graph) as ev:
        first = ev.evaluate(["S!B1", "S!B2"])
        assert first == {"S!B1": 20.0, "S!B2": 15.0}

        graph.set_node_value("S!A1", 20)
        second = ev.evaluate(["S!B1", "S!B2"])
        assert second == {"S!B1": 40.0, "S!B2": 15.0}


# --- auto_detect_changes tests ---


def test_auto_detect_changes_detects_mutated_leaf() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!B1", "=S!A1*2", None),
    )
    graph.add_edge("S!B1", "S!A1")

    with FormulaEvaluator(graph, auto_detect_changes=True) as ev:
        result1 = ev.evaluate(["S!B1"])
        assert result1["S!B1"] == 20.0

        graph.set_node_value("S!A1", 5)

        result2 = ev.evaluate(["S!B1"])
        assert result2["S!B1"] == 10.0


def test_auto_detect_changes_false_ignores_durable_leaf_update() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!B1", "=S!A1*2", None),
    )
    graph.add_edge("S!B1", "S!A1")

    with FormulaEvaluator(graph, auto_detect_changes=False) as ev:
        result1 = ev.evaluate(["S!B1"])
        assert result1["S!B1"] == 20.0

        graph.set_node_value("S!A1", 5)

        result2 = ev.evaluate(["S!B1"])
        assert result2["S!B1"] == 20.0


# --- eager_invalidation tests ---


def test_eager_invalidation_checks_all_leaves_upfront() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!A2", None, 5),
        _make_node("S!B1", "=S!A1*2", None),
        _make_node("S!B2", "=S!A2*3", None),
    )
    graph.add_edge("S!B1", "S!A1")
    graph.add_edge("S!B2", "S!A2")

    with FormulaEvaluator(graph, auto_detect_changes=True, eager_invalidation=True) as ev:
        ev.evaluate(["S!B1", "S!B2"])

        graph.set_node_value("S!A1", 1)
        graph.set_node_value("S!A2", 2)

        ev.evaluate(["S!B1"])
        # Eager mode checks all leaves up-front, so A2's change is observed
        # even though we only evaluate B1.
        assert ev._cache.get("S!A2") in (None, 2)


def test_lazy_invalidation_only_checks_visited_leaves() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!A2", None, 5),
        _make_node("S!B1", "=S!A1*2", None),
        _make_node("S!B2", "=S!A2*3", None),
    )
    graph.add_edge("S!B1", "S!A1")
    graph.add_edge("S!B2", "S!A2")

    with FormulaEvaluator(graph, auto_detect_changes=True, eager_invalidation=False) as ev:
        ev.evaluate(["S!B1", "S!B2"])
        assert ev._cache["S!B2"] == 15.0

        graph.set_node_value("S!A2", 100)

        ev.evaluate(["S!B1"])
        # B2's cached value should still be stale in lazy mode
        assert ev._cache.get("S!B2") == 15.0


def test_lazy_invalidation_detects_changes_in_evaluation_path() -> None:
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!B1", "=S!A1*2", None),
    )
    graph.add_edge("S!B1", "S!A1")

    with FormulaEvaluator(graph, auto_detect_changes=True, eager_invalidation=False) as ev:
        result1 = ev.evaluate(["S!B1"])
        assert result1["S!B1"] == 20.0

        graph.set_node_value("S!A1", 5)

        result2 = ev.evaluate(["S!B1"])
        assert result2["S!B1"] == 10.0


# --- eager leaf-scan skipping (GitHub #816) ---


def _graph_with_unused_leaves(
    *, n_unused_leaves: int, n_targets: int
) -> tuple[DependencyGraph, list[str]]:
    """Formulas in column B depend only on S!A1; unused leaves live in column C."""
    graph = DependencyGraph()
    graph.add_node(_make_node("S!A1", None, 1))
    for i in range(n_unused_leaves):
        graph.add_node(_make_node(f"S!C{i + 1}", None, i))
    targets: list[str] = []
    for i in range(n_targets):
        row = i + 1
        addr = f"S!B{row}"
        graph.add_node(_make_node(addr, "=S!A1", None))
        graph.add_edge(addr, "S!A1")
        targets.append(addr)
    return graph, targets


def test_eager_evaluate_does_not_lookup_unused_leaves(monkeypatch: pytest.MonkeyPatch) -> None:
    graph, targets = _graph_with_unused_leaves(n_unused_leaves=200, n_targets=10)
    unused = {f"S!C{i + 1}" for i in range(200)}
    looked_up: list[str] = []
    original = graph.get_node

    def counting_get_node(key: NodeKey):
        looked_up.append(str(key))
        return original(key)

    monkeypatch.setattr(graph, "get_node", counting_get_node)

    with FormulaEvaluator(graph, auto_detect_changes=True, eager_invalidation=True) as ev:
        for addr in targets:
            assert ev.evaluate(addr) == 1

    assert unused.isdisjoint(looked_up)


def test_repeated_eager_evaluate_does_not_rescan_leaves_until_set_node_value(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    graph, targets = _graph_with_unused_leaves(n_unused_leaves=50, n_targets=5)
    scans = {"leaf_keys": 0, "leaf_node_items": 0}
    original_leaf_keys = graph.leaf_keys
    original_items = graph.leaf_node_items

    def counting_leaf_keys() -> list[NodeKey]:
        scans["leaf_keys"] += 1
        return original_leaf_keys()

    def counting_items() -> Iterator[tuple[NodeKey, Node]]:
        scans["leaf_node_items"] += 1
        yield from original_items()

    monkeypatch.setattr(graph, "leaf_keys", counting_leaf_keys)
    monkeypatch.setattr(graph, "leaf_node_items", counting_items)

    with FormulaEvaluator(graph, auto_detect_changes=True, eager_invalidation=True) as ev:
        for addr in targets:
            ev.evaluate(addr)
        assert scans["leaf_keys"] == 0
        assert scans["leaf_node_items"] == 0

        graph.set_node_value("S!A1", 9)
        assert ev.evaluate(targets[0]) == 9
        assert scans["leaf_keys"] + scans["leaf_node_items"] == 1

        ev.evaluate(targets[1])
        assert scans["leaf_keys"] + scans["leaf_node_items"] == 1


def test_lazy_evaluate_still_sees_set_node_value_without_full_leaf_scan(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    graph, targets = _graph_with_unused_leaves(n_unused_leaves=50, n_targets=2)
    scans = {"leaf_keys": 0, "leaf_node_items": 0}
    original_leaf_keys = graph.leaf_keys
    original_items = graph.leaf_node_items

    def counting_leaf_keys() -> list[NodeKey]:
        scans["leaf_keys"] += 1
        return original_leaf_keys()

    def counting_items() -> Iterator[tuple[NodeKey, Node]]:
        scans["leaf_node_items"] += 1
        yield from original_items()

    monkeypatch.setattr(graph, "leaf_keys", counting_leaf_keys)
    monkeypatch.setattr(graph, "leaf_node_items", counting_items)

    with FormulaEvaluator(graph, auto_detect_changes=True, eager_invalidation=False) as ev:
        assert ev.evaluate(targets[0]) == 1
        graph.set_node_value("S!A1", 4)
        assert ev.evaluate(targets[0]) == 4
        assert scans["leaf_keys"] == 0
        assert scans["leaf_node_items"] == 0
