"""Tests for incremental computation and cache invalidation."""

from excel_grapher import DependencyGraph, Node
from excel_grapher.evaluator.evaluator import FormulaEvaluator
from excel_grapher.evaluator.name_utils import parse_address


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


# --- set_value tests ---


def test_set_value_updates_node_value() -> None:
    """set_value should update the node's value in the graph."""
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!B1", "=S!A1*2", None),
    )
    with FormulaEvaluator(graph) as ev:
        ev.set_value("S!A1", 20)
        node = graph.get_node("S!A1")
        assert node is not None
        assert node.value == 20


def test_set_value_invalidates_cache_for_cell() -> None:
    """set_value should remove the cell from the cache."""
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!B1", "=S!A1*2", None),
    )
    with FormulaEvaluator(graph) as ev:
        # First evaluation populates cache
        ev.evaluate(["S!A1"])
        assert "S!A1" in ev._cache  # noqa: SLF001

        # set_value should invalidate
        ev.set_value("S!A1", 20)
        assert "S!A1" not in ev._cache  # noqa: SLF001


def test_set_value_invalidates_cache_for_dependents() -> None:
    """set_value should invalidate cache for all cells that depend on the changed cell."""
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!B1", "=S!A1*2", None),
        _make_node("S!C1", "=S!B1+1", None),
    )
    # Add edges so graph knows the dependencies
    graph.add_edge("S!B1", "S!A1")
    graph.add_edge("S!C1", "S!B1")

    with FormulaEvaluator(graph) as ev:
        # First evaluation populates cache
        ev.evaluate(["S!C1"])
        assert "S!A1" in ev._cache  # noqa: SLF001
        assert "S!B1" in ev._cache  # noqa: SLF001
        assert "S!C1" in ev._cache  # noqa: SLF001

        # set_value on A1 should invalidate B1 and C1 (dependents)
        ev.set_value("S!A1", 20)
        assert "S!A1" not in ev._cache  # noqa: SLF001
        assert "S!B1" not in ev._cache  # noqa: SLF001
        assert "S!C1" not in ev._cache  # noqa: SLF001


def test_set_value_does_not_invalidate_unrelated_cells() -> None:
    """set_value should not invalidate cells that don't depend on the changed cell."""
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!A2", None, 5),
        _make_node("S!B1", "=S!A1*2", None),
        _make_node("S!B2", "=S!A2*3", None),
    )
    graph.add_edge("S!B1", "S!A1")
    graph.add_edge("S!B2", "S!A2")

    with FormulaEvaluator(graph) as ev:
        # Evaluate both branches
        ev.evaluate(["S!B1", "S!B2"])
        assert "S!B1" in ev._cache  # noqa: SLF001
        assert "S!B2" in ev._cache  # noqa: SLF001

        # set_value on A1 should only invalidate B1, not B2
        ev.set_value("S!A1", 20)
        assert "S!B1" not in ev._cache  # noqa: SLF001
        assert "S!B2" in ev._cache  # noqa: SLF001


def test_reevaluation_after_set_value_uses_new_value() -> None:
    """After set_value, re-evaluation should use the new value."""
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!B1", "=S!A1*2", None),
    )
    graph.add_edge("S!B1", "S!A1")

    with FormulaEvaluator(graph) as ev:
        # First evaluation
        result1 = ev.evaluate(["S!B1"])
        assert result1["S!B1"] == 20.0

        # Change value and re-evaluate
        ev.set_value("S!A1", 5)
        result2 = ev.evaluate(["S!B1"])
        assert result2["S!B1"] == 10.0


# --- auto_detect_changes tests ---


def test_auto_detect_changes_detects_mutated_leaf() -> None:
    """With auto_detect_changes=True, direct mutation of node.value should be detected."""
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!B1", "=S!A1*2", None),
    )
    graph.add_edge("S!B1", "S!A1")

    with FormulaEvaluator(graph, auto_detect_changes=True) as ev:
        # First evaluation
        result1 = ev.evaluate(["S!B1"])
        assert result1["S!B1"] == 20.0

        # Directly mutate the node value (not using set_value)
        node = graph.get_node("S!A1")
        assert node is not None
        node.value = 5

        # Re-evaluation should detect the change and recompute
        result2 = ev.evaluate(["S!B1"])
        assert result2["S!B1"] == 10.0


def test_auto_detect_changes_false_ignores_direct_mutation() -> None:
    """With auto_detect_changes=False, direct mutation is not detected (returns stale cache)."""
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!B1", "=S!A1*2", None),
    )
    graph.add_edge("S!B1", "S!A1")

    with FormulaEvaluator(graph, auto_detect_changes=False) as ev:
        # First evaluation
        result1 = ev.evaluate(["S!B1"])
        assert result1["S!B1"] == 20.0

        # Directly mutate the node value (not using set_value)
        node = graph.get_node("S!A1")
        assert node is not None
        node.value = 5

        # Re-evaluation should return stale cached value
        result2 = ev.evaluate(["S!B1"])
        assert result2["S!B1"] == 20.0  # Stale!


def test_auto_detect_changes_false_still_works_with_set_value() -> None:
    """With auto_detect_changes=False, set_value still works correctly."""
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!B1", "=S!A1*2", None),
    )
    graph.add_edge("S!B1", "S!A1")

    with FormulaEvaluator(graph, auto_detect_changes=False) as ev:
        # First evaluation
        result1 = ev.evaluate(["S!B1"])
        assert result1["S!B1"] == 20.0

        # Use set_value (explicit invalidation)
        ev.set_value("S!A1", 5)

        # Re-evaluation should use new value
        result2 = ev.evaluate(["S!B1"])
        assert result2["S!B1"] == 10.0


# --- eager_invalidation tests ---


def test_eager_invalidation_checks_all_leaves_upfront() -> None:
    """With eager_invalidation=True, all leaves are checked at start of evaluate()."""
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!A2", None, 5),
        _make_node("S!B1", "=S!A1*2", None),
        _make_node("S!B2", "=S!A2*3", None),
    )
    graph.add_edge("S!B1", "S!A1")
    graph.add_edge("S!B2", "S!A2")

    with FormulaEvaluator(
        graph, auto_detect_changes=True, eager_invalidation=True
    ) as ev:
        # First evaluation
        ev.evaluate(["S!B1", "S!B2"])

        # Mutate both leaves
        node_a1 = graph.get_node("S!A1")
        node_a2 = graph.get_node("S!A2")
        assert node_a1 is not None and node_a2 is not None
        node_a1.value = 1
        node_a2.value = 2

        # Even if we only evaluate B1, eager mode should detect A2 change too
        ev.evaluate(["S!B1"])
        # A2's change should have been detected and cached value invalidated
        assert "S!A2" not in ev._cache or ev._cache.get("S!A2") != 5  # noqa: SLF001


def test_lazy_invalidation_only_checks_visited_leaves() -> None:
    """With eager_invalidation=False, only leaves in the evaluation path are checked."""
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!A2", None, 5),
        _make_node("S!B1", "=S!A1*2", None),
        _make_node("S!B2", "=S!A2*3", None),
    )
    graph.add_edge("S!B1", "S!A1")
    graph.add_edge("S!B2", "S!A2")

    with FormulaEvaluator(
        graph, auto_detect_changes=True, eager_invalidation=False
    ) as ev:
        # First evaluation of both
        ev.evaluate(["S!B1", "S!B2"])
        assert ev._cache["S!B2"] == 15.0  # noqa: SLF001

        # Mutate A2 (not in B1's dependency path)
        node = graph.get_node("S!A2")
        assert node is not None
        node.value = 100

        # Evaluate only B1 - should NOT detect A2's change (lazy mode)
        ev.evaluate(["S!B1"])
        # B2's cached value should still be stale (15.0, not 300.0)
        assert ev._cache.get("S!B2") == 15.0  # noqa: SLF001


def test_lazy_invalidation_detects_changes_in_evaluation_path() -> None:
    """With eager_invalidation=False, changes in the evaluation path ARE detected."""
    graph = _make_graph(
        _make_node("S!A1", None, 10),
        _make_node("S!B1", "=S!A1*2", None),
    )
    graph.add_edge("S!B1", "S!A1")

    with FormulaEvaluator(
        graph, auto_detect_changes=True, eager_invalidation=False
    ) as ev:
        # First evaluation
        result1 = ev.evaluate(["S!B1"])
        assert result1["S!B1"] == 20.0

        # Mutate A1 (IS in B1's dependency path)
        node = graph.get_node("S!A1")
        assert node is not None
        node.value = 5

        # Evaluate B1 - should detect A1's change even in lazy mode
        result2 = ev.evaluate(["S!B1"])
        assert result2["S!B1"] == 10.0
