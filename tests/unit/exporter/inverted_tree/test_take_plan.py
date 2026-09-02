"""Index planning: predecessor-closure is the lag graph, not `1..max` as text."""

from __future__ import annotations

from excel_grapher.exporter.inverted_tree.deps import predecessor_closure


def test_predecessor_closure_walks_lag_graph() -> None:
    assert predecessor_closure((2, 4)) == (0, 1, 2, 3, 4)
    assert predecessor_closure((0,)) == (0,)
    assert predecessor_closure((1, 2)) == (0, 1, 2)
    assert predecessor_closure(()) == ()


def test_predecessor_closure_closes_under_supplied_distances() -> None:
    assert predecessor_closure((4,), distances=(2,)) == (0, 2, 4)
    assert predecessor_closure((5, 7), distances=(2,)) == (1, 3, 5, 7)
    assert predecessor_closure((2, 4), distances=(1, 2)) == (0, 1, 2, 3, 4)
    assert predecessor_closure((3,), distances=()) == (3,)
