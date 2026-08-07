"""Unit tests for the shared SCC-condensation ranking helper."""

from __future__ import annotations

from excel_grapher.grapher.lightweight_viz import unconditional_scc_ranks


def test_empty_graph_returns_no_ranks_and_zero_sccs():
    assert unconditional_scc_ranks([], 0) == ([], 0)


def test_isolated_nodes_are_all_rank_zero_and_each_its_own_scc():
    ranks, scc_count = unconditional_scc_ranks([[], [], []], 3)
    assert ranks == [0, 0, 0]
    assert scc_count == 3


def test_chain_ranks_increase_along_dependency_direction():
    # 0 -> 1 -> 2 -> 3
    ranks, scc_count = unconditional_scc_ranks([[1], [2], [3], []], 4)
    assert ranks == [0, 1, 2, 3]
    assert scc_count == 4


def test_cycle_members_collapse_into_one_scc_and_share_a_rank():
    # 0 -> 1 -> 2 -> 1 (1 and 2 form a cycle), 2 -> 3
    ranks, scc_count = unconditional_scc_ranks([[1], [2], [1, 3], []], 4)
    assert scc_count == 3
    assert ranks[1] == ranks[2]
    assert ranks[0] < ranks[1] < ranks[3]


def test_diamond_uses_longest_path_not_shortest():
    # 0 -> 1 -> 3 and 0 -> 2 -> 3, plus a long leg 0 -> 4 -> 5 -> 3
    ranks, _ = unconditional_scc_ranks([[1, 2, 4], [3], [3], [], [5], [3]], 6)
    assert ranks[3] == 3
