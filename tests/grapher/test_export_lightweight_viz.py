from __future__ import annotations

import pytest

from excel_grapher.grapher import DependencyGraph, Literal, Node
from excel_grapher.grapher.export import (
    to_graphviz,
    to_lightweight_viz,
    to_mermaid,
    to_networkx,
)
from excel_grapher.grapher.lightweight_viz import (
    DENSE_BUCKET_THRESHOLD,
    _build_local_csr,
    select_local_force_subgraph,
    serialize_lightweight_viz_json,
)


def _n(sheet: str, col: str, row: int, *, leaf: bool, formula: str | None) -> Node:
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=formula,
        value=1 if leaf else None,
        is_leaf=leaf,
    )


def _chain_graph() -> DependencyGraph:
    """A3 -> A2 -> A1 (dependencies first toward leaves)."""
    g = DependencyGraph()
    n1 = _n("S", "A", 1, leaf=True, formula=None)
    n2 = _n("S", "A", 2, leaf=False, formula="=A1")
    n3 = _n("S", "A", 3, leaf=False, formula="=A2")
    for n in (n1, n2, n3):
        g.add_node(n)
    g.add_edge(n3.key, n2.key)
    g.add_edge(n2.key, n1.key)
    return g


def _fork_join_graph() -> DependencyGraph:
    """D depends on B and C; B and C depend on A."""
    g = DependencyGraph()
    na = _n("S", "A", 1, leaf=True, formula=None)
    nb = _n("S", "B", 1, leaf=False, formula="=A1")
    nc = _n("S", "C", 1, leaf=False, formula="=A1")
    nd = _n("S", "D", 1, leaf=False, formula="=B1+C1")
    for n in (na, nb, nc, nd):
        g.add_node(n)
    g.add_edge(nb.key, na.key)
    g.add_edge(nc.key, na.key)
    g.add_edge(nd.key, nb.key)
    g.add_edge(nd.key, nc.key)
    return g


def _two_cycle_graph() -> DependencyGraph:
    """Unconditional 2-cycle A <-> B."""
    g = DependencyGraph()
    na = _n("S", "A", 1, leaf=False, formula="=B1")
    nb = _n("S", "B", 1, leaf=False, formula="=A1")
    g.add_node(na)
    g.add_node(nb)
    g.add_edge(na.key, nb.key)
    g.add_edge(nb.key, na.key)
    return g


def _guarded_back_edge_graph() -> DependencyGraph:
    """B -> A unconditional; A -> B guarded only (no unconditional cycle)."""
    g = DependencyGraph()
    na = _n("S", "A", 1, leaf=False, formula="=B1")
    nb = _n("S", "B", 1, leaf=False, formula="=1")
    g.add_node(na)
    g.add_node(nb)
    g.add_edge(nb.key, na.key)
    g.add_edge(na.key, nb.key, guard=Literal(True))
    return g


def test_build_local_csr_hoisted_out_degree_sort() -> None:
    """Neighbor sort uses global out-degree; regression guard for O(n) precompute (not per-src)."""
    n = 4
    module_of = [0, 0, 0, 0]
    out_edges_by_src: list[list[tuple[int, bool]]] = [
        [(1, False), (2, False), (3, False)],
        [(2, False)],
        [(3, False)],
        [],
    ]
    mod_internal_edges = 5
    offsets, targets, guarded, complete = _build_local_csr(
        n,
        module_of,
        out_edges_by_src,
        [4],
        [mod_internal_edges],
        max_local_nodes=5000,
        max_local_edges=20000,
    )
    assert offsets == [0, 3, 4, 5, 5]
    assert targets == [1, 2, 3, 2, 3]
    assert guarded == [False] * 5
    assert all(complete)


def test_payload_contract_and_version() -> None:
    g = _chain_graph()
    p = to_lightweight_viz(g)
    assert p.version == 1
    assert p.stats.node_count == 3
    assert len(p.sheets) == 1 and p.sheets[0] == "S"
    assert len(p.nodes.sheet_index) == 3
    assert all(si == 0 for si in p.nodes.sheet_index)
    assert len(p.modules) >= 1
    assert p.local_edges.offsets[0] == 0
    assert p.local_edges.offsets[-1] == len(p.local_edges.targets)


def test_deterministic_ids_and_serialization() -> None:
    g = _fork_join_graph()
    a = serialize_lightweight_viz_json(to_lightweight_viz(g))
    b = serialize_lightweight_viz_json(to_lightweight_viz(g))
    assert a == b


def test_existing_exports_unchanged() -> None:
    g = _chain_graph()
    assert "digraph" in to_graphviz(g)
    assert "flowchart" in to_mermaid(g, max_nodes=10)
    G = to_networkx(g)
    assert G.number_of_nodes() == 3


def test_rank_chain_monotonic_along_flow() -> None:
    """Ranks are longest-path from condensation sources (roots); A3 -> A2 -> A1 implies non-increasing toward the leaf."""
    p = to_lightweight_viz(_chain_graph())
    ranks = p.nodes.rank
    keys = sorted(_chain_graph())
    idx = {k: i for i, k in enumerate(keys)}
    assert ranks[idx["S!A3"]] <= ranks[idx["S!A2"]] <= ranks[idx["S!A1"]]


def test_rank_fork_join() -> None:
    p = to_lightweight_viz(_fork_join_graph())
    r = list(p.nodes.rank)
    keys = sorted(_fork_join_graph())
    idx = {k: i for i, k in enumerate(keys)}
    assert r[idx["S!B1"]] == r[idx["S!C1"]]
    assert r[idx["S!D1"]] < r[idx["S!A1"]]


def test_cycle_single_scc_shared_rank() -> None:
    p = to_lightweight_viz(_two_cycle_graph())
    assert p.stats.scc_count == 1
    assert p.nodes.rank[0] == p.nodes.rank[1]


def test_guarded_does_not_create_uncond_cycle_scc() -> None:
    p = to_lightweight_viz(_guarded_back_edge_graph())
    assert p.stats.scc_count == 2


def test_module_edges_aggregate_matches_graph() -> None:
    g = _fork_join_graph()
    p = to_lightweight_viz(g, module_iterations=8)
    keys = sorted(g)
    key_id = {k: i for i, k in enumerate(keys)}
    mod = list(p.nodes.module_id)
    u_exp = 0
    g_exp = 0
    for fk in keys:
        fi = key_id[fk]
        for tk in sorted(g.dependencies(fk)):
            ti = key_id.get(tk)
            if ti is None or mod[fi] == mod[ti]:
                continue
            if g.edge_attrs(fk, tk).get("guard") is not None:
                g_exp += 1
            else:
                u_exp += 1
    u_act = sum(e.unconditional_weight for e in p.module_edges)
    g_act = sum(e.guarded_weight for e in p.module_edges)
    assert u_act == u_exp
    assert g_act == g_exp


def test_layout_x_monotone_with_rank() -> None:
    p = to_lightweight_viz(_fork_join_graph())
    n = p.stats.node_count
    for i in range(n):
        for j in range(n):
            if p.nodes.rank[i] < p.nodes.rank[j]:
                assert p.nodes.x[i] <= p.nodes.x[j]


def test_dense_bucket_metadata() -> None:
    g = DependencyGraph()
    for i in range(1, DENSE_BUCKET_THRESHOLD + 4):
        g.add_node(_n("S", "B", i, leaf=True, formula=None))
    hub = _n("S", "A", 1, leaf=False, formula="=B1")
    g.add_node(hub)
    for i in range(1, DENSE_BUCKET_THRESHOLD + 4):
        g.add_edge(hub.key, f"S!B{i}")
    p = to_lightweight_viz(g, module_iterations=8)
    assert p.stats.dense_bucket_count >= 1
    assert any(d > DENSE_BUCKET_THRESHOLD for d in p.nodes.bucket_density)


def test_small_module_local_edges_complete() -> None:
    p = to_lightweight_viz(_chain_graph(), max_local_nodes=100, max_local_edges=100)
    assert all(p.local_edges.complete)


def test_local_edges_valid_targets() -> None:
    p = to_lightweight_viz(_fork_join_graph())
    n = p.stats.node_count
    off = p.local_edges.offsets
    tg = p.local_edges.targets
    for i in range(n):
        for k in range(off[i], off[i + 1]):
            assert 0 <= tg[k] < n


def test_truncation_prefers_unconditional_first() -> None:
    g = DependencyGraph()
    na = _n("S", "A", 1, leaf=True, formula=None)
    root = _n("S", "A", 2, leaf=False, formula="=A1")
    g.add_node(na)
    g.add_node(root)
    leaves = []
    for i in range(30):
        n = _n("S", "B", i + 1, leaf=True, formula=None)
        g.add_node(n)
        leaves.append(n)
        g.add_edge(root.key, n.key, guard=Literal(True))
    g.add_edge(root.key, na.key)
    p = to_lightweight_viz(g, max_local_nodes=5000, max_local_edges=5)
    ri = sorted(g).index(root.key)
    off = p.local_edges.offsets
    tg = p.local_edges.targets
    first_targets = [tg[k] for k in range(off[ri], off[ri + 1])]
    ai = sorted(g).index(na.key)
    assert ai in first_targets


def test_select_local_force_module_scope() -> None:
    p = to_lightweight_viz(_chain_graph(), max_local_nodes=10, max_local_edges=50)
    sub = select_local_force_subgraph(p, node_id=2)
    assert sub.is_module_scope
    assert set(sub.node_ids) <= {0, 1, 2}


def test_select_local_force_neighborhood_when_large_module() -> None:
    p = to_lightweight_viz(_chain_graph(), max_local_nodes=2, max_local_edges=20)
    sub = select_local_force_subgraph(p, node_id=0)
    assert not sub.is_module_scope
    assert len(sub.node_ids) <= 2


@pytest.mark.slow
def test_lightweight_viz_large_chain_benchmark() -> None:
    """Scale sanity check: linear chain, columnar export stays within rough time/size bounds."""
    import time

    n_nodes = 15_000
    g = DependencyGraph()
    g.add_node(_n("S", "A", 1, leaf=True, formula=None))
    for r in range(2, n_nodes + 1):
        g.add_node(_n("S", "A", r, leaf=False, formula=f"=A{r - 1}"))
        g.add_edge(f"S!A{r}", f"S!A{r - 1}")
    t0 = time.perf_counter()
    p = to_lightweight_viz(g, max_local_nodes=500, max_local_edges=2000)
    elapsed = time.perf_counter() - t0
    raw = serialize_lightweight_viz_json(p)
    assert p.stats.node_count == n_nodes
    assert elapsed < 120.0
    assert len(raw.encode("utf-8")) < 50 * 1024 * 1024
