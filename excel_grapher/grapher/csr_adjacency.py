"""CSR+CSC adjacency sidecar for a `DependencyGraph` (issue #908 spike).

This module is a **storage prototype**, not production graph state. It builds
compressed sparse row / column arrays from an existing cell graph without
changing `DependencyGraph.add_edge`. Guard/provenance intern-id arrays of
length `nnz` can share this layout later (#124 / #491); they are not built
here.

Index space is the graph's current node set in `_nodes` insertion order.
Endpoints that are not graph nodes are dropped (`dropped_endpoints`). CSR
`col_idx[row_ptr[i]:row_ptr[i+1]]` holds dependency row ids; CSC
`row_idx[col_ptr[j]:col_ptr[j+1]]` holds dependent row ids. There is no
`values` array (unweighted).

Build streams one adjacency row at a time from `_edges` and derives CSC from
the CSR arrays. It does not copy neighbor `set`s and does not materialize a
second Python edge index (no COO list of tuples, no extra `dict[NodeKey, set]`).
Scratch is `O(n)` uint32, not `O(nnz)` Python objects. The input `_edges` map
already exists; a production landing must drop it after rebuild so peak RAM is
CSR+CSC rather than both representations.

Implied API breaks for a real landing are listed on `LANDING_API_BREAKS`.
"""

from __future__ import annotations

import array
from collections.abc import Callable, Iterable, Sequence
from dataclasses import dataclass

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import NodeKey

_UINT32_MAX = 0xFFFFFFFF
_ZERO = array.array("I", [0])

LANDING_API_BREAKS: tuple[str, ...] = (
    "Burst add_edge / _remove_edge goes away: callers rewrite formulas/nodes, then one rebuild.",
    "Direct _edges / _reverse_edges reads (formula rewrite, compression, subgraph, pickle, viz, consistency) must go through CSR/CSC or a rebuild.",
    "get_dependencies / get_dependents returning frozenset allocates; bulk walks should iterate uint32 slices.",
    "_all_adjacency / _unconditional_adjacency / cycle_report / evaluation_order must not rebuild dict[NodeKey, set] or the CSR win is erased.",
    "_copy_for_projection cannot clone per-node sets; clone arrays or rebuild after the rewrite burst.",
    "Dangling (non-node) endpoints are out of the node-indexed matrix; drop or intern them before rehydrate.",
    "Neighbor iteration order is no longer set-hash order; evaluation_order must sort by workbook order on read or store sorted rows.",
    "roots() and leaf detection that probe reverse/forward maps become degree checks on col_ptr / row_ptr.",
    "JSON/pickle already emit COO; they can write CSR frames directly and skip reconstituting sets.",
    "Guard/provenance stay on dict[EdgeKey, ...] until a follow-up nnz-aligned *_id array lands.",
)


def _uint32_zeros(length: int) -> array.array[int]:
    """Return a zero-filled uint32 array of `length` without a Python int list."""
    if length == 0:
        return array.array("I")
    return _ZERO * length


def _require_uint32(value: int, *, what: str) -> None:
    if value > _UINT32_MAX:
        raise OverflowError(f"{what} {value} exceeds uint32")


@dataclass(frozen=True, slots=True)
class CsrBuildStats:
    """Peak-relevant counters from a compact CSR+CSC build."""

    n: int
    nnz: int
    dropped_endpoints: int
    scratch_uint32: int
    held_input_adjacency: bool
    built_second_python_edge_index: bool


@dataclass(frozen=True, slots=True)
class CsrCscAdjacency:
    """Unweighted CSR + CSC adjacency over a dense node-id space.

    Attributes:
        keys: Node key for each row/column id (`keys[i]`).
        index: `NodeKey` -> row id. Keys are the same objects as `keys`.
        row_ptr: CSR row pointers, length `n + 1`.
        col_idx: CSR dependency ids, length `nnz`.
        col_ptr: CSC column pointers, length `n + 1`.
        row_idx: CSC dependent ids, length `nnz`.
        dropped_endpoints: Neighbor keys that were not graph nodes.
        stats: Build counters (scratch vs a second Python edge index).
    """

    keys: tuple[NodeKey, ...]
    index: dict[NodeKey, int]
    row_ptr: array.array[int]
    col_idx: array.array[int]
    col_ptr: array.array[int]
    row_idx: array.array[int]
    dropped_endpoints: int
    stats: CsrBuildStats

    @property
    def n(self) -> int:
        """Number of node rows (and columns)."""
        return len(self.keys)

    @property
    def nnz(self) -> int:
        """Number of stored in-graph arcs."""
        return len(self.col_idx)

    def dependencies(self, key: NodeKey) -> frozenset[NodeKey]:
        """Return in-graph dependency keys for `key` (empty if unknown)."""
        i = self.index.get(key)
        if i is None:
            return frozenset()
        start, end = self.row_ptr[i], self.row_ptr[i + 1]
        keys = self.keys
        col = self.col_idx
        return frozenset(keys[col[k]] for k in range(start, end))

    def dependents(self, key: NodeKey) -> frozenset[NodeKey]:
        """Return in-graph dependent keys for `key` (empty if unknown)."""
        j = self.index.get(key)
        if j is None:
            return frozenset()
        start, end = self.col_ptr[j], self.col_ptr[j + 1]
        keys = self.keys
        rows = self.row_idx
        return frozenset(keys[rows[k]] for k in range(start, end))

    def visit_forward_ids(self) -> int:
        """Visit every CSR arc; return `sum(col_idx)` as a checksum."""
        total = 0
        col = self.col_idx
        ptr = self.row_ptr
        for i in range(self.n):
            for k in range(ptr[i], ptr[i + 1]):
                total += col[k]
        return total

    def visit_reverse_ids(self) -> int:
        """Visit every CSC arc; return `sum(row_idx)` as a checksum."""
        total = 0
        rows = self.row_idx
        ptr = self.col_ptr
        for j in range(self.n):
            for k in range(ptr[j], ptr[j + 1]):
                total += rows[k]
        return total


def _csr_to_csc(
    row_ptr: array.array[int],
    col_idx: array.array[int],
    n: int,
) -> tuple[array.array[int], array.array[int]]:
    """Return `(col_ptr, row_idx)` for the transpose of CSR, using `O(n)` scratch."""
    nnz = len(col_idx)
    col_deg = _uint32_zeros(n)
    for j in col_idx:
        col_deg[j] += 1
    col_ptr = _uint32_zeros(n + 1)
    for j in range(n):
        col_ptr[j + 1] = col_ptr[j] + col_deg[j]
    row_idx = _uint32_zeros(nnz)
    next_slot = col_ptr * 1
    for i in range(n):
        for k in range(row_ptr[i], row_ptr[i + 1]):
            j = col_idx[k]
            slot = next_slot[j]
            row_idx[slot] = i
            next_slot[j] = slot + 1
    return col_ptr, row_idx


def build_csr_csc(
    keys: Sequence[NodeKey],
    neighbors_for: Callable[[NodeKey], Iterable[NodeKey]],
    *,
    is_node: Callable[[NodeKey], bool] | None = None,
    held_input_adjacency: bool = False,
) -> CsrCscAdjacency:
    """Build CSR+CSC from a per-row neighbor iterator.

    Materializes one source row at a time. CSC is derived from CSR, so a reverse
    adjacency map is not required. `is_node` defaults to membership in `keys`.
    """
    n = len(keys)
    _require_uint32(n, what="node count")
    key_tuple = tuple(keys)
    index = {key: i for i, key in enumerate(key_tuple)}
    in_graph = is_node if is_node is not None else index.__contains__

    degrees = _uint32_zeros(n)
    dropped = 0
    nnz = 0
    for i, key in enumerate(key_tuple):
        count = 0
        for dep in neighbors_for(key):
            if in_graph(dep):
                count += 1
            else:
                dropped += 1
        degrees[i] = count
        nnz += count
    _require_uint32(nnz, what="nnz")

    row_ptr = _uint32_zeros(n + 1)
    for i in range(n):
        row_ptr[i + 1] = row_ptr[i] + degrees[i]

    col_idx = _uint32_zeros(nnz)
    for i, key in enumerate(key_tuple):
        slot = row_ptr[i]
        for dep in neighbors_for(key):
            j = index.get(dep)
            if j is None:
                continue
            col_idx[slot] = j
            slot += 1

    col_ptr, row_idx = _csr_to_csc(row_ptr, col_idx, n)
    stats = CsrBuildStats(
        n=n,
        nnz=nnz,
        dropped_endpoints=dropped,
        scratch_uint32=n,  # next_slot copy of col_ptr; degrees reused as a pass
        held_input_adjacency=held_input_adjacency,
        built_second_python_edge_index=False,
    )
    return CsrCscAdjacency(
        keys=key_tuple,
        index=index,
        row_ptr=row_ptr,
        col_idx=col_idx,
        col_ptr=col_ptr,
        row_idx=row_idx,
        dropped_endpoints=dropped,
        stats=stats,
    )


def from_graph(graph: DependencyGraph) -> CsrCscAdjacency:
    """Build a CSR+CSC sidecar from `graph._edges`, one row at a time.

    Does not read `_reverse_edges`. CSC is the transpose of the filtered CSR.
    """
    nodes = graph._nodes
    edges = graph._edges
    empty: tuple[NodeKey, ...] = ()

    def neighbors_for(key: NodeKey) -> Iterable[NodeKey]:
        return edges.get(key, empty)

    return build_csr_csc(
        tuple(nodes),
        neighbors_for,
        is_node=nodes.__contains__,
        held_input_adjacency=True,
    )
