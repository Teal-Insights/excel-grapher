"""Compact uint32 CSR+CSC adjacency for `DependencyGraph`.

Index space is the graph's current node set. CSR `col_idx[row_ptr[i]:row_ptr[i+1]]`
holds dependency row ids; CSC `row_idx[col_ptr[j]:col_ptr[j+1]]` holds dependent
row ids. There is no `values` array (unweighted).

Build streams one adjacency row at a time and derives CSC from CSR. Scratch is
`O(n)` uint32, not `O(nnz)` Python objects. Endpoints that are not graph nodes
are dropped.
"""

from __future__ import annotations

import array
from collections.abc import Callable, Iterable, Sequence
from dataclasses import dataclass

from excel_grapher.grapher.node import NodeKey

_UINT32_MAX = 0xFFFFFFFF
_ZERO = array.array("I", [0])


def uint32_zeros(length: int) -> array.array[int]:
    """Return a zero-filled uint32 array of `length` without a Python int list."""
    if length == 0:
        return array.array("I")
    return _ZERO * length


def require_uint32(value: int, *, what: str) -> None:
    """Raise `OverflowError` when `value` does not fit in uint32."""
    if value > _UINT32_MAX:
        raise OverflowError(f"{what} {value} exceeds uint32")


@dataclass(frozen=True, slots=True)
class CsrCscArrays:
    """Unweighted CSR + CSC arrays over a dense node-id space.

    Attributes:
        keys: Node key for each row/column id (`keys[i]`).
        index: `NodeKey` -> row id. Keys are the same objects as `keys`.
        row_ptr: CSR row pointers, length `n + 1`.
        col_idx: CSR dependency ids, length `nnz`.
        col_ptr: CSC column pointers, length `n + 1`.
        row_idx: CSC dependent ids, length `nnz`.
        dropped_endpoints: Neighbor keys that were not graph nodes.
    """

    keys: list[NodeKey]
    index: dict[NodeKey, int]
    row_ptr: array.array[int]
    col_idx: array.array[int]
    col_ptr: array.array[int]
    row_idx: array.array[int]
    dropped_endpoints: int

    @property
    def n(self) -> int:
        """Number of node rows (and columns)."""
        return len(self.keys)

    @property
    def nnz(self) -> int:
        """Number of stored in-graph arcs."""
        return len(self.col_idx)


def csr_to_csc(
    row_ptr: array.array[int],
    col_idx: array.array[int],
    n: int,
) -> tuple[array.array[int], array.array[int]]:
    """Return `(col_ptr, row_idx)` for the transpose of CSR, using `O(n)` scratch."""
    nnz = len(col_idx)
    col_deg = uint32_zeros(n)
    for j in col_idx:
        col_deg[j] += 1
    col_ptr = uint32_zeros(n + 1)
    for j in range(n):
        col_ptr[j + 1] = col_ptr[j] + col_deg[j]
    row_idx = uint32_zeros(nnz)
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
) -> CsrCscArrays:
    """Build CSR+CSC from a per-row neighbor iterator.

    Materializes one source row at a time. CSC is derived from CSR, so a reverse
    adjacency map is not required. `is_node` defaults to membership in `keys`.
    """
    n = len(keys)
    require_uint32(n, what="node count")
    key_list = list(keys)
    index = {key: i for i, key in enumerate(key_list)}
    in_graph = is_node if is_node is not None else index.__contains__

    degrees = uint32_zeros(n)
    dropped = 0
    nnz = 0
    for i, key in enumerate(key_list):
        count = 0
        for dep in neighbors_for(key):
            if in_graph(dep):
                count += 1
            else:
                dropped += 1
        degrees[i] = count
        nnz += count
    require_uint32(nnz, what="nnz")

    row_ptr = uint32_zeros(n + 1)
    for i in range(n):
        row_ptr[i + 1] = row_ptr[i] + degrees[i]

    col_idx = uint32_zeros(nnz)
    for i, key in enumerate(key_list):
        slot = row_ptr[i]
        for dep in neighbors_for(key):
            j = index.get(dep)
            if j is None:
                continue
            col_idx[slot] = j
            slot += 1

    col_ptr, row_idx = csr_to_csc(row_ptr, col_idx, n)
    return CsrCscArrays(
        keys=key_list,
        index=index,
        row_ptr=row_ptr,
        col_idx=col_idx,
        col_ptr=col_ptr,
        row_idx=row_idx,
        dropped_endpoints=dropped,
    )
