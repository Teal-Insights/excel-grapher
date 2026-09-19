"""nnz-aligned intern-id sidecars for edge guards and provenance.

CSR `col_idx[k]` already stores the destination of forward arc `k`. These arrays
store intern ids in the same order: `0` means unset, any other value is an
index into the corresponding intern table.
"""

from __future__ import annotations

import array
from collections.abc import Callable
from dataclasses import dataclass
from typing import Any

from excel_grapher.grapher.csr_adjacency import CsrCscArrays, uint32_zeros
from excel_grapher.grapher.dependency_provenance import EdgeProvenance
from excel_grapher.grapher.guard import GuardExpr, intern_guard
from excel_grapher.grapher.node import NodeKey

EdgeMetaLookup = Callable[[NodeKey, NodeKey], tuple[GuardExpr | None, EdgeProvenance | None]]


def empty_intern_table() -> list[Any]:
    """Return a 1-based intern table whose index `0` slot is unused."""
    return [None]


@dataclass(frozen=True, slots=True)
class EdgeMetaArrays:
    """CSR-nnz intern-id arrays plus intern tables.

    Attributes:
        guard_id: Per-arc intern id; `0` means unguarded.
        prov_id: Per-arc intern id; `0` means no provenance.
        guard_exprs: Intern table; payload for id `i` is `guard_exprs[i]`.
        provenances: Intern table; payload for id `i` is `provenances[i]`.
    """

    guard_id: array.array[int]
    prov_id: array.array[int]
    guard_exprs: list[GuardExpr | None]
    provenances: list[EdgeProvenance | None]


def empty_edge_meta(nnz: int = 0) -> EdgeMetaArrays:
    """Return zero-filled intern-id arrays of length `nnz` and empty tables."""
    return EdgeMetaArrays(
        guard_id=uint32_zeros(nnz),
        prov_id=uint32_zeros(nnz),
        guard_exprs=empty_intern_table(),
        provenances=empty_intern_table(),
    )


def pack_edge_meta_ids(built: CsrCscArrays, lookup: EdgeMetaLookup) -> EdgeMetaArrays:
    """Fill nnz-parallel intern ids from `(src, dst) -> (guard, provenance)`.

    Guards are interned by `intern_guard` identity. Equal `EdgeProvenance`
    values share one intern id. Lookups that miss (including dangling
    endpoints dropped from `built`) leave id `0`.
    """
    nnz = built.nnz
    guard_id = uint32_zeros(nnz)
    prov_id = uint32_zeros(nnz)
    guard_exprs: list[GuardExpr | None] = empty_intern_table()
    provenances: list[EdgeProvenance | None] = empty_intern_table()
    guard_intern: dict[int, int] = {}
    prov_intern: dict[EdgeProvenance, int] = {}
    keys = built.keys
    col = built.col_idx
    row_ptr = built.row_ptr
    for i, src in enumerate(keys):
        for k in range(row_ptr[i], row_ptr[i + 1]):
            dst = keys[col[k]]
            guard, prov = lookup(src, dst)
            if guard is not None:
                interned = intern_guard(guard)
                gid = guard_intern.get(id(interned))
                if gid is None:
                    gid = len(guard_exprs)
                    guard_exprs.append(interned)
                    guard_intern[id(interned)] = gid
                guard_id[k] = gid
            if prov is not None:
                pid = prov_intern.get(prov)
                if pid is None:
                    pid = len(provenances)
                    provenances.append(prov)
                    prov_intern[prov] = pid
                prov_id[k] = pid
    return EdgeMetaArrays(
        guard_id=guard_id,
        prov_id=prov_id,
        guard_exprs=guard_exprs,
        provenances=provenances,
    )
