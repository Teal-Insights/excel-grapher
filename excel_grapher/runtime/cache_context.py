"""EvalContext definitions for export runtime (slim base and invalidating full)."""

from __future__ import annotations

from collections.abc import Callable, Hashable, Iterable
from dataclasses import dataclass, field
from typing import Any, TypeAlias

from excel_grapher.core import CellValue
from excel_grapher.core.address_keys import parse_cell_coords

from .leaves import LeafInputs, LeafStore, as_leaf_store, overlay_leaf_inputs

__all__ = ["EvalContext", "EvalContextBase", "HelperCacheKey", "RangeWatch"]

HelperCacheKey: TypeAlias = tuple[Hashable, tuple[tuple[str, Hashable], ...]]
RangeWatch: TypeAlias = tuple[str, int, int, int, int]


def _coalesce_range_watches(watches: list[RangeWatch]) -> list[RangeWatch]:
    """Merge same-sheet watches whose union is a hole-free rectangle."""
    if len(watches) <= 1:
        return list(watches)
    remaining = list(watches)
    merged = True
    while merged:
        merged = False
        out: list[RangeWatch] = []
        for watch in remaining:
            placed = False
            for index, existing in enumerate(out):
                combo = _try_merge_range_watches(existing, watch)
                if combo is not None:
                    out[index] = combo
                    placed = True
                    merged = True
                    break
            if not placed:
                out.append(watch)
        remaining = out
    return remaining


def _try_merge_range_watches(left: RangeWatch, right: RangeWatch) -> RangeWatch | None:
    if left[0] != right[0]:
        return None
    _sheet, ar1, ac1, ar2, ac2 = left
    _, br1, bc1, br2, bc2 = right
    r1, c1 = min(ar1, br1), min(ac1, bc1)
    r2, c2 = max(ar2, br2), max(ac2, bc2)
    union_area = (r2 - r1 + 1) * (c2 - c1 + 1)
    area_left = (ar2 - ar1 + 1) * (ac2 - ac1 + 1)
    area_right = (br2 - br1 + 1) * (bc2 - bc1 + 1)
    ir1, ic1 = max(ar1, br1), max(ac1, bc1)
    ir2, ic2 = min(ar2, br2), min(ac2, bc2)
    inter = (ir2 - ir1 + 1) * (ic2 - ic1 + 1) if ir1 <= ir2 and ic1 <= ic2 else 0
    if union_area != area_left + area_right - inter:
        return None
    return (_sheet, r1, c1, r2, c2)


def _try_parse_cell_coords(address: str) -> tuple[str, int, int] | None:
    try:
        return parse_cell_coords(address)
    except ValueError:
        return None


@dataclass(slots=True)
class EvalContextBase:
    """Per-run evaluation state without dependency-tracking fields."""

    inputs: Any
    resolver: Callable[[str], Callable[[EvalContext], CellValue] | None]
    cache: dict[str, CellValue] = field(default_factory=dict)
    computing: set[str] = field(default_factory=set)
    circular_warning_roots: set[str] = field(default_factory=set)
    helper_cache: dict[HelperCacheKey, CellValue] = field(default_factory=dict)
    helper_computing: set[HelperCacheKey] = field(default_factory=set)
    iterative_enabled: bool = False
    iterate_count: int = 100
    iterate_delta: float = 0.001
    iteration_values: dict[str, CellValue] = field(default_factory=dict)
    leaves: LeafStore = field(default_factory=dict)

    def __post_init__(self) -> None:
        store = as_leaf_store(self.inputs) if self.inputs else as_leaf_store(self.leaves)
        self.leaves = store
        self.inputs = LeafInputs(store)


@dataclass(slots=True)
class EvalContext(EvalContextBase):
    """Per-run evaluation state with dependency tracking for input invalidation.

    Formula-to-formula edges stay in `deps` / `reverse_deps`. Leaf-safe
    `xl_range` rectangles are recorded as `range_watches` and tested with
    point-in-rect on `invalidate` / `set_inputs`.
    """

    deps: dict[str, set[str]] = field(default_factory=dict)
    reverse_deps: dict[str, set[str]] = field(default_factory=dict)
    stack: list[str] = field(default_factory=list)
    range_watches: dict[str, list[RangeWatch]] = field(default_factory=dict)

    def _record_dependency(self, parent: str, child: str) -> None:
        if parent == child:
            return
        self.deps.setdefault(parent, set()).add(child)
        self.reverse_deps.setdefault(child, set()).add(parent)

    def _record_range_watch(
        self, parent: str, sheet: str, r1: int, c1: int, r2: int, c2: int
    ) -> None:
        watch: RangeWatch = (sheet, r1, c1, r2, c2)
        existing = self.range_watches.get(parent, [])
        self.range_watches[parent] = _coalesce_range_watches([*existing, watch])

    def _formulas_watching(self, sheet: str, row: int, col: int) -> list[str]:
        hit: list[str] = []
        for parent, watches in self.range_watches.items():
            for w_sheet, r1, c1, r2, c2 in watches:
                if w_sheet == sheet and r1 <= row <= r2 and c1 <= col <= c2:
                    hit.append(parent)
                    break
        return hit

    def invalidate(self, addresses: Iterable[str]) -> None:
        """Invalidate cached values for the given addresses and their dependents.

        Changed cells also wake formulas whose live range watches contain that
        point. Invalidated formulas drop outgoing cell deps and range watches
        so the next eval can re-bind OFFSET / INDIRECT / INDEX.

        Helper memos are not address-dep-tracked, so any address invalidation
        clears `helper_cache` and `helper_computing` entirely.
        """
        self.helper_cache.clear()
        self.helper_computing.clear()

        to_visit = list(addresses)
        seen_seed = set(to_visit)
        for addr in list(to_visit):
            coords = _try_parse_cell_coords(addr)
            if coords is None:
                continue
            sheet, row, col = coords
            for parent in self._formulas_watching(sheet, row, col):
                if parent not in seen_seed:
                    to_visit.append(parent)
                    seen_seed.add(parent)

        seen: set[str] = set()
        while to_visit:
            addr = to_visit.pop()
            if addr in seen:
                continue
            seen.add(addr)

            self.cache.pop(addr, None)
            self.circular_warning_roots.discard(addr)
            self.computing.discard(addr)

            dependents = list(self.reverse_deps.get(addr, set()))
            to_visit.extend(dependents)

            for dep in self.deps.get(addr, set()):
                parents = self.reverse_deps.get(dep)
                if parents is not None:
                    parents.discard(addr)
                    if not parents:
                        self.reverse_deps.pop(dep, None)

            self.deps.pop(addr, None)
            self.reverse_deps.pop(addr, None)
            self.range_watches.pop(addr, None)

    def set_inputs(self, inputs: dict[str, CellValue]) -> None:
        """Update input values and invalidate dependent cached results.

        `inputs` is NodeKey-keyed. Keys that cannot round-trip to
        `(sheet, row, col)` raise `ValueError`.
        """
        parsed = as_leaf_store(inputs)
        changed = [k for k, v in inputs.items() if self.inputs.get(k) != v]
        overlay_leaf_inputs(self.leaves, parsed)
        if changed:
            self.invalidate(changed)
