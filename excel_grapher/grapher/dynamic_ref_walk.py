"""Memoized static-ref walks over OFFSET / INDEX / INDIRECT argument subgraphs.

`create_dependency_graph` builds one `DynamicRefWalkContext` per graph and
passes it to both dependency extraction and provenance collection so the two
paths share `(formula, sheet)` ref sets, per-address nodes, and argument-subgraph
BFS results.
"""

from __future__ import annotations

from collections.abc import Callable, Collection, Iterable

from excel_grapher.core.address_keys import parse_address
from excel_grapher.core.cell_types import CellType

from .parser import (
    FormulaNormalizer,
    _find_function_calls_with_spans,
    expand_range,
    format_key,
    mask_ref_only_function_calls,
    mask_spans,
    parse_range_refs_with_spans,
    parse_standalone_cell_refs,
)

DYNAMIC_REF_FUNCTION_NAMES = frozenset({"OFFSET", "INDIRECT", "INDEX"})


class DynamicRefWalkContext:
    """Caches for one graph build's argument-subgraph static-ref walk.

    Returned ref sets are fresh mutable copies so callers that mutate in place
    (for example `expand_leaf_env_to_argument_env`'s `refs |= ...`) cannot
    poison later lookups.
    """

    def __init__(
        self,
        *,
        normalizer: FormulaNormalizer,
        max_range_cells: int,
        get_cell_value: Callable[[str, str], object],
        sheet_names: Collection[str],
        shared_cell_type_cache: dict[str, CellType] | None = None,
        stats: dict[str, int] | None = None,
    ) -> None:
        self._normalizer = normalizer
        self._max_range_cells = max_range_cells
        self._get_cell_value = get_cell_value
        self._sheet_names = (
            sheet_names if isinstance(sheet_names, (set, frozenset)) else set(sheet_names)
        )
        self.shared_cell_type_cache: dict[str, CellType] = (
            shared_cell_type_cache if shared_cell_type_cache is not None else {}
        )
        self._stats = stats
        self._refs_without_dynamic_cache: dict[tuple[str, str], frozenset[str]] = {}
        self._arg_subgraph_cache: dict[frozenset[str], tuple[frozenset[str], frozenset[str]]] = {}
        self._arg_node_cache: dict[str, tuple[frozenset[str], bool]] = {}

    def refs_in_formula_without_dynamic(self, formula_str: str, sheet_of_cell: str) -> set[str]:
        """Return static (non-dynamic-ref) cell addresses referenced by `formula_str`."""
        f = formula_str if formula_str.startswith("=") else "=" + formula_str
        cache_key = (f, sheet_of_cell)
        cached = self._refs_without_dynamic_cache.get(cache_key)
        if cached is not None:
            return set(cached)
        dyn = _find_function_calls_with_spans(f, DYNAMIC_REF_FUNCTION_NAMES)
        spans = [span for _fn, _inner, span in dyn]
        masked = mask_spans(f, spans)
        masked = mask_ref_only_function_calls(masked)
        norm = self._normalizer.normalize(masked, sheet_of_cell)
        out: set[str] = set()
        for ref in parse_standalone_cell_refs(norm):
            sh = ref.sheet if ref.sheet is not None else sheet_of_cell
            out.add(format_key(sh, f"{ref.column}{ref.row}"))
        for start, end, _span in parse_range_refs_with_spans(norm):
            sh = start.sheet if start.sheet is not None else sheet_of_cell
            for dep_sheet, dep_a1 in expand_range(
                sheet=sh,
                start_col=start.column,
                start_row=start.row,
                end_col=end.column,
                end_row=end.row,
                max_cells=self._max_range_cells,
            ):
                out.add(format_key(dep_sheet, dep_a1))
        self._refs_without_dynamic_cache[cache_key] = frozenset(out)
        return out

    def cell_formula(self, addr: str) -> str | None:
        """Return the normalized formula at `addr`, or None if it is not a formula cell."""
        sheet, a1 = parse_address(addr)
        if sheet not in self._sheet_names:
            return None
        value = self._get_cell_value(sheet, a1)
        if not isinstance(value, str) or not value.startswith("="):
            return None
        return self._normalizer.normalize(value, sheet)

    def argument_node(self, addr: str) -> tuple[frozenset[str], bool]:
        """Return static child refs and whether `addr` is a non-formula leaf."""
        cached = self._arg_node_cache.get(addr)
        if cached is not None:
            return cached
        children: frozenset[str] = frozenset()
        is_leaf = False
        sheet, a1 = parse_address(addr)
        if sheet in self._sheet_names:
            cell_val = self._get_cell_value(sheet, a1)
            if isinstance(cell_val, str) and cell_val.startswith("="):
                children = frozenset(self.refs_in_formula_without_dynamic(cell_val, sheet))
            else:
                is_leaf = True
        result = (children, is_leaf)
        self._arg_node_cache[addr] = result
        return result

    def argument_subgraph_refs(self, argument_addrs: Iterable[str]) -> tuple[set[str], set[str]]:
        """Return statically reachable refs and leaves feeding dynamic-ref arguments."""
        cache_key = frozenset(argument_addrs)
        cached = self._arg_subgraph_cache.get(cache_key)
        if cached is not None:
            if self._stats is not None:
                self._stats["arg_subgraph_hits"] = self._stats.get("arg_subgraph_hits", 0) + 1
            return set(cached[0]), set(cached[1])
        all_refs: set[str] = set()
        leaves: set[str] = set()
        to_visit = set(argument_addrs)
        while to_visit:
            addr = to_visit.pop()
            if addr in all_refs:
                continue
            all_refs.add(addr)
            children, is_leaf = self.argument_node(addr)
            if is_leaf:
                leaves.add(addr)
            else:
                to_visit.update(children)
        self._arg_subgraph_cache[cache_key] = (frozenset(all_refs), frozenset(leaves))
        return set(all_refs), set(leaves)
