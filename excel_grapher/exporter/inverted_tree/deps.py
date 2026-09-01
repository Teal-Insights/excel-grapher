"""First-level series dependencies, leaf closures, and fail-closed checks."""

from __future__ import annotations

from collections.abc import Iterable, Iterator
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from fastpyxl.utils.cell import get_column_letter

from excel_grapher.core.address_keys import (
    format_cell_key,
    parse_cell_coords,
)
from excel_grapher.core.address_keys import (
    normalize_key as normalize_address,
)
from excel_grapher.core.excel_function_names import normalize_excel_function_name
from excel_grapher.core.formula_ast import (
    AstNode,
    BinaryOpNode,
    CellRefNode,
    FunctionCallNode,
    NumberNode,
    RangeNode,
    UnaryOpNode,
    resolve_cell_ref,
)
from excel_grapher.exporter.inverted_tree.catalog import BoundSeries, SeriesCatalog, covering_series
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError

if TYPE_CHECKING:
    from excel_grapher.grapher.graph import DependencyGraph


def iter_range_addresses(start: str, end: str) -> list[str]:
    """Expand a same-sheet A1 range into canonical cell addresses (row-major)."""
    sheet1, row1, col1 = parse_cell_coords(start)
    sheet2, row2, col2 = parse_cell_coords(end)
    if sheet1 != sheet2:
        raise InvertedTreeExportError(f"cross-sheet range {start}:{end} is not supported")
    r1, r2 = min(row1, row2), max(row1, row2)
    c1, c2 = min(col1, col2), max(col1, col2)
    return [
        format_cell_key(sheet1, get_column_letter(col), row)
        for row in range(r1, r2 + 1)
        for col in range(c1, c2 + 1)
    ]


def range_column_addresses(start: str, end: str, col_index: int) -> list[str]:
    """Return the 1-based column `col_index` of the rectangle `start:end`."""
    sheet1, row1, col1 = parse_cell_coords(start)
    sheet2, row2, col2 = parse_cell_coords(end)
    if sheet1 != sheet2:
        raise InvertedTreeExportError(f"cross-sheet range {start}:{end} is not supported")
    r1, r2 = min(row1, row2), max(row1, row2)
    c1, c2 = min(col1, col2), max(col1, col2)
    width = c2 - c1 + 1
    if col_index < 1 or col_index > width:
        raise InvertedTreeExportError(f"INDEX column {col_index} is outside range {start}:{end}")
    col = c1 + col_index - 1
    return [format_cell_key(sheet1, get_column_letter(col), row) for row in range(r1, r2 + 1)]


def node_formula_ast(graph: DependencyGraph, address: str) -> AstNode:
    """Return the formula AST for `address`, or fail closed."""
    node = graph.get_node(normalize_address(address))
    if node is None:
        raise InvertedTreeExportError(f"graph is missing bound cell {address}")
    ast = getattr(node, "formula_ast", None)
    if ast is None:
        raise InvertedTreeExportError(
            f"bound cell {address} has no formula AST (cannot verify first-level refs)"
        )
    return ast


def predecessor_address(series: BoundSeries, index: int, catalog: SeriesCatalog) -> str | None:
    """Return the lagged predecessor of `series.cells[index]`, if any.

    For index > 0 this is the previous cell in the series. For index 0 it is the
    same-row previous column when that cell belongs to a different bound series
    (the year-0 scalar of a recursive path).
    """
    if index < 0 or index >= len(series.cells):
        return None
    if index > 0:
        return series.cells[index - 1]
    sheet, row, col = parse_cell_coords(series.cells[0])
    if col <= 1:
        return None
    prev = format_cell_key(sheet, get_column_letter(col - 1), row)
    owner = catalog.series_id_for(prev)
    if owner is None or owner == series.series_id:
        return None
    return prev


@dataclass
class SeriesDeps:
    """First-level bound-series dependencies of one formula series."""

    host_id: str
    param_ids: tuple[str, ...]
    is_scan: bool
    seed_id: str | None
    aligned_ids: frozenset[str]
    lookup_ids: frozenset[str]


@dataclass
class _DepCollector:
    host: BoundSeries
    catalog: SeriesCatalog
    params: dict[str, None] = field(default_factory=dict)
    lookup_ids: set[str] = field(default_factory=set)
    aligned_hits: dict[str, list[tuple[int, int]]] = field(default_factory=dict)
    saw_self_lag: bool = False
    seed_id: str | None = None

    def add_param(self, series_id: str, *, lookup: bool = False) -> None:
        if series_id == self.host.series_id:
            return
        self.params.setdefault(series_id, None)
        if lookup:
            self.lookup_ids.add(series_id)

    def note_cell(self, address: str, host_index: int) -> None:
        owner = self.catalog.require_series_for(address)
        if owner.series_id == self.host.series_id:
            pred = predecessor_address(self.host, host_index, self.catalog)
            if pred is not None and normalize_address(address) == normalize_address(pred):
                self.saw_self_lag = True
                return
            if normalize_address(address) == self.host.cells[host_index]:
                return
            raise InvertedTreeExportError(
                f"series {self.host.series_id!r} cell {self.host.cells[host_index]} "
                f"references non-lag cell {address} in the same series"
            )
        idx = owner.index_of(address)
        if idx is not None:
            self.aligned_hits.setdefault(owner.series_id, []).append((host_index, idx))
        if host_index == 0:
            pred = predecessor_address(self.host, 0, self.catalog)
            if pred is not None and normalize_address(address) == normalize_address(pred):
                self.seed_id = owner.series_id
                self.add_param(owner.series_id)
                return
        self.add_param(owner.series_id)

    def visit(self, node: AstNode, *, host_cell: str, host_index: int) -> None:
        match node:
            case CellRefNode():
                address = resolve_cell_ref(node, host_cell)
                self.note_cell(address, host_index)
            case RangeNode():
                start = resolve_cell_ref(node.start_ref, host_cell)
                end = resolve_cell_ref(node.end_ref, host_cell)
                addresses = iter_range_addresses(start, end)
                covered = covering_series(self.catalog, addresses)
                if covered is None:
                    missing = [
                        addr for addr in addresses if self.catalog.series_id_for(addr) is None
                    ]
                    raise InvertedTreeExportError(
                        f"series {self.host.series_id!r} range {start}:{end} is not a "
                        f"bound series (unbound cells: {missing[:8]})"
                    )
                self.add_param(covered.series_id, lookup=True)
            case FunctionCallNode():
                self._visit_function(node, host_cell=host_cell, host_index=host_index)
            case BinaryOpNode():
                self.visit(node.left, host_cell=host_cell, host_index=host_index)
                self.visit(node.right, host_cell=host_cell, host_index=host_index)
            case UnaryOpNode():
                self.visit(node.operand, host_cell=host_cell, host_index=host_index)
            case _:
                return

    def _visit_function(
        self,
        node: FunctionCallNode,
        *,
        host_cell: str,
        host_index: int,
    ) -> None:
        name = normalize_excel_function_name(node.name)
        if name == "OFFSET":
            self._visit_offset(node, host_cell=host_cell, host_index=host_index)
            return
        if name == "INDEX":
            self._visit_index(node, host_cell=host_cell, host_index=host_index)
            return
        if name == "MATCH":
            self._visit_match(node, host_cell=host_cell, host_index=host_index)
            return
        for arg in node.args:
            self.visit(arg, host_cell=host_cell, host_index=host_index)

    def _visit_offset(
        self,
        node: FunctionCallNode,
        *,
        host_cell: str,
        host_index: int,
    ) -> None:
        if not node.args:
            raise InvertedTreeExportError(
                f"series {self.host.series_id!r}: OFFSET with no arguments"
            )
        table = self._series_for_ref(node.args[0], host_cell)
        self.add_param(table.series_id, lookup=True)
        for arg in node.args[1:]:
            self.visit(arg, host_cell=host_cell, host_index=host_index)

    def _visit_index(
        self,
        node: FunctionCallNode,
        *,
        host_cell: str,
        host_index: int,
    ) -> None:
        if len(node.args) < 2:
            raise InvertedTreeExportError(
                f"series {self.host.series_id!r}: INDEX expects a range and row"
            )
        row_arg = node.args[1]
        col_arg = node.args[2] if len(node.args) > 2 else None
        if isinstance(row_arg, FunctionCallNode) and (
            normalize_excel_function_name(row_arg.name) == "MATCH"
        ):
            self._visit_match(row_arg, host_cell=host_cell, host_index=host_index)
        else:
            self.visit(row_arg, host_cell=host_cell, host_index=host_index)
        col_index = 1
        if isinstance(col_arg, NumberNode):
            col_index = int(col_arg.value)
        elif col_arg is not None:
            self.visit(col_arg, host_cell=host_cell, host_index=host_index)
        if isinstance(node.args[0], RangeNode):
            start = resolve_cell_ref(node.args[0].start_ref, host_cell)
            end = resolve_cell_ref(node.args[0].end_ref, host_cell)
            column_cells = range_column_addresses(start, end, col_index)
            covered = covering_series(self.catalog, column_cells)
            if covered is None:
                raise InvertedTreeExportError(
                    f"series {self.host.series_id!r}: INDEX column {col_index} of "
                    f"{start}:{end} is not a bound series"
                )
            self.add_param(covered.series_id, lookup=True)
        else:
            self.visit(node.args[0], host_cell=host_cell, host_index=host_index)

    def _visit_match(
        self,
        node: FunctionCallNode,
        *,
        host_cell: str,
        host_index: int,
    ) -> None:
        if len(node.args) < 2:
            raise InvertedTreeExportError(
                f"series {self.host.series_id!r}: MATCH expects lookup and array"
            )
        self.visit(node.args[0], host_cell=host_cell, host_index=host_index)
        array_series = self._series_for_ref(node.args[1], host_cell)
        self.add_param(array_series.series_id, lookup=True)
        for arg in node.args[2:]:
            self.visit(arg, host_cell=host_cell, host_index=host_index)

    def _series_for_ref(self, node: AstNode, host_cell: str) -> BoundSeries:
        if isinstance(node, CellRefNode):
            return self.catalog.require_series_for(resolve_cell_ref(node, host_cell))
        if isinstance(node, RangeNode):
            start = resolve_cell_ref(node.start_ref, host_cell)
            end = resolve_cell_ref(node.end_ref, host_cell)
            covered = covering_series(self.catalog, iter_range_addresses(start, end))
            if covered is None:
                raise InvertedTreeExportError(
                    f"series {self.host.series_id!r}: reference {start}:{end} is not a bound series"
                )
            return covered
        raise InvertedTreeExportError(
            f"series {self.host.series_id!r}: expected a cell or range reference, "
            f"got {type(node).__name__}"
        )


def collect_series_deps(
    series: BoundSeries,
    *,
    catalog: SeriesCatalog,
    graph: DependencyGraph,
) -> SeriesDeps:
    """Collect first-level bound-series dependencies of `series`."""
    collector = _DepCollector(host=series, catalog=catalog)
    for index, address in enumerate(series.cells):
        ast = node_formula_ast(graph, address)
        collector.visit(ast, host_cell=address, host_index=index)
    aligned: set[str] = set()
    for series_id, pairs in collector.aligned_hits.items():
        if series_id in collector.lookup_ids:
            continue
        if series_id == collector.seed_id:
            continue
        dep = catalog.get(series_id)
        if dep.is_scalar:
            continue
        if any(host_i == dep_i for host_i, dep_i in pairs):
            aligned.add(series_id)
    is_scan = collector.saw_self_lag or collector.seed_id is not None
    remaining = [sid for sid in catalog.order if sid in collector.params]
    if collector.seed_id is not None and collector.seed_id in remaining:
        remaining.remove(collector.seed_id)
        param_ids = (collector.seed_id, *remaining)
    else:
        param_ids = tuple(remaining)
    return SeriesDeps(
        host_id=series.series_id,
        param_ids=param_ids,
        is_scan=is_scan,
        seed_id=collector.seed_id,
        aligned_ids=frozenset(aligned),
        lookup_ids=frozenset(collector.lookup_ids),
    )


def collect_all_deps(
    catalog: SeriesCatalog,
    graph: DependencyGraph,
) -> dict[str, SeriesDeps]:
    """Collect first-level deps for every formula series."""
    return {
        series.series_id: collect_series_deps(series, catalog=catalog, graph=graph)
        for series in catalog.formula_series()
    }


def leaf_closure(
    root_id: str,
    *,
    catalog: SeriesCatalog,
    deps: dict[str, SeriesDeps],
) -> tuple[str, ...]:
    """Return input and constant series in the subgraph of `root_id`.

    Order follows the bindings catalog (`catalog.order`).
    """
    seen: set[str] = set()
    stack = [root_id]
    while stack:
        current = stack.pop()
        if current in seen:
            continue
        seen.add(current)
        info = deps.get(current)
        if info is None:
            continue
        for param_id in info.param_ids:
            stack.append(param_id)
    inputs = [sid for sid in catalog.order if sid in seen and catalog.get(sid).direction == "input"]
    constants = [
        sid for sid in catalog.order if sid in seen and catalog.get(sid).direction == "constant"
    ]
    return tuple(inputs + constants)


def formula_closure(
    root_id: str,
    *,
    catalog: SeriesCatalog,
    deps: dict[str, SeriesDeps],
) -> tuple[str, ...]:
    """Return formula series in the subgraph of `root_id` (bindings order, topo)."""
    seen: set[str] = set()
    stack = [root_id]
    while stack:
        current = stack.pop()
        if current in seen:
            continue
        seen.add(current)
        info = deps.get(current)
        if info is None:
            continue
        for param_id in info.param_ids:
            if catalog.get(param_id).is_formula_series:
                stack.append(param_id)
    formula_ids = [
        sid for sid in catalog.order if sid in seen and catalog.get(sid).is_formula_series
    ]
    return tuple(_topo_sort(formula_ids, deps=deps))


def _topo_sort(series_ids: Iterable[str], *, deps: dict[str, SeriesDeps]) -> list[str]:
    selected = set(series_ids)
    remaining = set(selected)
    ordered: list[str] = []
    while remaining:
        ready = [
            sid
            for sid in series_ids
            if sid in remaining
            and all(
                pid not in remaining
                for pid in (deps[sid].param_ids if sid in deps else ())
                if pid in selected
            )
        ]
        if not ready:
            raise InvertedTreeExportError(
                f"cyclic formula-series dependencies among {sorted(remaining)}"
            )
        for sid in ready:
            remaining.remove(sid)
            ordered.append(sid)
    return ordered


def assert_subgraph_bound(
    *,
    catalog: SeriesCatalog,
    graph: DependencyGraph,
    roots: Iterable[str],
) -> None:
    """Fail closed when an unbound formula cell sits in a bound target subgraph."""
    bound = catalog.bound_addresses()
    seen: set[str] = set()
    stack = [normalize_address(addr) for addr in roots]
    while stack:
        address = stack.pop()
        if address in seen:
            continue
        seen.add(address)
        node = graph.get_node(address)
        if node is None:
            continue
        has_formula = bool(getattr(node, "has_formula", False))
        if has_formula and address not in bound:
            raise InvertedTreeExportError(
                f"unbound formula cell {address} is in the target subgraph"
            )
        for dep in graph.get_dependencies(address):
            stack.append(normalize_address(dep))


def all_formula_root_cells(catalog: SeriesCatalog) -> Iterator[str]:
    """Yield every cell of every formula series."""
    for series in catalog.formula_series():
        yield from series.cells
