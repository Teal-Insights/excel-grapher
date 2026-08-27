from __future__ import annotations

import hashlib
import logging
import math
import re
import time
from collections.abc import Callable, Iterable, Iterator, Mapping
from contextlib import contextmanager
from contextvars import ContextVar
from dataclasses import dataclass, field
from functools import lru_cache
from itertools import product
from pathlib import Path
from typing import TYPE_CHECKING, Any, cast, get_args, get_origin

if TYPE_CHECKING:
    from excel_grapher.grapher.type_analysis_cache import TypeAnalysisCache

import fastpyxl
from fastpyxl.utils.cell import coordinate_from_string, coordinate_to_tuple

from excel_grapher.core.address_keys import parse_address
from excel_grapher.core.addressing import (
    WorkbookBoundsProtocol,
    indirect_text_to_range,
    offset_range,
)
from excel_grapher.core.cell_types import (
    CellKind,
    CellType,
    CellTypeEnv,
    EnumDomain,
    GreaterThanCell,
    IntervalDomain,
    NotEqualCell,
    constraints_to_cell_type_env,
    normalize_cell_type_env_key,
)
from excel_grapher.core.excel_function_meta import is_ref_only_arg
from excel_grapher.core.expr_eval import Unsupported, evaluate_expr
from excel_grapher.core.formula_ast import (
    AstNode,
    BinaryOpNode,
    BoolNode,
    CellRefNode,
    EmptyArgNode,
    ErrorNode,
    FormulaParseError,
    FunctionCallNode,
    NumberNode,
    RangeNode,
    StringNode,
    UnaryOpNode,
    WholeColumnNode,
    WholeRowNode,
    resolve_whole_row_ref,
)
from excel_grapher.core.formula_ast import (
    parse as parse_ast,
)
from excel_grapher.core.range_shorthand import expand_whole_column_deps, expand_whole_row_deps
from excel_grapher.core.types import ExcelRange, XlError

from .parser import _find_function_calls_with_spans, expand_range, format_key

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Tracing infrastructure
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class DynamicRefTraceEvent:
    """A single trace event emitted during dynamic-ref inference.

    Attributes:
        kind: Event type (`"infer"`, `"expand-env"`, `"build-domains"`,
            `"build-value-domains"`, `"offset-scalar-fallback"`,
            `"offset-scalar-wide"`, plus `"-error"` variants).
        name: Function that emitted the event.
        elapsed_s: Wall-clock seconds spent in the traced operation.  Defaults
            to `0.0` for structural events (e.g. `"index-abstract"`,
            `"index-enumerated"`) that report *what* happened rather than
            how long it took.
        detail: Flexible per-event payload (target counts, branch estimates,
            expressions, etc.).
    """

    kind: str
    name: str
    elapsed_s: float = 0.0
    detail: dict[str, Any] = field(default_factory=dict)


DynamicRefTraceFn = Callable[[DynamicRefTraceEvent], None]
"""Callback signature for receiving trace events."""

_active_tracer: ContextVar[DynamicRefTraceFn | None] = ContextVar("_active_tracer", default=None)


@contextmanager
def trace_dynamic_refs(callback: DynamicRefTraceFn) -> Iterator[None]:
    """Activate *callback* as the dynamic-ref tracer for the enclosed block.

    Nestable: an inner `trace_dynamic_refs` overrides the outer one; the
    outer tracer is restored when the inner block exits.
    """
    token = _active_tracer.set(callback)
    try:
        yield
    finally:
        _active_tracer.reset(token)


def _emit_trace(event: DynamicRefTraceEvent) -> None:
    """Deliver *event* to the active tracer, if any."""
    tracer = _active_tracer.get()
    if tracer is not None:
        tracer(event)


_EXPAND_PROGRESS_INTERVAL = 50_000
"""Emit an `expand-env-progress` trace every N `cell_type_for` entries.

Argument-env expansion over a large workbook can run for minutes inside a
single call; without periodic progress the caller sees nothing between
`workbook-loaded` and the terminal `expand-env` event.
"""

_RANGE_EXPANSION_CACHE_SIZE = 16
"""Memo size for static-range address expansion.

The pattern worth caching is a formula chain that mentions the same handful of
static ranges at every level, so a small cache captures nearly all of the reuse
while bounding how many expanded ranges (up to `max_range_cells` strings each)
are held alive.
"""

_MAX_ANALYSIS_DEPTH = 600
"""Max nested `cell_type_for` frames before raising `DynamicRefError`.

Each analysis level costs roughly one Python frame, so this trips well inside
CPython's default 1000-frame recursion limit: a pathologically deep argument
chain fails with an actionable error instead of a bare `RecursionError`.
"""


class DynamicRefError(ValueError):
    """Raised when dynamic reference analysis cannot proceed.

    When building a dependency graph, pass a `DynamicRefConfig` (e.g. via
    `DynamicRefConfig.from_constraints`) or set `use_cached_dynamic_refs=True`
    to resolve OFFSET/INDIRECT instead of raising.
    """


def _apply_constraint_to_schema(schema: dict[str, Any], address: str, annotation: Any) -> None:
    """Assign *annotation* to every sheet-qualified cell implied by *address* in *schema*."""
    sheet_name, range_a1 = _split_addr_sheet_coord(address)
    cells = _expand_sheet_qualified_range(sheet_name, range_a1)
    for key in cells:
        schema[key] = annotation


@dataclass(frozen=True)
class DynamicRefLimits:
    """Tuneable safety limits for dynamic-reference inference.

    Pass a custom instance via the `limits` parameter of
    `DynamicRefConfig.from_constraints` or
    `DynamicRefConfig.from_constraints_and_workbook` to override any of these defaults.

    Attributes:
        max_branches: Maximum number of discrete value assignments explored
            during constraint enumeration.  This cap is applied in two places:

            * **Per-dependency domain size** - a single cell constrained to an
              integer interval wider than *max_branches* values cannot be
              enumerated; the caller must either tighten the constraint or rely
              on the symbolic (abstract) analysis path.
            * **Cartesian-product size** - when a formula cell falls back to
              brute-force evaluation over all combinations of its dependencies'
              domains, the product of those domain sizes must not exceed
              *max_branches*.  If it does, a `DynamicRefError` is raised
              immediately (rather than hanging) with a breakdown of which
              dependencies contributed to the explosion.  Raise this limit or
              tighten the offending constraints to resolve the error.

            Default: `1024`.
        max_cells: Maximum number of cells collected when expanding a range
            reference.  Default: `10_000`.
        max_depth: Maximum AST-evaluation recursion depth.  Default: `10`.
    """

    max_branches: int = 1024
    max_cells: int = 10_000
    max_depth: int = 10


@dataclass(frozen=True)
class DynamicRefConfig:
    """Configuration for resolving OFFSET/INDIRECT via constraint-based inference.

    Prefer building via `from_constraints`; the constructor is for internal use.
    """

    cell_type_env: CellTypeEnv
    limits: DynamicRefLimits

    @classmethod
    def from_constraints(
        cls,
        constraints_schema: Mapping[str, Any],
        constraints_data: Mapping[str, Any],
        *,
        limits: DynamicRefLimits | None = None,
    ) -> DynamicRefConfig:
        r"""Build a config from a constraints schema (address keys -> type annotations) and instance data.

        *constraints_schema* must be a mapping (typically `dict[str, type]`) whose keys are
        sheet-qualified addresses (e.g. `\"Sheet1!B1\"`). Values are typing objects describing
        domains (`Annotated`, `Literal`, etc.). *constraints_data* may mirror keys for runtime
        validation elsewhere; it is not validated here.

        Raises:
            TypeError: If *constraints_schema* is not a mapping (e.g. a legacy `TypedDict` class passed instead of a dict).
        """
        if isinstance(constraints_schema, type):
            raise TypeError(
                "constraints_schema must be a dict[str, type] mapping addresses to annotations, "
                "not a TypedDict/class object."
            )
        if not isinstance(constraints_schema, Mapping):
            raise TypeError(
                f"constraints_schema must be a mapping, got {type(constraints_schema).__name__!r}"
            )
        env = constraints_to_cell_type_env(constraints_schema, constraints_data)
        return cls(cell_type_env=env, limits=limits or DynamicRefLimits())

    @classmethod
    def from_constraints_and_workbook(
        cls,
        constraints_schema: Mapping[str, Any],
        workbook_path: str | Path,
        *,
        limits: DynamicRefLimits | None = None,
        data_only: bool = True,
    ) -> DynamicRefConfig:
        """Build config from constraints schema plus workbook values for constant cells.

        Constraints whose annotations carry a `FromWorkbook` marker are
        treated as singleton domains derived from the current cached value in the
        workbook.  Other constraints are interpreted via
        `constraints_to_cell_type_env` as usual.

        **Performance note:** The workbook is opened once in `read_only` mode
        and values are read via fastpyxl's streaming parser.  `FromWorkbook`
        addresses are sorted by (sheet, row, column) before iteration so that
        each sheet's XML is parsed in a single forward pass.  For large sheets
        whose constrained cells sit far down (e.g. row 900+), the initial parse
        to that row can take tens of seconds; subsequent sequential reads on the
        same sheet are fast.  This is a deliberate tradeoff: `FromWorkbook`
        eliminates the maintenance burden of hardcoded `Literal` values at
        the cost of a longer config-build step.
        """
        if isinstance(constraints_schema, type):
            raise TypeError(
                "constraints_schema must be a dict[str, type] mapping addresses to annotations, "
                "not a TypedDict/class object."
            )
        if not isinstance(constraints_schema, Mapping):
            raise TypeError(
                f"constraints_schema must be a mapping, got {type(constraints_schema).__name__!r}"
            )

        hints = dict(constraints_schema)
        dummy_data: dict[str, Any] = {k: None for k in hints}
        env = constraints_to_cell_type_env(hints, dummy_data)

        from_wb_items: list[tuple[str, str, str]] = []
        for addr, annotated_type in hints.items():
            if not _has_from_workbook_marker(annotated_type):
                continue
            sheet_name, coord = _split_addr_sheet_coord(addr)
            from_wb_items.append((addr, sheet_name, coord))

        from_wb_items.sort(key=lambda item: (item[1], coordinate_to_tuple(item[2])))

        wb = fastpyxl.load_workbook(Path(workbook_path), data_only=data_only, read_only=True)
        try:
            for addr, sheet_name, coord in from_wb_items:
                if sheet_name not in wb.sheetnames:
                    raise DynamicRefError(
                        f"Sheet {sheet_name!r} (from constraint {addr!r}) not found in workbook"
                    )
                ws = wb[sheet_name]
                value = ws[coord].value
                if value is None:
                    continue
                kind = _infer_kind_from_value(value)
                norm = normalize_cell_type_env_key(addr)
                env[norm] = CellType(
                    kind=kind,
                    enum=EnumDomain(values=frozenset({value})),
                )
        finally:
            wb.close()

        return cls(cell_type_env=env, limits=limits or DynamicRefLimits())


@dataclass(frozen=True)
class FromWorkbook:
    """Metadata marker: resolve domain from the current cached workbook value.

    Use `Annotated[T, FromWorkbook()]` in the constraints schema to derive a
    singleton domain at config-build time instead of hardcoding a `Literal`.
    This eliminates maintenance when the workbook template changes, at the cost
    of a slower `DynamicRefConfig.from_constraints_and_workbook` call (the
    workbook must be opened and each marked cell read via a streaming parser).
    """


def _has_from_workbook_marker(annotated_type: Any) -> bool:
    try:
        from typing import Annotated
    except ImportError:  # pragma: no cover - Annotated always available in supported versions
        return False

    if get_origin(annotated_type) is not Annotated:
        return False
    args = get_args(annotated_type)
    if len(args) < 2:
        return False
    metadata = args[1:]
    return any(isinstance(m, FromWorkbook) for m in metadata)


def _split_addr_sheet_coord(addr: str) -> tuple[str, str]:
    """Split an address-style key into (sheet_name, coord) and normalize quoting."""
    try:
        return parse_address(addr)
    except ValueError as exc:
        raise DynamicRefError(str(exc)) from exc


def _infer_kind_from_value(value: Any) -> CellKind:
    if isinstance(value, (int, float)):
        return CellKind.NUMBER
    if isinstance(value, bool):
        return CellKind.BOOL
    if isinstance(value, str):
        return CellKind.STRING
    return CellKind.ANY


def _expand_sheet_qualified_range(sheet_name: str, range_a1: str) -> list[str]:
    """Expand a single-cell or A1 range into sheet-qualified keys."""
    range_a1 = range_a1.strip()
    if ":" in range_a1:
        start_a1, end_a1 = range_a1.split(":", 1)
        start_a1 = _strip_optional_sheet_prefix(start_a1.strip(), sheet_name)
        end_a1 = _strip_optional_sheet_prefix(end_a1.strip(), sheet_name)
    else:
        start_a1 = end_a1 = _strip_optional_sheet_prefix(range_a1, sheet_name)

    start_col, start_row = coordinate_from_string(start_a1)
    end_col, end_row = coordinate_from_string(end_a1)
    cells = expand_range(
        sheet=sheet_name,
        start_col=start_col,
        start_row=start_row,
        end_col=end_col,
        end_row=end_row,
        max_cells=10_000_000,
    )
    return [format_key(sheet, a1) for sheet, a1 in cells]


def _strip_optional_sheet_prefix(part: str, expected_sheet: str) -> str:
    """Accept `A1` and `Sheet!A1` forms for range endpoints."""
    if "!" not in part:
        return part
    sheet_name, coord = _split_addr_sheet_coord(part)
    if sheet_name != expected_sheet:
        raise DynamicRefError(f"Range endpoint {part!r} must use sheet {expected_sheet!r}")
    return coord


@dataclass(frozen=True)
class GlobalWorkbookBounds(WorkbookBoundsProtocol):
    """Simple bounds implementation using Excel's global sheet limits."""

    sheet: str
    min_row: int = 1
    max_row: int = 1_048_576  # Excel row limit
    min_col: int = 1
    max_col: int = 16_384  # Excel column limit


def _bounds_for_sheet(
    bounds: WorkbookBoundsProtocol | None,
    *,
    sheet: str,
) -> WorkbookBoundsProtocol:
    if bounds is None:
        return GlobalWorkbookBounds(sheet=sheet)
    return GlobalWorkbookBounds(
        sheet=sheet,
        min_row=bounds.min_row,
        max_row=bounds.max_row,
        min_col=bounds.min_col,
        max_col=bounds.max_col,
    )


def _sheet_from_addr(addr: str) -> str:
    """Return sheet part of address (e.g. 'Sheet1!A1' -> 'Sheet1')."""
    if "!" not in addr:
        return ""
    return parse_address(addr)[0]


def expand_leaf_env_to_argument_env(
    argument_refs: set[str],
    get_cell_formula: Callable[[str], str | None],
    get_refs_from_formula: Callable[[str, str], set[str]],
    leaf_env: CellTypeEnv,
    limits: DynamicRefLimits,
    named_ranges: Mapping[str, tuple[str, str]] | None = None,
    named_range_ranges: Mapping[str, tuple[str, str, str]] | None = None,
    *,
    max_range_cells: int = 5000,
    shared_cell_type_cache: dict[str, CellType] | None = None,
    type_analysis_cache: TypeAnalysisCache | None = None,
    workbook_sha256: str | None = None,
) -> dict[str, CellType]:
    """Build a CellTypeEnv for all refs in the argument chain from leaf constraints only.

    The cell env targets leaves: only leaf (non-formula) addresses need to be in
    leaf_env. Intermediate (formula) cells are inferred by evaluating their formulas
    over their dependencies' domains; they do not need to be constrained. If an
    intermediate matches a leaf_env entry under `excel_grapher.core.cell_types.normalize_cell_type_env_key`,
    that type is used and we do not traverse that branch.
    When an intermediate cannot be inferred (e.g. its formula is OFFSET/INDIRECT and
    refs are empty after masking), it is assigned CellType(ANY); enumeration may then
    require a constraint for that cell.

    `max_range_cells` must match the graph builder's range expansion limit so static
    ranges collected from the AST align with
    `excel_grapher.grapher.builder.create_dependency_graph` argument-subgraph BFS.

    When `shared_cell_type_cache` is provided, intermediate cell type inferences
    are persisted across multiple calls.  This avoids redundant work when many
    BFS nodes share intermediate formula cells in their argument subgraphs.
    Addresses already present in that cache are not re-walked: a later call
    whose `argument_refs` are a subset of a previous expansion returns without
    re-entering `cell_type_for` (issue #528).

    When `type_analysis_cache` and `workbook_sha256` are provided, successful
    intermediate formula-cell type results are persisted to SQLite across runs.
    """
    cache: dict[str, CellType] = (
        shared_cell_type_cache if shared_cell_type_cache is not None else {}
    )
    in_progress: set[str] = set()
    nr = named_ranges or {}
    nrr = named_range_ranges or {}

    # Persistent cache support
    from excel_grapher.grapher.type_analysis_cache import (
        _compute_leaf_env_subset_fingerprint,
        _compute_limits_fingerprint,
    )

    _tac = (
        type_analysis_cache
        if (type_analysis_cache is not None and workbook_sha256 is not None)
        else None
    )
    _limits_fp = _compute_limits_fingerprint(limits) if _tac else ""
    # Track which leaf_env keys each formula cell consumes during analysis.
    # Only `_persist_result` ever reads this, so skip the bookkeeping entirely
    # when there is no persistent cache to write to: maintaining it costs
    # O(stack_depth x consumed_leaves) per cache hit (issue #463).
    _track_consumed = _tac is not None
    _consumed_leaves: dict[str, set[str]] = {}
    # Addresses loaded from persistent cache (skip re-persisting in finally block)
    _loaded_from_persistent: set[str] = set()
    # Stack of formula cells being analysed (for recording consumed leaves)
    _analysis_stack: list[str] = []
    # Progress accounting for `expand-env-progress` traces.  `_calls` counts
    # `cell_type_for` entries; `_bulk_hits` counts refs served straight from
    # `cache` without re-entering it (issue #465).  Both are progress, so both
    # drive the trace interval.
    t0 = time.perf_counter()
    _calls = 0
    _bulk_hits = 0
    _next_progress = _EXPAND_PROGRESS_INTERVAL

    def _note_progress(last: str) -> None:
        """Emit an `expand-env-progress` trace once the counters cross the interval."""
        nonlocal _next_progress
        if _calls + _bulk_hits < _next_progress:
            return
        _next_progress = _calls + _bulk_hits + _EXPAND_PROGRESS_INTERVAL
        _emit_trace(
            DynamicRefTraceEvent(
                kind="expand-env-progress",
                name="expand_leaf_env_to_argument_env",
                elapsed_s=time.perf_counter() - t0,
                detail={
                    "calls": _calls,
                    "bulk_ref_hits": _bulk_hits,
                    "cache_size": len(cache),
                    "stack_depth": len(_analysis_stack),
                    "last": last,
                },
            )
        )

    def _formula_to_parse(raw: str) -> str:
        body = raw[1:] if raw.startswith("=") else raw
        qualified = _qualify_fragment(body, nr, nrr)
        return "=" + qualified

    def _values_to_cell_type(values: set[Any]) -> CellType:
        if not values:
            return CellType(kind=CellKind.ANY)
        kinds = {type(v) for v in values}
        if kinds <= {int, float}:
            return CellType(
                kind=CellKind.NUMBER,
                enum=EnumDomain(values=frozenset(values)),
            )
        if kinds <= {str}:
            return CellType(
                kind=CellKind.STRING,
                enum=EnumDomain(values=frozenset(values)),
            )
        if kinds <= {bool}:
            return CellType(
                kind=CellKind.BOOL,
                enum=EnumDomain(values=frozenset(values)),
            )
        return CellType(kind=CellKind.ANY, enum=EnumDomain(values=frozenset(values)))

    def _record_consumed_leaf(leaf_addr: str) -> None:
        """Record that every formula cell on the analysis stack consumed a leaf.

        A leaf referenced by a child formula transitively affects every
        ancestor on the stack, so all of them must include that leaf in
        their fingerprint for correct cache invalidation.

        No-op without a persistent cache, which is the only consumer.
        """
        if not _track_consumed:
            return
        for ancestor in _analysis_stack:
            _consumed_leaves.setdefault(ancestor, set()).add(leaf_addr)

    def _try_persistent_lookup(addr: str, formula: str) -> tuple[CellType, list[str]] | None:
        """Try to load a cached type from the persistent SQLite cache.

        Returns `(cell_type, consumed_leaf_keys)` on hit, or `None`.
        """
        if _tac is None or workbook_sha256 is None:
            return None
        norm_formula_sha = hashlib.sha256(formula.encode()).hexdigest()
        return _tac.get_formula_cell_type(
            workbook_sha256=workbook_sha256,
            address=addr,
            normalized_formula_sha256=norm_formula_sha,
            limits_fingerprint=_limits_fp,
            current_leaf_env=leaf_env,
        )

    def _persist_result(addr: str, formula: str, ct: CellType) -> None:
        """Write a successful formula-cell type to the persistent cache."""
        if _tac is None or workbook_sha256 is None:
            return
        # Don't cache ANY results in v1
        if ct.kind is CellKind.ANY and ct.enum is None:
            return
        norm_formula_sha = hashlib.sha256(formula.encode()).hexdigest()
        consumed = sorted(_consumed_leaves.get(addr, set()))
        fp = _compute_leaf_env_subset_fingerprint(consumed, leaf_env)
        _tac.put_formula_cell_type(
            workbook_sha256=workbook_sha256,
            address=addr,
            normalized_formula_sha256=norm_formula_sha,
            limits_fingerprint=_limits_fp,
            leaf_env_subset_fingerprint=fp,
            consumed_leaf_keys=consumed,
            cell_type=ct,
        )

    def _propagate_consumed_leaves_to_ancestors(addr: str) -> None:
        """Propagate a cell's consumed leaves to all ancestors on the stack.

        Called when `addr` is served from the in-memory cache so that
        ancestors still record its transitive leaf dependencies even though
        `_record_consumed_leaf` won't fire for the cached cell's refs.

        No-op without a persistent cache, which is the only consumer.
        """
        if not _track_consumed:
            return
        child_leaves = _consumed_leaves.get(addr)
        if child_leaves:
            for ancestor in _analysis_stack:
                _consumed_leaves.setdefault(ancestor, set()).update(child_leaves)

    def _propagate_consumed_leaves_bulk(addrs: list[str]) -> None:
        """Propagate the consumed leaves of many cached cells in one pass.

        Equivalent to calling `_propagate_consumed_leaves_to_ancestors` for each
        address, but merges the children first so the stack is walked once
        instead of once per address (issue #465).
        """
        if not _track_consumed or not _analysis_stack:
            return
        merged: set[str] = set()
        for addr in addrs:
            child_leaves = _consumed_leaves.get(addr)
            if child_leaves:
                merged |= child_leaves
        if not merged:
            return
        for ancestor in _analysis_stack:
            _consumed_leaves.setdefault(ancestor, set()).update(merged)

    def _resolve_ref_types(refs: Iterable[str]) -> dict[str, CellType]:
        """Resolve the type of every ref, serving already-inferred ones in bulk.

        A long formula chain that repeatedly mentions a large static range would
        otherwise re-enter `cell_type_for` for every cell of that range at every
        level of the chain -- `O(depth x range_cells)` entries for work that is a
        dict lookup after the first level (issue #465).  Cached refs are read
        straight out of `cache` here; only their consumed-leaf propagation still
        has to happen, and that is batched.
        """
        nonlocal _bulk_hits
        ref_types: dict[str, CellType] = {}
        bulk_served: list[str] = []
        for r in refs:
            cached_ct = cache.get(r)
            if cached_ct is not None:
                ref_types[r] = cached_ct
                bulk_served.append(r)
                continue
            ref_types[r] = cell_type_for(r)
        if bulk_served:
            _bulk_hits += len(bulk_served)
            _note_progress(bulk_served[-1])
            _propagate_consumed_leaves_bulk(bulk_served)
        return ref_types

    def cell_type_for(addr: str) -> CellType:
        nonlocal _calls
        _calls += 1
        _note_progress(addr)
        if addr in cache:
            _propagate_consumed_leaves_to_ancestors(addr)
            return cache[addr]
        if addr in in_progress:
            return CellType(kind=CellKind.ANY)
        ct_resolved = _lookup_cell_type(leaf_env, addr)
        if ct_resolved is not None:
            cache[addr] = ct_resolved
            if _track_consumed:
                norm_key = normalize_cell_type_env_key(addr)
                # Record the leaf as its own consumed-leaf so that cache hits
                # for this address can propagate it to future ancestors.
                _consumed_leaves.setdefault(addr, set()).add(norm_key)
                _record_consumed_leaf(norm_key)
            return cache[addr]
        in_progress.add(addr)
        formula = get_cell_formula(addr)
        try:
            if formula is None:
                raise DynamicRefError(
                    f"Missing constraint for leaf {addr!r} that feeds OFFSET/INDIRECT. "
                    "Add constraints only for leaf cells (non-formula) in the argument subgraph."
                )
            # Check persistent cache before expensive analysis
            _persistent_result = _try_persistent_lookup(addr, formula)
            if _persistent_result is not None:
                cached_ct, cached_consumed = _persistent_result
                cache[addr] = cached_ct
                _consumed_leaves[addr] = set(cached_consumed)
                _loaded_from_persistent.add(addr)
                _propagate_consumed_leaves_to_ancestors(addr)
                return cache[addr]
            if len(_analysis_stack) >= _MAX_ANALYSIS_DEPTH:
                raise DynamicRefError(
                    f"Argument-subgraph analysis exceeded the maximum depth of "
                    f"{_MAX_ANALYSIS_DEPTH} while inferring the type of {addr!r}. "
                    "The dynamic-ref argument chain is too deeply nested to analyse; "
                    "constrain an intermediate cell in the chain to cut it short, or "
                    "simplify the chain."
                )
            _analysis_stack.append(addr)
            refs = get_refs_from_formula(formula, _sheet_from_addr(addr))
            try:
                formula_parse = _formula_to_parse(formula)
                ast_root = parse_ast(formula_parse)
            except FormulaParseError:
                ast_root = None
            if ast_root is not None:
                refs |= _collect_static_addresses_from_ast(
                    ast_root, max_range_cells=max_range_cells
                )
            if not refs:
                if ast_root is None:
                    cache[addr] = CellType(kind=CellKind.ANY)
                    return cache[addr]
                try:
                    val = evaluate_expr(
                        ast_root,
                        get_cell_value=lambda _: None,
                        max_depth=limits.max_depth,
                    )
                except Exception:
                    val = None
                if isinstance(val, Unsupported):
                    cache[addr] = CellType(kind=CellKind.ANY)
                    return cache[addr]
                if val is None or isinstance(val, XlError):
                    cache[addr] = CellType(kind=CellKind.ANY)
                    return cache[addr]
                cache[addr] = _values_to_cell_type({val})
                return cache[addr]
            ref_types = _resolve_ref_types(sorted(refs))
            if ast_root is not None:
                infer_result = _infer_numeric_domain_result(
                    ast_root,
                    ref_types,
                    limits,
                    context=None,
                    current_sheet=_sheet_from_addr(addr),
                )
                if infer_result.diagnostic is not None:
                    detail = infer_result.diagnostic
                    refs_text = ", ".join(sorted(detail.refs))
                    refs_clause = f" Constrain one or more of: {refs_text}." if refs_text else ""
                    expr_clause = (
                        f" Divisor expression: {detail.expression}."
                        if detail.expression is not None
                        else ""
                    )
                    raise DynamicRefError(
                        f"Formula cell {addr!r} is not covered by numeric abstract analysis: "
                        f"{detail.reason}.{expr_clause}{refs_clause}"
                    )
                inferred = infer_result.domain
                if inferred is not None:
                    if isinstance(inferred, _FiniteInts):
                        cache[addr] = CellType(
                            kind=CellKind.NUMBER,
                            enum=EnumDomain(values=inferred.values),
                        )
                    else:
                        span = inferred.hi - inferred.lo + 1
                        if span <= limits.max_branches:
                            cache[addr] = CellType(
                                kind=CellKind.NUMBER,
                                enum=EnumDomain(
                                    values=frozenset(range(inferred.lo, inferred.hi + 1))
                                ),
                            )
                        else:
                            cache[addr] = CellType(
                                kind=CellKind.NUMBER,
                                interval=IntervalDomain(min=inferred.lo, max=inferred.hi),
                            )
                    return cache[addr]
            unsupported = _describe_unsupported_numeric_construct(ast_root)
            domains: dict[str, list[Any]] = {}
            for r, ct in ref_types.items():
                if ct.enum is not None:
                    domains[r] = list(ct.enum.values)
                elif ct.interval is not None:
                    try:
                        domains[r] = _interval_to_values(ct.interval, limits)
                    except DynamicRefError as exc:
                        detail = (
                            f" First unsupported construct: {unsupported}."
                            if unsupported is not None
                            else ""
                        )
                        raise DynamicRefError(
                            f"{exc} (while expanding types for formula cell {addr!r}, dependency {r!r}; "
                            f"this formula is not covered by numeric abstract analysis.{detail} "
                            f"constrain {r!r} more tightly, simplify the formula, or extend analysis "
                            f"for the unsupported construct)"
                        ) from exc
                elif ct.real_interval is not None:
                    cache[addr] = CellType(kind=CellKind.ANY)
                    return cache[addr]
                else:
                    cache[addr] = CellType(kind=CellKind.ANY)
                    return cache[addr]
            total_branches = math.prod(len(v) for v in domains.values())
            if total_branches > limits.max_branches:
                dep_sizes = ", ".join(f"{r!r}: {len(domains[r])}" for r in sorted(domains))
                unsupported_hint = (
                    f" First unsupported construct: {unsupported}."
                    if unsupported is not None
                    else ""
                )
                raise DynamicRefError(
                    f"Formula cell {addr!r} fallback enumeration would require "
                    f"{total_branches} branches (limit {limits.max_branches}). "
                    f"Dependency domain sizes: {dep_sizes}.{unsupported_hint} "
                    f"Tighten constraints on one or more dependencies, simplify "
                    f"the formula, or extend numeric abstract analysis to cover it."
                )
            result_values: set[Any] = set()
            last_unsupported: Unsupported | None = None
            for assignment in product(*(domains[r] for r in refs)):
                addr_to_val = dict(zip(refs, assignment, strict=False))

                def get_cell_value(a: str, _av=addr_to_val) -> Any:
                    return _av.get(a)

                try:
                    formula_parse = _formula_to_parse(formula)
                    ast = parse_ast(formula_parse)
                except FormulaParseError:
                    cache[addr] = CellType(kind=CellKind.ANY)
                    return cache[addr]
                val = evaluate_expr(
                    ast,
                    get_cell_value=get_cell_value,
                    max_depth=limits.max_depth,
                )
                if isinstance(val, Unsupported):
                    last_unsupported = val
                    continue
                if isinstance(val, XlError):
                    continue
                result_values.add(val)
            if not result_values:
                if last_unsupported is not None:
                    cache[addr] = CellType(kind=CellKind.ANY)
                    return cache[addr]
                cache[addr] = CellType(kind=CellKind.ANY)
                return cache[addr]
            cache[addr] = _values_to_cell_type(result_values)
            return cache[addr]
        finally:
            in_progress.discard(addr)
            if _analysis_stack and _analysis_stack[-1] == addr:
                _analysis_stack.pop()
                # Propagate this cell's consumed leaves to the parent on the
                # stack.  This covers the case where a child was served from
                # the in-memory cache (so _record_consumed_leaf was never
                # called for its leaves during *this* traversal).
                if _track_consumed and _analysis_stack:
                    parent = _analysis_stack[-1]
                    child_leaves = _consumed_leaves.get(addr)
                    if child_leaves:
                        _consumed_leaves.setdefault(parent, set()).update(child_leaves)
            # Persist successful result to SQLite cache (skip if loaded from persistent cache)
            if addr in cache and formula is not None and addr not in _loaded_from_persistent:
                _persist_result(addr, formula, cache[addr])

    skipped_cached_refs = 0
    try:
        pending = [addr for addr in argument_refs if addr not in cache]
        skipped_cached_refs = len(argument_refs) - len(pending)
        for addr in sorted(pending):
            cell_type_for(addr)
    except Exception as exc:
        _emit_trace(
            DynamicRefTraceEvent(
                kind="expand-env-error",
                name="expand_leaf_env_to_argument_env",
                elapsed_s=time.perf_counter() - t0,
                detail={
                    "argument_refs": len(argument_refs),
                    "error": f"{type(exc).__name__}: {exc}",
                },
            )
        )
        raise
    finally:
        if _tac is not None:
            _tac.flush()
    _emit_trace(
        DynamicRefTraceEvent(
            kind="expand-env",
            name="expand_leaf_env_to_argument_env",
            elapsed_s=time.perf_counter() - t0,
            detail={
                "argument_refs": len(argument_refs),
                "inferred_cells": len(cache),
                "calls": _calls,
                "bulk_ref_hits": _bulk_hits,
                "skipped_cached_refs": skipped_cached_refs,
                "consumed_leaf_tracking": _track_consumed,
            },
        )
    )
    return cache


def infer_dynamic_offset_targets(
    formula: str,
    *,
    current_sheet: str,
    cell_type_env: CellTypeEnv,
    limits: DynamicRefLimits | None = None,
    bounds: WorkbookBoundsProtocol | None = None,
    named_ranges: Mapping[str, tuple[str, str]] | None = None,
    named_range_ranges: Mapping[str, tuple[str, str, str]] | None = None,
    current_row: int | None = None,
    current_col: int | None = None,
) -> set[str]:
    """Infer the union of all possible OFFSET targets for a formula.

    This helper is intentionally focused and conservative:
    - Only OFFSET calls are analysed (INDIRECT is currently ignored).
    - Arguments may use a small Excel expression subset supported by
      `core.expr_eval.evaluate_expr`.
    - Leaf cells referenced by OFFSET/INDEX arguments must have a numeric
      domain in `cell_type_env` unless they appear only in ref_only
      argument positions (see `excel_grapher.core.excel_function_meta`).
    - Integer interval domains must be finite and small enough to enumerate.
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return set()

    t0 = time.perf_counter()
    lim = limits or DynamicRefLimits()
    out: set[str] = set()

    calls = _find_function_calls_with_spans(formula, frozenset({"OFFSET"}))
    for fn, inner, _span in calls:
        if fn != "OFFSET":
            continue
        targets = _infer_single_offset_call(
            inner,
            current_sheet=current_sheet,
            cell_type_env=cell_type_env,
            limits=lim,
            bounds=bounds,
            named_ranges=named_ranges,
            named_range_ranges=named_range_ranges,
            current_row=current_row,
            current_col=current_col,
        )
        out |= targets
        if len(out) > lim.max_cells:
            raise DynamicRefError(f"Dynamic ref cells exceed limit ({len(out)} > {lim.max_cells})")

    _emit_trace(
        DynamicRefTraceEvent(
            kind="infer",
            name="infer_dynamic_offset_targets",
            elapsed_s=time.perf_counter() - t0,
            detail={"targets": len(out), "formula": formula, "current_sheet": current_sheet},
        )
    )
    return out


def infer_dynamic_index_targets(
    formula: str,
    *,
    current_sheet: str,
    cell_type_env: CellTypeEnv,
    limits: DynamicRefLimits | None = None,
    bounds: WorkbookBoundsProtocol | None = None,
    named_ranges: Mapping[str, tuple[str, str]] | None = None,
    named_range_ranges: Mapping[str, tuple[str, str, str]] | None = None,
    current_row: int | None = None,
    current_col: int | None = None,
) -> set[str]:
    """Infer the union of all possible standalone INDEX targets for a formula.

    INDEX calls that appear as the first argument of OFFSET are skipped - those
    are already handled by `infer_dynamic_offset_targets`.
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return set()

    t0 = time.perf_counter()
    lim = limits or DynamicRefLimits()
    out: set[str] = set()

    # Find all INDEX and OFFSET calls so we can exclude INDEX-inside-OFFSET.
    index_calls = _find_function_calls_with_spans(formula, frozenset({"INDEX"}))
    offset_calls = _find_function_calls_with_spans(formula, frozenset({"OFFSET"}))

    nested_index_spans: set[tuple[int, int]] = set()
    offset_spans = [span for _fn, _inner, span in offset_calls]
    for fn, _inner, idx_span in index_calls:
        if fn != "INDEX":
            continue
        # Check if this INDEX is inside an OFFSET call
        is_nested = False
        for o_start, o_end in offset_spans:
            if idx_span[0] > o_start and idx_span[1] <= o_end:
                is_nested = True
                break
        if is_nested:
            nested_index_spans.add(idx_span)

    for fn, inner, span in index_calls:
        if fn != "INDEX":
            continue
        if span in nested_index_spans:
            continue
        targets = _infer_single_index_call(
            inner,
            current_sheet=current_sheet,
            cell_type_env=cell_type_env,
            limits=lim,
            named_ranges=named_ranges,
            named_range_ranges=named_range_ranges,
            current_row=current_row,
            current_col=current_col,
        )
        out |= targets
        if len(out) > lim.max_cells:
            raise DynamicRefError(f"Dynamic ref cells exceed limit ({len(out)} > {lim.max_cells})")

    _emit_trace(
        DynamicRefTraceEvent(
            kind="infer",
            name="infer_dynamic_index_targets",
            elapsed_s=time.perf_counter() - t0,
            detail={"targets": len(out), "formula": formula, "current_sheet": current_sheet},
        )
    )
    return out


def _infer_single_index_call(
    inner_args: str,
    *,
    current_sheet: str,
    cell_type_env: CellTypeEnv,
    limits: DynamicRefLimits,
    named_ranges: Mapping[str, tuple[str, str]] | None = None,
    named_range_ranges: Mapping[str, tuple[str, str, str]] | None = None,
    current_row: int | None = None,
    current_col: int | None = None,
) -> set[str]:
    """Infer targets for a single INDEX(...) call body."""
    args = _split_top_level_args(inner_args)
    if args is None or len(args) < 2 or len(args) > 3:
        raise DynamicRefError("INDEX expects 2 or 3 arguments (array, row_num, [column_num])")

    named_ranges = named_ranges or {}
    named_range_ranges = named_range_ranges or {}

    array_expr = _qualify_fragment(args[0], named_ranges, named_range_ranges)
    row_expr = _qualify_fragment(args[1], named_ranges, named_range_ranges)
    col_expr = (
        _qualify_fragment(args[2], named_ranges, named_range_ranges) if len(args) >= 3 else ""
    )

    try:
        array_ast = parse_ast("=" + array_expr)
        array_range = _base_to_range(array_ast, current_sheet=current_sheet)
    except (DynamicRefError, FormulaParseError) as exc:
        raise DynamicRefError(f"INDEX array argument must be a static range: {exc}") from exc

    row_ast = parse_ast("=" + row_expr)
    col_ast = parse_ast("=" + col_expr) if col_expr else None

    eval_context = (
        {"row": current_row, "column": current_col}
        if current_row is not None and current_col is not None
        else None
    )

    return _infer_index_targets_core(
        array_range,
        row_ast,
        col_ast,
        cell_type_env,
        limits,
        eval_context,
        current_sheet=current_sheet,
    )


@lru_cache(maxsize=4096)
def _defined_name_token_pattern(name: str) -> re.Pattern[str]:
    """Compiled word-bounded defined-name token that is not a sheet qualifier."""
    return re.compile(rf"\b{re.escape(name)}\b(?!\s*!)")


_QualifyReplacementPairs = tuple[tuple[str, str], ...]
_qualify_pairs_cache: dict[tuple[int, int], _QualifyReplacementPairs] = {}
_QUALIFY_PAIRS_CACHE_MAX = 64


def _qualify_replacement_pairs(
    named_ranges: Mapping[str, tuple[str, str]],
    named_range_ranges: Mapping[str, tuple[str, str, str]] | None,
) -> _QualifyReplacementPairs:
    """Return `(name, replacement)` pairs, longest name first."""
    all_names = sorted(
        set(named_ranges.keys())
        | (set(named_range_ranges.keys()) if named_range_ranges else set()),
        key=lambda n: (-len(n), n),
    )
    pairs: list[tuple[str, str]] = []
    for name in all_names:
        if name in named_ranges:
            sheet, addr = named_ranges[name]
            replacement = format_key(sheet, addr)
        elif named_range_ranges and name in named_range_ranges:
            sheet, start_a1, end_a1 = named_range_ranges[name]
            replacement = f"{format_key(sheet, start_a1)}:{format_key(sheet, end_a1)}"
        else:
            continue
        pairs.append((name, replacement))
    return tuple(pairs)


def _cached_qualify_replacement_pairs(
    named_ranges: Mapping[str, tuple[str, str]],
    named_range_ranges: Mapping[str, tuple[str, str, str]] | None,
) -> _QualifyReplacementPairs:
    """Reuse replacement pairs for a given pair of named-range mapping objects."""
    key = (
        id(named_ranges),
        id(named_range_ranges) if named_range_ranges is not None else 0,
    )
    cached = _qualify_pairs_cache.get(key)
    if cached is not None:
        return cached
    pairs = _qualify_replacement_pairs(named_ranges, named_range_ranges)
    if len(_qualify_pairs_cache) >= _QUALIFY_PAIRS_CACHE_MAX:
        _qualify_pairs_cache.clear()
    _qualify_pairs_cache[key] = pairs
    return pairs


def _qualify_fragment(
    expr: str,
    named_ranges: Mapping[str, tuple[str, str]],
    named_range_ranges: Mapping[str, tuple[str, str, str]] | None,
) -> str:
    """Replace named range tokens in a formula fragment with sheet-qualified refs.

    Patterns are compiled once per defined name and reused. Names whose literal
    text is absent from the fragment are skipped. Longer names are applied first
    so `Sales.Total` is not partially matched by `Sales`.
    """
    if not expr.strip():
        return expr
    pairs = _cached_qualify_replacement_pairs(named_ranges, named_range_ranges)
    if not pairs:
        return expr
    result = expr
    for name, replacement in pairs:
        if name not in result:
            continue
        result = _defined_name_token_pattern(name).sub(replacement, result)
    return result


def _infer_single_offset_call(
    inner_args: str,
    *,
    current_sheet: str,
    cell_type_env: CellTypeEnv,
    limits: DynamicRefLimits,
    bounds: WorkbookBoundsProtocol | None,
    named_ranges: Mapping[str, tuple[str, str]] | None = None,
    named_range_ranges: Mapping[str, tuple[str, str, str]] | None = None,
    current_row: int | None = None,
    current_col: int | None = None,
) -> set[str]:
    """Infer targets for a single OFFSET(...) call body."""
    args = _split_top_level_args(inner_args)
    if args is None or len(args) < 3 or len(args) > 5:
        raise DynamicRefError("OFFSET expects 3 to 5 arguments")

    named_ranges = named_ranges or {}
    named_range_ranges = named_range_ranges or {}

    base_expr = _qualify_fragment(args[0], named_ranges, named_range_ranges)
    rows_expr = _qualify_fragment(args[1], named_ranges, named_range_ranges)
    cols_expr = _qualify_fragment(args[2], named_ranges, named_range_ranges)
    height_expr = (
        _qualify_fragment(args[3], named_ranges, named_range_ranges)
        if len(args) >= 4 and args[3]
        else ""
    )
    width_expr = (
        _qualify_fragment(args[4], named_ranges, named_range_ranges)
        if len(args) >= 5 and args[4]
        else ""
    )

    try:
        base_ast = parse_ast("=" + base_expr)
        base_ranges = _resolve_offset_base(
            base_ast,
            current_sheet=current_sheet,
            cell_type_env=cell_type_env,
            limits=limits,
            current_row=current_row,
            current_col=current_col,
        )
    except DynamicRefError as exc:
        raise DynamicRefError(f"{exc} (OFFSET base expression {base_expr!r})") from exc

    rows_ast = parse_ast("=" + rows_expr)
    cols_ast = parse_ast("=" + cols_expr)
    height_ast = parse_ast("=" + height_expr) if height_expr else None
    width_ast = parse_ast("=" + width_expr) if width_expr else None

    eval_context = (
        {"row": current_row, "column": current_col}
        if current_row is not None and current_col is not None
        else None
    )

    rows_list = _infer_offset_scalar_domains(
        rows_ast, cell_type_env, limits, eval_context, current_sheet=current_sheet
    )
    cols_list = _infer_offset_scalar_domains(
        cols_ast, cell_type_env, limits, eval_context, current_sheet=current_sheet
    )
    use_infer = rows_list is not None and cols_list is not None
    height_vals: list[int | None]
    width_vals: list[int | None]
    if use_infer:
        if height_ast is None:
            height_vals = [None]
        else:
            hv = _infer_offset_scalar_domains(
                height_ast, cell_type_env, limits, eval_context, current_sheet=current_sheet
            )
            if hv is None:
                use_infer = False
            else:
                height_vals = cast(list[int | None], hv)
    if use_infer:
        if width_ast is None:
            width_vals = [None]
        else:
            wv = _infer_offset_scalar_domains(
                width_ast, cell_type_env, limits, eval_context, current_sheet=current_sheet
            )
            if wv is None:
                use_infer = False
            else:
                width_vals = cast(list[int | None], wv)

    targets: set[str] = set()
    if use_infer:
        assert rows_list is not None
        assert cols_list is not None
        for base_range in base_ranges:
            base_bounds = _bounds_for_sheet(bounds, sheet=base_range.sheet)
            for rv in rows_list:
                for cv in cols_list:
                    for hv in height_vals:
                        for wv in width_vals:
                            result = offset_range(
                                base_range,
                                rows=rv,
                                cols=cv,
                                height=hv,
                                width=wv,
                                bounds=base_bounds,
                            )
                            if isinstance(result, ExcelRange):
                                targets |= set(result.cell_addresses())
                                if len(targets) > limits.max_cells:
                                    raise DynamicRefError(
                                        f"Dynamic ref cells from single OFFSET call exceed limit "
                                        f"({len(targets)} > {limits.max_cells})"
                                    )
        return targets

    leaf_addrs: set[str] = set()
    leaf_addrs |= _collect_addresses(rows_ast)
    leaf_addrs |= _collect_addresses(cols_ast)
    if height_ast is not None:
        leaf_addrs |= _collect_addresses(height_ast)
    if width_ast is not None:
        leaf_addrs |= _collect_addresses(width_ast)

    domains = _build_domains(leaf_addrs, cell_type_env, limits)

    for base_range in base_ranges:
        base_bounds = _bounds_for_sheet(bounds, sheet=base_range.sheet)

        for assignment in _enumerate_assignments(domains.values(), limits):
            addr_to_value = dict(zip(domains.keys(), assignment, strict=False))

            def get_cell_value(addr: str, addr_to_value_map=addr_to_value) -> float:
                try:
                    return addr_to_value_map[addr]
                except KeyError as exc:
                    raise DynamicRefError(
                        f"OFFSET argument formula references cell without domain: {addr!r}"
                    ) from exc

            rows_val = _eval_arg(rows_ast, get_cell_value, limits, context=eval_context)
            cols_val = _eval_arg(cols_ast, get_cell_value, limits, context=eval_context)
            height_val = (
                _eval_arg(height_ast, get_cell_value, limits, context=eval_context)
                if height_ast is not None
                else None
            )
            width_val = (
                _eval_arg(width_ast, get_cell_value, limits, context=eval_context)
                if width_ast is not None
                else None
            )

            if isinstance(rows_val, XlError) or isinstance(cols_val, XlError):
                continue
            if isinstance(height_val, XlError) or isinstance(width_val, XlError):
                continue

            result = offset_range(
                base_range,
                rows=rows_val,
                cols=cols_val,
                height=height_val,
                width=width_val,
                bounds=base_bounds,
            )
            if isinstance(result, ExcelRange):
                targets |= set(result.cell_addresses())
                if len(targets) > limits.max_cells:
                    raise DynamicRefError(
                        f"Dynamic ref cells from single OFFSET call exceed limit "
                        f"({len(targets)} > {limits.max_cells})"
                    )

    return targets


def _eval_arg(
    node: AstNode | None,
    get_cell_value,
    limits: DynamicRefLimits,
    context: dict[str, int] | None = None,
) -> float | XlError:
    if node is None:
        return 0.0

    value = evaluate_expr(
        node,
        get_cell_value=get_cell_value,
        max_depth=limits.max_depth,
        context=context,
    )
    if isinstance(value, Unsupported):
        raise DynamicRefError(f"Unsupported argument expression: {value.reason or ''}")
    if isinstance(value, XlError):
        return value
    if isinstance(value, (int, float)):
        return float(value)
    raise DynamicRefError(f"Non-numeric OFFSET argument result: {value!r}")


def _resolve_offset_base(
    base_ast: AstNode,
    *,
    current_sheet: str,
    cell_type_env: CellTypeEnv,
    limits: DynamicRefLimits,
    current_row: int | None = None,
    current_col: int | None = None,
) -> list[ExcelRange]:
    """Resolve OFFSET base to a list of candidate ranges (one per cell when base is INDEX)."""
    if isinstance(base_ast, CellRefNode):
        r = _base_to_range(base_ast, current_sheet=current_sheet)
        return [r]
    if isinstance(base_ast, RangeNode):
        r = _base_to_range(base_ast, current_sheet=current_sheet)
        return [r]
    if isinstance(base_ast, FunctionCallNode) and base_ast.name == "INDEX":
        if len(base_ast.args) < 2 or len(base_ast.args) > 3:
            raise DynamicRefError(
                "OFFSET base INDEX must have 2 or 3 arguments (array, row_num, [column_num])"
            )
        array_ast, row_ast = base_ast.args[0], base_ast.args[1]
        col_ast = base_ast.args[2] if len(base_ast.args) >= 3 else None
        array_range = _base_to_range(array_ast, current_sheet=current_sheet)
        eval_context = (
            {"row": current_row, "column": current_col}
            if current_row is not None and current_col is not None
            else None
        )
        addr_set = _infer_index_targets_core(
            array_range,
            row_ast,
            col_ast,
            cell_type_env,
            limits,
            eval_context,
            current_sheet=current_sheet,
        )
        bases: list[ExcelRange] = []
        for addr in addr_set:
            sheet, coord = _split_address(addr, current_sheet=current_sheet)
            row, col = coordinate_to_tuple(coord)
            bases.append(
                ExcelRange(
                    sheet=sheet,
                    start_row=row,
                    start_col=col,
                    end_row=row,
                    end_col=col,
                )
            )
        return bases
    raise DynamicRefError("OFFSET base must be a cell or range reference")


def _base_to_range(base_ast: AstNode, *, current_sheet: str) -> ExcelRange:
    if isinstance(base_ast, CellRefNode):
        sheet, coord = _split_address(base_ast.address, current_sheet=current_sheet)
        row, col = coordinate_to_tuple(coord)
        return ExcelRange(
            sheet=sheet,
            start_row=row,
            start_col=col,
            end_row=row,
            end_col=col,
        )
    if isinstance(base_ast, RangeNode):
        try:
            s1, coord1 = base_ast.start.split("!", 1)
            s2, coord2 = base_ast.end.split("!", 1)
        except ValueError as exc:
            raise DynamicRefError("OFFSET base range must be a single-sheet A1 range") from exc
        if s1 != s2:
            raise DynamicRefError("OFFSET base range must be on a single sheet")
        row1, col1 = coordinate_to_tuple(coord1)
        row2, col2 = coordinate_to_tuple(coord2)
        start_row, end_row = sorted((row1, row2))
        start_col, end_col = sorted((col1, col2))
        return ExcelRange(
            sheet=s1,
            start_row=start_row,
            start_col=start_col,
            end_row=end_row,
            end_col=end_col,
        )
    raise DynamicRefError("OFFSET base must be a cell or range reference")


def _split_address(addr: str, *, current_sheet: str) -> tuple[str, str]:
    if "!" in addr:
        sheet, coord = addr.split("!", 1)
        return sheet, coord
    return current_sheet, addr


@dataclass(frozen=True, slots=True)
class _IntBounds:
    """Inclusive integer bounds for analysis-only numeric domains."""

    lo: int
    hi: int


@dataclass(frozen=True, slots=True)
class _FiniteInts:
    values: frozenset[int]


@dataclass(frozen=True, slots=True)
class _UnsupportedNumericDiagnostic:
    reason: str
    refs: frozenset[str] = frozenset()
    expression: str | None = None


@dataclass(frozen=True, slots=True)
class _NumericDomainInferenceResult:
    domain: _FiniteInts | _IntBounds | None
    diagnostic: _UnsupportedNumericDiagnostic | None = None


def _domain_result(
    domain: _FiniteInts | _IntBounds | None,
    diagnostic: _UnsupportedNumericDiagnostic | None = None,
) -> _NumericDomainInferenceResult:
    return _NumericDomainInferenceResult(domain=domain, diagnostic=diagnostic)


def _domain_from_cell_type(
    ct: CellType | None, limits: DynamicRefLimits
) -> _FiniteInts | _IntBounds | None:
    if ct is None:
        return None
    if ct.kind not in (CellKind.NUMBER, CellKind.ANY):
        return None
    if ct.enum is not None:
        if not ct.enum.values:
            return _FiniteInts(frozenset())
        ints: list[int] = []
        for v in ct.enum.values:
            if isinstance(v, bool):
                return None
            if not isinstance(v, int):
                return None
            ints.append(v)
        return _FiniteInts(frozenset(ints))
    if ct.interval is not None:
        if ct.interval.min is None or ct.interval.max is None:
            return None
        lo, hi = int(ct.interval.min), int(ct.interval.max)
        if hi < lo:
            return None
        span = hi - lo + 1
        if span <= limits.max_branches:
            return _FiniteInts(frozenset(range(lo, hi + 1)))
        return _IntBounds(lo, hi)
    return None


def _literal_positive_int(node: AstNode) -> int | None:
    """Return a positive whole number when `node` is a numeric literal >= 1."""
    if not isinstance(node, NumberNode):
        return None
    v = node.value
    if isinstance(v, bool):
        return None
    if isinstance(v, int):
        return v if v >= 1 else None
    if isinstance(v, float) and v.is_integer() and v >= 1:
        return int(v)
    return None


def _is_literal_zero(node: AstNode) -> bool:
    """Return True when `node` is the numeric literal 0 (not a boolean)."""
    if not isinstance(node, NumberNode):
        return False
    v = node.value
    if isinstance(v, bool):
        return False
    if isinstance(v, int):
        return v == 0
    if isinstance(v, float) and v.is_integer():
        return int(v) == 0
    return False


def _is_scalar_literal_ast(node: AstNode) -> bool:
    """Return True for literal scalars (numbers, bools, strings, errors)."""
    return isinstance(node, (NumberNode, BoolNode, StringNode, ErrorNode))


def _static_rect_bounds(node: AstNode) -> tuple[str, int, int, int, int] | None:
    """Return `(sheet, rlo, rhi, clo, chi)` for a static cell or rectangular range."""
    if isinstance(node, CellRefNode):
        try:
            sheet, coord = node.address.split("!", 1)
        except ValueError:
            return None
        row, col = coordinate_to_tuple(coord.replace("$", ""))
        return sheet, row, row, col, col
    if isinstance(node, RangeNode):
        try:
            s1, coord_start = node.start.split("!", 1)
            s2, coord_end = node.end.split("!", 1)
        except ValueError:
            return None
        if s1 != s2:
            return None
        row1, col1 = coordinate_to_tuple(coord_start.replace("$", ""))
        row2, col2 = coordinate_to_tuple(coord_end.replace("$", ""))
        rlo, rhi = sorted((row1, row2))
        clo, chi = sorted((col1, col2))
        return s1, rlo, rhi, clo, chi
    return None


def _array_expr_static_bounds(node: AstNode) -> tuple[str, int, int, int, int] | None:
    """Return static bounds for a range or shape-preserving array expression.

    Elementwise unary/binary ops over a static rectangle keep that rectangle's
    shape (e.g. `(N11:N14<>0)`), which MATCH needs for lookup extent even when
    cell values are projected.
    """
    bounds = _static_rect_bounds(node)
    if bounds is not None:
        return bounds
    if isinstance(node, UnaryOpNode):
        return _array_expr_static_bounds(node.operand)
    if isinstance(node, BinaryOpNode):
        left_bounds = _array_expr_static_bounds(node.left)
        right_bounds = _array_expr_static_bounds(node.right)
        if left_bounds is not None and _is_scalar_literal_ast(node.right):
            return left_bounds
        if right_bounds is not None and _is_scalar_literal_ast(node.left):
            return right_bounds
        if left_bounds is not None and right_bounds == left_bounds:
            return left_bounds
    return None


def _range_node_from_bounds(
    sheet: str, rlo: int, rhi: int, clo: int, chi: int
) -> CellRefNode | RangeNode:
    from fastpyxl.utils.cell import get_column_letter

    start = f"{sheet}!{get_column_letter(clo)}{rlo}"
    end = f"{sheet}!{get_column_letter(chi)}{rhi}"
    if start == end:
        return CellRefNode(address=start)
    return RangeNode(start=start, end=end)


def _index_vector_from_bounds(
    bounds: tuple[str, int, int, int, int],
    row_arg: AstNode,
    col_arg: AstNode | None,
) -> AstNode | None:
    """Map INDEX row/col selectors over static bounds to a rectangular result range.

    Supports omitted axes, literal positive selectors, and Excel's `0` form that
    returns an entire row/column/array.
    """
    sheet, rlo, rhi, clo, chi = bounds
    nrows = rhi - rlo + 1
    ncols = chi - clo + 1

    row_omitted = isinstance(row_arg, EmptyArgNode)
    col_omitted = col_arg is None or isinstance(col_arg, EmptyArgNode)
    row_zero = _is_literal_zero(row_arg)
    col_zero = col_arg is not None and _is_literal_zero(col_arg)

    if row_omitted and col_omitted:
        return None

    if row_omitted:
        assert col_arg is not None
        if col_zero:
            return _range_node_from_bounds(sheet, rlo, rhi, clo, chi)
        k = _literal_positive_int(col_arg)
        if k is None or k > ncols:
            return None
        c = clo + k - 1
        return _range_node_from_bounds(sheet, rlo, rhi, c, c)

    if col_omitted:
        if row_zero:
            return _range_node_from_bounds(sheet, rlo, rhi, clo, chi)
        k = _literal_positive_int(row_arg)
        if k is None:
            return None
        if nrows == 1:
            if k > ncols:
                return None
            c = clo + k - 1
            return _range_node_from_bounds(sheet, rlo, rhi, c, c)
        if ncols == 1:
            if k > nrows:
                return None
            r = rlo + k - 1
            return _range_node_from_bounds(sheet, r, r, clo, chi)
        if k > nrows:
            return None
        r = rlo + k - 1
        return _range_node_from_bounds(sheet, r, r, clo, chi)

    # Both axes present: 0 selects the full opposite axis (Excel INDEX).
    if row_zero and col_zero:
        return _range_node_from_bounds(sheet, rlo, rhi, clo, chi)
    if row_zero:
        assert col_arg is not None
        k = _literal_positive_int(col_arg)
        if k is None or k > ncols:
            return None
        c = clo + k - 1
        return _range_node_from_bounds(sheet, rlo, rhi, c, c)
    if col_zero:
        k = _literal_positive_int(row_arg)
        if k is None or k > nrows:
            return None
        r = rlo + k - 1
        return _range_node_from_bounds(sheet, r, r, clo, chi)
    return None


def _index_vector_lookup_array(
    node: AstNode, *, allow_shape_preserving: bool = False
) -> AstNode | None:
    """Resolve static INDEX forms to a rectangular range for MATCH lookup geometry.

    Handles `INDEX(range,,k)` / `INDEX(range,k[,])` and Excel's `INDEX(...,0[,k])`
    whole-axis form. When `allow_shape_preserving` is True, also accepts array
    expressions that keep a static rectangle's shape (e.g. `(range<>0)`).

    Value-preserving callers (exact MATCH cell enumeration) must leave
    `allow_shape_preserving` False so projected arrays are not treated as the
    underlying cells' raw values.
    """
    if not isinstance(node, FunctionCallNode) or node.name.upper() != "INDEX":
        return None
    if len(node.args) < 2:
        return None
    bounds = (
        _array_expr_static_bounds(node.args[0])
        if allow_shape_preserving
        else _static_rect_bounds(node.args[0])
    )
    if bounds is None:
        return None
    col_arg: AstNode | None = node.args[2] if len(node.args) >= 3 else None
    return _index_vector_from_bounds(bounds, node.args[1], col_arg)


def _static_match_lookup_extent(node: AstNode) -> int | None:
    """Return N so MATCH position is within [1, N] when lookup_array has static shape."""
    if isinstance(node, CellRefNode):
        return 1
    if isinstance(node, RangeNode):
        try:
            _s1, coord_start = node.start.split("!", 1)
            _s2, coord_end = node.end.split("!", 1)
        except ValueError:
            return None
        row1, col1 = coordinate_to_tuple(coord_start)
        row2, col2 = coordinate_to_tuple(coord_end)
        rlo, rhi = sorted((row1, row2))
        clo, chi = sorted((col1, col2))
        nrows = rhi - rlo + 1
        ncols = chi - clo + 1
        if nrows == 1:
            return ncols
        if ncols == 1:
            return nrows
        return nrows * ncols
    vector = _index_vector_lookup_array(node, allow_shape_preserving=True)
    if vector is not None:
        return _static_match_lookup_extent(vector)
    return None


def _ordered_match_lookup_cells(arg: AstNode, *, current_sheet: str) -> list[str] | None:
    """Return sheet-qualified addresses for a one-dimensional MATCH lookup_array (in scan order)."""
    from fastpyxl.utils.cell import get_column_letter

    if isinstance(arg, CellRefNode):
        addr = arg.address
        if "!" in addr:
            return [addr]
        return [format_key(current_sheet, addr)]

    if isinstance(arg, RangeNode):
        try:
            s1, coord_start = arg.start.split("!", 1)
            s2, coord_end = arg.end.split("!", 1)
        except ValueError:
            return None
        if s1 != s2:
            return None
        sheet = s1
        row1, col1 = coordinate_to_tuple(coord_start)
        row2, col2 = coordinate_to_tuple(coord_end)
        rlo, rhi = sorted((row1, row2))
        clo, chi = sorted((col1, col2))
        nrows = rhi - rlo + 1
        ncols = chi - clo + 1
        if nrows != 1 and ncols != 1:
            return None
        out: list[str] = []
        if ncols == 1:
            for r in range(rlo, rhi + 1):
                out.append(format_key(sheet, f"{get_column_letter(clo)}{r}"))
        else:
            for c in range(clo, chi + 1):
                out.append(format_key(sheet, f"{get_column_letter(c)}{rlo}"))
        return out
    # Value-preserving INDEX only: projected arrays must not expose raw cell domains.
    vector = _index_vector_lookup_array(arg)
    if vector is not None:
        return _ordered_match_lookup_cells(vector, current_sheet=current_sheet)
    return None


def _domains_may_equal_exact_match(
    a: _FiniteInts | _IntBounds | None,
    b: _FiniteInts | _IntBounds | None,
) -> bool:
    if a is None or b is None:
        return False
    if isinstance(a, _FiniteInts) and isinstance(b, _FiniteInts):
        return bool(a.values & b.values)
    if isinstance(a, _FiniteInts) and isinstance(b, _IntBounds):
        return any(b.lo <= x <= b.hi for x in a.values)
    if isinstance(b, _FiniteInts) and isinstance(a, _IntBounds):
        return any(a.lo <= x <= a.hi for x in b.values)
    if isinstance(a, _IntBounds) and isinstance(b, _IntBounds):
        lo = max(a.lo, b.lo)
        hi = min(a.hi, b.hi)
        return lo <= hi
    return False


def _infer_exact_match_position_domain(
    node: FunctionCallNode,
    env: CellTypeEnv,
    limits: DynamicRefLimits,
    *,
    context: dict[str, int],
    current_sheet: str,
    depth: int,
) -> _FiniteInts | None:
    """If MATCH(...,...,0) has exactly one feasible row/column index, return it."""
    if len(node.args) < 2:
        return None
    ordered = _ordered_match_lookup_cells(node.args[1], current_sheet=current_sheet)
    if not ordered or len(ordered) > limits.max_cells:
        return None

    lookup_res = _infer_numeric_domain_result(
        node.args[0],
        env,
        limits,
        context=context,
        current_sheet=current_sheet,
        depth=depth + 1,
    )
    if lookup_res.diagnostic is not None:
        return None
    lookup_dom = lookup_res.domain
    if lookup_dom is None:
        return None

    candidates: list[int] = []
    for idx, cell_addr in enumerate(ordered, start=1):
        cell_dom = _infer_numeric_domain_result(
            CellRefNode(address=cell_addr),
            env,
            limits,
            context=context,
            current_sheet=current_sheet,
            depth=depth + 1,
        ).domain
        if _domains_may_equal_exact_match(lookup_dom, cell_dom):
            candidates.append(idx)

    if len(candidates) == 1:
        return _FiniteInts(frozenset({candidates[0]}))
    return None


def _match_is_exact_match_type(node: FunctionCallNode) -> bool:
    if len(node.args) < 3:
        return False
    mt = node.args[2]
    if isinstance(mt, NumberNode):
        v = mt.value
        if isinstance(v, bool):
            return False
        return float(v) == 0.0
    return False


def _normalize_to_bounds(d: _FiniteInts | _IntBounds) -> _IntBounds:
    if isinstance(d, _IntBounds):
        return d
    if not d.values:
        return _IntBounds(0, -1)
    return _IntBounds(min(d.values), max(d.values))


def _ast_to_expr_string(node: AstNode) -> str:
    if isinstance(node, NumberNode):
        return str(
            int(node.value)
            if isinstance(node.value, float) and node.value.is_integer()
            else node.value
        )
    if isinstance(node, CellRefNode):
        return node.address
    if isinstance(node, RangeNode):
        return f"{node.start}:{node.end}"
    if isinstance(node, UnaryOpNode):
        return f"{node.op}{_ast_to_expr_string(node.operand)}"
    if isinstance(node, BinaryOpNode):
        return f"({_ast_to_expr_string(node.left)}{node.op}{_ast_to_expr_string(node.right)})"
    if isinstance(node, FunctionCallNode):
        args = ",".join(_ast_to_expr_string(arg) for arg in node.args)
        return f"{node.name.upper()}({args})"
    if isinstance(node, StringNode):
        return f'"{node.value}"'
    if isinstance(node, BoolNode):
        return "TRUE" if node.value else "FALSE"
    if isinstance(node, ErrorNode):
        return str(node.error)
    return type(node).__name__


def _domain_may_include_zero(domain: _FiniteInts | _IntBounds | None) -> bool:
    if domain is None:
        return True
    if isinstance(domain, _FiniteInts):
        return 0 in domain.values
    return domain.lo <= 0 <= domain.hi


def _domain_with_min(
    domain: _FiniteInts | _IntBounds | None,
    minimum: int,
    limits: DynamicRefLimits,
) -> _FiniteInts | _IntBounds | None:
    if domain is None:
        return None
    if isinstance(domain, _FiniteInts):
        values = frozenset(v for v in domain.values if v >= minimum)
        return _FiniteInts(values) if values else None
    lo = max(domain.lo, minimum)
    hi = domain.hi
    if hi < lo:
        return None
    span = hi - lo + 1
    if span <= limits.max_branches:
        return _FiniteInts(frozenset(range(lo, hi + 1)))
    return _IntBounds(lo, hi)


def _domain_with_max(
    domain: _FiniteInts | _IntBounds | None,
    maximum: int,
    limits: DynamicRefLimits,
) -> _FiniteInts | _IntBounds | None:
    if domain is None:
        return None
    if isinstance(domain, _FiniteInts):
        values = frozenset(v for v in domain.values if v <= maximum)
        return _FiniteInts(values) if values else None
    lo = domain.lo
    hi = min(domain.hi, maximum)
    if hi < lo:
        return None
    span = hi - lo + 1
    if span <= limits.max_branches:
        return _FiniteInts(frozenset(range(lo, hi + 1)))
    return _IntBounds(lo, hi)


def _domain_without_zero(
    domain: _FiniteInts | _IntBounds | None,
) -> _FiniteInts | _IntBounds | None:
    if domain is None:
        return None
    if isinstance(domain, _FiniteInts):
        values = frozenset(v for v in domain.values if v != 0)
        return _FiniteInts(values) if values else None
    return domain


def _lookup_cell_type(env: CellTypeEnv, address: str) -> CellType | None:
    """Resolve env entry; keys match `excel_grapher.core.cell_types.normalize_cell_type_env_key`."""
    return env.get(normalize_cell_type_env_key(address))


def _cell_has_relation(
    env: CellTypeEnv, addr: str, relation: type[GreaterThanCell | NotEqualCell], other: str
) -> bool:
    ct = _lookup_cell_type(env, addr)
    if ct is None:
        return False
    other_norm = normalize_cell_type_env_key(other)
    return any(isinstance(rel, relation) and rel.other == other_norm for rel in ct.relations)


def _cells_are_known_not_equal(env: CellTypeEnv, left: str, right: str) -> bool:
    return (
        _cell_has_relation(env, left, GreaterThanCell, right)
        or _cell_has_relation(env, right, GreaterThanCell, left)
        or _cell_has_relation(env, left, NotEqualCell, right)
        or _cell_has_relation(env, right, NotEqualCell, left)
    )


def _refine_difference_domain(
    node: BinaryOpNode,
    env: CellTypeEnv,
    domain: _FiniteInts | _IntBounds | None,
    limits: DynamicRefLimits,
) -> _FiniteInts | _IntBounds | None:
    if domain is None:
        return None
    if not isinstance(node.left, CellRefNode) or not isinstance(node.right, CellRefNode):
        return domain

    left = node.left.address
    right = node.right.address
    if _cell_has_relation(env, left, GreaterThanCell, right):
        return _domain_with_min(domain, 1, limits)
    if _cell_has_relation(env, right, GreaterThanCell, left):
        return _domain_with_max(domain, -1, limits)
    if _cells_are_known_not_equal(env, left, right):
        return _domain_without_zero(domain)
    return domain


def _expr_is_known_nonzero(node: AstNode, env: CellTypeEnv) -> bool:
    if isinstance(node, CellRefNode):
        ct = _lookup_cell_type(env, node.address)
        return (
            ct is not None
            and _domain_may_include_zero(_domain_from_cell_type(ct, DynamicRefLimits())) is False
        )
    if (
        isinstance(node, BinaryOpNode)
        and node.op == "-"
        and isinstance(node.left, CellRefNode)
        and isinstance(node.right, CellRefNode)
    ):
        return _cells_are_known_not_equal(env, node.left.address, node.right.address)
    return False


def _finite_or_hull(
    values: Iterable[int],
    limits: DynamicRefLimits,
) -> _FiniteInts | _IntBounds:
    """Keep `values` exact, or collapse to its bounding interval past `max_branches`."""
    vals = frozenset(values)
    if len(vals) <= limits.max_branches:
        return _FiniteInts(vals)
    return _IntBounds(min(vals), max(vals))


def _densify_or_bounds(
    lo: int,
    hi: int,
    limits: DynamicRefLimits,
) -> _FiniteInts | _IntBounds:
    """Enumerate `[lo, hi]` exactly when narrow enough, else keep it as bounds.

    An inverted span (`hi < lo`) is the empty domain.
    """
    if hi < lo:
        return _FiniteInts(frozenset())
    if hi - lo + 1 <= limits.max_branches:
        return _FiniteInts(frozenset(range(lo, hi + 1)))
    return _IntBounds(lo, hi)


def _pointwise_numeric_domains(
    a: _FiniteInts | _IntBounds | None,
    b: _FiniteInts | _IntBounds | None,
    limits: DynamicRefLimits,
    *,
    values_op: Callable[[int, int], int],
    bounds_op: Callable[[_IntBounds, _IntBounds], tuple[int, int]],
    densify: bool = True,
) -> _FiniteInts | _IntBounds | None:
    """Lift a scalar op to domains: exact over finite pairs, interval arithmetic otherwise."""
    if a is None or b is None:
        return None
    if isinstance(a, _FiniteInts) and isinstance(b, _FiniteInts):
        return _finite_or_hull((values_op(x, y) for x in a.values for y in b.values), limits)
    lo, hi = bounds_op(_normalize_to_bounds(a), _normalize_to_bounds(b))
    if not densify:
        return _FiniteInts(frozenset()) if hi < lo else _IntBounds(lo, hi)
    return _densify_or_bounds(lo, hi, limits)


def _union_numeric_domains(
    a: _FiniteInts | _IntBounds | None,
    b: _FiniteInts | _IntBounds | None,
    limits: DynamicRefLimits,
) -> _FiniteInts | _IntBounds | None:
    if a is None or b is None:
        return None
    if isinstance(a, _FiniteInts) and isinstance(b, _FiniteInts):
        return _finite_or_hull(a.values | b.values, limits)
    ba, bb = _normalize_to_bounds(a), _normalize_to_bounds(b)
    return _densify_or_bounds(min(ba.lo, bb.lo), max(ba.hi, bb.hi), limits)


def _add_numeric_domains(
    a: _FiniteInts | _IntBounds | None,
    b: _FiniteInts | _IntBounds | None,
    limits: DynamicRefLimits,
) -> _FiniteInts | _IntBounds | None:
    return _pointwise_numeric_domains(
        a,
        b,
        limits,
        values_op=lambda x, y: x + y,
        bounds_op=lambda ba, bb: (ba.lo + bb.lo, ba.hi + bb.hi),
    )


def _sub_numeric_domains(
    a: _FiniteInts | _IntBounds | None,
    b: _FiniteInts | _IntBounds | None,
    limits: DynamicRefLimits,
) -> _FiniteInts | _IntBounds | None:
    return _pointwise_numeric_domains(
        a,
        b,
        limits,
        values_op=lambda x, y: x - y,
        bounds_op=lambda ba, bb: (ba.lo - bb.hi, ba.hi - bb.lo),
    )


def _mul_bounds_corners(ba: _IntBounds, bb: _IntBounds) -> tuple[int, int]:
    corners = (ba.lo * bb.lo, ba.lo * bb.hi, ba.hi * bb.lo, ba.hi * bb.hi)
    return min(corners), max(corners)


def _mul_numeric_domains(
    a: _FiniteInts | _IntBounds | None,
    b: _FiniteInts | _IntBounds | None,
    limits: DynamicRefLimits,
) -> _FiniteInts | _IntBounds | None:
    # Unlike ADD/SUB, a narrow product interval is left as bounds rather than re-enumerated.
    return _pointwise_numeric_domains(
        a,
        b,
        limits,
        values_op=lambda x, y: x * y,
        bounds_op=_mul_bounds_corners,
        densify=False,
    )


def _trunc_div_int(numerator: int, denominator: int) -> int:
    return int(numerator / denominator)


def _div_numeric_domains(
    a: _FiniteInts | _IntBounds | None,
    b: _FiniteInts | _IntBounds | None,
    limits: DynamicRefLimits,
    *,
    known_nonzero: bool = False,
) -> _FiniteInts | _IntBounds | None:
    if a is None or b is None:
        return None
    if isinstance(a, _FiniteInts) and isinstance(b, _FiniteInts):
        quotients = {_trunc_div_int(x, y) for x in a.values for y in b.values if y != 0}
        if not quotients:
            return None
        return _finite_or_hull(quotients, limits)
    bb = _normalize_to_bounds(b)
    if bb.lo <= 0 <= bb.hi and not known_nonzero:
        return None
    ba = _normalize_to_bounds(a)
    divisors: tuple[int, ...]
    if bb.hi < 0 or bb.lo > 0:
        divisors = (bb.lo, bb.hi)
    else:
        divisors = (
            *((bb.lo, -1) if bb.lo < 0 else ()),
            *((1, bb.hi) if bb.hi > 0 else ()),
        )
    if not divisors:
        return None
    corners = [_trunc_div_int(num, den) for num in (ba.lo, ba.hi) for den in divisors]
    return _densify_or_bounds(min(corners), max(corners), limits)


def _div_numeric_result(
    left: _NumericDomainInferenceResult,
    right: _NumericDomainInferenceResult,
    node: BinaryOpNode,
    env: CellTypeEnv,
    limits: DynamicRefLimits,
) -> _NumericDomainInferenceResult:
    if left.diagnostic is not None:
        return left
    if right.diagnostic is not None:
        return right
    if left.domain is None or right.domain is None:
        return _domain_result(None)
    refined_right = (
        _refine_difference_domain(node.right, env, right.domain, limits)
        if isinstance(node.right, BinaryOpNode) and node.right.op == "-"
        else right.domain
    )
    known_nonzero = _expr_is_known_nonzero(node.right, env)
    if _domain_may_include_zero(refined_right) and not known_nonzero:
        return _domain_result(
            None,
            _UnsupportedNumericDiagnostic(
                reason="divisor may include zero",
                refs=frozenset(_collect_addresses(node.right)),
                expression=_ast_to_expr_string(node.right),
            ),
        )
    return _domain_result(
        _div_numeric_domains(left.domain, refined_right, limits, known_nonzero=known_nonzero)
    )


def _comparison_numeric_domain(
    op: str,
    a: _FiniteInts | _IntBounds | None,
    b: _FiniteInts | _IntBounds | None,
) -> _FiniteInts | _IntBounds | None:
    if a is None or b is None:
        return None
    if isinstance(a, _FiniteInts) and isinstance(b, _FiniteInts):
        predicates: dict[str, Callable[[int, int], bool]] = {
            "=": lambda x, y: x == y,
            "<>": lambda x, y: x != y,
            "<": lambda x, y: x < y,
            ">": lambda x, y: x > y,
            "<=": lambda x, y: x <= y,
            ">=": lambda x, y: x >= y,
        }
        pred = predicates.get(op)
        if pred is None:
            return None
        out = {1 if pred(x, y) else 0 for x in a.values for y in b.values}
        return _FiniteInts(frozenset(out))
    ba = _normalize_to_bounds(a)
    bb = _normalize_to_bounds(b)
    definitely_true = False
    definitely_false = False
    if op == "=":
        definitely_false = ba.hi < bb.lo or bb.hi < ba.lo
        definitely_true = ba.lo == ba.hi == bb.lo == bb.hi
    elif op == "<>":
        definitely_true = ba.hi < bb.lo or bb.hi < ba.lo
        definitely_false = ba.lo == ba.hi == bb.lo == bb.hi
    elif op == "<":
        definitely_true = ba.hi < bb.lo
        definitely_false = ba.lo >= bb.hi
    elif op == ">":
        definitely_true = ba.lo > bb.hi
        definitely_false = ba.hi <= bb.lo
    elif op == "<=":
        definitely_true = ba.hi <= bb.lo
        definitely_false = ba.lo > bb.hi
    elif op == ">=":
        definitely_true = ba.lo >= bb.hi
        definitely_false = ba.hi < bb.lo
    else:
        return None
    if definitely_true:
        return _FiniteInts(frozenset({1}))
    if definitely_false:
        return _FiniteInts(frozenset({0}))
    return _FiniteInts(frozenset({0, 1}))


def _neg_numeric_domain(d: _FiniteInts | _IntBounds | None) -> _FiniteInts | _IntBounds | None:
    if d is None:
        return None
    if isinstance(d, _FiniteInts):
        vals = {-v for v in d.values}
        return _FiniteInts(frozenset(vals))
    return _IntBounds(-d.hi, -d.lo)


# ---------------------------------------------------------------------------
# Phase 1: ABS / MIN / MAX transfer rules
# ---------------------------------------------------------------------------


def _abs_numeric_domain(
    d: _FiniteInts | _IntBounds | None,
) -> _FiniteInts | _IntBounds | None:
    if d is None:
        return None
    if isinstance(d, _FiniteInts):
        return _FiniteInts(frozenset(abs(v) for v in d.values))
    lo, hi = d.lo, d.hi
    if lo >= 0:
        return _IntBounds(lo, hi)
    if hi <= 0:
        return _IntBounds(-hi, -lo)
    # Mixed: crosses zero — hull [0, max(|lo|, hi)]
    return _IntBounds(0, max(-lo, hi))


def _exp_numeric_domain(
    d: _FiniteInts | _IntBounds | None,
) -> _FiniteInts | _IntBounds | None:
    """Map an integer numeric domain through ``EXP`` as conservative ``_IntBounds``."""
    if d is None:
        return None
    if isinstance(d, _FiniteInts):
        exp_lo = math.inf
        exp_hi = -math.inf
        for value in d.values:
            try:
                ev = math.exp(value)
            except OverflowError:
                return None
            exp_lo = min(exp_lo, ev)
            exp_hi = max(exp_hi, ev)
        return _IntBounds(int(math.floor(exp_lo)), int(math.ceil(exp_hi)))
    lo, hi = d.lo, d.hi
    try:
        exp_lo = math.exp(lo)
        exp_hi = math.exp(hi)
    except OverflowError:
        return None
    return _IntBounds(int(math.floor(exp_lo)), int(math.ceil(exp_hi)))


def _extremum_numeric_domains(
    a: _FiniteInts | _IntBounds,
    b: _FiniteInts | _IntBounds,
    limits: DynamicRefLimits,
    *,
    pick: Callable[[int, int], int],
) -> _FiniteInts | _IntBounds:
    """Apply `MIN`/`MAX` componentwise; both operands are known, so the result is too."""
    if isinstance(a, _FiniteInts) and isinstance(b, _FiniteInts):
        return _finite_or_hull((pick(x, y) for x in a.values for y in b.values), limits)
    ba, bb = _normalize_to_bounds(a), _normalize_to_bounds(b)
    return _densify_or_bounds(pick(ba.lo, bb.lo), pick(ba.hi, bb.hi), limits)


def _min_numeric_domains(
    a: _FiniteInts | _IntBounds,
    b: _FiniteInts | _IntBounds,
    limits: DynamicRefLimits,
) -> _FiniteInts | _IntBounds:
    return _extremum_numeric_domains(a, b, limits, pick=min)


def _max_numeric_domains(
    a: _FiniteInts | _IntBounds,
    b: _FiniteInts | _IntBounds,
    limits: DynamicRefLimits,
) -> _FiniteInts | _IntBounds:
    return _extremum_numeric_domains(a, b, limits, pick=max)


# ---------------------------------------------------------------------------
# Phase 2: Branch-local env refinement helpers for IF
# ---------------------------------------------------------------------------

_NARROW_PREDICATES: dict[str, Callable[[int, int], bool]] = {
    "=": lambda v, b: v == b,
    "<>": lambda v, b: v != b,
    "<": lambda v, b: v < b,
    ">": lambda v, b: v > b,
    "<=": lambda v, b: v <= b,
    ">=": lambda v, b: v >= b,
}
_NEG_OP: dict[str, str] = {"=": "<>", "<>": "=", "<": ">=", ">": "<=", "<=": ">", ">=": "<"}
_FLIP_OP: dict[str, str] = {"<": ">", ">": "<", "<=": ">=", ">=": "<=", "=": "=", "<>": "<>"}


def _narrow_domain(
    d: _FiniteInts | _IntBounds,
    op: str,
    bound: int,
) -> _FiniteInts | _IntBounds | None:
    """Narrow domain d by applying `d op bound` (e.g. d > 3, d <= 5)."""
    if isinstance(d, _FiniteInts):
        predicates = _NARROW_PREDICATES
        pred = predicates.get(op)
        if pred is None:
            return d
        return _FiniteInts(frozenset(v for v in d.values if pred(v, bound)))
    lo, hi = d.lo, d.hi
    if op == "=":
        if lo <= bound <= hi:
            return _FiniteInts(frozenset({bound}))
        return _FiniteInts(frozenset())
    elif op == "<>":
        return d  # can't tighten interval for not-equal; keep conservatively
    elif op == ">":
        lo = max(lo, bound + 1)
    elif op == ">=":
        lo = max(lo, bound)
    elif op == "<":
        hi = min(hi, bound - 1)
    elif op == "<=":
        hi = min(hi, bound)
    else:
        return d
    if hi < lo:
        return _FiniteInts(frozenset())
    return _IntBounds(lo, hi)


def _domain_to_cell_type(domain: _FiniteInts | _IntBounds, kind: CellKind) -> CellType:
    """Convert a narrowed numeric domain back to a CellType for env updates."""
    if isinstance(domain, _FiniteInts):
        return CellType(kind=kind, enum=EnumDomain(values=frozenset(domain.values)))
    return CellType(kind=kind, interval=IntervalDomain(min=domain.lo, max=domain.hi))


def _refine_env_for_condition(
    env: CellTypeEnv,
    cond_node: AstNode,
    limits: DynamicRefLimits,
    *,
    negate: bool = False,
) -> dict[str, CellType] | None:
    """Return a copy of env narrowed by the condition, or None if unsupported.

    Supports:
    - `CellRefNode op NumberNode` and `NumberNode op CellRefNode`
    - `AND(pred1, pred2, ...)` (intersects each predicate)

    The `negate` flag applies the opposite constraint (for the else-branch).
    """
    if isinstance(cond_node, FunctionCallNode) and cond_node.name.upper() == "AND":
        refined: dict[str, CellType] = dict(env)
        changed = False
        for arg in cond_node.args:
            sub = _refine_env_for_condition(refined, arg, limits, negate=negate)
            if sub is not None:
                refined = sub
                changed = True
        return refined if changed else None

    if not isinstance(cond_node, BinaryOpNode):
        return None
    op = cond_node.op
    if op not in {"=", "<>", "<", ">", "<=", ">="}:
        return None

    left, right = cond_node.left, cond_node.right

    if isinstance(left, CellRefNode) and isinstance(right, NumberNode):
        addr, bound_val = left.address, right.value
        if isinstance(bound_val, bool):
            return None
        if not isinstance(bound_val, (int, float)) or not float(bound_val).is_integer():
            return None
        bound = int(bound_val)
        if negate:
            op = _NEG_OP.get(op)
            if op is None:
                return None
        ct = env.get(addr)
        if ct is None:
            return None
        current_domain = _domain_from_cell_type(ct, limits)
        if current_domain is None:
            return None
        narrowed = _narrow_domain(current_domain, op, bound)
        if narrowed is None:
            return None
        return {**env, addr: _domain_to_cell_type(narrowed, ct.kind)}
    elif isinstance(left, NumberNode) and isinstance(right, CellRefNode):
        addr, bound_val = right.address, left.value
        if isinstance(bound_val, bool):
            return None
        if not isinstance(bound_val, (int, float)) or not float(bound_val).is_integer():
            return None
        bound = int(bound_val)
        # Flip op to express as: cell op bound
        op = _FLIP_OP.get(op, op)
        if negate:
            op = _NEG_OP.get(op)
            if op is None:
                return None
        ct = env.get(addr)
        if ct is None:
            return None
        current_domain = _domain_from_cell_type(ct, limits)
        if current_domain is None:
            return None
        narrowed = _narrow_domain(current_domain, op, bound)
        if narrowed is None:
            return None
        return {**env, addr: _domain_to_cell_type(narrowed, ct.kind)}
    elif isinstance(left, CellRefNode) and isinstance(right, CellRefNode):
        # cell op cell: use the other cell's domain bounds and add GreaterThanCell relations.
        laddr, raddr = left.address, right.address
        lct = env.get(laddr)
        rct = env.get(raddr)
        if lct is None or rct is None:
            return None
        ldom = _domain_from_cell_type(lct, limits)
        rdom = _domain_from_cell_type(rct, limits)
        if ldom is None or rdom is None:
            return None
        lb = _normalize_to_bounds(ldom)
        rb = _normalize_to_bounds(rdom)

        effective_op = op
        if negate:
            effective_op = _NEG_OP.get(op)
            if effective_op is None:
                return None

        result_env: dict[str, CellType] = dict(env)
        # Narrow left using right's bounds and vice-versa.
        if effective_op == ">":
            # left > right: left >= rb.lo+1, right <= lb.hi-1; add GreaterThanCell to left
            ln = _narrow_domain(ldom, ">=", rb.lo + 1)
            rn = _narrow_domain(rdom, "<=", lb.hi - 1)
            if ln is not None:
                new_relations = lct.relations + (GreaterThanCell(raddr),)
                result_env[laddr] = CellType(
                    kind=lct.kind,
                    enum=_domain_to_cell_type(ln, lct.kind).enum,
                    interval=_domain_to_cell_type(ln, lct.kind).interval,
                    relations=new_relations,
                )
            if rn is not None:
                result_env[raddr] = _domain_to_cell_type(rn, rct.kind)
        elif effective_op == ">=":
            ln = _narrow_domain(ldom, ">=", rb.lo)
            rn = _narrow_domain(rdom, "<=", lb.hi)
            if ln is not None:
                result_env[laddr] = _domain_to_cell_type(ln, lct.kind)
            if rn is not None:
                result_env[raddr] = _domain_to_cell_type(rn, rct.kind)
        elif effective_op == "<":
            ln = _narrow_domain(ldom, "<=", rb.hi - 1)
            rn = _narrow_domain(rdom, ">=", lb.lo + 1)
            if rn is not None:
                new_relations = rct.relations + (GreaterThanCell(laddr),)
                result_env[raddr] = CellType(
                    kind=rct.kind,
                    enum=_domain_to_cell_type(rn, rct.kind).enum,
                    interval=_domain_to_cell_type(rn, rct.kind).interval,
                    relations=new_relations,
                )
            if ln is not None:
                result_env[laddr] = _domain_to_cell_type(ln, lct.kind)
        elif effective_op == "<=":
            ln = _narrow_domain(ldom, "<=", rb.hi)
            rn = _narrow_domain(rdom, ">=", lb.lo)
            if ln is not None:
                result_env[laddr] = _domain_to_cell_type(ln, lct.kind)
            if rn is not None:
                result_env[raddr] = _domain_to_cell_type(rn, rct.kind)
        elif effective_op == "=":
            # Intersection of both domains (use tighter bounds)
            lo = max(lb.lo, rb.lo)
            hi = min(lb.hi, rb.hi)
            if lo > hi:
                eq_dom: _FiniteInts | _IntBounds = _FiniteInts(frozenset())
            elif hi - lo + 1 <= limits.max_branches:
                eq_dom = _FiniteInts(frozenset(range(lo, hi + 1)))
            else:
                eq_dom = _IntBounds(lo, hi)
            result_env[laddr] = _domain_to_cell_type(eq_dom, lct.kind)
            result_env[raddr] = _domain_to_cell_type(eq_dom, rct.kind)
        elif effective_op == "<>":
            new_relations_l = lct.relations + (NotEqualCell(raddr),)
            result_env[laddr] = CellType(
                kind=lct.kind, enum=lct.enum, interval=lct.interval, relations=new_relations_l
            )
        else:
            return None
        return result_env
    else:
        return None


def _describe_unsupported_numeric_construct(node: AstNode | None) -> str | None:
    if node is None:
        return None
    if isinstance(node, (NumberNode, CellRefNode)):
        return None
    if isinstance(node, (StringNode, BoolNode, ErrorNode, RangeNode)):
        return type(node).__name__
    if isinstance(node, UnaryOpNode):
        if node.op in {"-", "%"}:
            return _describe_unsupported_numeric_construct(node.operand)
        return f"unary operator {node.op!r}"
    if isinstance(node, BinaryOpNode):
        if node.op in {"+", "-", "*", "/", "=", "<>", "<", ">", "<=", ">="}:
            left = _describe_unsupported_numeric_construct(node.left)
            if left is not None:
                return left
            return _describe_unsupported_numeric_construct(node.right)
        return f"binary operator {node.op!r}"
    if isinstance(node, FunctionCallNode):
        if node.name.upper() in {
            "ROW",
            "COLUMN",
            "MATCH",
            "IF",
            "SUM",
            "ISNUMBER",
            "CHOOSE",
            "MIN",
            "MAX",
            "ABS",
            "EXP",
        }:
            for arg in node.args:
                reason = _describe_unsupported_numeric_construct(arg)
                if reason is not None:
                    return reason
            return None
        return f"function {node.name.upper()!r}"
    return type(node).__name__


@lru_cache(maxsize=_RANGE_EXPANSION_CACHE_SIZE)
def _range_node_cell_addresses(node: RangeNode) -> tuple[str, ...] | None:
    """Expand a single-sheet A1 range to sheet-qualified cell keys in row-major order.

    Memoized: a formula chain that mentions the same static range at every level
    re-expands it once per level otherwise (issue #465).
    """
    try:
        norm_start = normalize_cell_type_env_key(node.start)
        norm_end = normalize_cell_type_env_key(node.end)
        sheet, coord_start = norm_start.split("!", 1)
        sheet2, coord_end = norm_end.split("!", 1)
    except ValueError:
        return None
    if sheet != sheet2:
        return None
    row1, col1 = coordinate_to_tuple(coord_start)
    row2, col2 = coordinate_to_tuple(coord_end)
    rlo, rhi = sorted((row1, row2))
    clo, chi = sorted((col1, col2))
    from fastpyxl.utils.cell import get_column_letter

    return tuple(
        f"{sheet}!{get_column_letter(c)}{r}"
        for r in range(rlo, rhi + 1)
        for c in range(clo, chi + 1)
    )


def _infer_sum_argument_domain(
    arg: AstNode,
    env: CellTypeEnv,
    limits: DynamicRefLimits,
    *,
    context: dict[str, int],
    current_sheet: str,
    depth: int,
) -> _FiniteInts | _IntBounds | None:
    if isinstance(arg, CellRefNode):
        return _domain_from_cell_type(_lookup_cell_type(env, arg.address), limits)
    if isinstance(arg, RangeNode):
        addrs = _range_node_cell_addresses(arg)
        if addrs is None:
            return None
        acc: _FiniteInts | _IntBounds | None = _FiniteInts(frozenset({0}))
        for addr in addrs:
            # `_range_node_cell_addresses` already emits normalized env keys, so
            # look them up directly: re-normalizing every cell of the range at
            # every level of a formula chain is the same O(depth x range_cells)
            # cost issue #465 is about.
            d = _domain_from_cell_type(env.get(addr), limits)
            if d is None:
                return None
            acc = _add_numeric_domains(acc, d, limits)
            if acc is None:
                return None
        return acc
    return _infer_numeric_domain(
        arg,
        env,
        limits,
        context=context,
        current_sheet=current_sheet,
        depth=depth,
    )


def _infer_sum_numeric_domain(
    node: FunctionCallNode,
    env: CellTypeEnv,
    limits: DynamicRefLimits,
    *,
    context: dict[str, int],
    current_sheet: str,
    depth: int,
) -> _FiniteInts | _IntBounds | None:
    acc: _FiniteInts | _IntBounds | None = _FiniteInts(frozenset({0}))
    for arg in node.args:
        d = _infer_sum_argument_domain(
            arg,
            env,
            limits,
            context=context,
            current_sheet=current_sheet,
            depth=depth + 1,
        )
        if d is None:
            return None
        acc = _add_numeric_domains(acc, d, limits)
        if acc is None:
            return None
    return acc


def _infer_sum_numeric_domain_result(
    node: FunctionCallNode,
    env: CellTypeEnv,
    limits: DynamicRefLimits,
    *,
    context: dict[str, int],
    current_sheet: str,
    depth: int,
) -> _NumericDomainInferenceResult:
    return _domain_result(
        _infer_sum_numeric_domain(
            node,
            env,
            limits,
            context=context,
            current_sheet=current_sheet,
            depth=depth,
        )
    )


def _infer_choose_numeric_domain(
    node: FunctionCallNode,
    env: CellTypeEnv,
    limits: DynamicRefLimits,
    *,
    context: dict[str, int],
    current_sheet: str,
    depth: int,
) -> _FiniteInts | _IntBounds | None:
    if len(node.args) < 2:
        return None
    index_dom = _infer_numeric_domain(
        node.args[0],
        env,
        limits,
        context=context,
        current_sheet=current_sheet,
        depth=depth + 1,
    )
    if index_dom is None:
        return None
    option_count = len(node.args) - 1
    if isinstance(index_dom, _FiniteInts):
        selected = sorted(i for i in index_dom.values if 1 <= i <= option_count)
    else:
        lo = max(1, index_dom.lo)
        hi = min(option_count, index_dom.hi)
        if hi < lo:
            return None
        selected = list(range(lo, hi + 1))

    out: _FiniteInts | _IntBounds | None = None
    for idx in selected:
        option_dom = _infer_numeric_domain(
            node.args[idx],
            env,
            limits,
            context=context,
            current_sheet=current_sheet,
            depth=depth + 1,
        )
        if option_dom is None:
            return None
        out = option_dom if out is None else _union_numeric_domains(out, option_dom, limits)
        if out is None:
            return None
    return out


def _infer_choose_numeric_domain_result(
    node: FunctionCallNode,
    env: CellTypeEnv,
    limits: DynamicRefLimits,
    *,
    context: dict[str, int],
    current_sheet: str,
    depth: int,
) -> _NumericDomainInferenceResult:
    if len(node.args) < 2:
        return _domain_result(None)
    index_result = _infer_numeric_domain_result(
        node.args[0],
        env,
        limits,
        context=context,
        current_sheet=current_sheet,
        depth=depth + 1,
    )
    if index_result.diagnostic is not None:
        return index_result
    index_dom = index_result.domain
    if index_dom is None:
        return _domain_result(None)
    option_count = len(node.args) - 1
    if isinstance(index_dom, _FiniteInts):
        selected = sorted(i for i in index_dom.values if 1 <= i <= option_count)
    else:
        lo = max(1, index_dom.lo)
        hi = min(option_count, index_dom.hi)
        if hi < lo:
            return _domain_result(None)
        selected = list(range(lo, hi + 1))

    out: _FiniteInts | _IntBounds | None = None
    for idx in selected:
        option_result = _infer_numeric_domain_result(
            node.args[idx],
            env,
            limits,
            context=context,
            current_sheet=current_sheet,
            depth=depth + 1,
        )
        if option_result.diagnostic is not None:
            return option_result
        option_dom = option_result.domain
        if option_dom is None:
            return _domain_result(None)
        out = option_dom if out is None else _union_numeric_domains(out, option_dom, limits)
        if out is None:
            return _domain_result(None)
    return _domain_result(out)


def _infer_numeric_domain_result(
    node: AstNode,
    env: CellTypeEnv,
    limits: DynamicRefLimits,
    *,
    context: dict[str, int] | None = None,
    current_sheet: str = "",
    depth: int = 0,
) -> _NumericDomainInferenceResult:
    """Analysis-only numeric abstract interpretation for selector expressions.

    Returns `None` when the expression is unsupported or cannot be summarized
    soundly as integers. Must never raise for well-formed AST nodes.
    """
    if depth > limits.max_depth:
        return _domain_result(None)
    ctx = context or {}

    if isinstance(node, NumberNode):
        v = node.value
        if isinstance(v, bool):
            return _domain_result(_FiniteInts(frozenset({int(v)})))
        if isinstance(v, int):
            return _domain_result(_FiniteInts(frozenset({v})))
        if isinstance(v, float) and v.is_integer():
            return _domain_result(_FiniteInts(frozenset({int(v)})))
        return _domain_result(None)

    if isinstance(node, (StringNode, BoolNode, ErrorNode)):
        return _domain_result(None)

    if isinstance(node, CellRefNode):
        return _domain_result(_domain_from_cell_type(_lookup_cell_type(env, node.address), limits))

    if isinstance(node, RangeNode):
        return _domain_result(None)

    if isinstance(node, UnaryOpNode):
        if node.op == "-":
            inner = _infer_numeric_domain_result(
                node.operand, env, limits, context=ctx, current_sheet=current_sheet, depth=depth + 1
            )
            if inner.diagnostic is not None:
                return inner
            return _domain_result(_neg_numeric_domain(inner.domain))
        if node.op == "%":
            inner = _infer_numeric_domain_result(
                node.operand, env, limits, context=ctx, current_sheet=current_sheet, depth=depth + 1
            )
            if inner.diagnostic is not None:
                return inner
            return _domain_result(
                _div_numeric_domains(
                    inner.domain,
                    _FiniteInts(frozenset({100})),
                    limits,
                )
            )
        return _domain_result(None)

    if isinstance(node, BinaryOpNode):
        left = _infer_numeric_domain_result(
            node.left, env, limits, context=ctx, current_sheet=current_sheet, depth=depth + 1
        )
        right = _infer_numeric_domain_result(
            node.right, env, limits, context=ctx, current_sheet=current_sheet, depth=depth + 1
        )
        if left.diagnostic is not None:
            return left
        if right.diagnostic is not None:
            return right
        op = node.op
        if op == "+":
            return _domain_result(_add_numeric_domains(left.domain, right.domain, limits))
        if op == "-":
            return _domain_result(
                _refine_difference_domain(
                    node,
                    env,
                    _sub_numeric_domains(left.domain, right.domain, limits),
                    limits,
                )
            )
        if op == "*":
            return _domain_result(_mul_numeric_domains(left.domain, right.domain, limits))
        if op == "/":
            return _div_numeric_result(left, right, node, env, limits)
        if op in {"=", "<>", "<", ">", "<=", ">="}:
            return _domain_result(_comparison_numeric_domain(op, left.domain, right.domain))
        return _domain_result(None)

    if isinstance(node, FunctionCallNode):
        name = node.name.upper()
        if name == "ROW":
            if len(node.args) == 0:
                row = ctx.get("row")
                if row is None:
                    return _domain_result(None)
                return _domain_result(_FiniteInts(frozenset({int(row)})))
            if len(node.args) == 1 and isinstance(node.args[0], CellRefNode):
                cell = _cell_part_from_address_for_infer(node.args[0].address)
                _col_letter, row = coordinate_from_string(cell)
                return _domain_result(_FiniteInts(frozenset({row})))
            return _domain_result(None)
        if name == "COLUMN":
            if len(node.args) == 0:
                col = ctx.get("column")
                if col is None:
                    return _domain_result(None)
                return _domain_result(_FiniteInts(frozenset({int(col)})))
            if len(node.args) == 1 and isinstance(node.args[0], CellRefNode):
                cell = _cell_part_from_address_for_infer(node.args[0].address)
                col_letter, _row = coordinate_from_string(cell)
                from fastpyxl.utils.cell import column_index_from_string

                return _domain_result(
                    _FiniteInts(frozenset({column_index_from_string(col_letter)}))
                )
            return _domain_result(None)
        if name in {"COLUMNS", "ROWS"}:
            if len(node.args) != 1:
                return _domain_result(None)
            bounds = _static_rect_bounds(node.args[0])
            if bounds is None:
                return _domain_result(None)
            _sheet, rlo, rhi, clo, chi = bounds
            if name == "COLUMNS":
                return _domain_result(_FiniteInts(frozenset({chi - clo + 1})))
            return _domain_result(_FiniteInts(frozenset({rhi - rlo + 1})))
        if name == "MATCH":
            if len(node.args) < 2:
                return _domain_result(None)
            if _match_is_exact_match_type(node):
                refined = _infer_exact_match_position_domain(
                    node,
                    env,
                    limits,
                    context=ctx,
                    current_sheet=current_sheet,
                    depth=depth,
                )
                if refined is not None:
                    return _domain_result(refined)
            n = _static_match_lookup_extent(node.args[1])
            if n is None or n < 1:
                return _domain_result(None)
            return _domain_result(_IntBounds(1, n))
        if name == "IF":
            if len(node.args) < 2:
                return _domain_result(None)
            # Check if condition is provably true or false using Excel truthiness:
            # any non-zero value is truthy; zero is falsy.
            cond_result = _infer_numeric_domain_result(
                node.args[0], env, limits, context=ctx, current_sheet=current_sheet, depth=depth + 1
            )
            if cond_result.domain is not None and isinstance(cond_result.domain, _FiniteInts):
                if cond_result.domain.values and all(v != 0 for v in cond_result.domain.values):
                    # Provably truthy: all values are non-zero.
                    return _infer_numeric_domain_result(
                        node.args[1],
                        env,
                        limits,
                        context=ctx,
                        current_sheet=current_sheet,
                        depth=depth + 1,
                    )
                if all(v == 0 for v in cond_result.domain.values):
                    # Provably falsy: all values are zero.
                    if len(node.args) >= 3:
                        return _infer_numeric_domain_result(
                            node.args[2],
                            env,
                            limits,
                            context=ctx,
                            current_sheet=current_sheet,
                            depth=depth + 1,
                        )
                    return _domain_result(_FiniteInts(frozenset({0})))
            # Ambiguous condition: try branch-local environment refinement.
            then_env: CellTypeEnv = (
                _refine_env_for_condition(env, node.args[0], limits, negate=False) or env
            )
            else_env: CellTypeEnv = (
                _refine_env_for_condition(env, node.args[0], limits, negate=True) or env
            )
            then_result = _infer_numeric_domain_result(
                node.args[1],
                then_env,
                limits,
                context=ctx,
                current_sheet=current_sheet,
                depth=depth + 1,
            )
            if then_result.diagnostic is not None:
                # Diagnostic from one branch must not propagate through IF; fall back.
                return _domain_result(None)
            if len(node.args) >= 3:
                else_result = _infer_numeric_domain_result(
                    node.args[2],
                    else_env,
                    limits,
                    context=ctx,
                    current_sheet=current_sheet,
                    depth=depth + 1,
                )
                if else_result.diagnostic is not None:
                    return _domain_result(None)
                else_d = else_result.domain
            else:
                else_d = _FiniteInts(frozenset({0}))
            return _domain_result(_union_numeric_domains(then_result.domain, else_d, limits))
        if name == "SUM":
            return _infer_sum_numeric_domain_result(
                node, env, limits, context=ctx, current_sheet=current_sheet, depth=depth
            )
        if name == "CHOOSE":
            return _infer_choose_numeric_domain_result(
                node, env, limits, context=ctx, current_sheet=current_sheet, depth=depth
            )
        if name == "ISNUMBER":
            if len(node.args) != 1:
                return _domain_result(None)
            arg = node.args[0]
            if isinstance(arg, CellRefNode):
                ct = _lookup_cell_type(env, arg.address)
                if ct is None:
                    return _domain_result(None)
                if ct.kind is CellKind.NUMBER:
                    return _domain_result(_FiniteInts(frozenset({1})))
                if ct.kind is CellKind.ANY and (
                    ct.enum is not None or ct.interval is not None or ct.real_interval is not None
                ):
                    return _domain_result(_FiniteInts(frozenset({0, 1})))
                return _domain_result(_FiniteInts(frozenset({0})))
            arg_result = _infer_numeric_domain_result(
                arg, env, limits, context=ctx, current_sheet=current_sheet, depth=depth + 1
            )
            if arg_result.diagnostic is not None:
                return arg_result
            if arg_result.domain is None:
                return _domain_result(_FiniteInts(frozenset({0})))
            return _domain_result(_FiniteInts(frozenset({0, 1})))
        if name == "ABS":
            if len(node.args) != 1:
                return _domain_result(None)
            inner = _infer_numeric_domain_result(
                node.args[0], env, limits, context=ctx, current_sheet=current_sheet, depth=depth + 1
            )
            if inner.diagnostic is not None:
                return inner
            if inner.domain is None:
                return _domain_result(None)
            return _domain_result(_abs_numeric_domain(inner.domain))
        if name == "EXP":
            if len(node.args) != 1:
                return _domain_result(None)
            inner = _infer_numeric_domain_result(
                node.args[0], env, limits, context=ctx, current_sheet=current_sheet, depth=depth + 1
            )
            if inner.diagnostic is not None:
                return inner
            if inner.domain is None:
                return _domain_result(None)
            return _domain_result(_exp_numeric_domain(inner.domain))
        if name in {"MIN", "MAX"}:
            if len(node.args) < 1:
                return _domain_result(None)
            acc: _FiniteInts | _IntBounds | None = None
            for arg in node.args:
                arg_result = _infer_numeric_domain_result(
                    arg, env, limits, context=ctx, current_sheet=current_sheet, depth=depth + 1
                )
                if arg_result.diagnostic is not None:
                    return arg_result
                if arg_result.domain is None:
                    return _domain_result(None)
                if acc is None:
                    acc = arg_result.domain
                elif name == "MIN":
                    acc = _min_numeric_domains(acc, arg_result.domain, limits)
                else:
                    acc = _max_numeric_domains(acc, arg_result.domain, limits)
            return _domain_result(acc)
        if name == "CONCAT":
            return _domain_result(None)
        return _domain_result(None)

    return _domain_result(None)


def _infer_numeric_domain(
    node: AstNode,
    env: CellTypeEnv,
    limits: DynamicRefLimits,
    *,
    context: dict[str, int] | None = None,
    current_sheet: str = "",
    depth: int = 0,
) -> _FiniteInts | _IntBounds | None:
    return _infer_numeric_domain_result(
        node,
        env,
        limits,
        context=context,
        current_sheet=current_sheet,
        depth=depth,
    ).domain


def _cell_part_from_address_for_infer(addr: str) -> str:
    if "!" in addr:
        return addr.split("!", 1)[-1].strip()
    return addr.strip()


def _index_pair_to_addresses(
    array_range: ExcelRange,
    r: int,
    c: int,
    *,
    nrows: int,
    ncols: int,
) -> set[str]:
    from fastpyxl.utils.cell import get_column_letter

    out: set[str] = set()
    if r == 0 and c == 0:
        out.update(array_range.cell_addresses())
        return out
    if r == 0:
        if 1 <= c <= ncols:
            for row_off in range(nrows):
                cell_row = array_range.start_row + row_off
                cell_col = array_range.start_col + c - 1
                out.add(f"{array_range.sheet}!{get_column_letter(cell_col)}{cell_row}")
        return out
    if c == 0:
        if 1 <= r <= nrows:
            for col_off in range(ncols):
                cell_row = array_range.start_row + r - 1
                cell_col = array_range.start_col + col_off
                out.add(f"{array_range.sheet}!{get_column_letter(cell_col)}{cell_row}")
        return out
    if 1 <= r <= nrows and 1 <= c <= ncols:
        cell_row = array_range.start_row + r - 1
        cell_col = array_range.start_col + c - 1
        out.add(f"{array_range.sheet}!{get_column_letter(cell_col)}{cell_row}")
    return out


def _clamp_int(v: int, lo: int, hi: int) -> int:
    return max(lo, min(hi, v))


# Per-graph-build cache for INDEX target emission: avoids regenerating the same
# target set when multiple INDEX formulas resolve to identical
# (array_range, row_dom, col_dom) triples within a single graph build.
# Cleared at the start of each create_dependency_graph call via
# clear_index_target_cache().
_emit_index_cache: dict[
    tuple[ExcelRange, _FiniteInts | _IntBounds, _FiniteInts | _IntBounds],
    frozenset[str],
] = {}


def clear_index_target_cache() -> None:
    """Clear the INDEX target emission cache.

    Called at the start of each graph build to scope the cache to a single
    invocation and prevent unbounded memory retention across builds.
    """
    _emit_index_cache.clear()


def _emit_index_targets_from_domains(
    array_range: ExcelRange,
    row_dom: _FiniteInts | _IntBounds,
    col_dom: _FiniteInts | _IntBounds,
    limits: DynamicRefLimits,
) -> set[str]:
    nrows = array_range.end_row - array_range.start_row + 1
    ncols = array_range.end_col - array_range.start_col + 1
    targets: set[str] = set()

    if isinstance(row_dom, _FiniteInts) and isinstance(col_dom, _FiniteInts):
        rs = sorted(row_dom.values)
        cs = sorted(col_dom.values)
        for r in rs:
            for c in cs:
                targets |= _index_pair_to_addresses(array_range, r, c, nrows=nrows, ncols=ncols)
        if len(targets) > limits.max_cells:
            raise DynamicRefError(
                f"INDEX target cells exceed limit ({len(targets)} > {limits.max_cells})"
            )
        return targets

    if isinstance(row_dom, _FiniteInts) and isinstance(col_dom, _IntBounds):
        cb = col_dom
        clo = _clamp_int(cb.lo, 1, ncols)
        chi = _clamp_int(cb.hi, 1, ncols)
        if chi < clo:
            return set()
        for r in sorted(row_dom.values):
            if r == 0:
                for c in range(clo, chi + 1):
                    targets |= _index_pair_to_addresses(array_range, 0, c, nrows=nrows, ncols=ncols)
            else:
                cr_lo, cr_hi = clo, chi
                if 1 <= r <= nrows:
                    for c in range(cr_lo, cr_hi + 1):
                        targets |= _index_pair_to_addresses(
                            array_range, r, c, nrows=nrows, ncols=ncols
                        )
        if len(targets) > limits.max_cells:
            raise DynamicRefError(
                f"INDEX target cells exceed limit ({len(targets)} > {limits.max_cells})"
            )
        return targets

    if isinstance(row_dom, _IntBounds) and isinstance(col_dom, _FiniteInts):
        rb = row_dom
        rlo = _clamp_int(rb.lo, 1, nrows)
        rhi = _clamp_int(rb.hi, 1, nrows)
        if rhi < rlo:
            return set()
        for c in sorted(col_dom.values):
            if c == 0:
                for r in range(rlo, rhi + 1):
                    targets |= _index_pair_to_addresses(array_range, r, 0, nrows=nrows, ncols=ncols)
            else:
                cc_lo, cc_hi = rlo, rhi
                if 1 <= c <= ncols:
                    for r in range(cc_lo, cc_hi + 1):
                        targets |= _index_pair_to_addresses(
                            array_range, r, c, nrows=nrows, ncols=ncols
                        )
        if len(targets) > limits.max_cells:
            raise DynamicRefError(
                f"INDEX target cells exceed limit ({len(targets)} > {limits.max_cells})"
            )
        return targets

    rb = _normalize_to_bounds(row_dom)
    cb = _normalize_to_bounds(col_dom)
    r_has_zero = rb.lo <= 0 <= rb.hi
    c_has_zero = cb.lo <= 0 <= cb.hi
    r_pos_lo = max(1, rb.lo)
    r_pos_hi = max(1, rb.hi)
    c_pos_lo = max(1, cb.lo)
    c_pos_hi = max(1, cb.hi)
    r_pos_lo = _clamp_int(r_pos_lo, 1, nrows)
    r_pos_hi = _clamp_int(r_pos_hi, 1, nrows)
    c_pos_lo = _clamp_int(c_pos_lo, 1, ncols)
    c_pos_hi = _clamp_int(c_pos_hi, 1, ncols)

    if r_has_zero and c_has_zero:
        targets |= _index_pair_to_addresses(array_range, 0, 0, nrows=nrows, ncols=ncols)
    if r_has_zero and not c_has_zero and c_pos_hi >= c_pos_lo:
        for c in range(c_pos_lo, c_pos_hi + 1):
            targets |= _index_pair_to_addresses(array_range, 0, c, nrows=nrows, ncols=ncols)
    if c_has_zero and not r_has_zero and r_pos_hi >= r_pos_lo:
        for r in range(r_pos_lo, r_pos_hi + 1):
            targets |= _index_pair_to_addresses(array_range, r, 0, nrows=nrows, ncols=ncols)

    if r_pos_hi >= r_pos_lo and c_pos_hi >= c_pos_lo:
        from fastpyxl.utils.cell import get_column_letter

        for rr in range(array_range.start_row + r_pos_lo - 1, array_range.start_row + r_pos_hi):
            for cc in range(array_range.start_col + c_pos_lo - 1, array_range.start_col + c_pos_hi):
                targets.add(f"{array_range.sheet}!{get_column_letter(cc)}{rr}")

    if len(targets) > limits.max_cells:
        raise DynamicRefError(
            f"INDEX target cells exceed limit ({len(targets)} > {limits.max_cells})"
        )
    return targets


def _infer_index_targets_core(
    array_range: ExcelRange,
    row_ast: AstNode,
    col_ast: AstNode | None,
    cell_type_env: CellTypeEnv,
    limits: DynamicRefLimits,
    eval_context: dict[str, int] | None,
    *,
    current_sheet: str,
) -> set[str]:
    col_effective = col_ast if col_ast is not None else NumberNode(1)
    row_dom = _infer_numeric_domain(
        row_ast,
        cell_type_env,
        limits,
        context=eval_context,
        current_sheet=current_sheet,
    )
    col_dom = _infer_numeric_domain(
        col_effective,
        cell_type_env,
        limits,
        context=eval_context,
        current_sheet=current_sheet,
    )

    nrows, ncols = array_range.shape

    if row_dom is not None and col_dom is not None:
        _cache_key = (array_range, row_dom, col_dom)
        cached = _emit_index_cache.get(_cache_key)
        if cached is not None:
            return set(cached)
        targets = _emit_index_targets_from_domains(array_range, row_dom, col_dom, limits)
        _emit_trace(
            DynamicRefTraceEvent(
                kind="index-abstract",
                name="_emit_index_targets",
                detail={
                    "sheet": array_range.sheet,
                    "shape": f"{nrows}x{ncols}",
                    "row_dom": str(row_dom),
                    "col_dom": str(col_dom),
                    "targets": len(targets),
                },
            )
        )
        _emit_index_cache[_cache_key] = frozenset(targets)
        return targets

    leaf_addrs = _collect_addresses_needing_domain(row_ast)
    if col_ast is not None:
        leaf_addrs |= _collect_addresses_needing_domain(col_ast)
    domains = _build_domains(leaf_addrs, cell_type_env, limits)
    eval_ctx = eval_context

    targets: set[str] = set()
    for assignment in _enumerate_assignments(domains.values(), limits):
        addr_to_value = dict(zip(domains.keys(), assignment, strict=False))

        def get_cell_value(addr: str, m=addr_to_value) -> float:
            try:
                return m[addr]
            except KeyError as exc:
                raise DynamicRefError(
                    f"INDEX argument formula references cell without domain: {addr!r}"
                ) from exc

        row_val = _eval_arg(row_ast, get_cell_value, limits, context=eval_ctx)
        col_val = _eval_arg(col_effective, get_cell_value, limits, context=eval_ctx)
        if isinstance(row_val, XlError) or isinstance(col_val, XlError):
            continue
        r1, c1 = int(row_val), int(col_val)
        targets |= _index_pair_to_addresses(array_range, r1, c1, nrows=nrows, ncols=ncols)
    if len(targets) > limits.max_cells:
        raise DynamicRefError(
            f"INDEX target cells exceed limit ({len(targets)} > {limits.max_cells})"
        )
    _emit_trace(
        DynamicRefTraceEvent(
            kind="index-enumerated",
            name="_emit_index_targets",
            detail={
                "sheet": array_range.sheet,
                "shape": f"{nrows}x{ncols}",
                "leaves": len(leaf_addrs),
                "targets": len(targets),
            },
        )
    )
    return targets


def _numeric_domain_to_int_list(
    dom: _FiniteInts | _IntBounds,
    limits: DynamicRefLimits,
) -> list[int] | None:
    if isinstance(dom, _FiniteInts):
        if len(dom.values) > limits.max_branches:
            return None
        return sorted(dom.values)
    span = dom.hi - dom.lo + 1
    if span > limits.max_branches:
        return None
    return list(range(dom.lo, dom.hi + 1))


def _infer_offset_scalar_domains(
    node: AstNode,
    cell_type_env: CellTypeEnv,
    limits: DynamicRefLimits,
    eval_context: dict[str, int] | None,
    *,
    current_sheet: str,
) -> list[int] | None:
    t0 = time.perf_counter()
    expr_str = _ast_to_expr_string(node)

    result = _infer_offset_scalar_domains_core(
        node, cell_type_env, limits, eval_context, current_sheet=current_sheet
    )

    if result is None:
        _emit_trace(
            DynamicRefTraceEvent(
                kind="offset-scalar-fallback",
                name="_infer_offset_scalar_domains",
                elapsed_s=time.perf_counter() - t0,
                detail={"expr": expr_str},
            )
        )
    elif len(result) > 8:
        _emit_trace(
            DynamicRefTraceEvent(
                kind="offset-scalar-wide",
                name="_infer_offset_scalar_domains",
                elapsed_s=time.perf_counter() - t0,
                detail={"expr": expr_str, "count": len(result)},
            )
        )
    return result


def _infer_offset_scalar_domains_core(
    node: AstNode,
    cell_type_env: CellTypeEnv,
    limits: DynamicRefLimits,
    eval_context: dict[str, int] | None,
    *,
    current_sheet: str,
) -> list[int] | None:
    dom = _infer_numeric_domain(
        node,
        cell_type_env,
        limits,
        context=eval_context,
        current_sheet=current_sheet,
    )
    if dom is not None:
        listed = _numeric_domain_to_int_list(dom, limits)
        if listed is not None:
            return listed
    addrs = _collect_addresses(node)
    try:
        bd = _build_domains(addrs, cell_type_env, limits)
    except DynamicRefError:
        return None
    keys = sorted(bd.keys())
    if not keys:
        return None
    total = 1
    for k in keys:
        total *= len(bd[k])
        if total > limits.max_branches:
            return None
    out_vals: set[int] = set()
    for assignment in product(*(bd[k] for k in keys)):
        addr_to_value = dict(zip(keys, assignment, strict=False))

        def get_cell_value(addr: str, m=addr_to_value) -> float:
            return m[addr]

        v = _eval_arg(node, get_cell_value, limits, context=eval_context)
        if isinstance(v, XlError):
            continue
        out_vals.add(int(v))
    return sorted(out_vals) if out_vals else None


def _narrowed_branch_is_infeasible(
    original_env: CellTypeEnv,
    narrowed_env: dict[str, CellType],
    limits: DynamicRefLimits,
) -> bool:
    """Return True if any cell narrowed by the condition now has an empty domain.

    Only cells that actually changed (different object identity from original_env)
    are checked, so unrelated empty domains do not incorrectly trigger pruning.
    """
    for addr, ct in narrowed_env.items():
        if original_env.get(addr) is ct:
            continue  # unchanged entry
        d = _domain_from_cell_type(ct, limits)
        if isinstance(d, _FiniteInts) and not d.values:
            return True
    return False


def _collect_addresses_needing_domain(
    node: AstNode,
    env: CellTypeEnv | None = None,
    limits: DynamicRefLimits | None = None,
) -> set[str]:
    """Return cell/range addresses that appear in a value context (need a numeric domain).

    Refs that appear only as ref_only arguments (e.g. ROW(ref), COLUMN(ref)) are
    excluded; their implementations use only the reference, not the cell value.

    When `env` and `limits` are provided, IF and CHOOSE branches that are
    provably dead are skipped, reducing the set of required domains.
    """
    addrs: set[str] = set()
    _limits = limits or DynamicRefLimits()

    def visit(
        n: AstNode,
        parent: AstNode | None = None,
        arg_index: int | None = None,
        local_env: CellTypeEnv | None = None,
    ) -> None:
        eff_env = local_env if local_env is not None else env
        if isinstance(n, CellRefNode):
            if not (
                parent is not None
                and isinstance(parent, FunctionCallNode)
                and arg_index is not None
                and is_ref_only_arg(parent.name, arg_index)
            ):
                addrs.add(n.address)
            return
        if isinstance(n, RangeNode):
            if (
                parent is not None
                and isinstance(parent, FunctionCallNode)
                and arg_index is not None
                and is_ref_only_arg(parent.name, arg_index)
            ):
                return
            try:
                sheet, coord_start = n.start.split("!", 1)
                _sheet2, coord_end = n.end.split("!", 1)
            except ValueError:
                return
            row1, col1 = coordinate_to_tuple(coord_start)
            row2, col2 = coordinate_to_tuple(coord_end)
            rlo, rhi = sorted((row1, row2))
            clo, chi = sorted((col1, col2))
            from fastpyxl.utils.cell import get_column_letter

            for r in range(rlo, rhi + 1):
                for c in range(clo, chi + 1):
                    col_letter = get_column_letter(c)
                    addrs.add(f"{sheet}!{col_letter}{r}")
            return
        if isinstance(n, FunctionCallNode) and n.name.upper() == "MATCH" and len(n.args) >= 2:
            visit(n.args[0], n, 0, eff_env)
            if len(n.args) >= 3:
                visit(n.args[2], n, 2, eff_env)
            return
        if isinstance(n, FunctionCallNode) and n.name.upper() == "IF" and env is not None:
            cond_args = n.args
            if len(cond_args) >= 2:
                # Always collect refs from the condition expression.
                visit(cond_args[0], n, 0, eff_env)
                # Determine if the condition is provably truthy or falsy (Excel: non-zero = true).
                cond_dom = _infer_numeric_domain(cond_args[0], eff_env or {}, _limits)
                if (
                    isinstance(cond_dom, _FiniteInts)
                    and cond_dom.values
                    and all(v != 0 for v in cond_dom.values)
                ):
                    # Condition provably truthy: only visit then-branch.
                    visit(cond_args[1], n, 1, eff_env)
                elif isinstance(cond_dom, _FiniteInts) and all(v == 0 for v in cond_dom.values):
                    # Condition provably falsy: only visit else-branch.
                    if len(cond_args) >= 3:
                        visit(cond_args[2], n, 2, eff_env)
                else:
                    # Ambiguous: refine envs per branch and skip infeasible branches.
                    base_env = eff_env or {}
                    then_env_refined = _refine_env_for_condition(
                        base_env, cond_args[0], _limits, negate=False
                    )
                    else_env_refined = _refine_env_for_condition(
                        base_env, cond_args[0], _limits, negate=True
                    )
                    then_env = then_env_refined or base_env
                    else_env = else_env_refined or base_env
                    then_infeasible = (
                        then_env_refined is not None
                        and _narrowed_branch_is_infeasible(base_env, then_env_refined, _limits)
                    )
                    else_infeasible = (
                        else_env_refined is not None
                        and _narrowed_branch_is_infeasible(base_env, else_env_refined, _limits)
                    )
                    if not then_infeasible:
                        visit(cond_args[1], n, 1, then_env)
                    if not else_infeasible and len(cond_args) >= 3:
                        visit(cond_args[2], n, 2, else_env)
                return
        if (
            isinstance(n, FunctionCallNode)
            and n.name.upper() == "CHOOSE"
            and env is not None
            and len(n.args) >= 2
        ):
            # Always collect the index expression refs.
            visit(n.args[0], n, 0, eff_env)
            # Determine which branches are reachable.
            index_dom = _infer_numeric_domain(n.args[0], eff_env or {}, _limits)
            option_count = len(n.args) - 1
            if index_dom is not None:
                if isinstance(index_dom, _FiniteInts):
                    live = sorted(i for i in index_dom.values if 1 <= i <= option_count)
                else:
                    lo = max(1, index_dom.lo)
                    hi = min(option_count, index_dom.hi)
                    live = list(range(lo, hi + 1)) if lo <= hi else []
                for idx in live:
                    visit(n.args[idx], n, idx, eff_env)
                return
            # index_dom is None: fall through to generic handler to collect all option refs.
        if isinstance(n, FunctionCallNode):
            for i, arg in enumerate(n.args):
                visit(arg, n, i, eff_env)
            return
        if hasattr(n, "left") and hasattr(n, "right"):
            visit(cast(AstNode, n.left), n, None, eff_env)
            visit(cast(AstNode, n.right), n, None, eff_env)
        if hasattr(n, "operand"):
            visit(cast(AstNode, n.operand), n, None, eff_env)

    visit(node)
    return addrs


def _collect_addresses(node: AstNode) -> set[str]:
    addrs: set[str] = set()

    def visit(n: AstNode) -> None:
        if isinstance(n, CellRefNode):
            addrs.add(n.address)
            return
        if isinstance(n, RangeNode):
            try:
                sheet, coord_start = n.start.split("!", 1)
                _sheet2, coord_end = n.end.split("!", 1)
            except ValueError:
                return
            row1, col1 = coordinate_to_tuple(coord_start)
            row2, col2 = coordinate_to_tuple(coord_end)
            rlo, rhi = sorted((row1, row2))
            clo, chi = sorted((col1, col2))
            for r in range(rlo, rhi + 1):
                for c in range(clo, chi + 1):
                    # coordinate_to_tuple gives (row, col); we need back to A1
                    from fastpyxl.utils.cell import get_column_letter

                    col_letter = get_column_letter(c)
                    addrs.add(f"{sheet}!{col_letter}{r}")
            return
        if isinstance(n, FunctionCallNode):
            for arg in n.args:
                visit(arg)
            return
        # Binary/unary ops and other nodes: recurse into children where present.
        if hasattr(n, "left") and hasattr(n, "right"):
            visit(cast(AstNode, n.left))
            visit(cast(AstNode, n.right))
        if hasattr(n, "operand"):
            visit(cast(AstNode, n.operand))

    visit(node)
    return addrs


def _split_qualified_to_sheet_a1(qualified: str) -> tuple[str, str]:
    if "!" not in qualified:
        raise ValueError(f"Expected sheet-qualified reference, got {qualified!r}")
    sheet, a1 = parse_address(qualified)
    return sheet.strip(), a1.strip()


def _ast_address_to_ref_key(address: str) -> str:
    sheet, a1 = _split_qualified_to_sheet_a1(address)
    col_letter, row = coordinate_from_string(a1.replace("$", ""))
    return format_key(sheet, f"{col_letter}{row}")


@lru_cache(maxsize=_RANGE_EXPANSION_CACHE_SIZE)
def _expanded_range_keys(
    *,
    sheet: str,
    start_col: str,
    start_row: int,
    end_col: str,
    end_row: int,
    max_cells: int,
) -> tuple[str, ...]:
    """Return the sheet-qualified keys of a static range, memoized.

    Thin cached wrapper over `excel_grapher.grapher.parser.expand_range` (its
    `max_cells` truncation behaviour is preserved).  A long formula chain that
    mentions the same range at every level would otherwise re-derive all of its
    cells once per level (issue #465).
    """
    return tuple(
        format_key(dep_sheet, dep_a1)
        for dep_sheet, dep_a1 in expand_range(
            sheet=sheet,
            start_col=start_col,
            start_row=start_row,
            end_col=end_col,
            end_row=end_row,
            max_cells=max_cells,
        )
    )


def _collect_static_addresses_from_ast(
    node: AstNode,
    *,
    max_range_cells: int,
    sheet_bounds: dict[str, tuple[int, int]] | None = None,
) -> set[str]:
    """Collect static cell/range addresses while skipping dynamic-ref call subtrees.

    Range expansion uses the same `max_range_cells` policy as
    `excel_grapher.grapher.parser.expand_range` in the graph builder so the
    argument subgraph matches `expand_leaf_env_to_argument_env` traversal.
    """
    addrs: set[str] = set()

    def visit(n: AstNode) -> None:
        if isinstance(n, CellRefNode):
            addrs.add(_ast_address_to_ref_key(n.address))
            return
        if isinstance(n, WholeColumnNode):
            bounds = sheet_bounds or {}
            for dep_sheet, dep_a1 in expand_whole_column_deps(n.sheet, n.column, bounds):
                addrs.add(format_key(dep_sheet, dep_a1))
            return
        if isinstance(n, WholeRowNode):
            bounds = sheet_bounds or {}
            _sheet, row = resolve_whole_row_ref(n, None)
            for dep_sheet, dep_a1 in expand_whole_row_deps(_sheet, row, bounds):
                addrs.add(format_key(dep_sheet, dep_a1))
            return
        if isinstance(n, RangeNode):
            try:
                sheet_s, coord_s = _split_qualified_to_sheet_a1(n.start)
                sheet_e, coord_e = _split_qualified_to_sheet_a1(n.end)
            except ValueError:
                return
            if sheet_s != sheet_e:
                for raw in _collect_addresses(n):
                    try:
                        addrs.add(_ast_address_to_ref_key(raw))
                    except ValueError:
                        addrs.add(raw)
                return
            col_s, row_s = coordinate_from_string(coord_s.replace("$", ""))
            col_e, row_e = coordinate_from_string(coord_e.replace("$", ""))
            addrs.update(
                _expanded_range_keys(
                    sheet=sheet_s,
                    start_col=col_s,
                    start_row=row_s,
                    end_col=col_e,
                    end_row=row_e,
                    max_cells=max_range_cells,
                )
            )
            return
        if isinstance(n, FunctionCallNode) and n.name.upper() in {"OFFSET", "INDIRECT", "INDEX"}:
            return
        if isinstance(n, FunctionCallNode):
            for arg in n.args:
                visit(arg)
            return
        if isinstance(n, BinaryOpNode):
            visit(n.left)
            visit(n.right)
            return
        if isinstance(n, UnaryOpNode):
            visit(n.operand)

    visit(node)
    return addrs


def _build_domains(
    addrs: Iterable[str],
    env: CellTypeEnv,
    limits: DynamicRefLimits,
) -> dict[str, list[int]]:
    t0 = time.perf_counter()
    domains: dict[str, list[int]] = {}
    try:
        for addr in addrs:
            ct = _lookup_cell_type(env, addr)
            if ct is None:
                raise DynamicRefError(f"Missing CellType for {addr!r}")
            if ct.kind is not CellKind.NUMBER:
                raise DynamicRefError(f"CellType for {addr!r} must be numeric, got {ct.kind!r}")
            vals: list[int]
            if ct.enum is not None:
                vals = [int(v) for v in ct.enum.values]
            elif ct.interval is not None:
                vals = _interval_to_values(ct.interval, limits)
            elif ct.real_interval is not None:
                raise DynamicRefError(
                    f"CellType for {addr!r} uses a real interval (RealBetween); "
                    "integer enum or Between bounds are required to enumerate OFFSET/INDEX arguments."
                )
            else:
                raise DynamicRefError(
                    f"CellType for {addr!r} has no enum or interval domain (e.g. formula uses "
                    "OFFSET/INDIRECT and could not be inferred). Add a constraint for this cell."
                )
            if not vals:
                raise DynamicRefError(f"Empty domain for {addr!r}")
            domains[addr] = sorted(vals)

        branch_estimate = 1
        for vs in domains.values():
            branch_estimate *= len(vs)
            if branch_estimate > limits.max_branches:
                raise DynamicRefError(
                    f"Dynamic ref branches exceed limit ({branch_estimate} > {limits.max_branches})"
                )
    except Exception as exc:
        _emit_trace(
            DynamicRefTraceEvent(
                kind="build-domains-error",
                name="_build_domains",
                elapsed_s=time.perf_counter() - t0,
                detail={"refs": len(domains), "error": f"{type(exc).__name__}: {exc}"},
            )
        )
        raise
    _emit_trace(
        DynamicRefTraceEvent(
            kind="build-domains",
            name="_build_domains",
            elapsed_s=time.perf_counter() - t0,
            detail={"refs": len(domains), "branch_estimate": branch_estimate},
        )
    )
    return domains


def _interval_to_values(interval: IntervalDomain, limits: DynamicRefLimits) -> list[int]:
    if interval.min is None or interval.max is None:
        raise DynamicRefError("Unbounded intervals are not supported for dynamic refs")
    lo, hi = int(interval.min), int(interval.max)
    if hi < lo:
        raise DynamicRefError(f"Invalid interval domain [{lo}, {hi}]")
    count = hi - lo + 1
    if count > limits.max_branches:
        raise DynamicRefError(f"Interval size {count} exceeds branch limit {limits.max_branches}")
    return list(range(lo, hi + 1))


def _enumerate_assignments(
    domains: Iterable[list[int]],
    limits: DynamicRefLimits,
) -> Iterable[tuple[int, ...]]:
    # Domains size has already been checked in _build_domains; this is a thin wrapper.
    return product(*domains)


def infer_dynamic_indirect_targets(
    formula: str,
    *,
    current_sheet: str,
    cell_type_env: CellTypeEnv,
    limits: DynamicRefLimits | None = None,
    bounds: WorkbookBoundsProtocol | None = None,
    named_ranges: Mapping[str, tuple[str, str]] | None = None,
    named_range_ranges: Mapping[str, tuple[str, str, str]] | None = None,
) -> set[str]:
    """Infer the union of all possible INDIRECT targets for a formula."""
    if not isinstance(formula, str) or not formula.startswith("="):
        return set()

    t0 = time.perf_counter()
    lim = limits or DynamicRefLimits()
    out: set[str] = set()

    calls = _find_function_calls_with_spans(formula, frozenset({"INDIRECT"}))
    for fn, inner, _span in calls:
        if fn != "INDIRECT":
            continue
        targets = _infer_single_indirect_call(
            inner,
            current_sheet=current_sheet,
            cell_type_env=cell_type_env,
            limits=lim,
            bounds=bounds,
            named_ranges=named_ranges,
            named_range_ranges=named_range_ranges,
        )
        out |= targets
        if len(out) > lim.max_cells:
            raise DynamicRefError(f"Dynamic ref cells exceed limit ({len(out)} > {lim.max_cells})")

    _emit_trace(
        DynamicRefTraceEvent(
            kind="infer",
            name="infer_dynamic_indirect_targets",
            elapsed_s=time.perf_counter() - t0,
            detail={"targets": len(out), "formula": formula, "current_sheet": current_sheet},
        )
    )
    return out


def _infer_single_indirect_call(
    inner_args: str,
    *,
    current_sheet: str,
    cell_type_env: CellTypeEnv,
    limits: DynamicRefLimits,
    bounds: WorkbookBoundsProtocol | None,
    named_ranges: Mapping[str, tuple[str, str]] | None = None,
    named_range_ranges: Mapping[str, tuple[str, str, str]] | None = None,
) -> set[str]:
    args = _split_top_level_args(inner_args)
    if args is None or len(args) < 1 or len(args) > 2:
        raise DynamicRefError("INDIRECT expects 1 or 2 arguments")

    nr = named_ranges or {}
    nrr = named_range_ranges or {}
    text_expr = _qualify_fragment(args[0], nr, nrr)
    a1_expr = _qualify_fragment(args[1], nr, nrr) if len(args) == 2 else ""

    text_ast = parse_ast("=" + text_expr)
    a1_ast = parse_ast("=" + a1_expr) if a1_expr else None

    leaf_addrs: set[str] = set()
    leaf_addrs |= _collect_addresses(text_ast)
    if a1_ast is not None:
        leaf_addrs |= _collect_addresses(a1_ast)

    domains = _build_value_domains(leaf_addrs, cell_type_env, limits)

    targets: set[str] = set()

    for assignment in _enumerate_value_assignments(domains.values(), limits):
        addr_to_value = dict(zip(domains.keys(), assignment, strict=False))

        def get_cell_value(addr: str, addr_to_value_map=addr_to_value) -> Any:
            try:
                return addr_to_value_map[addr]
            except KeyError as exc:
                raise DynamicRefError(
                    f"INDIRECT argument formula references cell without domain: {addr!r}"
                ) from exc

        text_value = evaluate_expr(
            text_ast, get_cell_value=get_cell_value, max_depth=limits.max_depth
        )
        if isinstance(text_value, Unsupported):
            raise DynamicRefError(
                f"Unsupported INDIRECT text expression: {text_value.reason or ''}"
            )
        if isinstance(text_value, XlError):
            continue
        if not isinstance(text_value, str):
            raise DynamicRefError(
                f"INDIRECT text argument must be a string, got {type(text_value).__name__}"
            )

        if a1_ast is None:
            a1_flag = True
        else:
            a1_value = evaluate_expr(
                a1_ast, get_cell_value=get_cell_value, max_depth=limits.max_depth
            )
            if isinstance(a1_value, Unsupported):
                raise DynamicRefError(
                    f"Unsupported INDIRECT A1/R1C1 flag expression: {a1_value.reason or ''}"
                )
            if isinstance(a1_value, XlError):
                continue
            if isinstance(a1_value, bool):
                a1_flag = a1_value
            elif isinstance(a1_value, (int, float)):
                a1_flag = bool(a1_value)
            else:
                raise DynamicRefError(
                    f"INDIRECT A1/R1C1 flag must be boolean or numeric, got {type(a1_value).__name__}"
                )

        # Derive per-call bounds so sheet-qualified references are not rejected.
        sheet_for_bounds = _sheet_from_indirect_text(text_value, current_sheet=current_sheet)
        local_bounds = _bounds_for_sheet(bounds, sheet=sheet_for_bounds)

        result = indirect_text_to_range(text_value, a1_flag, bounds=local_bounds)
        if isinstance(result, ExcelRange):
            targets |= set(result.cell_addresses())

    return targets


def _sheet_from_indirect_text(text: str, *, current_sheet: str) -> str:
    raw = text.strip()
    if not raw:
        return current_sheet
    if "!" in raw:
        sheet_text, _addr = raw.split("!", 1)
        return sheet_text or current_sheet
    return current_sheet


def _build_value_domains(
    addrs: Iterable[str],
    env: CellTypeEnv,
    limits: DynamicRefLimits,
) -> dict[str, list[Any]]:
    t0 = time.perf_counter()
    domains: dict[str, list[Any]] = {}
    try:
        for addr in addrs:
            ct = _lookup_cell_type(env, addr)
            if ct is None:
                raise DynamicRefError(f"Missing CellType for {addr!r}")
            values: list[Any]
            if ct.enum is not None:
                values = list(ct.enum.values)
            elif ct.interval is not None:
                values = _interval_to_values(ct.interval, limits)
            elif ct.real_interval is not None:
                raise DynamicRefError(
                    f"CellType for {addr!r} uses a real interval (RealBetween); "
                    "use Literal / integer Between or an explicit enum for INDIRECT text arguments."
                )
            else:
                raise DynamicRefError(
                    f"CellType for {addr!r} must have an interval or enum domain for INDIRECT analysis"
                )
            if not values:
                raise DynamicRefError(f"Empty domain for {addr!r}")
            domains[addr] = values

        branch_estimate = 1
        for vs in domains.values():
            branch_estimate *= len(vs)
            if branch_estimate > limits.max_branches:
                raise DynamicRefError(
                    f"Dynamic ref branches exceed limit ({branch_estimate} > {limits.max_branches})"
                )
    except Exception as exc:
        _emit_trace(
            DynamicRefTraceEvent(
                kind="build-value-domains-error",
                name="_build_value_domains",
                elapsed_s=time.perf_counter() - t0,
                detail={"refs": len(domains), "error": f"{type(exc).__name__}: {exc}"},
            )
        )
        raise
    _emit_trace(
        DynamicRefTraceEvent(
            kind="build-value-domains",
            name="_build_value_domains",
            elapsed_s=time.perf_counter() - t0,
            detail={"refs": len(domains), "branch_estimate": branch_estimate},
        )
    )
    return domains


def _enumerate_value_assignments(
    domains: Iterable[list[Any]],
    limits: DynamicRefLimits,
) -> Iterable[tuple[Any, ...]]:
    return product(*domains)


def _split_top_level_args(s: str) -> list[str] | None:
    """Minimal top-level argument splitter mirroring parser._split_top_level_args."""
    buf: list[str] = []
    args: list[str] = []
    depth = 0
    in_str = False
    i = 0
    while i < len(s):
        ch = s[i]
        if ch == '"':
            in_str = not in_str
            buf.append(ch)
            i += 1
            continue
        if in_str:
            buf.append(ch)
            i += 1
            continue
        if ch == "(":
            depth += 1
            buf.append(ch)
            i += 1
            continue
        if ch == ")":
            if depth == 0:
                return None
            depth -= 1
            buf.append(ch)
            i += 1
            continue
        if ch == "," and depth == 0:
            args.append("".join(buf).strip())
            buf = []
            i += 1
            continue
        buf.append(ch)
        i += 1
    if in_str or depth != 0:
        return None
    args.append("".join(buf).strip())
    return [a for a in args if a != ""]
