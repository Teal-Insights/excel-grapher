from __future__ import annotations

import logging
import math
import re
import time
from collections.abc import Callable, Iterable, Iterator, Mapping, Sequence
from contextlib import contextmanager
from contextvars import ContextVar
from dataclasses import dataclass, field
from enum import Enum
from functools import lru_cache
from itertools import product
from pathlib import Path
from typing import TYPE_CHECKING, Any, cast

if TYPE_CHECKING:
    from excel_grapher.grapher.type_analysis_cache import TypeAnalysisCache

from fastpyxl.utils.cell import coordinate_from_string, coordinate_to_tuple

from excel_grapher.core.address_keys import format_range_key, parse_address
from excel_grapher.core.addressing import (
    WorkbookBoundsProtocol,
    indirect_text_to_range,
    offset_range,
)
from excel_grapher.core.cell_types import (
    CellKind,
    CellType,
    CellTypeEnv,
    CellTypeEnvDict,
    EnumDomain,
    GreaterThanCell,
    IntervalDomain,
    NotEqualCell,
    canonicalize_cell_type_env_keys,
    constraints_to_cell_type_env,
    lookup_cell_type,
    normalize_cell_type_env_key,
)
from excel_grapher.core.coercions import excel_casefold, to_string, try_coerce_string_to_float
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
    bind_axes,
    resolve_whole_column_ref,
    resolve_whole_row_ref,
)
from excel_grapher.core.formula_ast import (
    parse as parse_ast,
)
from excel_grapher.core.formula_ast_json import formula_identity_digest
from excel_grapher.core.lookup_funcs import _values_match
from excel_grapher.core.range_shorthand import (
    expand_whole_column_span_deps,
    expand_whole_row_span_deps,
)
from excel_grapher.core.text_funcs import value_from_text
from excel_grapher.core.types import ExcelRange, XlError

from .blank_ranges import BlankRangeRect, address_in_blank_ranges
from .parser import (
    DEFAULT_MAX_RANGE_CELLS,
    _find_function_calls_with_spans,
    expand_range,
    format_key,
    mask_ref_only_function_calls,
)

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Tracing infrastructure
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class DynamicRefTraceEvent:
    """A single trace event emitted during dynamic-ref inference.

    Attributes:
        kind: Event type (`"infer"`, `"expand-env"`, `"expand-env-skipped"`,
            `"build-domains"`, `"build-value-domains"`, `"offset-scalar-fallback"`,
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
"""Max nested argument-subgraph cells before raising `DynamicRefError`.

This is a logical depth on `_analysis_stack`, not a Python call-stack bound.
`expand_leaf_env_to_argument_env` walks the subgraph with an explicit worklist
(issue #716) so a 400-cell chain does not consume 400 CPython frames.
"""


class DynamicRefError(ValueError):
    """Raised when dynamic reference analysis cannot proceed.

    When building a dependency graph, pass a `DynamicRefConfig` (e.g. via
    `DynamicRefConfig.from_bindings` or `DynamicRefConfig.from_constraints`)
    or set `use_cached_dynamic_refs=True` to resolve OFFSET/INDIRECT instead
    of raising.
    """


class DynamicRefCellLimitError(DynamicRefError):
    """Raised when inferred dynamic-ref targets exceed `max_cells`.

    Candidate scanning propagates this error instead of returning a silently
    incomplete leaf list. Raise `DynamicRefLimits.max_cells` or tighten the
    selector domain.
    """


def _raise_cell_limit(count: int, limit: int, *, what: str) -> None:
    """Raise `DynamicRefCellLimitError` for an over-budget target set."""
    raise DynamicRefCellLimitError(
        f"{what} exceed limit ({count} > {limit}). "
        "Refusing to drop inferred targets. Raise DynamicRefLimits.max_cells "
        "or tighten the selector domain."
    )


@dataclass(frozen=True)
class DynamicRefLimits:
    """Tuneable safety limits for dynamic-reference inference.

    Pass a custom instance via the `limits` parameter of
    `DynamicRefConfig.from_bindings` or `DynamicRefConfig.from_constraints`
    to override any of these defaults.

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

    Prefer building via `from_bindings` (series `domain`) or `from_constraints`
    (address -> annotation mapping); the constructor is for internal use.
    """

    cell_type_env: CellTypeEnv
    limits: DynamicRefLimits

    @classmethod
    def from_constraints(
        cls,
        constraints_schema: Mapping[str, Any],
        *,
        limits: DynamicRefLimits | None = None,
    ) -> DynamicRefConfig:
        r"""Build a config from a constraints schema (address keys -> type annotations).

        *constraints_schema* must be a mapping (typically `dict[str, type]`) whose keys are
        sheet-qualified addresses (e.g. `Sheet1!B1`). Values are typing objects describing
        domains (`Annotated`, `Literal`, etc.). Prefer `from_bindings` when domains live on
        a series binding sidecar.

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
        env = constraints_to_cell_type_env(constraints_schema, {})
        return cls(cell_type_env=env, limits=limits or DynamicRefLimits())

    @classmethod
    def from_bindings(
        cls,
        bindings: Mapping[str, Any],
        workbook: str | Path,
        *,
        limits: DynamicRefLimits | None = None,
        bindings_path: str | Path | None = None,
    ) -> DynamicRefConfig:
        """Build a config whose env is derived from a series binding manifest.

        Args:
            bindings: Loaded (merged) series binding manifest.
            workbook: Workbook path; read for range expansion and `from_workbook`.
            limits: Optional inference limits (defaults to `DynamicRefLimits()`).
            bindings_path: Sidecar path stored on the domain-index pickle handle.
        """
        from excel_grapher.series_bindings.domains import SeriesDomainIndex

        env = SeriesDomainIndex.from_bindings(
            bindings, workbook=workbook, bindings_path=bindings_path
        )
        return cls(cell_type_env=env, limits=limits or DynamicRefLimits())

    def overlay(self, other: DynamicRefConfig) -> tuple[DynamicRefConfig, tuple[str, ...]]:
        """Union this env with `other`; `other` wins per overlapping key.

        Args:
            other: Config whose cell types replace this env on shared addresses.

        Returns:
            The merged config (using `other.limits`) and the sorted keys present
            in both envs.
        """
        base = dict(self.cell_type_env)
        extra = dict(other.cell_type_env)
        overrides = tuple(sorted(key for key in extra if key in base))
        base.update(extra)
        return (
            DynamicRefConfig(cell_type_env=base, limits=other.limits),
            overrides,
        )


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


_BLANK_RANGE_LEAF_TYPE = CellType(kind=CellKind.NUMBER, enum=EnumDomain(values=frozenset({0})))
"""Excel blank used as an OFFSET/INDEX selector: numeric `0` (empty cell coerce)."""


class _NormalizedCellTypeCache:
    """Write-through view that stores `CellTypeEnv` keys in canonical form.

    Used only when the caller shares a plain `dict`. A `CellTypeEnvDict` is
    used directly so already-canonical membership stays a raw hash lookup.
    """

    __slots__ = ("_store",)

    def __init__(self, store: dict[str, CellType]) -> None:
        self._store = store

    def __setitem__(self, key: str, value: CellType) -> None:
        self._store[normalize_cell_type_env_key(key)] = value

    def __getitem__(self, key: str) -> CellType:
        return self._store[normalize_cell_type_env_key(key)]

    def __contains__(self, key: object) -> bool:
        return isinstance(key, str) and normalize_cell_type_env_key(key) in self._store

    def get(self, key: str, default: CellType | None = None) -> CellType | None:
        return self._store.get(normalize_cell_type_env_key(key), default)

    def __len__(self) -> int:
        return len(self._store)


def expand_leaf_env_to_argument_env(
    argument_refs: set[str],
    get_cell_formula: Callable[[str], str | None],
    get_refs_from_formula: Callable[[str, str], set[str]],
    leaf_env: CellTypeEnv,
    limits: DynamicRefLimits,
    named_ranges: Mapping[str, tuple[str, str]] | None = None,
    named_range_ranges: Mapping[str, tuple[str, str, str]] | None = None,
    *,
    max_range_cells: int = DEFAULT_MAX_RANGE_CELLS,
    shared_cell_type_cache: dict[str, CellType] | None = None,
    type_analysis_cache: TypeAnalysisCache | None = None,
    workbook_sha256: str | None = None,
    get_cell_ast: Callable[[str], AstNode | None] | None = None,
    blank_rects: Sequence[BlankRangeRect] | None = None,
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

    Leaves inside `blank_rects` are typed as a numeric `0` singleton (Excel
    blank-as-zero) so declared structural pads do not need a series `CellType`.

    `max_range_cells` must match the graph builder's range expansion limit so static
    ranges collected from the AST align with
    `excel_grapher.grapher.builder.create_dependency_graph` argument-subgraph BFS.

    Returned (and shared-cache) keys are `normalize_cell_type_env_key` of each
    address so `_lookup_cell_type` / `lookup_cell_type` can resolve graph
    `format_key` addresses after expand (issue #972). A `CellTypeEnvDict`
    shared cache is used in place (no O(|cache|) re-canonicalize); a plain
    dict is rewritten once, then writes go through a normalizing view.

    When `shared_cell_type_cache` is provided, intermediate cell type inferences
    are persisted across multiple calls.  This avoids redundant work when many
    BFS nodes share intermediate formula cells in their argument subgraphs.
    Addresses already present in that cache are not re-walked: a later call
    whose `argument_refs` are a subset of a previous expansion returns without
    re-entering `cell_type_for` (issue #528).

    When `type_analysis_cache` and `workbook_sha256` are provided, successful
    intermediate formula-cell type results are persisted to SQLite across runs.

    `get_cell_ast`, when provided, supplies the stored per-cell `formula_ast`.
    Cache identity then uses that tree so relative vs absolute formulas that
    share A1 text do not collide. The fallback path keys by formula string and
    parses only after a cache miss.
    """
    backing: dict[str, CellType] = (
        shared_cell_type_cache if shared_cell_type_cache is not None else CellTypeEnvDict()
    )
    if isinstance(backing, CellTypeEnvDict):
        cache: CellTypeEnvDict | _NormalizedCellTypeCache = backing
    else:
        canonicalize_cell_type_env_keys(backing)
        cache = _NormalizedCellTypeCache(backing)
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

    def _try_persistent_lookup(
        addr: str, formula: str, formula_ast: object | None
    ) -> tuple[CellType, list[str]] | None:
        """Try to load a cached type from the persistent SQLite cache.

        Returns `(cell_type, consumed_leaf_keys)` on hit, or `None`.
        """
        if _tac is None or workbook_sha256 is None:
            return None
        ast = formula_ast if isinstance(formula_ast, AstNode) else None
        norm_formula_sha = formula_identity_digest(formula=formula, formula_ast=ast)
        return _tac.get_formula_cell_type(
            workbook_sha256=workbook_sha256,
            address=addr,
            normalized_formula_sha256=norm_formula_sha,
            limits_fingerprint=_limits_fp,
            current_leaf_env=leaf_env,
        )

    def _persist_result(addr: str, formula: str, formula_ast: object | None, ct: CellType) -> None:
        """Write a successful formula-cell type to the persistent cache."""
        if _tac is None or workbook_sha256 is None:
            return
        # Don't cache ANY results in v1
        if ct.kind is CellKind.ANY and ct.enum is None:
            return
        ast = formula_ast if isinstance(formula_ast, AstNode) else None
        norm_formula_sha = formula_identity_digest(formula=formula, formula_ast=ast)
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

        Cached refs are read straight out of `cache`; cycle back-edges (still
        on `_analysis_stack`) resolve to `ANY`. Uncached deps are typed by the
        worklist before `_infer_cell_type` runs, so this is a dict lookup.
        """
        nonlocal _bulk_hits
        ref_types: dict[str, CellType] = {}
        bulk_served: list[str] = []
        for r in refs:
            norm = normalize_cell_type_env_key(r)
            cached_ct = cache.get(norm)
            if cached_ct is not None:
                ref_types[norm] = cached_ct
                bulk_served.append(norm)
                continue
            # Cycle back-edge (still on the analysis stack) or a worklist miss.
            ref_types[norm] = CellType(kind=CellKind.ANY)
        if bulk_served:
            _bulk_hits += len(bulk_served)
            _note_progress(bulk_served[-1])
            _propagate_consumed_leaves_bulk(bulk_served)
        return ref_types

    def _exit_cell(addr: str, formula: str | None, stored_ast: AstNode | None) -> None:
        norm = normalize_cell_type_env_key(addr)
        in_progress.discard(norm)
        if _analysis_stack and _analysis_stack[-1] == norm:
            _analysis_stack.pop()
            if _track_consumed and _analysis_stack:
                parent = _analysis_stack[-1]
                child_leaves = _consumed_leaves.get(norm)
                if child_leaves:
                    _consumed_leaves.setdefault(parent, set()).update(child_leaves)
        if norm in cache and formula is not None and norm not in _loaded_from_persistent:
            _persist_result(norm, formula, stored_ast, cache[norm])

    def _infer_cell_type(
        addr: str,
        formula: str,
        ast_root: AstNode | None,
        refs: set[str],
    ) -> None:
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
                            enum=EnumDomain(values=frozenset(range(inferred.lo, inferred.hi + 1))),
                        )
                    else:
                        cache[addr] = CellType(
                            kind=CellKind.NUMBER,
                            interval=IntervalDomain(min=inferred.lo, max=inferred.hi),
                        )
                return
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
                return
            else:
                cache[addr] = CellType(kind=CellKind.ANY)
                return
        total_branches = math.prod(len(v) for v in domains.values())
        if total_branches > limits.max_branches:
            dep_sizes = ", ".join(f"{r!r}: {len(domains[r])}" for r in sorted(domains))
            unsupported_hint = (
                f" First unsupported construct: {unsupported}." if unsupported is not None else ""
            )
            raise DynamicRefError(
                f"Formula cell {addr!r} fallback enumeration would require "
                f"{total_branches} branches (limit {limits.max_branches}). "
                f"Dependency domain sizes: {dep_sizes}.{unsupported_hint} "
                f"Tighten constraints on one or more dependencies, simplify "
                f"the formula, or extend numeric abstract analysis to cover it."
            )
        result_values: set[Any] = set()
        ordered_refs = list(ref_types)
        for assignment in product(*(domains[r] for r in ordered_refs)):
            addr_to_val = dict(zip(ordered_refs, assignment, strict=False))

            def get_cell_value(a: str, _av=addr_to_val) -> Any:
                try:
                    return _av.get(normalize_cell_type_env_key(a))
                except (IndexError, ValueError):
                    return _av.get(a)

            try:
                formula_parse = _formula_to_parse(formula)
                ast = parse_ast(formula_parse)
            except FormulaParseError:
                cache[addr] = CellType(kind=CellKind.ANY)
                return
            val = evaluate_expr(
                ast,
                get_cell_value=get_cell_value,
                max_depth=limits.max_depth,
            )
            if isinstance(val, Unsupported):
                continue
            if isinstance(val, XlError):
                continue
            result_values.add(val)
        if not result_values:
            cache[addr] = CellType(kind=CellKind.ANY)
            return
        cache[addr] = _values_to_cell_type(result_values)

    def _enter_cell(
        addr: str,
    ) -> tuple[str, set[str], AstNode | None, AstNode | None] | None:
        """Type `addr` or return a finish payload when deps still need walking.

        Returns None when `addr` is already typed (or a cycle back-edge).
        """
        nonlocal _calls
        norm = normalize_cell_type_env_key(addr)
        if norm in cache:
            _propagate_consumed_leaves_to_ancestors(norm)
            return None
        _calls += 1
        _note_progress(addr)
        if norm in in_progress:
            return None
        ct_resolved = _lookup_cell_type(leaf_env, addr)
        if ct_resolved is not None:
            cache[norm] = ct_resolved
            if _track_consumed:
                _consumed_leaves.setdefault(norm, set()).add(norm)
                _record_consumed_leaf(norm)
            return None
        in_progress.add(norm)
        formula = get_cell_formula(addr)
        stored_ast: AstNode | None = None
        ast_root: AstNode | None = None
        keep_open = False
        try:
            if formula is None:
                if blank_rects and address_in_blank_ranges(addr, blank_rects):
                    cache[norm] = _BLANK_RANGE_LEAF_TYPE
                    if _track_consumed:
                        _consumed_leaves.setdefault(norm, set()).add(norm)
                        _record_consumed_leaf(norm)
                    return None
                raise DynamicRefError(
                    f"Missing constraint for leaf {addr!r} that feeds OFFSET/INDIRECT. "
                    "Add constraints only for leaf cells (non-formula) in the argument subgraph."
                )
            if get_cell_ast is not None:
                fetched = get_cell_ast(addr)
                stored_ast = fetched if isinstance(fetched, AstNode) else None
            _persistent_result = _try_persistent_lookup(norm, formula, stored_ast)
            if _persistent_result is not None:
                cached_ct, cached_consumed = _persistent_result
                cache[norm] = cached_ct
                _consumed_leaves[norm] = set(cached_consumed)
                _loaded_from_persistent.add(norm)
                _propagate_consumed_leaves_to_ancestors(norm)
                return None
            ast_root = stored_ast
            if ast_root is None:
                try:
                    formula_parse = _formula_to_parse(formula)
                    ast_root = parse_ast(formula_parse)
                except FormulaParseError:
                    ast_root = None
            elif stored_ast is not None:
                ast_root = bind_axes(stored_ast, addr)
            if len(_analysis_stack) >= _MAX_ANALYSIS_DEPTH:
                raise DynamicRefError(
                    f"Argument-subgraph analysis exceeded the maximum depth of "
                    f"{_MAX_ANALYSIS_DEPTH} while inferring the type of {addr!r}. "
                    "The dynamic-ref argument chain is too deeply nested to analyse; "
                    "constrain an intermediate cell in the chain to cut it short, or "
                    "simplify the chain."
                )
            _analysis_stack.append(norm)
            refs = get_refs_from_formula(formula, _sheet_from_addr(addr))
            if ast_root is not None:
                refs |= _collect_static_addresses_from_ast(
                    ast_root, max_range_cells=max_range_cells
                )
            if not refs:
                if ast_root is None:
                    cache[addr] = CellType(kind=CellKind.ANY)
                    return None
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
                    return None
                if val is None or isinstance(val, XlError):
                    cache[addr] = CellType(kind=CellKind.ANY)
                    return None
                cache[addr] = _values_to_cell_type({val})
                return None
            keep_open = True
            return (formula, refs, stored_ast, ast_root)
        finally:
            if not keep_open:
                _exit_cell(addr, formula, stored_ast)

    def cell_type_for(addr: str) -> CellType:
        """Type `addr` with an explicit worklist (issue #716).

        Argument subgraphs in LIC-DSF-scale workbooks are hundreds of cells
        deep. Walking them with Python recursion blew the CPython stack; this
        worklist keeps `_analysis_stack` as a logical depth counter only.
        """
        stack: list[tuple[str, str]] = [("enter", addr)]
        finish_payload: dict[str, tuple[str, set[str], AstNode | None, AstNode | None]] = {}
        while stack:
            phase, current = stack.pop()
            if phase == "enter":
                payload = _enter_cell(current)
                if payload is None:
                    continue
                formula, refs, stored_ast, ast_root = payload
                finish_payload[current] = (formula, refs, stored_ast, ast_root)
                stack.append(("finish", current))
                # Push larger keys first so sorted order is popped first, matching
                # the historical `_resolve_ref_types(sorted(refs))` walk. Re-push
                # uncached deps even if another frame already queued them so they
                # sit above this finish and are typed before inference.
                for dep in reversed(sorted(refs)):
                    if dep not in cache:
                        stack.append(("enter", dep))
                continue
            formula, refs, stored_ast, ast_root = finish_payload.pop(current)
            try:
                _infer_cell_type(current, formula, ast_root, refs)
            finally:
                _exit_cell(current, formula, stored_ast)
        if addr in cache:
            return cache[addr]
        return CellType(kind=CellKind.ANY)

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
    return backing


def dynamic_ref_selectors_boundable_without_expand(
    formula: str,
    *,
    current_sheet: str,
    limits: DynamicRefLimits | None = None,
    current_row: int | None = None,
    current_col: int | None = None,
    ast: AstNode | None = None,
) -> bool:
    """Return True when INDEX/OFFSET selectors need no argument-env expansion.

    `MATCH` over a static rectangular lookup, `ROWS`/`COLUMNS`, and numeric
    literals get integer domains from range geometry without reading
    `cell_type_env`. An omitted INDEX axis (`EmptyArg`, including
    `INDEX(array,,k)` and `INDEX(array,k,)`) is that same geometry when another
    selector densifies. Excel's literal `0` whole-axis form is already a
    numeric domain. Nested static INDEX slices are rewritten to rectangular
    refs first, so `INDEX(INDEX(array,,k), MATCH(...))` is checked as a range
    array. `INDIRECT`, cell selectors, and OFFSET height/width that cannot be
    densified from geometry still need expand.
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return False
    if "INDEX" in formula.upper():
        narrowed = narrow_static_index_lookup_vectors(formula, current_sheet)
        if narrowed and not narrowed.startswith("="):
            narrowed = "=" + narrowed
        if narrowed != formula:
            formula = narrowed
            ast = None
    lim = limits or DynamicRefLimits()
    eval_context = (
        {"row": current_row, "column": current_col}
        if current_row is not None and current_col is not None
        else None
    )
    empty_env: CellTypeEnv = {}
    try:
        root = ast if ast is not None else parse_ast(formula)
    except FormulaParseError:
        return False

    found = False

    def offset_base_ok(node: AstNode) -> bool:
        if isinstance(node, (CellRefNode, RangeNode)):
            return True
        return isinstance(node, FunctionCallNode) and node.name.upper() == "INDEX"

    def selector_ok(node: AstNode, *, for_offset: bool) -> bool:
        if isinstance(node, EmptyArgNode):
            return False
        if for_offset:
            return (
                _infer_offset_scalar_domains_core(
                    node,
                    empty_env,
                    lim,
                    eval_context,
                    current_sheet=current_sheet,
                )
                is not None
            )
        return (
            _infer_numeric_domain(
                node,
                empty_env,
                lim,
                context=eval_context,
                current_sheet=current_sheet,
            )
            is not None
        )

    def visit(node: AstNode) -> bool:
        nonlocal found
        if isinstance(node, FunctionCallNode):
            name = node.name.upper()
            if name == "INDIRECT":
                found = True
                return False
            if name == "INDEX":
                found = True
                if len(node.args) < 2 or len(node.args) > 3:
                    return False
                if not isinstance(node.args[0], (CellRefNode, RangeNode)):
                    return False
                selectors = node.args[1:]
                # Every selector omitted is the whole array. Infer does not
                # treat that form as a static slice, so keep the expand path.
                if not any(not isinstance(sel, EmptyArgNode) for sel in selectors):
                    return False
                for sel in selectors:
                    if isinstance(sel, EmptyArgNode):
                        continue
                    if not selector_ok(sel, for_offset=False):
                        return False
            elif name == "OFFSET":
                found = True
                if len(node.args) < 3 or len(node.args) > 5:
                    return False
                if not offset_base_ok(node.args[0]):
                    return False
                for sel in node.args[1:]:
                    if not selector_ok(sel, for_offset=True):
                        return False
            return all(visit(arg) for arg in node.args)
        if isinstance(node, BinaryOpNode):
            return visit(node.left) and visit(node.right)
        if isinstance(node, UnaryOpNode):
            return visit(node.operand)
        return True

    ok = visit(root)
    return found and ok


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
    allow_wide_bounds: bool = False,
) -> set[str]:
    """Infer the union of all possible OFFSET targets for a formula.

    This helper is intentionally focused and conservative:
    - Only OFFSET calls are analysed (INDIRECT is currently ignored).
    - Arguments may use a small Excel expression subset supported by
      `core.expr_eval.evaluate_expr`.
    - Leaf cells referenced by OFFSET/INDEX arguments must have a numeric
      domain in `cell_type_env` unless they appear only in ref_only
      argument positions (see `excel_grapher.core.excel_function_meta`).
    - Integer interval domains must be finite and small enough to enumerate,
      unless `allow_wide_bounds` is True. In that case intervals wider than
      `max_branches` become a bounding rectangle of possible OFFSET results
      instead of raising `DynamicRefError`, so constraint-candidate scanning
      can still reach downstream leaves.
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
            allow_wide_bounds=allow_wide_bounds,
        )
        out |= targets
        if len(out) > lim.max_cells:
            _raise_cell_limit(len(out), lim.max_cells, what="Dynamic ref cells")

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
            _raise_cell_limit(len(out), lim.max_cells, what="Dynamic ref cells")

    _emit_trace(
        DynamicRefTraceEvent(
            kind="infer",
            name="infer_dynamic_index_targets",
            elapsed_s=time.perf_counter() - t0,
            detail={"targets": len(out), "formula": formula, "current_sheet": current_sheet},
        )
    )
    return out


def _parse_index_axis(expr: str) -> AstNode:
    """Parse one INDEX row or column selector.

    A blank selector is Excel's whole-axis form (`0`). Two-argument INDEX
    never calls this for the column; that form still defaults to column 1.
    """
    if not expr.strip():
        return NumberNode(0)
    return parse_ast("=" + expr)


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
    args = _split_top_level_args(inner_args, keep_empty=True)
    if args is None or len(args) < 2 or len(args) > 3:
        raise DynamicRefError("INDEX expects 2 or 3 arguments (array, row_num, [column_num])")

    named_ranges = named_ranges or {}
    named_range_ranges = named_range_ranges or {}

    array_expr = _qualify_fragment(args[0], named_ranges, named_range_ranges)
    row_expr = _qualify_fragment(args[1], named_ranges, named_range_ranges)
    has_col = len(args) >= 3
    col_expr = _qualify_fragment(args[2], named_ranges, named_range_ranges) if has_col else ""

    try:
        array_ast = parse_ast("=" + array_expr)
        array_range = _base_to_range(array_ast, current_sheet=current_sheet)
    except (DynamicRefError, FormulaParseError) as exc:
        raise DynamicRefError(f"INDEX array argument must be a static range: {exc}") from exc

    row_ast = _parse_index_axis(row_expr)
    col_ast = _parse_index_axis(col_expr) if has_col else None

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
    allow_wide_bounds: bool = False,
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
                                    _raise_cell_limit(
                                        len(targets),
                                        limits.max_cells,
                                        what="Dynamic ref cells from single OFFSET call",
                                    )
        return targets

    if allow_wide_bounds:
        rows_dom = _infer_numeric_domain(
            rows_ast,
            cell_type_env,
            limits,
            context=eval_context,
            current_sheet=current_sheet,
        )
        cols_dom = _infer_numeric_domain(
            cols_ast,
            cell_type_env,
            limits,
            context=eval_context,
            current_sheet=current_sheet,
        )
        height_dom = (
            None
            if height_ast is None
            else _infer_numeric_domain(
                height_ast,
                cell_type_env,
                limits,
                context=eval_context,
                current_sheet=current_sheet,
            )
        )
        width_dom = (
            None
            if width_ast is None
            else _infer_numeric_domain(
                width_ast,
                cell_type_env,
                limits,
                context=eval_context,
                current_sheet=current_sheet,
            )
        )
        height_ok = height_ast is None or height_dom is not None
        width_ok = width_ast is None or width_dom is not None
        if rows_dom is not None and cols_dom is not None and height_ok and width_ok:
            return _emit_offset_targets_from_domains(
                base_ranges,
                rows_dom,
                cols_dom,
                height_dom,
                width_dom,
                bounds=bounds,
                limits=limits,
            )

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
                    _raise_cell_limit(
                        len(targets),
                        limits.max_cells,
                        what="Dynamic ref cells from single OFFSET call",
                    )

    return targets


def _emit_offset_targets_from_domains(
    base_ranges: list[ExcelRange],
    rows_dom: _FiniteInts | _IntBounds,
    cols_dom: _FiniteInts | _IntBounds,
    height_dom: _FiniteInts | _IntBounds | None,
    width_dom: _FiniteInts | _IntBounds | None,
    *,
    bounds: WorkbookBoundsProtocol | None,
    limits: DynamicRefLimits,
) -> set[str]:
    """Emit OFFSET targets from numeric domains without enumerating assignments.

    Wide integer intervals become the bounding rectangle of possible OFFSET
    results. That is exact for contiguous offset intervals of a fixed-size base
    and a conservative over-approximation when a finite enum has holes.
    """
    rb = _normalize_to_bounds(rows_dom)
    cb = _normalize_to_bounds(cols_dom)
    targets: set[str] = set()
    for base_range in base_ranges:
        base_bounds = _bounds_for_sheet(bounds, sheet=base_range.sheet)
        base_h = base_range.end_row - base_range.start_row + 1
        base_w = base_range.end_col - base_range.start_col + 1
        if height_dom is None:
            h_hi = base_h
        else:
            h_hi = _normalize_to_bounds(height_dom).hi
            if h_hi < 1:
                continue
        if width_dom is None:
            w_hi = base_w
        else:
            w_hi = _normalize_to_bounds(width_dom).hi
            if w_hi < 1:
                continue
        start_row = max(base_range.start_row + rb.lo, base_bounds.min_row)
        start_col = max(base_range.start_col + cb.lo, base_bounds.min_col)
        end_row = min(base_range.start_row + rb.hi + h_hi - 1, base_bounds.max_row)
        end_col = min(base_range.start_col + cb.hi + w_hi - 1, base_bounds.max_col)
        if start_row > end_row or start_col > end_col:
            continue
        n_cells = (end_row - start_row + 1) * (end_col - start_col + 1)
        if len(targets) + n_cells > limits.max_cells:
            _raise_cell_limit(
                len(targets) + n_cells,
                limits.max_cells,
                what="Dynamic ref cells from single OFFSET call",
            )
        rng = ExcelRange(
            sheet=base_range.sheet,
            start_row=start_row,
            start_col=start_col,
            end_row=end_row,
            end_col=end_col,
        )
        targets.update(rng.cell_addresses())
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


@dataclass(frozen=True, slots=True)
class _IndexAxisEnv:
    """Type environment used to densify non-literal INDEX row/col selectors."""

    env: CellTypeEnv
    limits: DynamicRefLimits
    context: dict[str, int]
    current_sheet: str
    depth: int


def _singleton_int(domain: _FiniteInts | _IntBounds | None) -> int | None:
    """Return the only integer in `domain` when it is a singleton."""
    if isinstance(domain, _FiniteInts):
        if len(domain.values) != 1:
            return None
        return next(iter(domain.values))
    if isinstance(domain, _IntBounds) and domain.lo == domain.hi:
        return domain.lo
    return None


def _resolved_index_axis(node: AstNode, axes: _IndexAxisEnv | None) -> int | None:
    """Return a singleton INDEX axis, including Excel's `0` whole-axis form.

    Literals resolve without `axes`. Non-literals densify only when `axes` is
    provided and numeric inference yields exactly one integer.
    """
    if _is_literal_zero(node):
        return 0
    literal = _literal_positive_int(node)
    if literal is not None:
        return literal
    if axes is None:
        return None
    result = _infer_numeric_domain_result(
        node,
        axes.env,
        axes.limits,
        context=axes.context,
        current_sheet=axes.current_sheet,
        depth=axes.depth + 1,
    )
    if result.diagnostic is not None:
        return None
    resolved = _singleton_int(result.domain)
    if resolved is None or resolved < 0:
        return None
    return resolved


def _index_vector_from_bounds(
    bounds: tuple[str, int, int, int, int],
    row_arg: AstNode,
    col_arg: AstNode | None,
    *,
    axes: _IndexAxisEnv | None = None,
) -> AstNode | None:
    """Map INDEX row/col selectors over static bounds to a rectangular result range.

    Supports omitted axes, singleton positive selectors (literals or densified
    under `axes`), and Excel's `0` form that returns an entire row/column/array.
    Two-arg `INDEX(array, k)` indexes a 1-D vector, and does not resolve when
    `array` has more than one row and column (Excel `#REF!`). A trailing empty
    column (`INDEX(array, k,)`) is the 2-D form and selects that row.
    """
    sheet, rlo, rhi, clo, chi = bounds
    nrows = rhi - rlo + 1
    ncols = chi - clo + 1

    row_omitted = isinstance(row_arg, EmptyArgNode)
    col_missing = col_arg is None
    col_empty = isinstance(col_arg, EmptyArgNode)
    col_omitted = col_missing or col_empty
    row_sel = None if row_omitted else _resolved_index_axis(row_arg, axes)
    col_sel = None if col_arg is None or col_empty else _resolved_index_axis(col_arg, axes)
    row_zero = row_sel == 0
    col_zero = col_sel == 0

    if row_omitted and col_omitted:
        return None

    if row_omitted:
        if col_zero:
            return _range_node_from_bounds(sheet, rlo, rhi, clo, chi)
        if col_sel is None or col_sel > ncols:
            return None
        c = clo + col_sel - 1
        return _range_node_from_bounds(sheet, rlo, rhi, c, c)

    if col_omitted:
        if row_zero:
            return _range_node_from_bounds(sheet, rlo, rhi, clo, chi)
        if row_sel is None:
            return None
        # 2-arg INDEX(array, k) indexes a 1-D vector. A trailing empty
        # column (`INDEX(array, k,)`) is the 2-D form and selects that row.
        if col_missing and nrows == 1:
            if row_sel > ncols:
                return None
            c = clo + row_sel - 1
            return _range_node_from_bounds(sheet, rlo, rhi, c, c)
        if col_missing and ncols == 1:
            if row_sel > nrows:
                return None
            r = rlo + row_sel - 1
            return _range_node_from_bounds(sheet, r, r, clo, chi)
        # Two-arg INDEX on a 2-D block is #REF! in Excel.
        if col_missing:
            return None
        if row_sel > nrows:
            return None
        r = rlo + row_sel - 1
        return _range_node_from_bounds(sheet, r, r, clo, chi)

    # Both axes present: 0 selects the full opposite axis (Excel INDEX).
    if row_zero and col_zero:
        return _range_node_from_bounds(sheet, rlo, rhi, clo, chi)
    if row_zero:
        if col_sel is None or col_sel > ncols:
            return None
        c = clo + col_sel - 1
        return _range_node_from_bounds(sheet, rlo, rhi, c, c)
    if col_zero:
        if row_sel is None or row_sel > nrows:
            return None
        r = rlo + row_sel - 1
        return _range_node_from_bounds(sheet, r, r, clo, chi)
    return None


def _index_vector_lookup_array(
    node: AstNode,
    *,
    allow_shape_preserving: bool = False,
    axes: _IndexAxisEnv | None = None,
) -> AstNode | None:
    """Resolve INDEX forms to a rectangular range for MATCH lookup geometry.

    Handles `INDEX(range,,k)` / `INDEX(range,k[,])` and Excel's `INDEX(...,0[,k])`
    whole-axis form. `k` may be a literal or, when `axes` is provided, any
    selector that densifies to a singleton integer. When
    `allow_shape_preserving` is True, also accepts array expressions that keep
    a static rectangle's shape (e.g. `(range<>0)`).

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
    return _index_vector_from_bounds(bounds, node.args[1], col_arg, axes=axes)


def _span_strictly_inside(inner: tuple[int, int], outer: tuple[int, int]) -> bool:
    """Return True when `inner` lies strictly inside `outer`."""
    return (
        outer[0] <= inner[0]
        and inner[1] <= outer[1]
        and (outer[0] < inner[0] or inner[1] < outer[1])
    )


def _static_index_vector_text(call_text: str, current_sheet: str) -> str | None:
    """Return the rectangular ref for a static INDEX vector, or None."""
    anchor = format_key(current_sheet, "A1") if current_sheet else None
    try:
        node = parse_ast(call_text, anchor=anchor)
    except (FormulaParseError, ValueError):
        return None
    vector = _index_vector_lookup_array(node)
    if isinstance(vector, CellRefNode):
        return vector.address
    if not isinstance(vector, RangeNode):
        return None
    try:
        sheet, start = parse_address(vector.start)
        end_sheet, end = parse_address(vector.end)
    except ValueError:
        return None
    if sheet != end_sheet:
        return None
    return format_range_key(sheet, start, end)


@lru_cache(maxsize=4096)
def narrow_static_index_lookup_vectors(formula: str, current_sheet: str = "") -> str:
    """Replace static INDEX row/column slices with the cells they select.

    `INDEX(array,,k)` and `INDEX(array,k,)` (including Excel's `0` whole-axis
    form) read only that vector. Callers that expand every range in an INDEX
    selector otherwise pull the whole rectangle into the dependency graph.
    Dynamic selectors are left unchanged. Nested static slices rewrite from
    the inside out.

    Args:
        formula: Formula or fragment, with or without a leading `=`.
        current_sheet: Sheet used to qualify unqualified ranges inside `INDEX`.

    Returns:
        `formula` with each resolved static INDEX vector replaced by its
        rectangular reference.
    """
    if "INDEX" not in formula.upper():
        return formula
    current = formula
    for _ in range(16):
        calls = _find_function_calls_with_spans(current, frozenset({"INDEX"}), include_nested=True)
        if not calls:
            return current
        spans = [span for _fn, _inner, span in calls]
        replacements: list[tuple[tuple[int, int], str]] = []
        for _fn, _inner, span in calls:
            # Rewrite innermost calls first so an outer INDEX sees the vector.
            if any(_span_strictly_inside(other, span) for other in spans):
                continue
            call_text = current[span[0] : span[1]]
            replacement = _static_index_vector_text(call_text, current_sheet)
            if replacement is None or replacement == call_text:
                continue
            replacements.append((span, replacement))
        if not replacements:
            return current
        pieces: list[str] = []
        cursor = 0
        for (start, end), replacement in sorted(replacements, key=lambda item: item[0][0]):
            pieces.append(current[cursor:start])
            pieces.append(replacement)
            cursor = end
        pieces.append(current[cursor:])
        current = "".join(pieces)
    return current


def prepare_dynamic_selector_expr(formula: str, *, current_sheet: str) -> str:
    """Mask address-only calls, then narrow static INDEX lookup vectors.

    Selector text is expanded into value dependencies. `ROW`/`COLUMN`/`ROWS`/
    `COLUMNS` contribute no values, and a static `INDEX` slice contributes
    only its vector.
    """
    return narrow_static_index_lookup_vectors(
        mask_ref_only_function_calls(formula),
        current_sheet=current_sheet,
    )


def _index_dynamic_axis_match_extent(node: AstNode) -> int | None:
    """Return MATCH lookup length when INDEX shape is known but an axis is dynamic.

    An omitted or `0` axis still yields a vector of known length. A specific
    (unknown) axis over a 1-D array, or both specific axes over a 2-D array,
    yields a scalar. Two-arg `INDEX` on a 2-D block has no lookup (`#REF!`).
    """
    if not isinstance(node, FunctionCallNode) or node.name.upper() != "INDEX":
        return None
    if len(node.args) < 2:
        return None
    bounds = _array_expr_static_bounds(node.args[0])
    if bounds is None:
        return None
    _sheet, rlo, rhi, clo, chi = bounds
    nrows = rhi - rlo + 1
    ncols = chi - clo + 1
    row_arg = node.args[1]
    col_arg: AstNode | None = node.args[2] if len(node.args) >= 3 else None
    row_omitted = isinstance(row_arg, EmptyArgNode)
    col_missing = col_arg is None
    col_empty = isinstance(col_arg, EmptyArgNode)
    row_zero = _is_literal_zero(row_arg)
    col_zero = col_arg is not None and _is_literal_zero(col_arg)
    if row_omitted and (col_missing or col_empty):
        return None
    if (row_omitted or row_zero) and (col_missing or col_empty or col_zero):
        return nrows * ncols
    if row_omitted or row_zero:
        return nrows
    if col_missing:
        if nrows == 1 or ncols == 1:
            return 1
        return None
    if col_empty or col_zero:
        return ncols
    return 1


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
    return _index_dynamic_axis_match_extent(node)


def _ordered_match_lookup_cells(
    arg: AstNode,
    *,
    current_sheet: str,
    env: CellTypeEnv | None = None,
    limits: DynamicRefLimits | None = None,
    context: dict[str, int] | None = None,
    depth: int = 0,
) -> list[str] | None:
    """Return sheet-qualified addresses for a one-dimensional MATCH lookup_array (in scan order).

    When `env` and `limits` are provided, INDEX axes that densify to a singleton
    integer are treated like literals.
    """
    from fastpyxl.utils.cell import get_column_letter

    if isinstance(arg, CellRefNode):
        addr = arg.address
        if "!" in addr:
            return [addr]
        return [format_key(current_sheet, addr)]

    if isinstance(arg, RangeNode):
        try:
            # `RangeNode.start` is already quoted when the sheet needs it.
            # Re-quoting that text would miss `CellTypeEnv` keys.
            sheet, coord_start = parse_address(arg.start)
            end_sheet, coord_end = parse_address(arg.end)
        except ValueError:
            return None
        if sheet != end_sheet:
            return None
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
    axes = (
        _IndexAxisEnv(
            env=env,
            limits=limits,
            context=context or {},
            current_sheet=current_sheet,
            depth=depth,
        )
        if env is not None and limits is not None
        else None
    )
    vector = _index_vector_lookup_array(arg, axes=axes)
    if vector is not None:
        return _ordered_match_lookup_cells(
            vector,
            current_sheet=current_sheet,
            env=env,
            limits=limits,
            context=context,
            depth=depth,
        )
    return None


def _exact_match_lookup_scan_limit(limits: DynamicRefLimits) -> int:
    """Return how many lookup cells exact MATCH may compare.

    `DynamicRefLimits.max_cells` caps emitted targets. A singleton needle can
    collapse a longer lookup to one position, so the scan budget stays at
    least the default target cap when the caller tightens `max_cells`.
    """
    return max(limits.max_cells, DynamicRefLimits().max_cells)


def _match_string_uses_wildcards(value: str) -> bool:
    """Return True when `value` is not a literal under Excel MATCH type 0.

    Unescaped `*` and `?` are wildcards. A `~` escape of `*`, `?`, or `~`
    also blocks refinement: this comparison does not unescape the needle.
    """
    index = 0
    while index < len(value):
        char = value[index]
        if char == "~" and index + 1 < len(value) and value[index + 1] in "*?~":
            return True
        if char in "*?":
            return True
        index += 1
    return False


def _finite_exact_match_values(node: AstNode, env: CellTypeEnv) -> frozenset[object] | None:
    """Return a finite equality domain for an exact MATCH needle, if known."""
    if isinstance(node, StringNode):
        return frozenset({node.value})
    if isinstance(node, CellRefNode):
        cell_type = _lookup_cell_type(env, node.address)
        if cell_type is None or cell_type.enum is None or not cell_type.enum.values:
            return None
        return cell_type.enum.values
    return None


def _needle_blocks_exact_match_refine(values: frozenset[object]) -> bool:
    return any(isinstance(value, str) and _match_string_uses_wildcards(value) for value in values)


def _value_may_equal_number(value: object) -> bool:
    """Return True when `value` can compare equal to a number under MATCH."""
    if isinstance(value, (int, float)):
        return True
    if isinstance(value, str):
        return try_coerce_string_to_float(value) is not None
    return False


def _string_values(values: frozenset[object]) -> frozenset[str] | None:
    """Return `values` when every member is a string."""
    strings: list[str] = []
    for value in values:
        if not isinstance(value, str):
            return None
        strings.append(value)
    return frozenset(strings)


_BOOL_AS_NUMBER_DOMAIN = _FiniteInts(frozenset({0, 1}))


class _ExactMatchVerdict(Enum):
    """How a lookup cell relates to an exact-MATCH needle."""

    MISS = "miss"
    MAYBE = "maybe"
    CERTAIN = "certain"


def _match_number(value: object) -> float | None:
    """Return the float exact MATCH uses for `value`.

    Non-numeric text yields `None`. Non-finite results such as `nan` and `inf`
    are left for the caller to reject.
    """
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        return try_coerce_string_to_float(value)
    return None


def _scalar_in_int_domain(value: object, domain: _FiniteInts | _IntBounds) -> bool:
    """Return True when `value` may equal an integer in `domain` under MATCH."""
    number = _match_number(value)
    if number is None or not math.isfinite(number):
        return False
    as_int = int(number)
    if as_int != number:
        return False
    if isinstance(domain, _FiniteInts):
        return as_int in domain.values
    return domain.lo <= as_int <= domain.hi


def _values_may_equal_numeric_domain(
    values: frozenset[object],
    domain: _FiniteInts | _IntBounds,
) -> bool:
    """Return True when any value may equal `domain` under exact MATCH."""
    return any(_scalar_in_int_domain(value, domain) for value in values)


def _int_domain_singleton(domain: _FiniteInts | _IntBounds) -> int | None:
    """Return the only integer in `domain` when it has exactly one."""
    if isinstance(domain, _FiniteInts):
        if len(domain.values) != 1:
            return None
        return next(iter(domain.values))
    if domain.lo == domain.hi:
        return domain.lo
    return None


def _values_vs_int_domain_verdict(
    values: frozenset[object],
    domain: _FiniteInts | _IntBounds,
) -> _ExactMatchVerdict:
    """Compare finite needle values with an integer domain."""
    if not _values_may_equal_numeric_domain(values, domain):
        return _ExactMatchVerdict.MISS
    singleton = _int_domain_singleton(domain)
    if singleton is None:
        return _ExactMatchVerdict.MAYBE
    only = _FiniteInts(frozenset({singleton}))
    if all(_scalar_in_int_domain(value, only) for value in values):
        return _ExactMatchVerdict.CERTAIN
    return _ExactMatchVerdict.MAYBE


def _enum_match_verdict(
    needles: frozenset[object],
    cell_values: frozenset[object],
    *,
    folded_string_needles: frozenset[str] | None,
) -> _ExactMatchVerdict:
    """Compare two finite equality sets under exact MATCH."""
    if not cell_values:
        return _ExactMatchVerdict.MISS
    if folded_string_needles is not None:
        cell_strings = _string_values(cell_values)
        if cell_strings is not None:
            cell_folded = {excel_casefold(value) for value in cell_strings}
            if cell_folded.isdisjoint(folded_string_needles):
                return _ExactMatchVerdict.MISS
            if len(cell_folded) == 1 and cell_folded == folded_string_needles:
                return _ExactMatchVerdict.CERTAIN
            return _ExactMatchVerdict.MAYBE
    certain = True
    any_hit = False
    for needle in needles:
        for cell in cell_values:
            if _values_match(needle, cell):
                any_hit = True
            else:
                certain = False
    if not any_hit:
        return _ExactMatchVerdict.MISS
    if certain:
        return _ExactMatchVerdict.CERTAIN
    return _ExactMatchVerdict.MAYBE


def _folded_string_needles(needles: frozenset[object]) -> frozenset[str] | None:
    """Return casefolded needles when every needle is a string."""
    strings = _string_values(needles)
    if strings is None:
        return None
    return frozenset(excel_casefold(value) for value in strings)


def _cell_may_equal_exact_match_values(
    cell_type: CellType | None,
    needles: frozenset[object],
    limits: DynamicRefLimits,
    *,
    folded_string_needles: frozenset[str] | None = None,
) -> _ExactMatchVerdict:
    """Return how `cell_type` compares with `needles`.

    `MAYBE` keeps the cell: its domain does not decide equality. `CERTAIN`
    means every allowed value equals every needle value, so MATCH cannot
    continue past it. `MISS` means every allowed value is unequal.
    """
    if cell_type is None:
        return _ExactMatchVerdict.MAYBE
    if cell_type.enum is not None:
        if not cell_type.enum.values:
            return _ExactMatchVerdict.MISS
        return _enum_match_verdict(
            needles,
            cell_type.enum.values,
            folded_string_needles=folded_string_needles,
        )
    if cell_type.kind in (CellKind.NUMBER, CellKind.ANY):
        numeric = _domain_from_cell_type(cell_type, limits)
        if numeric is not None:
            return _values_vs_int_domain_verdict(needles, numeric)
        if cell_type.kind is CellKind.NUMBER and not any(
            _value_may_equal_number(value) for value in needles
        ):
            return _ExactMatchVerdict.MISS
        return _ExactMatchVerdict.MAYBE
    if cell_type.kind is CellKind.ERROR:
        return _ExactMatchVerdict.MISS
    if cell_type.kind is CellKind.BOOL:
        if _values_may_equal_numeric_domain(needles, _BOOL_AS_NUMBER_DOMAIN):
            return _ExactMatchVerdict.MAYBE
        return _ExactMatchVerdict.MISS
    if cell_type.kind is CellKind.DATE and not any(
        _value_may_equal_number(value) for value in needles
    ):
        return _ExactMatchVerdict.MISS
    return _ExactMatchVerdict.MAYBE


def _numeric_domain_verdict(
    needle_dom: _FiniteInts | _IntBounds,
    cell_dom: _FiniteInts | _IntBounds,
) -> _ExactMatchVerdict:
    """Compare two integer domains under exact MATCH."""
    if not _domains_may_equal_exact_match(needle_dom, cell_dom):
        return _ExactMatchVerdict.MISS
    needle = _int_domain_singleton(needle_dom)
    cell = _int_domain_singleton(cell_dom)
    if needle is not None and needle == cell:
        return _ExactMatchVerdict.CERTAIN
    return _ExactMatchVerdict.MAYBE


def _cell_may_equal_numeric_needle(
    cell_type: CellType | None,
    needle_dom: _FiniteInts | _IntBounds,
    limits: DynamicRefLimits,
) -> _ExactMatchVerdict:
    """Return how `cell_type` compares with a numeric exact-MATCH needle.

    Missing domains stay possible. String enums that cannot coerce into
    `needle_dom` are misses. A singleton on both sides is a certain hit.
    """
    if cell_type is None:
        return _ExactMatchVerdict.MAYBE
    numeric = _domain_from_cell_type(cell_type, limits)
    if numeric is not None:
        return _numeric_domain_verdict(needle_dom, numeric)
    if cell_type.enum is not None:
        if not cell_type.enum.values:
            return _ExactMatchVerdict.MISS
        return _values_vs_int_domain_verdict(cell_type.enum.values, needle_dom)
    if cell_type.kind is CellKind.ERROR:
        return _ExactMatchVerdict.MISS
    if cell_type.kind is CellKind.BOOL:
        if _domains_may_equal_exact_match(needle_dom, _BOOL_AS_NUMBER_DOMAIN):
            return _ExactMatchVerdict.MAYBE
        return _ExactMatchVerdict.MISS
    return _ExactMatchVerdict.MAYBE


def _exact_match_position_domain(
    candidates: list[int],
    lookup_len: int,
) -> _FiniteInts | _IntBounds | None:
    """Return positions that may match, when that is stricter than `1..lookup_len`.

    `candidates` is ascending. Contiguous positions stay an interval so INDEX
    emission does not materialize each index. The empty set is not a refinement:
    a total miss is `#N/A` in Excel, and emitting no cells would drop every
    INDEX target if a cell were classified unequal by mistake. The full extent
    is omitted so the caller can keep the lookup bounds.
    """
    count = len(candidates)
    if count == 0 or count == lookup_len:
        return None
    first = candidates[0]
    last = candidates[-1]
    if last - first + 1 == count:
        return _IntBounds(first, last)
    return _FiniteInts(frozenset(candidates))


def _refine_exact_match_positions(
    ordered: list[str],
    verdict_of: Callable[[str], _ExactMatchVerdict],
) -> _FiniteInts | _IntBounds | None:
    """Drop lookup positions proven unequal to the needle.

    A certain hit ends the scan: MATCH returns the first match, so later
    cells are unreachable. One pass, bounded by the caller's scan limit.
    """
    candidates: list[int] = []
    for index, cell_addr in enumerate(ordered, start=1):
        verdict = verdict_of(cell_addr)
        if verdict is _ExactMatchVerdict.MISS:
            continue
        candidates.append(index)
        if verdict is _ExactMatchVerdict.CERTAIN:
            break
    return _exact_match_position_domain(candidates, len(ordered))


def _infer_enum_exact_match_position(
    needle: AstNode,
    ordered: list[str],
    env: CellTypeEnv,
    limits: DynamicRefLimits,
) -> _FiniteInts | _IntBounds | None:
    """Return lookup positions that may equal a non-numeric exact-MATCH needle."""
    needles = _finite_exact_match_values(needle, env)
    if needles is None or _needle_blocks_exact_match_refine(needles):
        return None
    folded = _folded_string_needles(needles)

    def verdict_of(cell_addr: str) -> _ExactMatchVerdict:
        return _cell_may_equal_exact_match_values(
            _lookup_cell_type(env, cell_addr),
            needles,
            limits,
            folded_string_needles=folded,
        )

    return _refine_exact_match_positions(ordered, verdict_of)


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
) -> _FiniteInts | _IntBounds | None:
    """Return exact-MATCH positions that are not proven unequal to the needle.

    Numeric needles compare through integer overlap. Other finite enums compare
    through exact MATCH equality. Cells with no deciding domain stay candidates,
    so one untyped lookup cell narrows the axis to that cell plus any real hits
    instead of the whole vector. A singleton is the common fully-typed case.

    Lookup addresses are materialized only after the needle has a finite
    domain and the lookup extent fits the scan budget. Callers that only need
    `IntBounds(1, n)` (an empty env, an unpinned needle) never walk the vector.
    """
    if len(node.args) < 2:
        return None
    extent = _static_match_lookup_extent(node.args[1])
    if extent is None or extent < 1 or extent > _exact_match_lookup_scan_limit(limits):
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
        needles = _finite_exact_match_values(node.args[0], env)
        if needles is None or _needle_blocks_exact_match_refine(needles):
            return None

    ordered = _ordered_match_lookup_cells(
        node.args[1],
        current_sheet=current_sheet,
        env=env,
        limits=limits,
        context=context,
        depth=depth,
    )
    if not ordered:
        return None
    if lookup_dom is None:
        return _infer_enum_exact_match_position(node.args[0], ordered, env, limits)

    def verdict_of(cell_addr: str) -> _ExactMatchVerdict:
        return _cell_may_equal_numeric_needle(
            _lookup_cell_type(env, cell_addr),
            lookup_dom,
            limits,
        )

    return _refine_exact_match_positions(ordered, verdict_of)


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
    return lookup_cell_type(env, address)


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


# Excel rejects text results longer than one cell. Checked before concatenation
# so an oversized fragment fails closed without building the product.
_EXCEL_CELL_TEXT_LIMIT = 32_767

# VALUE of an already-numeric expression is the identity. These calls never
# produce digit-text that would need value_from_text.
_VALUE_NUMERIC_IDENTITY_FUNCS = frozenset(
    {
        "ABS",
        "COLUMN",
        "COLUMNS",
        "EXP",
        "ISNUMBER",
        "MATCH",
        "MAX",
        "MIN",
        "ROW",
        "ROWS",
        "SUM",
        "VALUE",
    }
)


def _integer_from_excel_text(text: str) -> int | None:
    """Return the integer `VALUE` of `text`.

    Uses `value_from_text` so analysis matches the evaluator. Non-integral
    results (`12.5`, `#VALUE!`) fail closed.
    """
    parsed = value_from_text(text)
    if isinstance(parsed, bool) or not isinstance(parsed, (int, float)):
        return None
    if isinstance(parsed, float):
        if not math.isfinite(parsed) or not parsed.is_integer():
            return None
        return int(parsed)
    return parsed


def _numeric_domain_from_text_fragments(
    texts: frozenset[str] | None,
) -> _NumericDomainInferenceResult:
    """Map an exact text set through `VALUE`.

    `texts is None` means the text is unknown. A known set that is not
    uniformly integral also yields no numeric domain: dropping the failures
    would under-approximate a MATCH needle.
    """
    if texts is None:
        return _domain_result(None)
    if not texts:
        return _domain_result(_FiniteInts(frozenset()))
    values: set[int] = set()
    for text in texts:
        parsed = _integer_from_excel_text(text)
        if parsed is None:
            return _domain_result(None)
        values.add(parsed)
    return _domain_result(_FiniteInts(frozenset(values)))


def _integer_from_canonical_excel_text(text: str) -> int | None:
    """Return the integer when `text` is already its general-format spelling.

    `VALUE` may parse `"05"` or `"12.0"`. The `&` operator returns that text,
    so a later concatenation must see the original spelling. Caching the
    parsed integer would re-stringify it as `"5"` or `"12"`.
    """
    parsed = _integer_from_excel_text(text)
    if parsed is None:
        return None
    if to_string(parsed) != text:
        return None
    return parsed


def _numeric_domain_from_canonical_text_fragments(
    texts: frozenset[str] | None,
) -> _NumericDomainInferenceResult:
    """Integer domain of concat text that round-trips through general format.

    One non-canonical fragment fails the whole set. Keeping only the
    canonical members would under-approximate a MATCH needle.
    """
    if texts is None:
        return _domain_result(None)
    if not texts:
        return _domain_result(_FiniteInts(frozenset()))
    values: set[int] = set()
    for text in texts:
        parsed = _integer_from_canonical_excel_text(text)
        if parsed is None:
            return _domain_result(None)
        values.add(parsed)
    return _domain_result(_FiniteInts(frozenset(values)))


def _first_operand_diagnostic(
    operands: Sequence[AstNode],
    env: CellTypeEnv,
    limits: DynamicRefLimits,
    *,
    context: dict[str, int] | None,
    current_sheet: str,
    depth: int,
) -> _NumericDomainInferenceResult | None:
    """Return the first operand diagnostic, so concat does not hide it."""
    for operand in operands:
        result = _infer_numeric_domain_result(
            operand,
            env,
            limits,
            context=context,
            current_sheet=current_sheet,
            depth=depth + 1,
        )
        if result.diagnostic is not None:
            return result
    return None


def _join_text_fragment_sets(
    left: frozenset[str],
    right: frozenset[str],
    limits: DynamicRefLimits,
) -> frozenset[str] | None:
    """Concatenate two exact text sets, or return None past the branch cap.

    The product size is checked before any strings are allocated. The result
    stays a set: concatenated integers are sparse, so collapsing to a bounding
    interval would be unsound for MATCH.
    """
    if not left or not right:
        return frozenset()
    if len(left) * len(right) > limits.max_branches:
        return None
    if max(len(text) for text in left) + max(len(text) for text in right) > _EXCEL_CELL_TEXT_LIMIT:
        return None
    return frozenset(left_text + right_text for left_text in left for right_text in right)


def _union_text_fragment_sets(
    parts: Sequence[frozenset[str]],
    limits: DynamicRefLimits,
) -> frozenset[str] | None:
    """Union exact text sets, stopping once the cap is exceeded."""
    merged: set[str] = set()
    for part in parts:
        merged.update(part)
        if len(merged) > limits.max_branches:
            return None
    return frozenset(merged)


def _enum_number_fragment(value: object) -> str | None:
    """General-format text of one integral enum member, if it is one."""
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return to_string(value)
    if isinstance(value, float) and math.isfinite(value) and value.is_integer():
        return to_string(value)
    return None


def _cell_text_fragments(ct: CellType | None, limits: DynamicRefLimits) -> frozenset[str] | None:
    """Exact general-format text of a cell, when the domain is small enough.

    Wide intervals are not densified: concatenation is not interval arithmetic,
    and stringifying every integer would dominate analysis.
    """
    if ct is None:
        return None
    if ct.enum is not None:
        values = ct.enum.values
        if len(values) > limits.max_branches:
            return None
        if not values:
            return frozenset()
        if all(isinstance(value, str) for value in values):
            return frozenset(cast(str, value) for value in values)
        fragments: list[str] = []
        for value in values:
            fragment = _enum_number_fragment(value)
            if fragment is None:
                return None
            fragments.append(fragment)
        return frozenset(fragments)
    if ct.kind not in (CellKind.NUMBER, CellKind.ANY) or ct.interval is None:
        return None
    if ct.interval.min is None or ct.interval.max is None:
        return None
    lo = int(ct.interval.min)
    hi = int(ct.interval.max)
    if hi < lo:
        return frozenset()
    if hi - lo + 1 > limits.max_branches:
        return None
    return frozenset(to_string(value) for value in range(lo, hi + 1))


def _domain_provably_true(domain: _FiniteInts | _IntBounds) -> bool:
    if isinstance(domain, _FiniteInts):
        return bool(domain.values) and all(value != 0 for value in domain.values)
    return domain.lo > 0 or domain.hi < 0


def _domain_provably_false(domain: _FiniteInts | _IntBounds) -> bool:
    if isinstance(domain, _FiniteInts):
        return bool(domain.values) and all(value == 0 for value in domain.values)
    return domain.lo == domain.hi == 0


def _infer_text_fragments(
    node: AstNode,
    env: CellTypeEnv,
    limits: DynamicRefLimits,
    *,
    context: dict[str, int] | None = None,
    current_sheet: str = "",
    depth: int = 0,
) -> frozenset[str] | None:
    """Return the exact strings `node` can produce, or None if unknown.

    None is fail-closed. Callers that turn the set into a MATCH needle must
    not substitute an over-approximation: one extra integer can collapse onto
    the wrong row.
    """
    if depth > limits.max_depth:
        return None
    if isinstance(node, StringNode):
        if len(node.value) > _EXCEL_CELL_TEXT_LIMIT:
            return None
        return frozenset({node.value})
    if isinstance(node, (NumberNode, BoolNode)):
        return frozenset({to_string(node.value)})
    if isinstance(node, (ErrorNode, RangeNode, EmptyArgNode)):
        return None
    if isinstance(node, CellRefNode):
        return _cell_text_fragments(_lookup_cell_type(env, node.address), limits)
    if isinstance(node, BinaryOpNode) and node.op == "&":
        return _concat_text_fragments(
            (node.left, node.right),
            env,
            limits,
            context=context,
            current_sheet=current_sheet,
            depth=depth,
        )
    if isinstance(node, FunctionCallNode):
        name = node.name.upper()
        if name == "IF":
            return _if_text_fragments(
                node,
                env,
                limits,
                context=context,
                current_sheet=current_sheet,
                depth=depth,
            )
        if name == "CHOOSE":
            return _choose_text_fragments(
                node,
                env,
                limits,
                context=context,
                current_sheet=current_sheet,
                depth=depth,
            )
        if name in {"CONCAT", "CONCATENATE"}:
            if not node.args:
                return None
            return _concat_text_fragments(
                node.args,
                env,
                limits,
                context=context,
                current_sheet=current_sheet,
                depth=depth,
            )
    numeric = _infer_numeric_domain_result(
        node,
        env,
        limits,
        context=context,
        current_sheet=current_sheet,
        depth=depth + 1,
    )
    domain = numeric.domain
    if isinstance(domain, _FiniteInts) and len(domain.values) <= limits.max_branches:
        return frozenset(to_string(value) for value in domain.values)
    return None


def _concat_text_fragments(
    parts: Sequence[AstNode],
    env: CellTypeEnv,
    limits: DynamicRefLimits,
    *,
    context: dict[str, int] | None,
    current_sheet: str,
    depth: int,
) -> frozenset[str] | None:
    """Fold `&` / `CONCAT` left to right, capping the running product."""
    acc: frozenset[str] = frozenset({""})
    for part in parts:
        side = _infer_text_fragments(
            part,
            env,
            limits,
            context=context,
            current_sheet=current_sheet,
            depth=depth + 1,
        )
        if side is None:
            return None
        joined = _join_text_fragment_sets(acc, side, limits)
        if joined is None:
            return None
        acc = joined
    return acc


def _if_text_fragments(
    node: FunctionCallNode,
    env: CellTypeEnv,
    limits: DynamicRefLimits,
    *,
    context: dict[str, int] | None,
    current_sheet: str,
    depth: int,
) -> frozenset[str] | None:
    """Exact text of `IF`, dropping a branch only when the condition proves it dead."""
    if len(node.args) < 2:
        return None
    cond = _infer_numeric_domain_result(
        node.args[0],
        env,
        limits,
        context=context,
        current_sheet=current_sheet,
        depth=depth + 1,
    )
    if cond.diagnostic is not None or cond.domain is None:
        return None
    if isinstance(cond.domain, _FiniteInts) and not cond.domain.values:
        return frozenset()
    if _domain_provably_true(cond.domain):
        return _infer_text_fragments(
            node.args[1],
            env,
            limits,
            context=context,
            current_sheet=current_sheet,
            depth=depth + 1,
        )
    if _domain_provably_false(cond.domain):
        if len(node.args) < 3:
            return frozenset({"FALSE"})
        return _infer_text_fragments(
            node.args[2],
            env,
            limits,
            context=context,
            current_sheet=current_sheet,
            depth=depth + 1,
        )
    then_text = _infer_text_fragments(
        node.args[1],
        env,
        limits,
        context=context,
        current_sheet=current_sheet,
        depth=depth + 1,
    )
    else_text = (
        _infer_text_fragments(
            node.args[2],
            env,
            limits,
            context=context,
            current_sheet=current_sheet,
            depth=depth + 1,
        )
        if len(node.args) >= 3
        else frozenset({"FALSE"})
    )
    if then_text is None or else_text is None:
        return None
    return _union_text_fragment_sets((then_text, else_text), limits)


def _choose_text_fragments(
    node: FunctionCallNode,
    env: CellTypeEnv,
    limits: DynamicRefLimits,
    *,
    context: dict[str, int] | None,
    current_sheet: str,
    depth: int,
) -> frozenset[str] | None:
    """Exact text of the `CHOOSE` options selected by a known index domain."""
    if len(node.args) < 2:
        return None
    index = _infer_numeric_domain_result(
        node.args[0],
        env,
        limits,
        context=context,
        current_sheet=current_sheet,
        depth=depth + 1,
    )
    if index.diagnostic is not None or index.domain is None:
        return None
    option_count = len(node.args) - 1
    if isinstance(index.domain, _FiniteInts):
        selected = [i for i in sorted(index.domain.values) if 1 <= i <= option_count]
    else:
        lo = max(1, index.domain.lo)
        hi = min(option_count, index.domain.hi)
        selected = list(range(lo, hi + 1)) if lo <= hi else []
    if not selected:
        return frozenset()
    parts: list[frozenset[str]] = []
    for option_index in selected:
        text = _infer_text_fragments(
            node.args[option_index],
            env,
            limits,
            context=context,
            current_sheet=current_sheet,
            depth=depth + 1,
        )
        if text is None:
            return None
        parts.append(text)
    return _union_text_fragment_sets(parts, limits)


def _value_arg_keeps_numeric_domain(
    node: AstNode,
    env: CellTypeEnv,
    limits: DynamicRefLimits,
) -> bool:
    """Return True when `VALUE(node)` is the numeric domain of `node`.

    Digit-text (string cells, `&`, `CONCAT`) is excluded: `VALUE` must parse
    that text instead of treating it as already numeric.
    """
    if isinstance(node, NumberNode):
        return not isinstance(node.value, bool)
    if isinstance(node, UnaryOpNode):
        return node.op in {"-", "%"}
    if isinstance(node, BinaryOpNode):
        return node.op in {"+", "-", "*", "/"}
    if isinstance(node, CellRefNode):
        return _domain_from_cell_type(_lookup_cell_type(env, node.address), limits) is not None
    if isinstance(node, FunctionCallNode):
        return node.name.upper() in _VALUE_NUMERIC_IDENTITY_FUNCS
    return False


def _infer_value_numeric_domain(
    arg: AstNode,
    env: CellTypeEnv,
    limits: DynamicRefLimits,
    *,
    context: dict[str, int] | None,
    current_sheet: str,
    depth: int,
) -> _NumericDomainInferenceResult:
    """Integer domain of `VALUE(arg)`.

    Already-numeric arguments pass through, including wide bounds, without
    enumerating them. Digit-text is parsed only when the fragment set is finite.
    """
    if _value_arg_keeps_numeric_domain(arg, env, limits):
        return _infer_numeric_domain_result(
            arg,
            env,
            limits,
            context=context,
            current_sheet=current_sheet,
            depth=depth + 1,
        )
    texts = _infer_text_fragments(
        arg,
        env,
        limits,
        context=context,
        current_sheet=current_sheet,
        depth=depth + 1,
    )
    if texts is not None:
        return _numeric_domain_from_text_fragments(texts)
    return _infer_numeric_domain_result(
        arg,
        env,
        limits,
        context=context,
        current_sheet=current_sheet,
        depth=depth + 1,
    )


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
        if node.op == "&":
            diagnostic = _first_operand_diagnostic(
                (node.left, node.right),
                env,
                limits,
                context=ctx,
                current_sheet=current_sheet,
                depth=depth,
            )
            if diagnostic is not None:
                return diagnostic
            return _numeric_domain_from_canonical_text_fragments(
                _concat_text_fragments(
                    (node.left, node.right),
                    env,
                    limits,
                    context=ctx,
                    current_sheet=current_sheet,
                    depth=depth,
                )
            )
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
        if name == "VALUE":
            if len(node.args) != 1:
                return _domain_result(None)
            return _infer_value_numeric_domain(
                node.args[0],
                env,
                limits,
                context=ctx,
                current_sheet=current_sheet,
                depth=depth,
            )
        if name in {"CONCAT", "CONCATENATE"}:
            if not node.args:
                return _domain_result(None)
            diagnostic = _first_operand_diagnostic(
                node.args,
                env,
                limits,
                context=ctx,
                current_sheet=current_sheet,
                depth=depth,
            )
            if diagnostic is not None:
                return diagnostic
            return _numeric_domain_from_canonical_text_fragments(
                _concat_text_fragments(
                    node.args,
                    env,
                    limits,
                    context=ctx,
                    current_sheet=current_sheet,
                    depth=depth,
                )
            )
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
            _raise_cell_limit(len(targets), limits.max_cells, what="INDEX target cells")
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
            _raise_cell_limit(len(targets), limits.max_cells, what="INDEX target cells")
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
            _raise_cell_limit(len(targets), limits.max_cells, what="INDEX target cells")
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
        _raise_cell_limit(len(targets), limits.max_cells, what="INDEX target cells")
    return targets


def _finite_needle_values(
    needle: AstNode,
    env: CellTypeEnv,
    limits: DynamicRefLimits,
    *,
    context: dict[str, int] | None,
    current_sheet: str,
) -> frozenset[object] | None:
    """Return a small finite needle domain, or None when joint filtering should stop."""
    literal = _finite_exact_match_values(needle, env)
    if literal is not None:
        if _needle_blocks_exact_match_refine(literal) or len(literal) > limits.max_branches:
            return None
        return literal
    result = _infer_numeric_domain_result(
        needle,
        env,
        limits,
        context=context,
        current_sheet=current_sheet,
    )
    if result.diagnostic is not None or result.domain is None:
        return None
    domain = result.domain
    if isinstance(domain, _FiniteInts):
        if len(domain.values) > limits.max_branches:
            return None
        return frozenset(domain.values)
    span = domain.hi - domain.lo + 1
    if span < 1 or span > limits.max_branches:
        return None
    return frozenset(range(domain.lo, domain.hi + 1))


def _match_position_list(domain: _FiniteInts | _IntBounds, limit: int) -> list[int] | None:
    """Return positive positions in `domain` when the list is at most `limit` long."""
    if isinstance(domain, _FiniteInts):
        positions = [value for value in domain.values if value >= 1]
        if len(positions) > limit:
            return None
        positions.sort()
        return positions
    if domain.hi < max(1, domain.lo):
        return []
    start = max(1, domain.lo)
    span = domain.hi - start + 1
    if span > limit:
        return None
    return list(range(start, domain.hi + 1))


def _exact_match_lookup(
    node: AstNode, *, current_sheet: str
) -> tuple[FunctionCallNode, list[str]] | None:
    """Return an exact MATCH call and its lookup cells when both are static."""
    if not isinstance(node, FunctionCallNode) or node.name.upper() != "MATCH":
        return None
    if not _match_is_exact_match_type(node):
        return None
    ordered = _ordered_match_lookup_cells(node.args[1], current_sheet=current_sheet)
    if not ordered:
        return None
    return node, ordered


def _lookup_position_index(ordered: list[str]) -> dict[str, int]:
    """Map each lookup address to its first 1-based scan position."""
    found: dict[str, int] = {}
    for index, address in enumerate(ordered, start=1):
        found.setdefault(normalize_cell_type_env_key(address), index)
    return found


def _axis_requirement(position: int | None, chosen: int) -> bool | None:
    """Return whether `position` must match for MATCH to return `chosen`.

    `None` means MATCH stops before that cell. `True` is the returned hit.
    `False` is an earlier cell, which must miss.
    """
    if position is None or position > chosen:
        return None
    return position == chosen


def _untyped_joint_possible(
    row_must: bool | None,
    col_must: bool | None,
    row_needle: object,
    col_needle: object,
) -> bool:
    """Return whether some value can meet both axis constraints."""
    if row_must is True and col_must is True:
        return _values_match(row_needle, col_needle)
    if row_must is True and col_must is False:
        return not _values_match(row_needle, col_needle)
    if row_must is False and col_must is True:
        return not _values_match(row_needle, col_needle)
    return True


def _joint_cell_values(cell_type: CellType, limits: DynamicRefLimits) -> frozenset[object] | None:
    """Return a finite value set, or None when it is too wide to refute a pair."""
    if cell_type.enum is not None:
        if not cell_type.enum.values:
            return None
        return cell_type.enum.values
    if cell_type.kind is CellKind.BOOL:
        return frozenset({False, True})
    numeric = _domain_from_cell_type(cell_type, limits)
    if isinstance(numeric, _FiniteInts):
        if not numeric.values:
            return None
        return frozenset(numeric.values)
    if isinstance(numeric, _IntBounds):
        span = numeric.hi - numeric.lo + 1
        if 1 <= span <= limits.max_branches:
            return frozenset(range(numeric.lo, numeric.hi + 1))
    return None


def _shared_cell_possible(
    cell_type: CellType | None,
    limits: DynamicRefLimits,
    row_needle: object,
    col_needle: object,
    row_must: bool | None,
    col_must: bool | None,
) -> bool:
    """Return False only when `cell_type` cannot meet both axis constraints."""
    if row_must is None and col_must is None:
        return True
    if cell_type is not None and cell_type.enum is None and cell_type.kind is CellKind.ERROR:
        return row_must is not True and col_must is not True
    if cell_type is None:
        return _untyped_joint_possible(row_must, col_must, row_needle, col_needle)
    values = _joint_cell_values(cell_type, limits)
    if not values:
        return True
    return any(
        (row_must is None or _values_match(value, row_needle) is row_must)
        and (col_must is None or _values_match(value, col_needle) is col_must)
        for value in values
    )


def _pair_possible(
    shared: list[str],
    row_index: dict[str, int],
    col_index: dict[str, int],
    row_values: frozenset[object],
    col_values: frozenset[object],
    row_pos: int,
    col_pos: int,
    env: CellTypeEnv,
    limits: DynamicRefLimits,
) -> bool:
    """Return whether one needle assignment makes this INDEX pair feasible."""
    for row_needle in row_values:
        for col_needle in col_values:
            if all(
                _shared_cell_possible(
                    _lookup_cell_type(env, address),
                    limits,
                    row_needle,
                    col_needle,
                    _axis_requirement(row_index[address], row_pos),
                    _axis_requirement(col_index[address], col_pos),
                )
                for address in shared
            ):
                return True
    return False


def _joint_exact_match_targets(
    array_range: ExcelRange,
    row_ast: AstNode,
    col_ast: AstNode,
    row_dom: _FiniteInts | _IntBounds,
    col_dom: _FiniteInts | _IntBounds,
    env: CellTypeEnv,
    limits: DynamicRefLimits,
    *,
    context: dict[str, int] | None,
    current_sheet: str,
) -> set[str] | None:
    """Drop INDEX pairs that two exact MATCH calls cannot return together.

    Independent axes over-approximate when one cell is on both lookups: it
    cannot be a hit for both needles unless one value equals both. Returns
    None when the filter does not apply or does not remove any pair, so the
    caller keeps the rectangular emission. An empty result is also refused.
    """
    row_match = _exact_match_lookup(row_ast, current_sheet=current_sheet)
    col_match = _exact_match_lookup(col_ast, current_sheet=current_sheet)
    if row_match is None or col_match is None:
        return None
    row_call, row_ordered = row_match
    col_call, col_ordered = col_match
    row_index = _lookup_position_index(row_ordered)
    col_index = _lookup_position_index(col_ordered)
    shared = [address for address in row_index if address in col_index]
    if not shared:
        return None
    row_values = _finite_needle_values(
        row_call.args[0],
        env,
        limits,
        context=context,
        current_sheet=current_sheet,
    )
    col_values = _finite_needle_values(
        col_call.args[0],
        env,
        limits,
        context=context,
        current_sheet=current_sheet,
    )
    if (
        row_values is None
        or col_values is None
        or len(row_values) * len(col_values) > limits.max_branches
    ):
        return None
    cap = max(limits.max_cells, DynamicRefLimits().max_cells)
    row_positions = _match_position_list(row_dom, cap)
    col_positions = _match_position_list(col_dom, cap)
    if row_positions is None or col_positions is None:
        return None
    pair_count = len(row_positions) * len(col_positions)
    if pair_count == 0 or pair_count > cap:
        return None

    feasible: list[tuple[int, int]] = []
    for row_pos in row_positions:
        for col_pos in col_positions:
            if _pair_possible(
                shared,
                row_index,
                col_index,
                row_values,
                col_values,
                row_pos,
                col_pos,
                env,
                limits,
            ):
                feasible.append((row_pos, col_pos))
    if not feasible or len(feasible) == pair_count:
        return None

    nrows = array_range.end_row - array_range.start_row + 1
    ncols = array_range.end_col - array_range.start_col + 1
    targets: set[str] = set()
    for row_pos, col_pos in feasible:
        targets |= _index_pair_to_addresses(array_range, row_pos, col_pos, nrows=nrows, ncols=ncols)
    if len(targets) > limits.max_cells:
        _raise_cell_limit(len(targets), limits.max_cells, what="INDEX target cells")
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
        if col_ast is not None:
            joint = _joint_exact_match_targets(
                array_range,
                row_ast,
                col_ast,
                row_dom,
                col_dom,
                cell_type_env,
                limits,
                context=eval_context,
                current_sheet=current_sheet,
            )
            if joint is not None:
                return joint
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
        _raise_cell_limit(len(targets), limits.max_cells, what="INDEX target cells")
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
    `max_cells` budget is preserved, including `ValueError` when the rectangle
    is larger). A long formula chain that mentions the same range at every
    level would otherwise re-derive all of its cells once per level (issue
    #465).
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
            sheet, start_letter, end_letter = resolve_whole_column_ref(n, None)
            for dep_sheet, dep_a1 in expand_whole_column_span_deps(
                sheet, start_letter, end_letter, bounds
            ):
                addrs.add(format_key(dep_sheet, dep_a1))
            return
        if isinstance(n, WholeRowNode):
            bounds = sheet_bounds or {}
            _sheet, start_row, end_row = resolve_whole_row_ref(n, None)
            for dep_sheet, dep_a1 in expand_whole_row_span_deps(_sheet, start_row, end_row, bounds):
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
            _raise_cell_limit(len(out), lim.max_cells, what="Dynamic ref cells")

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


def _split_top_level_args(s: str, *, keep_empty: bool = False) -> list[str] | None:
    """Split `s` on top-level commas.

    Mirrors `parser._split_top_level_args`. Empty arguments are dropped unless
    `keep_empty` is set. INDEX inference keeps blanks so an omitted axis is
    Excel's whole-axis `0`. OFFSET and static-INDEX classification still drop
    them.

    Args:
        s: Argument text inside a call, without the surrounding parentheses.
        keep_empty: When True, retain blank arguments.

    Returns:
        Argument strings, or None when parentheses or quotes are unbalanced.
    """
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
    if keep_empty:
        return args
    return [a for a in args if a != ""]
