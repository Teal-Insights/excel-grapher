from __future__ import annotations

import hashlib
import logging
import re
import time
import warnings
from collections import deque
from collections.abc import Iterable, Mapping
from pathlib import Path

import fastpyxl
import fastpyxl.utils.cell
from fastpyxl.worksheet.formula import ArrayFormula
from fastpyxl.worksheet.worksheet import Worksheet

from excel_grapher.core.address_keys import CellKey, format_range_key, parse_address, sort_node_keys
from excel_grapher.core.cell_types import CellType, leaves_missing_cell_type_constraints
from excel_grapher.core.formula_ast import (
    AstNode,
    intern_formula_ast,
    parse_preserving_axes_optional,
)

from .blank_ranges import (
    address_in_blank_ranges,
    cell_in_blank_ranges,
    normalize_blank_range_specs,
)
from .dependency_provenance import EdgeProvenance
from .dynamic_ref_walk import DynamicRefWalkContext
from .dynamic_refs import (
    DynamicRefConfig,
    DynamicRefError,
    DynamicRefTraceEvent,
    GlobalWorkbookBounds,
    _emit_trace,
    clear_index_target_cache,
    expand_leaf_env_to_argument_env,
    infer_dynamic_index_targets,
    infer_dynamic_indirect_targets,
    infer_dynamic_offset_targets,
)
from .graph import DependencyGraph, NodeHook
from .guard import (
    And,
    Compare,
    GuardExpr,
    Literal,
    Not,
    and_guard,
    guard_range_shape,
    instantiate_element_guard,
    or_guard,
)
from .node import Node
from .parser import (
    CellRef,
    FormulaNormalizer,
    _find_function_calls_with_spans,
    _split_function_args,
    element_aligned_range_cells,
    expand_range,
    expand_range_ref,
    format_key,
    mask_ref_only_function_calls,
    mask_spans,
    parse_dynamic_range_refs_with_spans,
    parse_guard_expr,
    parse_range_refs_with_spans,
    parse_standalone_cell_refs,
    split_top_level_choose,
    split_top_level_function,
    split_top_level_if,
    split_top_level_ifs,
    split_top_level_switch,
)
from .provenance_collect import collect_provenance_for_formula
from .resolver import build_named_range_map
from .target_expansion import expand_targets_to_roots
from .type_analysis_cache import TypeAnalysisCache

_logger = logging.getLogger(__name__)
_CANDIDATES_BFS_PROGRESS_INTERVAL = 5000
"""Emit a `bfs-progress` trace every N outer BFS nodes in candidate listing."""
_CANDIDATES_ARG_PROGRESS_INTERVAL = 50_000
"""Emit a `candidates-arg-progress` trace every N argument-subgraph visits."""
_VOLATILE_DYNAMIC_REF_FUNCS = frozenset(
    {"NOW", "TODAY", "RAND", "RANDBETWEEN", "RANDARRAY", "INFO"}
)
_CONDITIONAL_FN_NAMES = frozenset({"IF", "IFS", "CHOOSE", "SWITCH"})
_DYNAMIC_REF_FN_NAMES = frozenset({"OFFSET", "INDIRECT", "INDEX"})
# Functions that consume a whole array argument and reduce or reshape it. A
# conditional nested inside one of these is evaluated element-wise by Excel
# (`SUM(IF(A1:A10>0,B1:B10,0))`), which is the array context that turns a
# range-typed condition into per-element edge guards. See issue #483.
#
# Element alignment is positional, matching CSE and dynamic-array evaluation.
# The one form it does not describe is legacy implicit intersection, where the
# formula's own row/column selects a single element instead. Excel 2019+ writes
# that reading explicitly as `@A1:A10`, which never parses into a range template
# (so those formulas stay conservative); only a pre-2019 save of a non-CSE
# `SUM(IF(range,...))` — a formula that already returns a one-element result in
# that engine — could be read the other way.
_ARRAY_CONSUMING_FN_NAMES = frozenset(
    {
        "AGGREGATE",
        "AVEDEV",
        "AVERAGE",
        "AVERAGEA",
        "CONCAT",
        "COUNT",
        "COUNTA",
        "DEVSQ",
        "GEOMEAN",
        "HARMEAN",
        "LARGE",
        "MAX",
        "MAXA",
        "MEDIAN",
        "MIN",
        "MINA",
        "MMULT",
        "MODE",
        "MODE.MULT",
        "MODE.SNGL",
        "PERCENTILE",
        "PERCENTILE.EXC",
        "PERCENTILE.INC",
        "PRODUCT",
        "QUARTILE",
        "QUARTILE.EXC",
        "QUARTILE.INC",
        "SMALL",
        "STDEV",
        "STDEV.P",
        "STDEV.S",
        "STDEVA",
        "STDEVP",
        "SUM",
        "SUMPRODUCT",
        "SUMSQ",
        "TEXTJOIN",
        "TRANSPOSE",
        "VAR",
        "VAR.P",
        "VAR.S",
        "VARA",
        "VARP",
    }
)
# Matches volatile builtins with optional Excel compatibility prefixes. Add new
# volatile function names to ``_VOLATILE_DYNAMIC_REF_FUNCS`` only (not prefix
# variants); ``_XLFN.`` / ``_XLUDF.`` are handled generically here.
_VOLATILE_DYNAMIC_REF_PATTERN = re.compile(
    r"(?<![A-Z0-9_])(?:_XLFN\.|_XLUDF\.)?(?:NOW|TODAY|RANDBETWEEN|RANDARRAY|RAND|INFO)\s*\("
)
_POSITION_DEPENDENT_DYNAMIC_REF_PATTERN = re.compile(
    r"(?<![A-Z0-9_])(?:ROW|COLUMN)\s*\(\s*\)",
    re.IGNORECASE,
)
# Position-independent formulas: (normalized_formula, sheet).
# Formulas with argument-less ROW()/COLUMN(): (normalized_formula, sheet, cell_a1).
_DynamicRefCacheKey = tuple[str, str] | tuple[str, str, str]
_DynamicRefTargets = tuple[set[str], set[str], set[str]]
# Deps and inferred targets are keyed by normalized formula; masking spans are not
# cached because they refer to raw formula text offsets.
_DynamicRefDependencyCacheValue = tuple[list[tuple[str, str]], _DynamicRefTargets]


def _dynamic_ref_cache_key(
    formula_for_infer: str,
    current_sheet: str,
    current_a1: str,
) -> _DynamicRefCacheKey:
    """Return the safest per-build cache key for dynamic-ref expansion."""
    if _POSITION_DEPENDENT_DYNAMIC_REF_PATTERN.search(formula_for_infer):
        return (formula_for_infer, current_sheet, current_a1)
    return (formula_for_infer, current_sheet)


def _workbook_sorted_sheet_a1_pairs(
    pairs: Iterable[tuple[str, str]], *, sheet_order: list[str]
) -> list[tuple[str, str]]:
    """Return `(sheet, a1)` pairs in workbook sheet/row/column order."""
    materialized = list(pairs)
    if not materialized:
        return []
    sorted_keys = sort_node_keys(
        [format_key(sh, a1) for sh, a1 in materialized],
        sheet_order=sheet_order,
    )
    return [parse_address(key) for key in sorted_keys]


def _sorted_guard_deps(
    dep_map: Mapping[tuple[str, str], GuardExpr | None], *, sheet_order: list[str]
) -> list[tuple[str, str, GuardExpr | None]]:
    return [
        (sh, a1, dep_map[(sh, a1)])
        for sh, a1 in _workbook_sorted_sheet_a1_pairs(dep_map.keys(), sheet_order=sheet_order)
    ]


def _conjoin_guards(outer: GuardExpr | None, inner: GuardExpr | None) -> GuardExpr | None:
    """AND an enclosing branch guard with a nested guard.

    `None` (unconditional or opaque) acts as identity: the other guard is still a
    sound necessary condition for the dependency to be active.
    """
    if outer is None:
        return inner
    if inner is None:
        return outer
    return and_guard(outer, inner)


def _merge_guarded_dep(
    out: dict[tuple[str, str], GuardExpr | None],
    key: tuple[str, str],
    guard: GuardExpr | None,
) -> None:
    """Merge a guarded dep into `out`, ORing guards when the dep repeats.

    `None` (unconditional/opaque) always wins, mirroring `DependencyGraph.add_edge`.
    """
    if key not in out:
        out[key] = guard
        return
    existing = out[key]
    if existing is None or guard is None:
        out[key] = None
    elif existing != guard:
        out[key] = or_guard(existing, guard)


def _sequential_default_guard(negations: list[GuardExpr]) -> GuardExpr:
    """Guard for a sequential default branch: NOT(cond_1) AND ... NOT(cond_n)."""
    if not negations:
        return Literal(True)
    if len(negations) == 1:
        return negations[0]
    return And(tuple(negations))


def _span_contains(outer: tuple[int, int], inner: tuple[int, int]) -> bool:
    """Return True when `inner` lies strictly inside `outer`."""
    return outer[0] <= inner[0] and inner[1] <= outer[1] and outer != inner


def _is_whole_formula_call(formula: str, span: tuple[int, int]) -> bool:
    """Return True when `span` covers the entire formula body after a leading `=`."""
    body_start = 1 if formula.startswith("=") else 0
    while body_start < len(formula) and formula[body_start].isspace():
        body_start += 1
    body_end = len(formula)
    while body_end > body_start and formula[body_end - 1].isspace():
        body_end -= 1
    return span[0] <= body_start and span[1] >= body_end


def _outermost_embedded_conditional_spans(
    formula: str,
) -> list[tuple[int, int]]:
    """Return spans of outermost conditional calls embedded in a larger expression.

    Skips:

    - the formula's own top-level call (avoids re-entering when splitters reject it,
      e.g. a two-argument `IF`);
    - conditionals nested inside another conditional (handled by recursive
      top-level branching);
    - conditionals whose entire call lies inside an `OFFSET`/`INDIRECT`/`INDEX`
      span (dynamic-ref argument analysis owns those deps).
    """
    dyn_spans = [
        span for _, _, span in _find_function_calls_with_spans(formula, _DYNAMIC_REF_FN_NAMES)
    ]
    cond_spans = [
        span
        for _, _, span in _find_function_calls_with_spans(formula, _CONDITIONAL_FN_NAMES)
        if not _is_whole_formula_call(formula, span)
        and not any(_span_contains(dyn, span) for dyn in dyn_spans)
    ]
    return [
        span for span in cond_spans if not any(_span_contains(other, span) for other in cond_spans)
    ]


def _spans_in_array_context(formula: str) -> list[tuple[int, int]]:
    """Return spans of calls that impose array context on nested conditionals."""
    return [
        span for _, _, span in _find_function_calls_with_spans(formula, _ARRAY_CONSUMING_FN_NAMES)
    ]


def _expand_targets_to_roots(
    targets: Iterable[str],
    *,
    sheetnames: list[str],
    named_ranges: dict[str, tuple[str, str]],
    named_range_ranges: dict[str, tuple[str, str, str]],
    max_range_cells: int,
) -> list[tuple[str, str]]:
    return expand_targets_to_roots(
        targets,
        sheetnames=sheetnames,
        named_ranges=named_ranges,
        named_range_ranges=named_range_ranges,
        max_range_cells=max_range_cells,
    )


def _format_missing_leaves(missing_leaves: set[str]) -> list[str]:
    """Format missing leaf cells for error messages.

    Per sheet, references are summarized without overstating coverage:

    - Within a column, contiguous rows become one vertical range.
    - Adjacent columns with the same row runs merge into one rectangle per run.
    """
    from fastpyxl.utils import range_boundaries
    from fastpyxl.utils.cell import coordinate_to_tuple, get_column_letter

    by_sheet: dict[str, dict[int, list[int]]] = {}
    others: list[str] = []

    for addr in missing_leaves:
        if "!" not in addr:
            others.append(addr)
            continue
        try:
            sheet, a1 = parse_address(addr)
        except ValueError:
            others.append(addr)
            continue
        if ":" in a1:
            try:
                min_col, min_row, max_col, max_row = range_boundaries(a1)
            except ValueError:
                others.append(addr)
                continue
            col_map = by_sheet.setdefault(sheet, {})
            for c in range(min_col, max_col + 1):
                col_map.setdefault(c, []).extend(range(min_row, max_row + 1))
            continue
        try:
            row, col = coordinate_to_tuple(a1)
        except ValueError:
            others.append(addr)
            continue
        by_sheet.setdefault(sheet, {}).setdefault(col, []).append(row)

    parts: list[str] = []
    for sheet in sorted(by_sheet.keys()):
        col_map = by_sheet[sheet]
        intervals_by_col: dict[int, tuple[tuple[int, int], ...]] = {}
        for col_idx, rows in col_map.items():
            rows_sorted = sorted(set(rows))
            merged: list[tuple[int, int]] = []
            if rows_sorted:
                run_start = prev = rows_sorted[0]
                for r in rows_sorted[1:]:
                    if r == prev + 1:
                        prev = r
                        continue
                    merged.append((run_start, prev))
                    run_start = prev = r
                merged.append((run_start, prev))
            intervals_by_col[col_idx] = tuple(merged)

        cols_sorted = sorted(intervals_by_col.keys())
        i = 0
        while i < len(cols_sorted):
            c0 = cols_sorted[i]
            ivals = intervals_by_col[c0]
            j = i + 1
            while (
                j < len(cols_sorted)
                and cols_sorted[j] == cols_sorted[j - 1] + 1
                and intervals_by_col[cols_sorted[j]] == ivals
            ):
                j += 1
            c_last = cols_sorted[j - 1]
            c_start_letter = get_column_letter(c0)
            c_end_letter = get_column_letter(c_last)
            for r1, r2 in ivals:
                if c0 == c_last and r1 == r2:
                    parts.append(format_key(sheet, f"{c_start_letter}{r1}"))
                else:
                    parts.append(
                        format_range_key(
                            sheet,
                            f"{c_start_letter}{r1}",
                            f"{c_end_letter}{r2}",
                        )
                    )
            i = j

    parts.extend(sorted(others))
    return sorted(parts)


def create_dependency_graph(
    workbook: Path | str,
    targets: Iterable[str],
    *,
    max_depth: int = 50,
    expand_ranges: bool = True,
    max_range_cells: int = 5000,
    hooks: list[NodeHook] | None = None,
    load_values: bool = True,
    dynamic_refs: DynamicRefConfig | None = None,
    use_cached_dynamic_refs: bool = False,
    capture_dependency_provenance: bool = False,
    store_raw_formula: bool = False,
    blank_ranges: Iterable[str] | None = None,
    type_analysis_cache: TypeAnalysisCache | None = None,
    warm_ast_cache: bool = False,
    warm_formula_shapes: bool = False,
) -> DependencyGraph:
    r"""Build a dependency graph starting from target cells.

    `targets` accepts any mix of:

    - sheet-qualified single cells (`"Sheet1!A1"`, `"'My Sheet'!B2"`);
    - sheet-qualified rectangular ranges (`"Sheet1!B12:F12"`,
      `"Sheet1!A1:B2"`, `"'My Sheet'!A1:B2"`); and
    - defined names that resolve to a single cell or a rectangular range
      (`"MyInput"`, `"DataRange"`).

    Range and named-range targets are expanded to one BFS root per cell
    (subject to `max_range_cells`); the deduplicated union seeds traversal.
    Targets that are neither sheet-qualified nor a known defined name raise
    `ValueError`.

    When `load_values` is True (default), the workbook is loaded once with
    `keep_formula_cache=True` so formula text and Excel's last-calculated
    cache are available without a second `data_only` parse. When False,
    formula nodes keep `value=None` and the formula-cache side map is not
    populated.

    Supports basic A1 references, sheet-qualified references, and dynamic references
    (OFFSET/INDIRECT/INDEX). For OFFSET/INDIRECT/INDEX:

    - **use_cached_dynamic_refs=True**: Resolve using cached workbook values (existing path).
      `dynamic_refs` is ignored.
    - **use_cached_dynamic_refs=False** (default), **dynamic_refs=None**: On any formula
      that contains OFFSET, INDIRECT, or INDEX requiring resolution, raise `DynamicRefError`.
      Callers can pass a `DynamicRefConfig` or set `use_cached_dynamic_refs=True` to avoid.
    - **use_cached_dynamic_refs=False**, **dynamic_refs** set: Resolve OFFSET/INDIRECT/INDEX via
      the config's `cell_type_env` and `limits`; missing or invalid domains raise
      `DynamicRefError`.

    To build a config from a `dict[str, type]` constraints schema, use
    `DynamicRefConfig.from_constraints`.

    When `capture_dependency_provenance` is True, each edge stores merged
    `excel_grapher.grapher.dependency_provenance.EdgeProvenance` on
    `DependencyGraph._edge_provenance` (how the dependency arises: direct
    reference, static range, dynamic OFFSET/INDIRECT/INDEX). Direct-reference spans
    are recorded against `Node.normalized_formula` only.

    When `store_raw_formula` is True, each formula cell also keeps the workbook
    formula text on `Node.formula`. It defaults to False because the string is
    not needed to evaluate, compress, or export a graph: `formula_ast` is the
    in-memory source of truth, `normalized_formula` is a derived
    `render_formula(..., style=A1_ABSOLUTE)` view (not stored on the node), and
    dropping the raw copy is a sizeable memory saving on large workbooks.
    Enable it when you need the original text -- audit / display use, and TACO
    range compression (`excel_grapher.grapher.range_compression`), which infers
    stride patterns from the relative/absolute (`$`) markers that
    normalization strips.

    Raw formulas are audit records of extraction: compression and projection
    rewrite `formula_ast` (the derived `normalized_formula` view follows), so on a
    compressed graph `Node.formula` still shows the pre-compression workbook
    text and must not be re-parsed as the node's current definition.

    Extraction always parses each formula cell into `Node.formula_ast` from the
    raw workbook text, preserving per-axis relative/absolute intent. Distinct
    ASTs share one interned tree (keyed by the frozen tree itself, not by a
    JSON encoding or by formula text). `normalized_formula` is derived
    absolute A1 via `render_formula`. Formulas the AST parser cannot handle leave
    `formula_ast` unset; extraction still records the cell and its
    dependencies.

    `blank_ranges` is an optional iterable of sheet-qualified A1 rectangles
    (e.g. `\"Sheet1!B2:D10\"`) treated as structurally empty: no nodes are
    created for those cells (edges into them are kept), and dynamic-ref leaf
    constraints are not required for addresses inside these ranges. Pair with the
    same declarations on `excel_grapher.FormulaEvaluator` and
    `excel_grapher.exporter.codegen.CodeGenerator.generate` for **evaluator
    <-> export** parity (consistent behavior between evaluation and generated code).

    When `warm_ast_cache` is True, each distinct derived `normalized_formula`
    in the built graph is stored on `DependencyGraph.preparsed_formulas` as an
    absolute-bound AST (`bind_axes` of `Node.formula_ast`). The mapping is
    keyed by absolute A1 text, not `NodeKey`. `FormulaEvaluator` evaluates
    per-node trees directly and seeds its string-keyed fallback cache from
    those bound ASTs, so first evaluation does not re-parse unless formulas
    change after extraction. Seeding is best-effort when distinct formulas
    exceed `FormulaEvaluator.ast_cache_maxsize` (oldest warmed entries may be
    evicted). `preparsed_formulas` is not stored in JSON graph caches; call
    `warm_preparsed_formulas` after cache load or formula mutation if you need
    the string-keyed overlay.

    When `warm_formula_shapes` is True, the same parse pass interns punched AST
    skeletons on `DependencyGraph.formula_shapes`, keyed by `NodeKey`.
    `FormulaEvaluator` compiles each shape once and `CodeGenerator` emits a
    shared helper per shape when profitable. The table is not JSON/pickle
    serialized; call `warm_formula_shapes` after cache load. Formula rewrite
    invalidates it.

    **Cost model**: constraint-based dynamic-ref expansion (`dynamic_refs` set,
    `use_cached_dynamic_refs=False`) runs `expand_leaf_env_to_argument_env`
    only when some argument-subgraph ref (a cell feeding OFFSET / INDIRECT /
    INDEX arguments) is not yet in the shared cell-type cache.  The first
    formula that needs a given set of argument cells pays for expansion;
    later formulas whose argument refs are already typed skip the expand call
    and reuse the env (issue #528).  That includes row-wise INDEX/MATCH and
    OFFSET variants that share a MATCH lookup.  INDEX / OFFSET / INDIRECT
    *target* inference still runs per formula so shifted arrays and bases
    keep distinct deps.

    Provenance collection (`capture_dependency_provenance=True`) reads the
    per-cell `_dyn_cache` of inferred targets filled during extraction
    (keyed by normalized formula, sheet, and A1).  Row-wise variants miss
    that key across cells, but extraction populates it before provenance
    runs, so provenance does not re-expand.  Extraction and provenance also
    share one `DynamicRefWalkContext` for argument-subgraph BFS and static
    ref parsing, so a `_dyn_cache` hit skips the walk entirely and a miss
    reuses extract's memoized `(formula, sheet)` ref sets (issue #539).
    Callers doing iterative constraint-tuning workflows can still set
    `capture_dependency_provenance=False` to avoid remaining provenance
    overhead (formula-string span collection, branch-union merging, etc.).
    """
    if not isinstance(workbook, (str, Path)):
        raise TypeError(
            "create_dependency_graph requires a path or path-like string for "
            f"`workbook`; got {type(workbook).__name__}. Pass a path so the "
            "builder can load formulas and cached values together via "
            "`keep_formula_cache=True` (or load formulas alone when "
            "`load_values=False`)."
        )

    blank_rects = normalize_blank_range_specs(blank_ranges)

    def load_wb(
        *,
        data_only: bool = False,
        keep_formula_cache: bool = False,
    ) -> fastpyxl.Workbook:
        path = Path(workbook)
        keep_vba = path.suffix.lower() == ".xlsm"
        return fastpyxl.load_workbook(
            path,
            data_only=data_only,
            keep_vba=keep_vba,
            keep_formula_cache=keep_formula_cache,
        )

    _t0 = time.perf_counter()
    # One parse supplies both formulas (`cell.value`) and Excel's last-calculated
    # cache (`cell.cached_value`) when values are requested.
    wb_formulas = load_wb(keep_formula_cache=load_values)
    _emit_trace(
        DynamicRefTraceEvent(
            kind="workbook-loaded",
            name="create_dependency_graph",
            elapsed_s=time.perf_counter() - _t0,
            detail={"keep_formula_cache": load_values},
        )
    )
    # Lazy fallback for resolve_cached_value when the initial load omitted caches
    # (`load_values=False`) but dynamic-ref resolution later needs them.
    wb_values: fastpyxl.Workbook | None = None

    # Compute workbook SHA-256 for persistent type-analysis cache
    _wb_sha256: str | None = None
    if type_analysis_cache is not None:
        with open(workbook, "rb") as _f:
            _wb_sha256 = hashlib.file_digest(_f, "sha256").hexdigest()

    graph = DependencyGraph(sheet_order=list(wb_formulas.sheetnames))
    sheet_bounds: dict[str, tuple[int, int]] = {}
    for h in hooks or []:
        graph.register_hook(h)

    named_range_maps = build_named_range_map(wb_formulas)
    named_ranges = named_range_maps.cell_map
    named_range_ranges = named_range_maps.range_map
    normalizer = FormulaNormalizer(named_ranges, named_range_ranges)
    defined_names: set[str] = {str(name) for name in wb_formulas.defined_names}
    # Clear per-graph-build caches from previous invocations.
    clear_index_target_cache()

    # Per-graph cache: (normalized_formula, current_sheet, current_a1) -> (offset_targets, indirect_targets, index_targets).
    # Populated by extract_expr_deps (constraint path); consumed by collect_provenance_for_formula
    # to avoid re-running the expensive expand_leaf_env_to_argument_env call.
    _dyn_cache: dict[tuple[str, str, str], _DynamicRefTargets] = {}
    # Per-graph cache for the full dependency contribution of constraint-based
    # dynamic refs. Position-independent repeated formulas can share this before
    # rebuilding argument domains and rerunning INDEX/MATCH inference.
    _dyn_dep_cache: dict[_DynamicRefCacheKey, _DynamicRefDependencyCacheValue] = {}
    # Shared cell-type cache for expand_leaf_env_to_argument_env: intermediate
    # formula cells inferred once are reused across BFS nodes, avoiding redundant
    # recursive domain inference when many dynamic-ref formulas share intermediates.
    _shared_cell_type_cache: dict[str, CellType] = {}
    _dyn_stats = {
        "infer_calls": 0,
        "cache_hits": 0,
        "dep_cache_hits": 0,
        "env_cache_hits": 0,
        "arg_subgraph_hits": 0,
    }
    _NAME_TOKEN_RE = re.compile(r"\b([A-Za-z_][A-Za-z0-9_]*)\b(?!\s*!)")

    # Worksheet caches: avoid repeated O(#sheets) __getitem__ scans on every BFS node.
    _ws_f_cache: dict[str, Worksheet] = {}
    _ws_v_cache: dict[str, Worksheet] = {}

    def _get_ws_f(sheet: str) -> Worksheet:
        ws = _ws_f_cache.get(sheet)
        if ws is None:
            ws = wb_formulas[sheet]
            _ws_f_cache[sheet] = ws
            max_row = getattr(ws, "max_row", None) or 1
            max_col = getattr(ws, "max_column", None) or 1
            if max_row < 1:
                max_row = 1
            if max_col < 1:
                max_col = 1
            sheet_bounds[sheet] = (max_row, max_col)
        return ws

    def _ensure_sheet_bounds(sheet: str) -> None:
        if sheet not in sheet_bounds:
            _get_ws_f(sheet)

    def _cell_value(sheet: str, a1: str) -> object:
        return _get_ws_f(sheet)[a1].value

    ref_walk = DynamicRefWalkContext(
        normalizer=normalizer,
        max_range_cells=max_range_cells,
        get_cell_value=_cell_value,
        sheet_names=wb_formulas.sheetnames,
        shared_cell_type_cache=_shared_cell_type_cache,
        stats=_dyn_stats,
        named_ranges=named_ranges,
        named_range_ranges=named_range_ranges,
    )

    def _get_ws_v(sheet: str) -> Worksheet:
        # Only used for the lazy data_only fallback workbook.
        ws = _ws_v_cache.get(sheet)
        if ws is None:
            assert wb_values is not None
            ws = wb_values[sheet]
            _ws_v_cache[sheet] = ws
        return ws

    def _cached_value_from_formula_cell(cell: object) -> object | None:
        """Return Excel's cached result for a formula cell, else the cell value."""
        data_type = getattr(cell, "data_type", None)
        raw = getattr(cell, "value", None)
        is_formula = data_type == "f" or isinstance(raw, ArrayFormula)
        if is_formula:
            return getattr(cell, "cached_value", None)
        return raw

    def resolve_cached_value(sheet: str, a1: str) -> object | None:
        nonlocal wb_values
        if wb_formulas.keep_formula_cache:
            return _cached_value_from_formula_cell(_get_ws_f(sheet)[a1])
        if wb_values is None:
            wb_values = load_wb(data_only=True)
        return _get_ws_v(sheet)[a1].value

    def _contains_volatile_function(formula_or_expr: str) -> bool:
        formula_text = formula_or_expr if formula_or_expr.startswith("=") else "=" + formula_or_expr
        upper_formula = formula_text.upper()
        if "RANDARRAY(" in upper_formula:
            return True
        return bool(_VOLATILE_DYNAMIC_REF_PATTERN.search(upper_formula))

    def _collect_refs_for_volatile_scan(formula_text: str, sheet_of_cell: str) -> set[str]:
        normalized = normalizer.normalize(
            formula_text if formula_text.startswith("=") else "=" + formula_text,
            sheet_of_cell,
        )
        normalized = mask_ref_only_function_calls(normalized)
        out: set[str] = set()
        for ref in parse_standalone_cell_refs(normalized):
            sh = ref.sheet if ref.sheet is not None else sheet_of_cell
            out.add(format_key(sh, f"{ref.column}{ref.row}"))
        for start, end, _span in parse_range_refs_with_spans(normalized):
            sh = start.sheet if start.sheet is not None else sheet_of_cell
            _ensure_sheet_bounds(sh)
            for dep_sheet, dep_a1 in expand_range_ref(
                start=start,
                end=end,
                default_sheet=sh,
                max_cells=max_range_cells,
                sheet_bounds=sheet_bounds,
            ):
                out.add(format_key(dep_sheet, dep_a1))
        return out

    def _call_has_volatile_dependency_chain(
        fn_name: str,
        inner_args: str,
        *,
        current_sheet: str,
    ) -> bool:
        args = _split_function_args(inner_args)
        if args is None:
            return False
        if fn_name == "OFFSET":
            relevant_args = args[1:]
        elif fn_name == "INDIRECT":
            relevant_args = args[:1]
        else:
            relevant_args = []

        to_visit: set[str] = set()
        visited: set[str] = set()
        for arg in relevant_args:
            expr = "=" + arg
            if _contains_volatile_function(expr):
                return True
            normalized = normalizer.normalize(expr, current_sheet)
            normalized = mask_ref_only_function_calls(normalized)
            for ref in parse_standalone_cell_refs(normalized):
                sh = ref.sheet if ref.sheet is not None else current_sheet
                to_visit.add(format_key(sh, f"{ref.column}{ref.row}"))

        while to_visit:
            addr = to_visit.pop()
            if addr in visited:
                continue
            visited.add(addr)
            sh, a1 = parse_address(addr)
            if sh not in wb_formulas.sheetnames:
                continue
            cell_val = _get_ws_f(sh)[a1].value
            if isinstance(cell_val, ArrayFormula):
                cell_val = cell_val.text or ""
                if cell_val and not cell_val.startswith("="):
                    cell_val = f"={cell_val}"
            if not isinstance(cell_val, str) or not cell_val.startswith("="):
                continue
            if _contains_volatile_function(cell_val):
                return True
            to_visit.update(_collect_refs_for_volatile_scan(cell_val, sh))
        return False

    def extract_deps_with_guards(
        formula: str, current_sheet: str, current_a1: str, *, array_formula: bool = False
    ) -> list[tuple[str, str, GuardExpr | None]]:
        if not formula.startswith("="):
            return []
        try:
            return _extract_deps_with_guards_inner(
                formula, current_sheet, current_a1, array_formula=array_formula
            )
        except DynamicRefError:
            raise
        except ValueError as exc:
            raise ValueError(f"{current_sheet}!{current_a1}: {exc}") from exc

    def _extract_deps_with_guards_inner(
        formula: str, current_sheet: str, current_a1: str, *, array_formula: bool = False
    ) -> list[tuple[str, str, GuardExpr | None]]:
        sheet_order = list(wb_formulas.sheetnames)

        def extract_expr_deps(expr: str) -> list[tuple[str, str]]:
            """Extract dependencies from an expression fragment (no leading '=')."""
            f = "=" + expr if not expr.startswith("=") else expr
            deps: list[tuple[str, str]] = []

            masked = f

            # 0) Dynamic refs (OFFSET/INDIRECT/INDEX): cached, raise, or constraint-based.
            dyn_spans: list[tuple[int, int]] = []
            if use_cached_dynamic_refs:
                for start, end, span, arg_refs in parse_dynamic_range_refs_with_spans(
                    f,
                    current_sheet=current_sheet,
                    current_cell_a1=current_a1,
                    named_ranges=named_ranges,
                    named_range_ranges=named_range_ranges,
                    normalizer=normalizer,
                    value_resolver=resolve_cached_value,
                ):
                    dyn_spans.append(span)
                    sheet = start.sheet if start.sheet is not None else current_sheet
                    deps.extend(
                        expand_range(
                            sheet=sheet,
                            start_col=start.column,
                            start_row=start.row,
                            end_col=end.column,
                            end_row=end.row,
                            max_cells=max_range_cells,
                        )
                    )
                    for ref in arg_refs:
                        arg_sheet = ref.sheet if ref.sheet is not None else current_sheet
                        deps.append((arg_sheet, f"{ref.column}{ref.row}"))
            else:
                calls = _find_function_calls_with_spans(
                    f, frozenset({"OFFSET", "INDIRECT", "INDEX"})
                )
                if dynamic_refs is None:
                    static_dynamic_calls_by_span: dict[
                        tuple[int, int], tuple[CellRef, CellRef, list[CellRef]]
                    ] = {}
                    try:
                        static_calls = parse_dynamic_range_refs_with_spans(
                            f,
                            current_sheet=current_sheet,
                            current_cell_a1=current_a1,
                            named_ranges=named_ranges,
                            named_range_ranges=named_range_ranges,
                            normalizer=normalizer,
                            value_resolver=None,
                        )
                        static_dynamic_calls_by_span = {
                            span: (start, end, arg_refs)
                            for start, end, span, arg_refs in static_calls
                        }
                    except ValueError:
                        static_dynamic_calls_by_span = {}
                    # Filter out INDEX calls that only have literal args (no dynamic resolution needed).
                    dynamic_calls = []
                    saw_volatile_in_dynamic_ref_chain = False
                    for fn_name_check, inner_check, span_check in calls:
                        is_volatile_dynamic_call = False
                        if fn_name_check in {"OFFSET", "INDIRECT"}:
                            is_volatile_dynamic_call = _call_has_volatile_dependency_chain(
                                fn_name_check,
                                inner_check,
                                current_sheet=current_sheet,
                            )
                            if is_volatile_dynamic_call:
                                saw_volatile_in_dynamic_ref_chain = True
                        if fn_name_check == "INDEX":
                            # INDEX only needs dynamic resolution when row/col args are non-literal
                            idx_args = _split_function_args(inner_check)
                            if idx_args is not None and len(idx_args) >= 2:
                                has_non_literal = False
                                for j, idx_arg in enumerate(idx_args):
                                    if j == 0:
                                        continue  # skip array arg
                                    try:
                                        float(idx_arg.strip())
                                    except ValueError:
                                        has_non_literal = True
                                        break
                                if not has_non_literal:
                                    continue
                        if fn_name_check in {"OFFSET", "INDIRECT"}:
                            static_call = static_dynamic_calls_by_span.get(span_check)
                            if static_call is not None:
                                if is_volatile_dynamic_call:
                                    dynamic_calls.append((fn_name_check, inner_check, span_check))
                                    continue
                                start_ref, end_ref, arg_refs = static_call
                                dyn_spans.append(span_check)
                                call_sheet = (
                                    start_ref.sheet
                                    if start_ref.sheet is not None
                                    else current_sheet
                                )
                                deps.extend(
                                    expand_range(
                                        sheet=call_sheet,
                                        start_col=start_ref.column,
                                        start_row=start_ref.row,
                                        end_col=end_ref.column,
                                        end_row=end_ref.row,
                                        max_cells=max_range_cells,
                                    )
                                )
                                for ref in arg_refs:
                                    arg_sheet = (
                                        ref.sheet if ref.sheet is not None else current_sheet
                                    )
                                    deps.append((arg_sheet, f"{ref.column}{ref.row}"))
                                continue
                        dynamic_calls.append((fn_name_check, inner_check, span_check))
                    calls = dynamic_calls
                    if calls:
                        if saw_volatile_in_dynamic_ref_chain:
                            warnings.warn(
                                "Detected volatile function in OFFSET/INDIRECT dependency chain. "
                                "use_cached_dynamic_refs=True is still required for resolution, and "
                                "FormulaEvaluator.evaluate or CodeGenerator output may hit runtime errors "
                                "until volatile dynamic-ref support is fully implemented.",
                                UserWarning,
                                stacklevel=2,
                            )
                        cell_key = format_key(current_sheet, current_a1)
                        fn_names = sorted({fn for fn, _, _ in calls})
                        raise DynamicRefError(
                            f"Formula at {cell_key} contains {'/'.join(fn_names)} that require resolution. "
                            "Pass dynamic_refs=DynamicRefConfig.from_constraints(...) or set "
                            "use_cached_dynamic_refs=True."
                        )
                else:
                    bounds = GlobalWorkbookBounds(sheet=current_sheet)
                    if calls:
                        formula_for_infer = normalizer.normalize(
                            f if f.startswith("=") else "=" + f,
                            current_sheet,
                        )
                        _cell_cache_key = (formula_for_infer, current_sheet, current_a1)
                        _dyn_dep_cache_key = _dynamic_ref_cache_key(
                            formula_for_infer,
                            current_sheet,
                            current_a1,
                        )
                        _cached_dynamic_deps = _dyn_dep_cache.get(_dyn_dep_cache_key)
                        if _cached_dynamic_deps is not None:
                            cached_deps, cached_targets = _cached_dynamic_deps
                            deps.extend(cached_deps)
                            dyn_spans.extend(span for _, _, span in calls)
                            _dyn_cache[_cell_cache_key] = cached_targets
                            _dyn_stats["dep_cache_hits"] += 1
                            calls = []
                    argument_addrs: set[str] = set()
                    if calls:
                        _deps_start = len(deps)
                        for fn_name, inner, span in calls:
                            dyn_spans.append(span)
                            args = _split_function_args(inner)
                            if args is None:
                                continue
                            for i, arg in enumerate(args):
                                normalized = normalizer.normalize(
                                    "=" + arg,
                                    current_sheet,
                                )
                                # Variable args: always traverse to leaves for domain expansion.
                                # OFFSET base (i==0): only traverse when base is an expression (e.g. INDEX(...))
                                # INDEX: array arg (i==0) is not variable; row/col args (i>=1) are.
                                is_variable = (
                                    (fn_name == "OFFSET" and i >= 1)
                                    or (fn_name == "OFFSET" and i == 0 and "(" in normalized)
                                    or fn_name == "INDIRECT"
                                    or (fn_name == "INDEX" and i >= 1)
                                )
                                # INDEX array arg: do not add refs/ranges here — nested parens mean the
                                # array is an expression (OFFSET/INDEX/...) and still needs traversal.
                                # Static ranges are handled by infer_dynamic_index_targets (GH-156).
                                if fn_name == "INDEX" and i == 0 and "(" not in normalized:
                                    continue
                                value_expr = mask_ref_only_function_calls(normalized)
                                for ref in parse_standalone_cell_refs(value_expr):
                                    sh = ref.sheet if ref.sheet is not None else current_sheet
                                    a1 = f"{ref.column}{ref.row}"
                                    deps.append((sh, a1))
                                    if is_variable:
                                        argument_addrs.add(format_key(sh, a1))
                                if is_variable:
                                    for start, end, _span in parse_range_refs_with_spans(
                                        value_expr
                                    ):
                                        range_sheet = (
                                            start.sheet
                                            if start.sheet is not None
                                            else current_sheet
                                        )
                                        for dep_sheet, dep_a1 in expand_range(
                                            sheet=range_sheet,
                                            start_col=start.column,
                                            start_row=start.row,
                                            end_col=end.column,
                                            end_row=end.row,
                                            max_cells=max_range_cells,
                                        ):
                                            deps.append((dep_sheet, dep_a1))
                                            argument_addrs.add(format_key(dep_sheet, dep_a1))
                    if calls:
                        all_refs, leaves = ref_walk.argument_subgraph_refs(argument_addrs)
                        missing_leaves = leaves_missing_cell_type_constraints(
                            leaves, dynamic_refs.cell_type_env
                        )
                        if blank_rects:
                            missing_leaves = {
                                a
                                for a in missing_leaves
                                if not address_in_blank_ranges(a, blank_rects)
                            }
                        if missing_leaves:
                            cell_key = format_key(current_sheet, current_a1)
                            formatted_missing = _format_missing_leaves(missing_leaves)
                            raise DynamicRefError(
                                f"Formula at {cell_key} contains OFFSET, INDIRECT, or INDEX; the following leaf "
                                f"cells that feed them have no constraint: {formatted_missing}. "
                                "Add constraints only for leaf (non-formula) cells."
                            )
                        formula_for_infer = normalizer.normalize(
                            f if f.startswith("=") else "=" + f,
                            current_sheet,
                        )
                        _col_letter, _current_row = fastpyxl.utils.cell.coordinate_from_string(
                            current_a1
                        )
                        _current_col = fastpyxl.utils.cell.column_index_from_string(_col_letter)
                        _cache_key = (formula_for_infer, current_sheet, current_a1)
                        if _cache_key in _dyn_cache:
                            offset_targets, indirect_targets, index_targets = _dyn_cache[_cache_key]
                            _dyn_stats["cache_hits"] += 1
                        else:
                            _dyn_stats["infer_calls"] += 1
                            if all_refs and all(
                                addr in _shared_cell_type_cache for addr in all_refs
                            ):
                                expanded_env = _shared_cell_type_cache
                                _dyn_stats["env_cache_hits"] += 1
                            else:
                                expanded_env = expand_leaf_env_to_argument_env(
                                    all_refs,
                                    ref_walk.cell_formula,
                                    ref_walk.refs_in_formula_without_dynamic,
                                    dynamic_refs.cell_type_env,
                                    dynamic_refs.limits,
                                    named_ranges=named_ranges,
                                    named_range_ranges=named_range_ranges,
                                    max_range_cells=max_range_cells,
                                    shared_cell_type_cache=_shared_cell_type_cache,
                                    type_analysis_cache=type_analysis_cache,
                                    workbook_sha256=_wb_sha256,
                                    get_cell_ast=ref_walk.cell_ast,
                                )
                            try:
                                offset_targets = infer_dynamic_offset_targets(
                                    formula_for_infer,
                                    current_sheet=current_sheet,
                                    cell_type_env=expanded_env,
                                    limits=dynamic_refs.limits,
                                    bounds=bounds,
                                    named_ranges=named_ranges,
                                    named_range_ranges=named_range_ranges,
                                    current_row=_current_row,
                                    current_col=_current_col,
                                )
                                indirect_targets = infer_dynamic_indirect_targets(
                                    formula_for_infer,
                                    current_sheet=current_sheet,
                                    cell_type_env=expanded_env,
                                    limits=dynamic_refs.limits,
                                    bounds=bounds,
                                    named_ranges=named_ranges,
                                    named_range_ranges=named_range_ranges,
                                )
                                index_targets = infer_dynamic_index_targets(
                                    formula_for_infer,
                                    current_sheet=current_sheet,
                                    cell_type_env=expanded_env,
                                    limits=dynamic_refs.limits,
                                    bounds=bounds,
                                    named_ranges=named_ranges,
                                    named_range_ranges=named_range_ranges,
                                    current_row=_current_row,
                                    current_col=_current_col,
                                )
                            except DynamicRefError as exc:
                                cell_key = format_key(current_sheet, current_a1)
                                raise DynamicRefError(
                                    f"{exc} (while analyzing dynamic OFFSET/INDIRECT/INDEX for {cell_key}; "
                                    f"normalized formula {formula_for_infer!r})"
                                ) from exc
                            _dyn_cache[_cache_key] = (
                                offset_targets,
                                indirect_targets,
                                index_targets,
                            )
                        for addr in sort_node_keys(
                            offset_targets | indirect_targets | index_targets,
                            sheet_order=sheet_order,
                        ):
                            sh, a1 = parse_address(addr)
                            deps.append((sh, a1))
                        _dyn_dep_cache[_dyn_dep_cache_key] = (
                            deps[_deps_start:],
                            (offset_targets, indirect_targets, index_targets),
                        )
            masked = mask_spans(masked, dyn_spans)
            masked = mask_ref_only_function_calls(masked)

            # Expand ranges when requested, then always parse standalone cells
            # (range spans masked) so bare endpoints in single-prefix forms are
            # never attributed to the formula's local sheet.
            if expand_ranges:
                for start, end, _span in parse_range_refs_with_spans(masked):
                    sheet = start.sheet if start.sheet is not None else current_sheet
                    _ensure_sheet_bounds(sheet)
                    deps.extend(
                        expand_range_ref(
                            start=start,
                            end=end,
                            default_sheet=sheet,
                            max_cells=max_range_cells,
                            sheet_bounds=sheet_bounds,
                        )
                    )

            for ref in parse_standalone_cell_refs(masked):
                sh = ref.sheet if ref.sheet is not None else current_sheet
                deps.append((sh, f"{ref.column}{ref.row}"))

            # 3) Named ranges
            for m in _NAME_TOKEN_RE.finditer(masked):
                token = m.group(1)
                resolved = named_ranges.get(token)
                if resolved is not None:
                    deps.append(resolved)
                    continue
                resolved_range = named_range_ranges.get(token)
                if resolved_range is not None:
                    if expand_ranges:
                        sheet, start_a1, end_a1 = resolved_range
                        start_col, start_row = fastpyxl.utils.cell.coordinate_from_string(start_a1)
                        end_col, end_row = fastpyxl.utils.cell.coordinate_from_string(end_a1)
                        deps.extend(
                            expand_range(
                                sheet=sheet,
                                start_col=start_col,
                                start_row=int(start_row),
                                end_col=end_col,
                                end_row=int(end_row),
                                max_cells=max_range_cells,
                            )
                        )
                    continue
                if token in defined_names:
                    raise ValueError(f"Unsupported defined name referenced in formula: {token}")

            # Deduplicate while preserving order
            seen: set[tuple[str, str]] = set()
            out: list[tuple[str, str]] = []
            for d in deps:
                if d in seen:
                    continue
                seen.add(d)
                out.append(d)
            return _workbook_sorted_sheet_a1_pairs(out, sheet_order=sheet_order)

        def extract_array_if_deps(f: str) -> dict[tuple[str, str], GuardExpr | None] | None:
            """Extract per-element deps for an array-context `IF`, else `None`.

            In array context Excel evaluates `IF` element-wise over its range
            arguments, so element `i` of a value range is read only under element
            `i` of the condition. The condition is parsed once into a template
            (`RangeRef` placeholders) and instantiated per element; branch cells
            that are not element-aligned with the condition — differently shaped
            ranges, scalars, ranges under an aggregate — stay unconditional.

            Returns `None` when the form is not an array-context `IF` (no
            parseable range-typed condition), leaving the scalar handling below
            to run unchanged.
            """
            args = split_top_level_function(f, "IF")
            if args is None or len(args) not in (2, 3) or not args[0] or not args[1]:
                return None
            cond_s = args[0]
            template = parse_guard_expr(
                cond_s,
                current_sheet=current_sheet,
                named_ranges=named_ranges,
                allow_ranges=True,
            )
            if template is None:
                return None
            shape = guard_range_shape(template)
            if shape is None:
                return None

            out: dict[tuple[str, str], GuardExpr | None] = {}
            # The whole condition range is read to build the boolean array.
            for sh, a1 in extract_expr_deps(cond_s):
                _merge_guarded_dep(out, (sh, a1), None)

            def add_array_branch(branch_expr: str, *, negated: bool) -> None:
                if not branch_expr:
                    return
                aligned = element_aligned_range_cells(
                    branch_expr,
                    current_sheet=current_sheet,
                    shape=shape,
                    max_cells=max_range_cells,
                )
                for key, inner_guard in extract_expr_deps_guarded(
                    branch_expr, array_context=True
                ).items():
                    offset = aligned.get(key)
                    element_guard: GuardExpr | None = None
                    if offset is not None:
                        element_guard = instantiate_element_guard(
                            template, row_offset=offset[0], col_offset=offset[1]
                        )
                        if element_guard is not None and negated:
                            element_guard = Not(element_guard)
                    _merge_guarded_dep(out, key, _conjoin_guards(element_guard, inner_guard))

            add_array_branch(args[1], negated=False)
            if len(args) == 3:
                add_array_branch(args[2], negated=True)
            return out

        def extract_expr_deps_guarded(
            expr: str, *, array_context: bool = False
        ) -> dict[tuple[str, str], GuardExpr | None]:
            """Extract guarded deps from an expression, recursing into conditionals.

            Deps of a branch that is itself a conditional carry the conjunction of
            the enclosing branch guard and their own (inner) guard. Conditionals
            embedded inside a larger expression (arithmetic, aggregates, etc.) are
            scanned the same way; surrounding refs stay unconditional. A dep
            reachable through several branches gets the disjunction of its
            per-branch guards (`None` — unconditional/opaque — always wins).

            Args:
                expr: Expression text, with or without a leading `=`.
                array_context: Whether Excel evaluates `expr` element-wise (a CSE
                    array formula, or nesting inside an array-consuming call), in
                    which case a range-typed `IF` condition yields per-element
                    guards instead of one guard for the whole range.
            """
            f = expr if expr.startswith("=") else "=" + expr
            out: dict[tuple[str, str], GuardExpr | None] = {}

            def add_branch(branch_expr: str, branch_guard: GuardExpr | None) -> None:
                for key, inner_guard in extract_expr_deps_guarded(
                    branch_expr, array_context=array_context
                ).items():
                    _merge_guarded_dep(out, key, _conjoin_guards(branch_guard, inner_guard))

            # 0) Array-context IF(range_condition, ...): per-element guards.
            if array_context:
                array_out = extract_array_if_deps(f)
                if array_out is not None:
                    return array_out

            # 1) IF(cond, then, else)
            if_parts = split_top_level_if(f)
            if if_parts is not None:
                cond_s, then_s, else_s = if_parts
                cond_guard = parse_guard_expr(
                    cond_s, current_sheet=current_sheet, named_ranges=named_ranges
                )
                for sh, a1 in extract_expr_deps(cond_s):
                    _merge_guarded_dep(out, (sh, a1), None)

                # If the condition can't be parsed, branch deps are still conditional,
                # but opaque.
                add_branch(then_s, cond_guard)
                if else_s:
                    add_branch(else_s, None if cond_guard is None else Not(cond_guard))
                return out

            # 2) IFS(cond1, value1, cond2, value2, ..., [default])
            ifs_args = split_top_level_ifs(f)
            if ifs_args is not None:
                conditions: list[str] = []
                values: list[str] = []
                ifs_default: str | None = None
                if len(ifs_args) >= 2:
                    pairs = ifs_args
                    if len(pairs) % 2 == 1:
                        ifs_default = pairs[-1]
                        pairs = pairs[:-1]
                    for i in range(0, len(pairs), 2):
                        conditions.append(pairs[i])
                        values.append(pairs[i + 1])

                # All condition expressions may be evaluated (sequentially), so include
                # deps from all conditions as unconditional to avoid missing required
                # inputs.
                for c in conditions:
                    for sh, a1 in extract_expr_deps(c):
                        _merge_guarded_dep(out, (sh, a1), None)

                prev_negations: list[GuardExpr] = []
                for cond_s, val_s in zip(conditions, values, strict=False):
                    cond_guard = parse_guard_expr(
                        cond_s, current_sheet=current_sheet, named_ranges=named_ranges
                    )
                    # Build sequential guard: cond_i AND NOT(cond_1) AND ... NOT(cond_{i-1})
                    g: GuardExpr | None
                    if cond_guard is None:
                        g = None
                    else:
                        ops: list[GuardExpr] = [cond_guard, *prev_negations]
                        g = ops[0] if len(ops) == 1 else And(tuple(ops))
                        prev_negations.append(Not(cond_guard))
                    add_branch(val_s, g)

                if ifs_default is not None:
                    add_branch(ifs_default, _sequential_default_guard(prev_negations))
                return out

            # 3) CHOOSE(index, value1, value2, ...)
            choose_args = split_top_level_choose(f)
            if choose_args is not None and len(choose_args) >= 2:
                index_s = choose_args[0]
                index_expr = parse_guard_expr(
                    index_s, current_sheet=current_sheet, named_ranges=named_ranges
                )
                for sh, a1 in extract_expr_deps(index_s):
                    _merge_guarded_dep(out, (sh, a1), None)

                for i, choice_s in enumerate(choose_args[1:], start=1):
                    guard: GuardExpr | None = None
                    if index_expr is not None:
                        guard = Compare(left=index_expr, op="=", right=Literal(i))
                    add_branch(choice_s, guard)
                return out

            # 4) SWITCH(expr, value1, result1, ..., [default])
            switch_args = split_top_level_switch(f)
            if switch_args is not None and len(switch_args) >= 3:
                expr_s = switch_args[0]
                expr_ge = parse_guard_expr(
                    expr_s, current_sheet=current_sheet, named_ranges=named_ranges
                )
                for sh, a1 in extract_expr_deps(expr_s):
                    _merge_guarded_dep(out, (sh, a1), None)

                switch_pairs = switch_args[1:]
                switch_default: str | None = None
                if len(switch_pairs) % 2 == 1:
                    switch_default = switch_pairs[-1]
                    switch_pairs = switch_pairs[:-1]

                match_negations: list[GuardExpr] = []
                for i in range(0, len(switch_pairs), 2):
                    val_s = switch_pairs[i]
                    res_s = switch_pairs[i + 1]
                    val_ge = parse_guard_expr(
                        val_s, current_sheet=current_sheet, named_ranges=named_ranges
                    )
                    match: GuardExpr | None = None
                    if expr_ge is not None and val_ge is not None:
                        match = Compare(left=expr_ge, op="=", right=val_ge)

                    case_guard: GuardExpr | None = None
                    if match is not None:
                        ops2: list[GuardExpr] = [match, *match_negations]
                        case_guard = ops2[0] if len(ops2) == 1 else And(tuple(ops2))
                        match_negations.append(Not(match))
                    add_branch(res_s, case_guard)

                if switch_default is not None:
                    add_branch(switch_default, _sequential_default_guard(match_negations))
                return out

            # 5) Conditionals embedded in a larger expression (e.g. `1+IF(...)`,
            # `SUM(IF(...),E1)`). Recurse into each outermost call for guarded
            # deps, then treat remaining surrounding refs as unconditional.
            embedded_spans = _outermost_embedded_conditional_spans(f)
            if embedded_spans:
                array_spans = _spans_in_array_context(f)
                for span in embedded_spans:
                    start, end = span
                    nested_array_context = array_context or any(
                        _span_contains(outer, span) for outer in array_spans
                    )
                    for key, guard in extract_expr_deps_guarded(
                        f[start:end], array_context=nested_array_context
                    ).items():
                        _merge_guarded_dep(out, key, guard)
                for sh, a1 in extract_expr_deps(mask_spans(f, embedded_spans)):
                    _merge_guarded_dep(out, (sh, a1), None)
                return out

            for sh, a1 in extract_expr_deps(f):
                _merge_guarded_dep(out, (sh, a1), None)
            return out

        return _sorted_guard_deps(
            extract_expr_deps_guarded(formula, array_context=array_formula),
            sheet_order=sheet_order,
        )

    visited: set[str] = set()
    q: deque[tuple[str, str, int]] = deque()
    target_roots = _expand_targets_to_roots(
        targets,
        sheetnames=list(wb_formulas.sheetnames),
        named_ranges=named_ranges,
        named_range_ranges=named_range_ranges,
        max_range_cells=max_range_cells,
    )
    target_root_keys = {format_key(sh, a1) for sh, a1 in target_roots}
    for sh, a1 in target_roots:
        q.append((sh, a1, 0))

    _bfs_t0 = time.perf_counter()
    _bfs_count = 0
    _bfs_next_log = 5000
    formula_ast_intern: dict[AstNode, AstNode] = {}

    try:
        while q:
            sheet, a1, depth = q.popleft()
            key = format_key(sheet, a1)
            if key in visited:
                continue
            visited.add(key)
            _bfs_count += 1
            if _bfs_count >= _bfs_next_log:
                _emit_trace(
                    DynamicRefTraceEvent(
                        kind="bfs-progress",
                        name="create_dependency_graph",
                        elapsed_s=time.perf_counter() - _bfs_t0,
                        detail={
                            "nodes": _bfs_count,
                            "queue": len(q),
                            "depth": depth,
                            "last": key,
                            "infer_calls": _dyn_stats["infer_calls"],
                            "cache_hits": _dyn_stats["cache_hits"],
                            "dep_cache_hits": _dyn_stats["dep_cache_hits"],
                            "env_cache_hits": _dyn_stats["env_cache_hits"],
                            "arg_subgraph_hits": _dyn_stats["arg_subgraph_hits"],
                            "env_cache_size": len(_shared_cell_type_cache),
                        },
                    )
                )
                _bfs_next_log += 5000
            if depth > max_depth:
                continue

            if blank_rects:
                col_str, row_i = fastpyxl.utils.cell.coordinate_from_string(a1)
                col_idx = fastpyxl.utils.cell.column_index_from_string(col_str)
                if cell_in_blank_ranges(sheet, int(row_i), col_idx, blank_rects):
                    continue

            ws_f = _get_ws_f(sheet)
            cell = ws_f[a1]
            raw = cell.value
            # CSE array formulas are evaluated element-wise, which is one of the
            # array contexts that make range-typed `IF` conditions per-element.
            is_array_formula = isinstance(raw, ArrayFormula)
            if is_array_formula:
                raw = raw.text or ""
                if raw and not raw.startswith("="):
                    raw = f"={raw}"
            is_formula = isinstance(raw, str) and raw.startswith("=")

            if is_formula:
                formula_str = str(raw)
                formula = formula_str if store_raw_formula else None
                normalized = normalizer.normalize(formula_str, sheet)
                value = None
                if load_values:
                    value = _cached_value_from_formula_cell(cell)
                is_leaf = False
                formula_ast = parse_preserving_axes_optional(
                    formula_str,
                    anchor=CellKey(key),
                    named_ranges=named_ranges,
                    named_range_ranges=named_range_ranges,
                )
                if formula_ast is not None:
                    formula_ast = intern_formula_ast(formula_ast, formula_ast_intern)
            else:
                formula_str = ""
                formula = None
                normalized = None
                value = raw
                is_leaf = True
                formula_ast = None

            col, row = fastpyxl.utils.cell.coordinate_from_string(a1)
            col_idx = fastpyxl.utils.cell.column_index_from_string(col)
            node = Node(
                sheet=sheet,
                column=col,
                row=int(row),
                formula=formula,
                normalized_formula=normalized,
                value=value,
                is_leaf=is_leaf,
                is_target=key in target_root_keys,
                formula_ast=formula_ast,
            )
            graph.add_node(node)

            if not is_formula:
                continue

            # Run extraction first so that the constraint-based dynamic-ref expansion
            # (_dyn_cache) is populated before provenance collection reads from it.
            deps_and_guards = extract_deps_with_guards(
                formula_str, sheet, a1, array_formula=is_array_formula
            )

            prov_map: dict[str, EdgeProvenance] | None = None
            if capture_dependency_provenance:
                prov_map = collect_provenance_for_formula(
                    formula_str,
                    normalized_formula=normalized,
                    current_sheet=sheet,
                    current_a1=a1,
                    named_ranges=named_ranges,
                    named_range_ranges=named_range_ranges,
                    normalizer=normalizer,
                    defined_names=defined_names,
                    expand_ranges=expand_ranges,
                    max_range_cells=max_range_cells,
                    use_cached_dynamic_refs=use_cached_dynamic_refs,
                    dynamic_refs=dynamic_refs,
                    wb_formulas=wb_formulas,
                    resolve_cached_value=resolve_cached_value,
                    dynamic_expansion_cache=_dyn_cache,
                    type_analysis_cache=type_analysis_cache,
                    workbook_sha256=_wb_sha256,
                    ref_walk=ref_walk,
                )

            for dep_sheet, dep_a1, guard in deps_and_guards:
                dep_key = format_key(dep_sheet, dep_a1)
                if prov_map is not None:
                    p = prov_map.get(dep_key)
                    if p is None:
                        p = EdgeProvenance.empty()
                    graph.add_edge(key, dep_key, guard=guard, provenance=p)
                else:
                    graph.add_edge(key, dep_key, guard=guard)
                if dep_key not in visited:
                    if dep_sheet not in wb_formulas.sheetnames:
                        continue
                    q.append((dep_sheet, dep_a1, depth + 1))

            if not graph.get_dependencies(key):
                node.is_leaf = True
    finally:
        if (
            _dyn_stats["infer_calls"]
            or _dyn_stats["cache_hits"]
            or _dyn_stats["dep_cache_hits"]
            or _dyn_stats["env_cache_hits"]
            or _dyn_stats["arg_subgraph_hits"]
        ):
            _emit_trace(
                DynamicRefTraceEvent(
                    kind="bfs-done",
                    name="create_dependency_graph",
                    elapsed_s=time.perf_counter() - _bfs_t0,
                    detail={
                        "nodes": _bfs_count,
                        "infer_calls": _dyn_stats["infer_calls"],
                        "cache_hits": _dyn_stats["cache_hits"],
                        "dep_cache_hits": _dyn_stats["dep_cache_hits"],
                        "env_cache_hits": _dyn_stats["env_cache_hits"],
                        "arg_subgraph_hits": _dyn_stats["arg_subgraph_hits"],
                        "env_cache_size": len(_shared_cell_type_cache),
                    },
                )
            )
        if wb_values is not None:
            wb_values.close()
        if not isinstance(workbook, fastpyxl.Workbook):
            wb_formulas.close()

    graph.named_ranges = dict(named_ranges)
    graph.named_range_ranges = dict(named_range_ranges)
    graph.sheet_bounds = dict(sheet_bounds)
    if warm_ast_cache or warm_formula_shapes:
        from .formula_shapes import warm_formula_shapes as intern_graph_formula_shapes
        from .preparsed_formulas import warm_preparsed_formulas

        parsed = warm_preparsed_formulas(graph)
        if warm_ast_cache:
            graph.preparsed_formulas = parsed
        if warm_formula_shapes:
            graph.formula_shapes = intern_graph_formula_shapes(graph, parsed=parsed)
    return graph


def list_dynamic_ref_constraint_candidates(
    workbook: Path | str | fastpyxl.Workbook,
    targets: Iterable[str],
    *,
    dynamic_refs: DynamicRefConfig | None = None,
    max_depth: int = 50,
    max_range_cells: int = 5000,
    type_analysis_cache: TypeAnalysisCache | None = None,
) -> list[str]:
    """Return sorted leaf cells missing dynamic-ref constraint entries.

    These are leaf cell addresses that feed dynamic-ref arguments
    (OFFSET/INDIRECT/INDEX) but have no entry in `dynamic_refs.cell_type_env`.

    Unlike `create_dependency_graph`, this function does **not** raise
    `DynamicRefError` when constraints are missing.  Instead it collects all
    missing leaf addresses in a single pass and returns them sorted.

    When `dynamic_refs` is `None` the function treats it as an empty constraint
    environment: all leaf cells that feed dynamic-ref arguments are returned as
    candidates.

    **Completeness caveat**: Cells reachable only through unresolvable dynamic refs
    will not be visited, so their constraint candidates won't appear in the output.
    A second call after adding the first batch of constraints will quickly find any
    remaining missing entries.
    """
    if isinstance(workbook, fastpyxl.Workbook):
        wb_formulas = workbook
        _owns_wb = False
    else:
        path = Path(workbook)
        keep_vba = path.suffix.lower() == ".xlsm"
        wb_formulas = fastpyxl.load_workbook(path, data_only=False, keep_vba=keep_vba)
        _owns_wb = True

    _wb_sha256_cand: str | None = None
    if type_analysis_cache is not None and isinstance(workbook, (str, Path)):
        with open(workbook, "rb") as _f:
            _wb_sha256_cand = hashlib.file_digest(_f, "sha256").hexdigest()

    # Shared across every dynamic-ref call site so intermediates in overlapping
    # argument subgraphs are inferred once, as in `create_dependency_graph`.
    _shared_cell_type_cache_cand: dict[str, CellType] = {}
    # Worksheet cache: avoid repeated O(#sheets) __getitem__ scans (issue #484).
    _ws_f_cache: dict[str, Worksheet] = {}
    # Memoize static-ref extraction across candidate BFS / argument walks.
    # Store frozensets and return a fresh mutable set on every call so callers
    # that mutate in place (e.g. expand_leaf_env_to_argument_env's `refs |= …`)
    # cannot poison later lookups.
    _refs_cache: dict[tuple[str, str], frozenset[str]] = {}
    _arg_subgraph_cache_cand: dict[frozenset[str], tuple[frozenset[str], frozenset[str]]] = {}
    _arg_node_cache_cand: dict[str, tuple[frozenset[str], bool]] = {}

    def _get_ws_f(sheet: str) -> Worksheet:
        ws = _ws_f_cache.get(sheet)
        if ws is None:
            ws = wb_formulas[sheet]
            _ws_f_cache[sheet] = ws
        return ws

    def _cell_value(sheet: str, a1: str) -> object | None:
        return _get_ws_f(sheet)[a1].value

    try:
        named_range_maps = build_named_range_map(wb_formulas)
        named_ranges = named_range_maps.cell_map
        named_range_ranges = named_range_maps.range_map
        normalizer = FormulaNormalizer(named_ranges, named_range_ranges)
        cell_type_env = {} if dynamic_refs is None else dynamic_refs.cell_type_env
        sheetnames = list(wb_formulas.sheetnames)
        sheetname_set = set(sheetnames)
        _NAME_TOKEN_RE = re.compile(r"\b([A-Za-z_][A-Za-z0-9_]*)\b(?!\s*!)")

        def _refs_without_dynamic(formula_str: str, sheet: str) -> set[str]:
            """Static (non-dynamic-ref) cell addresses referenced by *formula_str*."""
            f = formula_str if formula_str.startswith("=") else "=" + formula_str
            cache_key = (f, sheet)
            cached = _refs_cache.get(cache_key)
            if cached is not None:
                return set(cached)
            dyn = _find_function_calls_with_spans(f, frozenset({"OFFSET", "INDIRECT", "INDEX"}))
            spans = [span for _fn, _inner, span in dyn]
            masked = mask_spans(f, spans)
            masked = mask_ref_only_function_calls(masked)
            norm = normalizer.normalize(masked, sheet)
            out: set[str] = set()
            for ref in parse_standalone_cell_refs(norm):
                sh = ref.sheet if ref.sheet is not None else sheet
                out.add(format_key(sh, f"{ref.column}{ref.row}"))
            for start, end, _span in parse_range_refs_with_spans(norm):
                sh = start.sheet if start.sheet is not None else sheet
                for dep_sheet, dep_a1 in expand_range(
                    sheet=sh,
                    start_col=start.column,
                    start_row=start.row,
                    end_col=end.column,
                    end_row=end.row,
                    max_cells=max_range_cells,
                ):
                    out.add(format_key(dep_sheet, dep_a1))
            for m in _NAME_TOKEN_RE.finditer(norm):
                token = m.group(1)
                resolved = named_ranges.get(token)
                if resolved is not None:
                    out.add(format_key(resolved[0], resolved[1]))
                    continue
                resolved_range = named_range_ranges.get(token)
                if resolved_range is not None:
                    rsh, start_a1, end_a1 = resolved_range
                    s_col, s_row = fastpyxl.utils.cell.coordinate_from_string(start_a1)
                    e_col, e_row = fastpyxl.utils.cell.coordinate_from_string(end_a1)
                    for dep_sheet, dep_a1 in expand_range(
                        sheet=rsh,
                        start_col=s_col,
                        start_row=int(s_row),
                        end_col=e_col,
                        end_row=int(e_row),
                        max_cells=max_range_cells,
                    ):
                        out.add(format_key(dep_sheet, dep_a1))
            _refs_cache[cache_key] = frozenset(out)
            return out

        collected: set[str] = set()
        visited: set[str] = set()
        queue: deque[tuple[str, str, int]] = deque()

        target_roots = _expand_targets_to_roots(
            targets,
            sheetnames=sheetnames,
            named_ranges=named_ranges,
            named_range_ranges=named_range_ranges,
            max_range_cells=max_range_cells,
        )
        for sh, a1 in target_roots:
            queue.append((sh, a1, 0))

        _bfs_t0 = time.perf_counter()
        _bfs_count = 0
        _bfs_next_log = _CANDIDATES_BFS_PROGRESS_INTERVAL
        _arg_walk_count = 0
        _arg_next_log = _CANDIDATES_ARG_PROGRESS_INTERVAL

        while queue:
            current_sheet, current_a1, depth = queue.popleft()
            key = format_key(current_sheet, current_a1)
            if key in visited:
                continue
            visited.add(key)
            _bfs_count += 1
            if _bfs_count >= _bfs_next_log:
                _emit_trace(
                    DynamicRefTraceEvent(
                        kind="bfs-progress",
                        name="list_dynamic_ref_constraint_candidates",
                        elapsed_s=time.perf_counter() - _bfs_t0,
                        detail={
                            "nodes": _bfs_count,
                            "queue": len(queue),
                            "depth": depth,
                            "last": key,
                            "collected": len(collected),
                            "refs_cache_size": len(_refs_cache),
                            "arg_visited": _arg_walk_count,
                        },
                    )
                )
                _bfs_next_log += _CANDIDATES_BFS_PROGRESS_INTERVAL

            if depth >= max_depth or current_sheet not in sheetname_set:
                continue

            cell_val = _cell_value(current_sheet, current_a1)
            if isinstance(cell_val, ArrayFormula):
                cell_val = cell_val.text or ""
                if cell_val and not cell_val.startswith("="):
                    cell_val = f"={cell_val}"
            if not isinstance(cell_val, str) or not cell_val.startswith("="):
                continue  # leaf cell — nothing to do

            f = cell_val

            # Find dynamic calls (OFFSET/INDIRECT/INDEX), filter static INDEX.
            calls = _find_function_calls_with_spans(f, frozenset({"OFFSET", "INDIRECT", "INDEX"}))
            dynamic_calls = []
            for fn_name, inner, span in calls:
                if fn_name == "INDEX":
                    idx_args = _split_function_args(inner)
                    if idx_args is not None and len(idx_args) >= 2:
                        has_non_literal = False
                        for j, idx_arg in enumerate(idx_args):
                            if j == 0:
                                continue  # skip array arg
                            try:
                                float(idx_arg.strip())
                            except ValueError:
                                has_non_literal = True
                                break
                        if not has_non_literal:
                            continue  # static INDEX — skip
                dynamic_calls.append((fn_name, inner, span))

            if dynamic_calls:
                # Collect variable argument addresses for leaf discovery.
                argument_addrs: set[str] = set()
                for fn_name, inner, _span in dynamic_calls:
                    args = _split_function_args(inner)
                    if args is None:
                        continue
                    for i, arg in enumerate(args):
                        normalized_arg = normalizer.normalize("=" + arg, current_sheet)
                        is_variable = (
                            (fn_name == "OFFSET" and i >= 1)
                            or (fn_name == "OFFSET" and i == 0 and "(" in normalized_arg)
                            or fn_name == "INDIRECT"
                            or (fn_name == "INDEX" and i >= 1)
                        )
                        if is_variable:
                            value_arg = mask_ref_only_function_calls(normalized_arg)
                            for ref in parse_standalone_cell_refs(value_arg):
                                sh = ref.sheet if ref.sheet is not None else current_sheet
                                argument_addrs.add(format_key(sh, f"{ref.column}{ref.row}"))
                            for start, end, _span in parse_range_refs_with_spans(value_arg):
                                range_sheet = (
                                    start.sheet if start.sheet is not None else current_sheet
                                )
                                for dep_sheet, dep_a1 in expand_range(
                                    sheet=range_sheet,
                                    start_col=start.column,
                                    start_row=start.row,
                                    end_col=end.column,
                                    end_row=end.row,
                                    max_cells=max_range_cells,
                                ):
                                    argument_addrs.add(format_key(dep_sheet, dep_a1))

                # Walk argument_addrs to statically-reachable leaves.
                arg_key = frozenset(argument_addrs)
                cached_subgraph = _arg_subgraph_cache_cand.get(arg_key)
                if cached_subgraph is not None:
                    all_refs, leaves = set(cached_subgraph[0]), set(cached_subgraph[1])
                else:
                    all_refs = set()
                    leaves = set()
                    to_visit_inner = set(argument_addrs)
                    while to_visit_inner:
                        addr = to_visit_inner.pop()
                        if addr in all_refs:
                            continue
                        all_refs.add(addr)
                        node = _arg_node_cache_cand.get(addr)
                        if node is None:
                            _arg_walk_count += 1
                            if _arg_walk_count >= _arg_next_log:
                                _emit_trace(
                                    DynamicRefTraceEvent(
                                        kind="candidates-arg-progress",
                                        name="list_dynamic_ref_constraint_candidates",
                                        elapsed_s=time.perf_counter() - _bfs_t0,
                                        detail={
                                            "visited": _arg_walk_count,
                                            "pending": len(to_visit_inner),
                                            "last": addr,
                                            "bfs_nodes": _bfs_count,
                                            "refs_cache_size": len(_refs_cache),
                                        },
                                    )
                                )
                                _arg_next_log += _CANDIDATES_ARG_PROGRESS_INTERVAL
                            children: frozenset[str] = frozenset()
                            is_leaf = False
                            sh, a1 = parse_address(addr)
                            if sh in sheetname_set:
                                inner_val = _cell_value(sh, a1)
                                if isinstance(inner_val, str) and inner_val.startswith("="):
                                    children = frozenset(_refs_without_dynamic(inner_val, sh))
                                else:
                                    is_leaf = True
                            node = (children, is_leaf)
                            _arg_node_cache_cand[addr] = node
                        children, is_leaf = node
                        if is_leaf:
                            leaves.add(addr)
                        else:
                            to_visit_inner.update(children)
                    _arg_subgraph_cache_cand[arg_key] = (
                        frozenset(all_refs),
                        frozenset(leaves),
                    )

                missing = leaves_missing_cell_type_constraints(leaves, cell_type_env)
                if missing:
                    collected.update(missing)
                    # Skip infer — dynamic targets unknown without full constraints.
                elif dynamic_refs is not None:
                    # All leaves constrained — run infer to discover dynamic targets.
                    try:
                        bounds = GlobalWorkbookBounds(sheet=current_sheet)
                        formula_for_infer = normalizer.normalize(f, current_sheet)
                        _col_letter, _current_row = fastpyxl.utils.cell.coordinate_from_string(
                            current_a1
                        )
                        _current_col = fastpyxl.utils.cell.column_index_from_string(_col_letter)

                        if all_refs and all(
                            addr in _shared_cell_type_cache_cand for addr in all_refs
                        ):
                            expanded_env = _shared_cell_type_cache_cand
                        else:

                            def _get_cell_formula(addr: str) -> str | None:
                                sh2, a1_2 = parse_address(addr)
                                if sh2 not in sheetname_set:
                                    return None
                                v = _cell_value(sh2, a1_2)
                                if not isinstance(v, str) or not v.startswith("="):
                                    return None
                                return normalizer.normalize(v, sh2)

                            def _get_cell_ast(addr: str) -> AstNode | None:
                                sh2, a1_2 = parse_address(addr)
                                if sh2 not in sheetname_set:
                                    return None
                                v = _cell_value(sh2, a1_2)
                                if not isinstance(v, str) or not v.startswith("="):
                                    return None
                                return parse_preserving_axes_optional(
                                    v,
                                    anchor=addr,
                                    named_ranges=named_ranges,
                                    named_range_ranges=named_range_ranges,
                                )

                            expanded_env = expand_leaf_env_to_argument_env(
                                all_refs,
                                _get_cell_formula,
                                _refs_without_dynamic,
                                dynamic_refs.cell_type_env,
                                dynamic_refs.limits,
                                named_ranges=named_ranges,
                                named_range_ranges=named_range_ranges,
                                max_range_cells=max_range_cells,
                                shared_cell_type_cache=_shared_cell_type_cache_cand,
                                type_analysis_cache=type_analysis_cache,
                                workbook_sha256=_wb_sha256_cand,
                                get_cell_ast=_get_cell_ast,
                            )
                        offset_targets = infer_dynamic_offset_targets(
                            formula_for_infer,
                            current_sheet=current_sheet,
                            cell_type_env=expanded_env,
                            limits=dynamic_refs.limits,
                            bounds=bounds,
                            named_ranges=named_ranges,
                            named_range_ranges=named_range_ranges,
                            current_row=_current_row,
                            current_col=_current_col,
                        )
                        indirect_targets = infer_dynamic_indirect_targets(
                            formula_for_infer,
                            current_sheet=current_sheet,
                            cell_type_env=expanded_env,
                            limits=dynamic_refs.limits,
                            bounds=bounds,
                            named_ranges=named_ranges,
                            named_range_ranges=named_range_ranges,
                        )
                        index_targets = infer_dynamic_index_targets(
                            formula_for_infer,
                            current_sheet=current_sheet,
                            cell_type_env=expanded_env,
                            limits=dynamic_refs.limits,
                            bounds=bounds,
                            named_ranges=named_ranges,
                            named_range_ranges=named_range_ranges,
                            current_row=_current_row,
                            current_col=_current_col,
                        )
                        for addr in sort_node_keys(
                            offset_targets | indirect_targets | index_targets,
                            sheet_order=sheetnames,
                        ):
                            sh, a1 = parse_address(addr)
                            queue.append((sh, a1, depth + 1))
                    except DynamicRefError:
                        pass  # best-effort: skip dynamic targets for this formula

            # Queue static (non-dynamic-ref) deps for continued BFS.
            for addr in _refs_without_dynamic(f, current_sheet):
                sh, a1 = parse_address(addr)
                if sh in sheetname_set:
                    queue.append((sh, a1, depth + 1))

        _emit_trace(
            DynamicRefTraceEvent(
                kind="bfs-done",
                name="list_dynamic_ref_constraint_candidates",
                elapsed_s=time.perf_counter() - _bfs_t0,
                detail={
                    "nodes": _bfs_count,
                    "collected": len(collected),
                    "refs_cache_size": len(_refs_cache),
                    "arg_visited": _arg_walk_count,
                },
            )
        )

    finally:
        if _owns_wb:
            wb_formulas.close()

    return sorted(collected)
