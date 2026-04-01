from __future__ import annotations

import hashlib
import logging
import re
import time
from collections import deque
from collections.abc import Iterable
from pathlib import Path

import fastpyxl
import fastpyxl.utils.cell
from fastpyxl.worksheet.formula import ArrayFormula
from fastpyxl.worksheet.worksheet import Worksheet

from excel_grapher.core.cell_types import CellType, leaves_missing_cell_type_constraints

from .blank_ranges import (
    address_in_blank_ranges,
    cell_in_blank_ranges,
    normalize_blank_range_specs,
)
from .dependency_provenance import EdgeProvenance
from .dynamic_refs import (
    DynamicRefConfig,
    DynamicRefError,
    GlobalWorkbookBounds,
    clear_index_target_cache,
    expand_leaf_env_to_argument_env,
    infer_dynamic_index_targets,
    infer_dynamic_indirect_targets,
    infer_dynamic_offset_targets,
)
from .graph import DependencyGraph, NodeHook
from .guard import And, Compare, GuardExpr, Literal, Not
from .node import Node
from .parser import (
    FormulaNormalizer,
    _find_function_calls_with_spans,
    _split_function_args,
    expand_range,
    format_key,
    mask_spans,
    parse_cell_refs,
    parse_dynamic_range_refs_with_spans,
    parse_guard_expr,
    parse_range_refs_with_spans,
    split_top_level_choose,
    split_top_level_if,
    split_top_level_ifs,
    split_top_level_switch,
)
from .provenance_collect import collect_provenance_for_formula
from .resolver import build_named_range_map
from .type_analysis_cache import TypeAnalysisCache

_logger = logging.getLogger(__name__)


def _parse_address_to_sheet_a1(addr: str) -> tuple[str, str]:
    """Parse a sheet-qualified address (e.g. Sheet1!A1 or 'My Sheet'!A1) to (sheet, a1)."""
    if "!" not in addr:
        raise ValueError(f"Address must be sheet-qualified: {addr}")
    if addr.startswith("'"):
        end_quote = addr.index("'", 1)
        sheet = addr[1:end_quote]
        a1 = addr[end_quote + 2 :]
        return sheet, a1
    sheet, a1 = addr.split("!", 1)
    return sheet, a1


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
            sheet, a1 = _parse_address_to_sheet_a1(addr)
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
                    parts.append(f"{sheet}!{c_start_letter}{r1}")
                elif c0 == c_last:
                    parts.append(f"{sheet}!{c_start_letter}{r1}:{sheet}!{c_end_letter}{r2}")
                else:
                    parts.append(f"{sheet}!{c_start_letter}{r1}:{sheet}!{c_end_letter}{r2}")
            i = j

    parts.extend(sorted(others))
    return sorted(parts)


def create_dependency_graph(
    workbook: Path | str | fastpyxl.Workbook,
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
    blank_ranges: Iterable[str] | None = None,
    type_analysis_cache: TypeAnalysisCache | None = None,
) -> DependencyGraph:
    """
    Build a dependency graph starting from target cells.

    Supports basic A1 references, sheet-qualified references, and dynamic references
    (OFFSET/INDIRECT). For OFFSET/INDIRECT:

    - **use_cached_dynamic_refs=True**: Resolve using cached workbook values (existing path).
      ``dynamic_refs`` is ignored.
    - **use_cached_dynamic_refs=False** (default), **dynamic_refs=None**: On any formula
      that contains OFFSET or INDIRECT requiring resolution, raise :exc:`DynamicRefError`.
      Callers can pass a ``DynamicRefConfig`` or set ``use_cached_dynamic_refs=True`` to avoid.
    - **use_cached_dynamic_refs=False**, **dynamic_refs** set: Resolve OFFSET/INDIRECT via
      the config's ``cell_type_env`` and ``limits``; missing or invalid domains raise
      :exc:`DynamicRefError`.

    To build a config from a TypedDict of constraints, use
    :meth:`DynamicRefConfig.from_constraints`.

    When ``capture_dependency_provenance`` is True, each edge stores merged
    :class:`~excel_grapher.grapher.dependency_provenance.EdgeProvenance` under the
    ``\"provenance\"`` key in :meth:`DependencyGraph.edge_attrs` (how the dependency
    arises: direct reference, static range, dynamic OFFSET/INDIRECT).

    ``blank_ranges`` is an optional iterable of sheet-qualified A1 rectangles
    (e.g. ``\"Sheet1!B2:D10\"``) treated as structurally empty: no nodes are
    created for those cells (edges into them are kept), and dynamic-ref leaf
    constraints are not required for addresses inside these ranges. Pair with the
    same declarations on :class:`~excel_grapher.FormulaEvaluator` and
    :meth:`~excel_grapher.evaluator.codegen.CodeGenerator.generate` for evaluation
    and export parity.

    **Cost model**: constraint-based dynamic-ref expansion (``dynamic_refs`` set,
    ``use_cached_dynamic_refs=False``) runs :func:`expand_leaf_env_to_argument_env`
    once per formula regardless of ``capture_dependency_provenance``.  A shared
    per-graph cache ensures provenance collection reuses the already-computed
    expansion instead of repeating it.  Callers doing iterative constraint-tuning
    workflows can still set ``capture_dependency_provenance=False`` to avoid any
    provenance overhead (formula-string span collection, branch-union merging, etc.).
    """

    blank_rects = normalize_blank_range_specs(blank_ranges)

    def load_wb(data_only: bool) -> fastpyxl.Workbook:
        if isinstance(workbook, fastpyxl.Workbook):
            if data_only:
                raise ValueError(
                    "load_values=True is not supported when passing a Workbook instance"
                )
            return workbook
        path = Path(workbook)
        keep_vba = path.suffix.lower() == ".xlsm"
        return fastpyxl.load_workbook(path, data_only=data_only, keep_vba=keep_vba)

    _t0 = time.perf_counter()
    wb_formulas = load_wb(data_only=False)
    print(f"[DIAG] Loaded formula workbook in {time.perf_counter() - _t0:.2f}s", flush=True)
    _t0 = time.perf_counter()
    wb_values = (
        load_wb(data_only=True)
        if load_values and not isinstance(workbook, fastpyxl.Workbook)
        else None
    )
    if wb_values is not None:
        print(f"[DIAG] Loaded value workbook in {time.perf_counter() - _t0:.2f}s", flush=True)

    # Compute workbook SHA-256 for persistent type-analysis cache
    _wb_sha256: str | None = None
    if type_analysis_cache is not None and isinstance(workbook, (str, Path)):
        with open(workbook, "rb") as _f:
            _wb_sha256 = hashlib.file_digest(_f, "sha256").hexdigest()

    graph = DependencyGraph()
    for h in hooks or []:
        graph.register_hook(h)

    named_range_maps = build_named_range_map(wb_formulas)
    named_ranges = named_range_maps.cell_map
    named_range_ranges = named_range_maps.range_map
    normalizer = FormulaNormalizer(named_ranges, named_range_ranges)
    defined_names: set[str] = {str(name) for name in wb_formulas.defined_names}
    # Clear per-graph-build caches from previous invocations.
    clear_index_target_cache()

    # Per-graph cache: (normalized_formula, current_sheet, current_a1) → (offset_targets, indirect_targets, index_targets).
    # Populated by extract_expr_deps (constraint path); consumed by collect_provenance_for_formula
    # to avoid re-running the expensive expand_leaf_env_to_argument_env call.
    _dyn_cache: dict[tuple[str, str, str], tuple[set[str], set[str], set[str]]] = {}
    # Shared cell-type cache for expand_leaf_env_to_argument_env: intermediate
    # formula cells inferred once are reused across BFS nodes, avoiding redundant
    # recursive domain inference when many dynamic-ref formulas share intermediates.
    _shared_cell_type_cache: dict[str, CellType] = {}
    _NAME_TOKEN_RE = re.compile(r"\b([A-Za-z_][A-Za-z0-9_]*)\b(?!\s*!)")

    # Worksheet caches: avoid repeated O(#sheets) __getitem__ scans on every BFS node.
    _ws_f_cache: dict[str, Worksheet] = {}
    _ws_v_cache: dict[str, Worksheet] = {}

    def _get_ws_f(sheet: str) -> Worksheet:
        ws = _ws_f_cache.get(sheet)
        if ws is None:
            ws = wb_formulas[sheet]
            _ws_f_cache[sheet] = ws
        return ws

    def _get_ws_v(sheet: str) -> Worksheet:
        # Only called when wb_values is not None.
        ws = _ws_v_cache.get(sheet)
        if ws is None:
            assert wb_values is not None
            ws = wb_values[sheet]
            _ws_v_cache[sheet] = ws
        return ws

    def resolve_cached_value(sheet: str, a1: str) -> object | None:
        nonlocal wb_values
        if wb_values is None and not isinstance(workbook, fastpyxl.Workbook):
            wb_values = load_wb(data_only=True)
        if wb_values is None:
            return None
        return _get_ws_v(sheet)[a1].value

    def parse_target(t: str) -> tuple[str, str]:
        if "!" not in t:
            raise ValueError(f"Target must be sheet-qualified: {t}")
        # Handle quoted sheet names: 'Sheet Name'!A1 or 'Sheet!Name'!A1
        if t.startswith("'"):
            # Find the closing quote
            end_quote = t.index("'", 1)
            sheet = t[1:end_quote]
            a1 = t[end_quote + 2 :]  # Skip '!
        else:
            sheet, a1 = t.split("!", 1)
        if sheet not in wb_formulas.sheetnames:
            raise ValueError(f"Sheet not found: {sheet}")
        return sheet, a1

    def extract_deps_with_guards(
        formula: str, current_sheet: str, current_a1: str
    ) -> list[tuple[str, str, GuardExpr | None]]:
        if not formula.startswith("="):
            return []
        try:
            return _extract_deps_with_guards_inner(formula, current_sheet, current_a1)
        except DynamicRefError:
            raise
        except ValueError as exc:
            raise ValueError(f"{current_sheet}!{current_a1}: {exc}") from exc

    def _extract_deps_with_guards_inner(
        formula: str, current_sheet: str, current_a1: str
    ) -> list[tuple[str, str, GuardExpr | None]]:
        def extract_expr_deps(expr: str) -> list[tuple[str, str]]:
            """
            Extract dependencies from an expression fragment (no leading '=').
            """
            f = "=" + expr if not expr.startswith("=") else expr
            deps: list[tuple[str, str]] = []

            masked = f

            # 0) Dynamic refs (OFFSET/INDIRECT): cached, raise, or constraint-based.
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
                    # Filter out INDEX calls that only have literal args (no dynamic resolution needed).
                    dynamic_calls = []
                    for fn_name_check, inner_check, span_check in calls:
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
                        dynamic_calls.append((fn_name_check, inner_check, span_check))
                    calls = dynamic_calls
                    if calls:
                        cell_key = format_key(current_sheet, current_a1)
                        fn_names = sorted({fn for fn, _, _ in calls})
                        raise DynamicRefError(
                            f"Formula at {cell_key} contains {'/'.join(fn_names)} that require resolution. "
                            "Pass dynamic_refs=DynamicRefConfig.from_constraints(...) or set "
                            "use_cached_dynamic_refs=True."
                        )
                else:
                    bounds = GlobalWorkbookBounds(sheet=current_sheet)
                    argument_addrs: set[str] = set()
                    if calls:
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
                                for ref in parse_cell_refs(normalized):
                                    sh = ref.sheet if ref.sheet is not None else current_sheet
                                    a1 = f"{ref.column}{ref.row}"
                                    deps.append((sh, a1))
                                    if is_variable:
                                        argument_addrs.add(format_key(sh, a1))
                    if calls:

                        def _refs_in_formula_without_dynamic(
                            formula_str: str, sheet_of_cell: str
                        ) -> set[str]:
                            dyn = _find_function_calls_with_spans(
                                formula_str if formula_str.startswith("=") else "=" + formula_str,
                                frozenset({"OFFSET", "INDIRECT", "INDEX"}),
                            )
                            spans = [span for _fn, _inner, span in dyn]
                            masked = mask_spans(
                                formula_str if formula_str.startswith("=") else "=" + formula_str,
                                spans,
                            )
                            norm = normalizer.normalize(masked, sheet_of_cell)
                            out: set[str] = set()
                            for ref in parse_cell_refs(norm):
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
                                    max_cells=max_range_cells,
                                ):
                                    out.add(format_key(dep_sheet, dep_a1))
                            return out

                        all_refs: set[str] = set()
                        to_visit = set(argument_addrs)
                        while to_visit:
                            addr = to_visit.pop()
                            if addr in all_refs:
                                continue
                            all_refs.add(addr)
                            sh, a1 = _parse_address_to_sheet_a1(addr)
                            if sh not in wb_formulas.sheetnames:
                                continue
                            cell_val = _get_ws_f(sh)[a1].value
                            if isinstance(cell_val, str) and cell_val.startswith("="):
                                to_visit.update(_refs_in_formula_without_dynamic(cell_val, sh))
                        leaves = set()
                        for addr in all_refs:
                            sh, a1 = _parse_address_to_sheet_a1(addr)
                            if sh not in wb_formulas.sheetnames:
                                continue
                            cell_val = _get_ws_f(sh)[a1].value
                            if not (isinstance(cell_val, str) and cell_val.startswith("=")):
                                leaves.add(addr)
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

                            def _get_cell_formula(addr: str) -> str | None:
                                sh, a1 = _parse_address_to_sheet_a1(addr)
                                if sh not in wb_formulas.sheetnames:
                                    return None
                                v = _get_ws_f(sh)[a1].value
                                if not isinstance(v, str) or not v.startswith("="):
                                    return None
                                return normalizer.normalize(v, sh)

                            expanded_env = expand_leaf_env_to_argument_env(
                                all_refs,
                                _get_cell_formula,
                                _refs_in_formula_without_dynamic,
                                dynamic_refs.cell_type_env,
                                dynamic_refs.limits,
                                named_ranges=named_ranges,
                                named_range_ranges=named_range_ranges,
                                max_range_cells=max_range_cells,
                                shared_cell_type_cache=_shared_cell_type_cache,
                                type_analysis_cache=type_analysis_cache,
                                workbook_sha256=_wb_sha256,
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
                        for addr in offset_targets | indirect_targets | index_targets:
                            sh, a1 = _parse_address_to_sheet_a1(addr)
                            deps.append((sh, a1))
            masked = mask_spans(masked, dyn_spans)

            # 1) Expand ranges, then mask them so later cell-ref parsing doesn't
            # misinterpret the range endpoint as a same-sheet ref.
            if expand_ranges:
                spans: list[tuple[int, int]] = []
                for start, end, span in parse_range_refs_with_spans(masked):
                    spans.append(span)
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
                masked = mask_spans(masked, spans)

            for ref in parse_cell_refs(masked):
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
            return out

        # 1) IF(cond, then, else)
        if_parts = split_top_level_if(formula)
        if if_parts is not None:
            cond_s, then_s, else_s = if_parts
            cond_guard = parse_guard_expr(
                cond_s, current_sheet=current_sheet, named_ranges=named_ranges
            )

            unconditional = set(extract_expr_deps(cond_s))
            out: dict[tuple[str, str], GuardExpr | None] = {
                (sh, a1): None for (sh, a1) in unconditional
            }

            # If the condition can't be parsed, branch deps are still conditional, but opaque.
            then_guard: GuardExpr | None = cond_guard
            else_guard: GuardExpr | None = None if cond_guard is None else Not(cond_guard)

            for sh, a1 in extract_expr_deps(then_s):
                key = (sh, a1)
                if key in out:
                    continue
                out[key] = then_guard

            if else_s:
                for sh, a1 in extract_expr_deps(else_s):
                    key = (sh, a1)
                    if key in out:
                        continue
                    out[key] = else_guard

            return [(sh, a1, g) for (sh, a1), g in out.items()]

        # 2) IFS(cond1, value1, cond2, value2, ..., [default])
        ifs_args = split_top_level_ifs(formula)
        if ifs_args is not None:
            # All condition expressions may be evaluated (sequentially), so include deps from all
            # conditions as unconditional to avoid missing required inputs.
            conditions: list[str] = []
            values: list[str] = []
            default_expr: str | None = None
            if len(ifs_args) >= 2:
                pairs = ifs_args
                if len(pairs) % 2 == 1:
                    default_expr = pairs[-1]
                    pairs = pairs[:-1]
                for i in range(0, len(pairs), 2):
                    conditions.append(pairs[i])
                    values.append(pairs[i + 1])

            unconditional: set[tuple[str, str]] = set()
            for c in conditions:
                unconditional |= set(extract_expr_deps(c))

            out: dict[tuple[str, str], GuardExpr | None] = {
                (sh, a1): None for (sh, a1) in unconditional
            }

            prev_negations: list[GuardExpr] = []
            for _idx, (cond_s, val_s) in enumerate(zip(conditions, values, strict=False), start=1):
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

                for sh, a1 in extract_expr_deps(val_s):
                    key = (sh, a1)
                    if key in out:
                        continue
                    out[key] = g

            if default_expr is not None:
                if prev_negations:
                    default_guard: GuardExpr = (
                        prev_negations[0]
                        if len(prev_negations) == 1
                        else And(tuple(prev_negations))
                    )
                else:
                    default_guard = Literal(True)
                for sh, a1 in extract_expr_deps(default_expr):
                    key = (sh, a1)
                    if key in out:
                        continue
                    out[key] = default_guard

            return [(sh, a1, g) for (sh, a1), g in out.items()]

        # 3) CHOOSE(index, value1, value2, ...)
        choose_args = split_top_level_choose(formula)
        if choose_args is not None and len(choose_args) >= 2:
            index_s = choose_args[0]
            choices = choose_args[1:]

            index_expr = parse_guard_expr(
                index_s, current_sheet=current_sheet, named_ranges=named_ranges
            )
            unconditional = set(extract_expr_deps(index_s))
            out: dict[tuple[str, str], GuardExpr | None] = {
                (sh, a1): None for (sh, a1) in unconditional
            }

            for i, choice_s in enumerate(choices, start=1):
                guard: GuardExpr | None = None
                if index_expr is not None:
                    guard = Compare(left=index_expr, op="=", right=Literal(i))
                for sh, a1 in extract_expr_deps(choice_s):
                    key = (sh, a1)
                    if key in out:
                        continue
                    out[key] = guard

            return [(sh, a1, g) for (sh, a1), g in out.items()]

        # 4) SWITCH(expr, value1, result1, ..., [default])
        switch_args = split_top_level_switch(formula)
        if switch_args is not None and len(switch_args) >= 3:
            expr_s = switch_args[0]
            expr_ge = parse_guard_expr(
                expr_s, current_sheet=current_sheet, named_ranges=named_ranges
            )
            unconditional = set(extract_expr_deps(expr_s))
            out: dict[tuple[str, str], GuardExpr | None] = {
                (sh, a1): None for (sh, a1) in unconditional
            }

            pairs = switch_args[1:]
            default_expr: str | None = None
            if len(pairs) % 2 == 1:
                default_expr = pairs[-1]
                pairs = pairs[:-1]

            prev_negations: list[GuardExpr] = []
            for i in range(0, len(pairs), 2):
                val_s = pairs[i]
                res_s = pairs[i + 1]
                val_ge = parse_guard_expr(
                    val_s, current_sheet=current_sheet, named_ranges=named_ranges
                )
                match: GuardExpr | None = None
                if expr_ge is not None and val_ge is not None:
                    match = Compare(left=expr_ge, op="=", right=val_ge)

                guard: GuardExpr | None = None
                if match is not None:
                    ops2: list[GuardExpr] = [match, *prev_negations]
                    guard = ops2[0] if len(ops2) == 1 else And(tuple(ops2))
                    prev_negations.append(Not(match))

                for sh, a1 in extract_expr_deps(res_s):
                    key = (sh, a1)
                    if key in out:
                        continue
                    out[key] = guard

            if default_expr is not None:
                if prev_negations:
                    default_guard2: GuardExpr = (
                        prev_negations[0]
                        if len(prev_negations) == 1
                        else And(tuple(prev_negations))
                    )
                else:
                    default_guard2 = Literal(True)
                for sh, a1 in extract_expr_deps(default_expr):
                    key = (sh, a1)
                    if key in out:
                        continue
                    out[key] = default_guard2

            return [(sh, a1, g) for (sh, a1), g in out.items()]

        return [(sh, a1, None) for (sh, a1) in extract_expr_deps(formula)]

    visited: set[str] = set()
    q: deque[tuple[str, str, int]] = deque()
    for t in targets:
        sh, a1 = parse_target(str(t))
        q.append((sh, a1, 0))

    _bfs_t0 = time.perf_counter()
    _bfs_count = 0
    _bfs_next_log = 5000
    _dyn_stats = {"infer_calls": 0, "cache_hits": 0}

    try:
        while q:
            sheet, a1, depth = q.popleft()
            key = format_key(sheet, a1)
            if key in visited:
                continue
            visited.add(key)
            _bfs_count += 1
            if _bfs_count >= _bfs_next_log:
                print(
                    f"[DIAG] BFS: {_bfs_count} nodes, queue={len(q)}, depth={depth}, "
                    f"{time.perf_counter() - _bfs_t0:.1f}s, last={key}, "
                    f"dyn_infer={_dyn_stats['infer_calls']}, dyn_cache_hits={_dyn_stats['cache_hits']}, "
                    f"env_cache_size={len(_shared_cell_type_cache)}",
                    flush=True,
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
            raw = ws_f[a1].value
            if isinstance(raw, ArrayFormula):
                raw = raw.text or ""
                if raw and not raw.startswith("="):
                    raw = f"={raw}"
            is_formula = isinstance(raw, str) and raw.startswith("=")

            if is_formula:
                formula_str = str(raw)
                formula = formula_str
                normalized = normalizer.normalize(formula_str, sheet)
                value = None
                if wb_values is not None:
                    value = _get_ws_v(sheet)[a1].value
                is_leaf = False
            else:
                formula_str = ""
                formula = None
                normalized = None
                value = raw
                is_leaf = True

            col, row = fastpyxl.utils.cell.coordinate_from_string(a1)
            node = Node(
                sheet=sheet,
                column=col,
                row=int(row),
                formula=formula,
                normalized_formula=normalized,
                value=value,
                is_leaf=is_leaf,
            )
            graph.add_node(node)

            if not is_formula:
                continue

            # Run extraction first so that the constraint-based dynamic-ref expansion
            # (_dyn_cache) is populated before provenance collection reads from it.
            deps_and_guards = extract_deps_with_guards(formula_str, sheet, a1)

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
    finally:
        if _dyn_stats["infer_calls"] or _dyn_stats["cache_hits"]:
            print(
                f"[DIAG] BFS done: {_bfs_count} nodes, {time.perf_counter() - _bfs_t0:.2f}s, "
                f"dyn_infer={_dyn_stats['infer_calls']}, dyn_cache_hits={_dyn_stats['cache_hits']}, "
                f"env_cache_size={len(_shared_cell_type_cache)}",
                flush=True,
            )
        if wb_values is not None:
            wb_values.close()
        if not isinstance(workbook, fastpyxl.Workbook):
            wb_formulas.close()

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
    """Return a sorted list of leaf cell addresses that feed dynamic-ref arguments
    (OFFSET/INDIRECT/INDEX) but have no entry in ``dynamic_refs.cell_type_env``.

    Unlike :func:`create_dependency_graph`, this function does **not** raise
    :exc:`DynamicRefError` when constraints are missing.  Instead it collects all
    missing leaf addresses in a single pass and returns them sorted.

    When ``dynamic_refs`` is ``None`` the function treats it as an empty constraint
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

    try:
        named_range_maps = build_named_range_map(wb_formulas)
        named_ranges = named_range_maps.cell_map
        named_range_ranges = named_range_maps.range_map
        normalizer = FormulaNormalizer(named_ranges, named_range_ranges)
        cell_type_env = {} if dynamic_refs is None else dynamic_refs.cell_type_env
        _NAME_TOKEN_RE = re.compile(r"\b([A-Za-z_][A-Za-z0-9_]*)\b(?!\s*!)")

        def _refs_without_dynamic(formula_str: str, sheet: str) -> set[str]:
            """Static (non-dynamic-ref) cell addresses referenced by *formula_str*."""
            f = formula_str if formula_str.startswith("=") else "=" + formula_str
            dyn = _find_function_calls_with_spans(f, frozenset({"OFFSET", "INDIRECT", "INDEX"}))
            spans = [span for _fn, _inner, span in dyn]
            masked = mask_spans(f, spans)
            norm = normalizer.normalize(masked, sheet)
            out: set[str] = set()
            for ref in parse_cell_refs(norm):
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
            return out

        collected: set[str] = set()
        visited: set[str] = set()
        queue: deque[tuple[str, str, int]] = deque()

        for t in targets:
            t_str = str(t)
            if "!" not in t_str:
                raise ValueError(f"Target must be sheet-qualified: {t_str}")
            sh, a1 = _parse_address_to_sheet_a1(t_str)
            if sh not in wb_formulas.sheetnames:
                raise ValueError(f"Sheet not found: {sh}")
            queue.append((sh, a1, 0))

        while queue:
            current_sheet, current_a1, depth = queue.popleft()
            key = format_key(current_sheet, current_a1)
            if key in visited:
                continue
            visited.add(key)

            if depth >= max_depth or current_sheet not in wb_formulas.sheetnames:
                continue

            cell_val = wb_formulas[current_sheet][current_a1].value
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
                            for ref in parse_cell_refs(normalized_arg):
                                sh = ref.sheet if ref.sheet is not None else current_sheet
                                argument_addrs.add(format_key(sh, f"{ref.column}{ref.row}"))

                # Walk argument_addrs to statically-reachable leaves.
                all_refs: set[str] = set()
                to_visit_inner = set(argument_addrs)
                while to_visit_inner:
                    addr = to_visit_inner.pop()
                    if addr in all_refs:
                        continue
                    all_refs.add(addr)
                    sh, a1 = _parse_address_to_sheet_a1(addr)
                    if sh not in wb_formulas.sheetnames:
                        continue
                    inner_val = wb_formulas[sh][a1].value
                    if isinstance(inner_val, str) and inner_val.startswith("="):
                        to_visit_inner.update(_refs_without_dynamic(inner_val, sh))

                leaves: set[str] = set()
                for addr in all_refs:
                    sh, a1 = _parse_address_to_sheet_a1(addr)
                    if sh not in wb_formulas.sheetnames:
                        continue
                    inner_val = wb_formulas[sh][a1].value
                    if not (isinstance(inner_val, str) and inner_val.startswith("=")):
                        leaves.add(addr)

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

                        def _get_cell_formula(addr: str) -> str | None:
                            sh2, a1_2 = _parse_address_to_sheet_a1(addr)
                            if sh2 not in wb_formulas.sheetnames:
                                return None
                            v = wb_formulas[sh2][a1_2].value
                            if not isinstance(v, str) or not v.startswith("="):
                                return None
                            return normalizer.normalize(v, sh2)

                        expanded_env = expand_leaf_env_to_argument_env(
                            all_refs,
                            _get_cell_formula,
                            _refs_without_dynamic,
                            dynamic_refs.cell_type_env,
                            dynamic_refs.limits,
                            named_ranges=named_ranges,
                            named_range_ranges=named_range_ranges,
                            max_range_cells=max_range_cells,
                            type_analysis_cache=type_analysis_cache,
                            workbook_sha256=_wb_sha256_cand,
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
                        for addr in offset_targets | indirect_targets | index_targets:
                            sh, a1 = _parse_address_to_sheet_a1(addr)
                            queue.append((sh, a1, depth + 1))
                    except DynamicRefError:
                        pass  # best-effort: skip dynamic targets for this formula

            # Queue static (non-dynamic-ref) deps for continued BFS.
            for addr in _refs_without_dynamic(f, current_sheet):
                sh, a1 = _parse_address_to_sheet_a1(addr)
                if sh in wb_formulas.sheetnames:
                    queue.append((sh, a1, depth + 1))

    finally:
        if _owns_wb:
            wb_formulas.close()

    return sorted(collected)
