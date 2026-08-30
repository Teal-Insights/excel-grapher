from __future__ import annotations

import re
from collections.abc import Callable
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from excel_grapher.grapher.type_analysis_cache import TypeAnalysisCache

import fastpyxl
import fastpyxl.utils.cell

from excel_grapher.core.cell_types import leaves_missing_cell_type_constraints

from .dependency_provenance import DependencyCause, EdgeProvenance, merge_provenance_maps
from .dynamic_ref_walk import DynamicRefWalkContext
from .dynamic_refs import (
    DynamicRefConfig,
    DynamicRefError,
    GlobalWorkbookBounds,
    expand_leaf_env_to_argument_env,
    infer_dynamic_index_targets,
    infer_dynamic_indirect_targets,
    infer_dynamic_offset_targets,
)
from .parser import (
    FormulaNormalizer,
    _find_function_calls_with_spans,
    _split_function_args,
    expand_range,
    format_key,
    mask_ref_only_function_calls,
    mask_spans,
    parse_dynamic_range_refs_with_spans,
    parse_range_refs_with_spans,
    parse_standalone_cell_refs,
    parse_standalone_cell_refs_with_spans,
    split_top_level_choose,
    split_top_level_if,
    split_top_level_ifs,
    split_top_level_switch,
)

_NAME_TOKEN_RE = re.compile(r"\b([A-Za-z_][A-Za-z0-9_]*)\b(?!\s*!)")
_DYNAMIC_REF_FNS = frozenset({"OFFSET", "INDIRECT", "INDEX"})


def _dynamic_cause(fn_name: str) -> DependencyCause:
    if fn_name == "OFFSET":
        return DependencyCause.dynamic_offset
    if fn_name == "INDIRECT":
        return DependencyCause.dynamic_indirect
    if fn_name == "INDEX":
        return DependencyCause.dynamic_index
    return DependencyCause.dynamic_offset


def _index_has_non_literal_row_col(inner: str) -> bool:
    idx_args = _split_function_args(inner)
    if idx_args is None or len(idx_args) < 2:
        return False
    for j, idx_arg in enumerate(idx_args):
        if j == 0:
            continue
        try:
            float(idx_arg.strip())
        except ValueError:
            return True
    return False


def _merge_into(
    acc: dict[str, EdgeProvenance],
    dep_key: str,
    prov: EdgeProvenance,
) -> None:
    prev = acc.get(dep_key)
    if prev is None:
        acc[dep_key] = prov
    else:
        acc[dep_key] = prev.merge(prov)


def _prov_with_direct_span(
    span: tuple[int, int],
    *,
    collect_spans: bool,
) -> EdgeProvenance:
    """Build direct-ref provenance, recording `span` only for the dialect pass.

    Spans are stored against the string that becomes `Node.normalized_formula`
    (AST `render_formula` when parseable, regex fallback otherwise). A pass
    over a different spelling of the same formula may contribute cause flags
    but no positions.
    """
    if not collect_spans:
        return EdgeProvenance(causes=DependencyCause.direct_ref)
    return EdgeProvenance(
        causes=DependencyCause.direct_ref,
        direct_sites_normalized=(span,),
    )


def _flat_provenance_one_string(
    f: str,
    *,
    current_sheet: str,
    current_a1: str,
    named_ranges: dict[str, tuple[str, str]],
    named_range_ranges: dict[str, tuple[str, str, str]],
    normalizer: FormulaNormalizer | None = None,
    defined_names: set[str],
    expand_ranges: bool,
    max_range_cells: int,
    use_cached_dynamic_refs: bool,
    dynamic_refs: DynamicRefConfig | None,
    wb_formulas: fastpyxl.Workbook,
    resolve_cached_value: Callable[[str, str], object | None],
    collect_spans: bool,
    dynamic_expansion_cache: dict[tuple[str, str, str], tuple[set[str], set[str], set[str]]]
    | None = None,
    type_analysis_cache: TypeAnalysisCache | None = None,
    workbook_sha256: str | None = None,
    ref_walk: DynamicRefWalkContext | None = None,
) -> dict[str, EdgeProvenance]:
    """Mirror extract_expr_deps masking pipeline; accumulate provenance for one formula string starting with '='."""
    if normalizer is None:
        normalizer = FormulaNormalizer(named_ranges, named_range_ranges)
    acc: dict[str, EdgeProvenance] = {}

    if not f.startswith("="):
        f = "=" + f

    masked = f
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
            cause_dyn = _dynamic_cause(_call_kind_at_span(f, span))
            sheet = start.sheet if start.sheet is not None else current_sheet
            for dep_sheet, dep_a1 in expand_range(
                sheet=sheet,
                start_col=start.column,
                start_row=start.row,
                end_col=end.column,
                end_row=end.row,
                max_cells=max_range_cells,
            ):
                k = format_key(dep_sheet, dep_a1)
                _merge_into(acc, k, EdgeProvenance(causes=cause_dyn))
            for ref in arg_refs:
                arg_sheet = ref.sheet if ref.sheet is not None else current_sheet
                k = format_key(arg_sheet, f"{ref.column}{ref.row}")
                _merge_into(acc, k, EdgeProvenance(causes=cause_dyn))
    else:
        calls = _find_function_calls_with_spans(f, _DYNAMIC_REF_FNS)
        if dynamic_refs is None:
            calls = [c for c in calls if c[0] != "INDEX" or _index_has_non_literal_row_col(c[1])]
            if calls:
                raise DynamicRefError(
                    "Provenance collection requires dynamic ref resolution for "
                    "OFFSET/INDIRECT/INDEX in this formula."
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
                    dyn_cause = _dynamic_cause(fn_name)
                    for i, arg in enumerate(args):
                        norm_arg = normalizer.normalize(
                            "=" + arg,
                            current_sheet,
                        )
                        is_variable = (
                            (fn_name == "OFFSET" and i >= 1)
                            or (fn_name == "OFFSET" and i == 0 and "(" in norm_arg)
                            or fn_name == "INDIRECT"
                            or (fn_name == "INDEX" and i >= 1)
                        )
                        if fn_name == "INDEX" and i == 0 and "(" not in norm_arg:
                            continue
                        value_expr = mask_ref_only_function_calls(norm_arg)
                        for ref in parse_standalone_cell_refs(value_expr):
                            sh = ref.sheet if ref.sheet is not None else current_sheet
                            a1 = f"{ref.column}{ref.row}"
                            k = format_key(sh, a1)
                            _merge_into(acc, k, EdgeProvenance(causes=dyn_cause))
                            if is_variable:
                                argument_addrs.add(k)
                        if is_variable:
                            for start, end, _span in parse_range_refs_with_spans(value_expr):
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
                                    k = format_key(dep_sheet, dep_a1)
                                    _merge_into(acc, k, EdgeProvenance(causes=dyn_cause))
                                    argument_addrs.add(k)
            if calls:
                formula_for_infer = normalizer.normalize(f, current_sheet)
                _col_letter, _current_row = fastpyxl.utils.cell.coordinate_from_string(current_a1)
                _current_col = fastpyxl.utils.cell.column_index_from_string(_col_letter)
                _cache_key = (formula_for_infer, current_sheet, current_a1)
                if dynamic_expansion_cache is not None and _cache_key in dynamic_expansion_cache:
                    offset_targets, indirect_targets, index_targets = dynamic_expansion_cache[
                        _cache_key
                    ]
                else:
                    walk = ref_walk
                    if walk is None:
                        walk = DynamicRefWalkContext(
                            normalizer=normalizer,
                            max_range_cells=max_range_cells,
                            get_cell_value=lambda sh, a1: wb_formulas[sh][a1].value,
                            sheet_names=wb_formulas.sheetnames,
                            named_ranges=named_ranges,
                            named_range_ranges=named_range_ranges,
                        )
                    all_refs, leaves = walk.argument_subgraph_refs(argument_addrs)
                    missing_leaves = leaves_missing_cell_type_constraints(
                        leaves, dynamic_refs.cell_type_env
                    )
                    if missing_leaves:
                        raise DynamicRefError(
                            "Provenance: leaf cells feeding OFFSET/INDIRECT/INDEX have no "
                            f"constraint: {sorted(missing_leaves)}"
                        )
                    expanded_env = expand_leaf_env_to_argument_env(
                        all_refs,
                        walk.cell_formula,
                        walk.refs_in_formula_without_dynamic,
                        dynamic_refs.cell_type_env,
                        dynamic_refs.limits,
                        named_ranges=named_ranges,
                        named_range_ranges=named_range_ranges,
                        max_range_cells=max_range_cells,
                        shared_cell_type_cache=walk.shared_cell_type_cache,
                        type_analysis_cache=type_analysis_cache,
                        workbook_sha256=workbook_sha256,
                        get_cell_ast=walk.cell_ast,
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
                    if dynamic_expansion_cache is not None:
                        dynamic_expansion_cache[_cache_key] = (
                            offset_targets,
                            indirect_targets,
                            index_targets,
                        )
                for addr in offset_targets:
                    _merge_into(
                        acc,
                        addr,
                        EdgeProvenance(causes=DependencyCause.dynamic_offset),
                    )
                for addr in indirect_targets:
                    _merge_into(
                        acc,
                        addr,
                        EdgeProvenance(causes=DependencyCause.dynamic_indirect),
                    )
                for addr in index_targets:
                    _merge_into(
                        acc,
                        addr,
                        EdgeProvenance(causes=DependencyCause.dynamic_index),
                    )

    masked = mask_spans(masked, dyn_spans)
    masked = mask_ref_only_function_calls(masked)

    if expand_ranges:
        for start, end, _span in parse_range_refs_with_spans(masked):
            sheet = start.sheet if start.sheet is not None else current_sheet
            for dep_sheet, dep_a1 in expand_range(
                sheet=sheet,
                start_col=start.column,
                start_row=start.row,
                end_col=end.column,
                end_row=end.row,
                max_cells=max_range_cells,
            ):
                k = format_key(dep_sheet, dep_a1)
                _merge_into(acc, k, EdgeProvenance(causes=DependencyCause.static_range))

    for ref, span in parse_standalone_cell_refs_with_spans(masked):
        sh = ref.sheet if ref.sheet is not None else current_sheet
        k = format_key(sh, f"{ref.column}{ref.row}")
        _merge_into(acc, k, _prov_with_direct_span(span, collect_spans=collect_spans))

    for m in _NAME_TOKEN_RE.finditer(masked):
        token = m.group(1)
        resolved = named_ranges.get(token)
        if resolved is not None:
            sh, a1 = resolved
            k = format_key(sh, a1)
            span = m.span()
            _merge_into(acc, k, _prov_with_direct_span(span, collect_spans=collect_spans))
            continue
        resolved_range = named_range_ranges.get(token)
        if resolved_range is not None:
            if expand_ranges:
                sheet, start_a1, end_a1 = resolved_range
                start_col, start_row = fastpyxl.utils.cell.coordinate_from_string(start_a1)
                end_col, end_row = fastpyxl.utils.cell.coordinate_from_string(end_a1)
                for dep_sheet, dep_a1 in expand_range(
                    sheet=sheet,
                    start_col=start_col,
                    start_row=int(start_row),
                    end_col=end_col,
                    end_row=int(end_row),
                    max_cells=max_range_cells,
                ):
                    k = format_key(dep_sheet, dep_a1)
                    _merge_into(
                        acc,
                        k,
                        EdgeProvenance(causes=DependencyCause.static_range),
                    )
            continue
        if token in defined_names:
            raise ValueError(f"Unsupported defined name referenced in formula: {token}")

    return acc


def _call_kind_at_span(formula: str, span: tuple[int, int]) -> str:
    """Return OFFSET, INDIRECT, or INDEX for the dynamic call covering span."""
    calls = _find_function_calls_with_spans(formula, _DYNAMIC_REF_FNS)
    for fn, _inner, sp in calls:
        if sp == span:
            return fn
    return "OFFSET"


def _flat_provenance_formula_and_normalized(
    formula_str: str,
    normalized: str | None,
    *,
    current_sheet: str,
    current_a1: str,
    named_ranges: dict[str, tuple[str, str]],
    named_range_ranges: dict[str, tuple[str, str, str]],
    normalizer: FormulaNormalizer | None = None,
    defined_names: set[str],
    expand_ranges: bool,
    max_range_cells: int,
    use_cached_dynamic_refs: bool,
    dynamic_refs: DynamicRefConfig | None,
    wb_formulas: fastpyxl.Workbook,
    resolve_cached_value: Callable[[str, str], object | None],
    type_analysis_cache: TypeAnalysisCache | None = None,
    workbook_sha256: str | None = None,
    dynamic_expansion_cache: dict[tuple[str, str, str], tuple[set[str], set[str], set[str]]]
    | None = None,
    ref_walk: DynamicRefWalkContext | None = None,
) -> dict[str, EdgeProvenance]:
    # When the extraction string already is the `normalized_formula` dialect
    # (AST render, or regex fallback when the AST parser failed), the single
    # pass collects spans. `normalized` is None on branch recursion in
    # `collect_provenance_for_formula`, where `formula_str` is a sub-expression:
    # spans there would be offset against the branch rather than the node's
    # formula, so that path stays span-free.
    normalized_matches_raw = bool(normalized) and normalized == formula_str
    raw_map = _flat_provenance_one_string(
        formula_str,
        current_sheet=current_sheet,
        current_a1=current_a1,
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
        collect_spans=normalized_matches_raw,
        dynamic_expansion_cache=dynamic_expansion_cache,
        type_analysis_cache=type_analysis_cache,
        workbook_sha256=workbook_sha256,
        ref_walk=ref_walk,
    )
    if not normalized or normalized_matches_raw:
        return raw_map

    norm_map = _flat_provenance_one_string(
        normalized,
        current_sheet=current_sheet,
        current_a1=current_a1,
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
        collect_spans=True,
        dynamic_expansion_cache=dynamic_expansion_cache,
        type_analysis_cache=type_analysis_cache,
        workbook_sha256=workbook_sha256,
        ref_walk=ref_walk,
    )
    out: dict[str, EdgeProvenance] = {}
    all_keys = set(raw_map) | set(norm_map)
    for k in all_keys:
        r = raw_map.get(k)
        n = norm_map.get(k)
        if r is None:
            if n is not None:
                out[k] = n
            continue
        if n is None:
            out[k] = r
            continue
        # The non-dialect pass contributes cause flags only; spans always come
        # from the `normalized_formula` string.
        out[k] = EdgeProvenance(
            causes=r.causes | n.causes,
            direct_sites_normalized=n.direct_sites_normalized,
        )
    return out


def _ensure_leading_equals(s: str) -> str:
    t = s.strip()
    return t if t.startswith("=") else "=" + t


def collect_provenance_for_formula(
    formula: str,
    *,
    normalized_formula: str | None,
    current_sheet: str,
    current_a1: str,
    named_ranges: dict[str, tuple[str, str]],
    named_range_ranges: dict[str, tuple[str, str, str]],
    normalizer: FormulaNormalizer | None = None,
    defined_names: set[str],
    expand_ranges: bool,
    max_range_cells: int,
    use_cached_dynamic_refs: bool,
    dynamic_refs: DynamicRefConfig | None,
    wb_formulas: fastpyxl.Workbook,
    resolve_cached_value: Callable[[str, str], object | None],
    dynamic_expansion_cache: dict[tuple[str, str, str], tuple[set[str], set[str], set[str]]]
    | None = None,
    type_analysis_cache: TypeAnalysisCache | None = None,
    workbook_sha256: str | None = None,
    ref_walk: DynamicRefWalkContext | None = None,
) -> dict[str, EdgeProvenance]:
    """Build a dependency-key to `EdgeProvenance` map for one formula.

    Includes IF/IFS/CHOOSE/SWITCH branch union semantics.
    """
    if normalizer is None:
        normalizer = FormulaNormalizer(named_ranges, named_range_ranges)
    if ref_walk is None:

        def _cell_value(sheet: str, a1: str) -> object:
            return wb_formulas[sheet][a1].value

        ref_walk = DynamicRefWalkContext(
            normalizer=normalizer,
            max_range_cells=max_range_cells,
            get_cell_value=_cell_value,
            sheet_names=wb_formulas.sheetnames,
            named_ranges=named_ranges,
            named_range_ranges=named_range_ranges,
        )
    f = _ensure_leading_equals(formula)

    def _collect_branch(branch_formula: str) -> dict[str, EdgeProvenance]:
        return collect_provenance_for_formula(
            _ensure_leading_equals(branch_formula),
            normalized_formula=None,
            current_sheet=current_sheet,
            current_a1=current_a1,
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
            dynamic_expansion_cache=dynamic_expansion_cache,
            type_analysis_cache=type_analysis_cache,
            workbook_sha256=workbook_sha256,
            ref_walk=ref_walk,
        )

    if_parts = split_top_level_if(f)
    if if_parts is not None:
        cond_s, then_s, else_s = if_parts
        maps = [_collect_branch(cond_s), _collect_branch(then_s)]
        if else_s:
            maps.append(_collect_branch(else_s))
        return merge_provenance_maps(maps)

    ifs_args = split_top_level_ifs(f)
    if ifs_args is not None and len(ifs_args) >= 2:
        pairs: list[str] = list(ifs_args)
        default_ifs: str | None = None
        if len(pairs) % 2 == 1:
            default_ifs = pairs[-1]
            pairs = pairs[:-1]
        maps = []
        for i in range(0, len(pairs), 2):
            cond_s, val_s = pairs[i], pairs[i + 1]
            maps.append(_collect_branch(cond_s))
            maps.append(_collect_branch(val_s))
        if default_ifs is not None:
            maps.append(_collect_branch(default_ifs))
        return merge_provenance_maps(maps)

    choose_args = split_top_level_choose(f)
    if choose_args is not None and len(choose_args) >= 2:
        maps = [_collect_branch(choose_args[0])]
        for choice_s in choose_args[1:]:
            maps.append(_collect_branch(choice_s))
        return merge_provenance_maps(maps)

    switch_args = split_top_level_switch(f)
    if switch_args is not None and len(switch_args) >= 3:
        maps = [_collect_branch(switch_args[0])]
        pairs = switch_args[1:]
        default_expr: str | None = None
        if len(pairs) % 2 == 1:
            default_expr = pairs[-1]
            pairs = pairs[:-1]
        for i in range(0, len(pairs), 2):
            val_s, res_s = pairs[i], pairs[i + 1]
            maps.append(_collect_branch(val_s))
            maps.append(_collect_branch(res_s))
        if default_expr is not None:
            maps.append(_collect_branch(default_expr))
        return merge_provenance_maps(maps)

    return _flat_provenance_formula_and_normalized(
        formula_str=f,
        normalized=normalized_formula,
        current_sheet=current_sheet,
        current_a1=current_a1,
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
        dynamic_expansion_cache=dynamic_expansion_cache,
        type_analysis_cache=type_analysis_cache,
        workbook_sha256=workbook_sha256,
        ref_walk=ref_walk,
    )
