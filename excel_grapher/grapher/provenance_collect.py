from __future__ import annotations

import re
from collections.abc import Callable
from typing import Literal

import fastpyxl
import fastpyxl.utils.cell

from excel_grapher.core.cell_types import leaves_missing_cell_type_constraints

from .dependency_provenance import DependencyCause, EdgeProvenance, merge_provenance_maps
from .dynamic_refs import (
    DynamicRefConfig,
    DynamicRefError,
    GlobalWorkbookBounds,
    expand_leaf_env_to_argument_env,
    infer_dynamic_indirect_targets,
    infer_dynamic_offset_targets,
)
from .parser import (
    FormulaNormalizer,
    _find_function_calls_with_spans,
    _split_function_args,
    expand_range,
    format_key,
    mask_spans,
    parse_cell_refs,
    parse_cell_refs_with_spans,
    parse_dynamic_range_refs_with_spans,
    parse_range_refs_with_spans,
    split_top_level_choose,
    split_top_level_if,
    split_top_level_ifs,
    split_top_level_switch,
)


def _parse_address_to_sheet_a1(addr: str) -> tuple[str, str]:
    if "!" not in addr:
        raise ValueError(f"Address must be sheet-qualified: {addr}")
    if addr.startswith("'"):
        end_quote = addr.index("'", 1)
        sheet = addr[1:end_quote]
        a1 = addr[end_quote + 2 :]
        return sheet, a1
    sheet, a1 = addr.split("!", 1)
    return sheet, a1


_NAME_TOKEN_RE = re.compile(r"\b([A-Za-z_][A-Za-z0-9_]*)\b(?!\s*!)")


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
    span_target: Literal["formula", "normalized"],
) -> EdgeProvenance:
    if span_target == "formula":
        return EdgeProvenance(
            causes=frozenset({DependencyCause.direct_ref}),
            direct_sites_formula=(span,),
        )
    return EdgeProvenance(
        causes=frozenset({DependencyCause.direct_ref}),
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
    span_target: Literal["formula", "normalized"],
    dynamic_expansion_cache: dict[tuple[str, str, str], tuple[set[str], set[str], set[str]]] | None = None,
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
            fn = _call_kind_at_span(f, span)
            cause_dyn = (
                DependencyCause.dynamic_offset
                if fn == "OFFSET"
                else DependencyCause.dynamic_indirect
            )
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
                _merge_into(acc, k, EdgeProvenance(causes=frozenset({cause_dyn})))
            for ref in arg_refs:
                arg_sheet = ref.sheet if ref.sheet is not None else current_sheet
                k = format_key(arg_sheet, f"{ref.column}{ref.row}")
                _merge_into(acc, k, EdgeProvenance(causes=frozenset({cause_dyn})))
    else:
        calls = _find_function_calls_with_spans(f, {"OFFSET", "INDIRECT"})
        if dynamic_refs is None:
            if calls:
                raise DynamicRefError(
                    "Provenance collection requires dynamic ref resolution for OFFSET/INDIRECT in this formula."
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
                        norm_arg = normalizer.normalize(
                            "=" + arg,
                            current_sheet,
                        )
                        is_variable = (
                            (fn_name == "OFFSET" and i >= 1)
                            or (fn_name == "OFFSET" and i == 0 and "(" in norm_arg)
                            or fn_name == "INDIRECT"
                        )
                        for ref in parse_cell_refs(norm_arg):
                            sh = ref.sheet if ref.sheet is not None else current_sheet
                            a1 = f"{ref.column}{ref.row}"
                            k = format_key(sh, a1)
                            dyn_cause = (
                                DependencyCause.dynamic_offset
                                if fn_name == "OFFSET"
                                else DependencyCause.dynamic_indirect
                            )
                            _merge_into(acc, k, EdgeProvenance(causes=frozenset({dyn_cause})))
                            if is_variable:
                                argument_addrs.add(k)
            if calls:
                def _refs_in_formula_without_dynamic(formula_str: str, sheet_of_cell: str) -> set[str]:
                    dyn = _find_function_calls_with_spans(
                        formula_str if formula_str.startswith("=") else "=" + formula_str,
                        {"OFFSET", "INDIRECT"},
                    )
                    spans = [span for _fn, _inner, span in dyn]
                    masked2 = mask_spans(
                        formula_str if formula_str.startswith("=") else "=" + formula_str,
                        spans,
                    )
                    norm = normalizer.normalize(masked2, sheet_of_cell)
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
                    cell_val = wb_formulas[sh][a1].value
                    if isinstance(cell_val, str) and cell_val.startswith("="):
                        to_visit.update(_refs_in_formula_without_dynamic(cell_val, sh))
                leaves: set[str] = set()
                for addr in all_refs:
                    sh, a1 = _parse_address_to_sheet_a1(addr)
                    if sh not in wb_formulas.sheetnames:
                        continue
                    cell_val = wb_formulas[sh][a1].value
                    if not (isinstance(cell_val, str) and cell_val.startswith("=")):
                        leaves.add(addr)
                missing_leaves = leaves_missing_cell_type_constraints(
                    leaves, dynamic_refs.cell_type_env
                )
                if missing_leaves:
                    raise DynamicRefError(
                        f"Provenance: leaf cells feeding OFFSET/INDIRECT have no constraint: {sorted(missing_leaves)}"
                    )

                formula_for_infer = normalizer.normalize(f, current_sheet)
                _col_letter, _current_row = fastpyxl.utils.cell.coordinate_from_string(current_a1)
                _current_col = fastpyxl.utils.cell.column_index_from_string(_col_letter)
                _cache_key = (formula_for_infer, current_sheet, current_a1)
                if dynamic_expansion_cache is not None and _cache_key in dynamic_expansion_cache:
                    offset_targets, indirect_targets, _ = dynamic_expansion_cache[_cache_key]
                else:
                    def _get_cell_formula(addr: str) -> str | None:
                        sh, a1 = _parse_address_to_sheet_a1(addr)
                        if sh not in wb_formulas.sheetnames:
                            return None
                        v = wb_formulas[sh][a1].value
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
                for addr in offset_targets:
                    _merge_into(
                        acc,
                        addr,
                        EdgeProvenance(causes=frozenset({DependencyCause.dynamic_offset})),
                    )
                for addr in indirect_targets:
                    _merge_into(
                        acc,
                        addr,
                        EdgeProvenance(causes=frozenset({DependencyCause.dynamic_indirect})),
                    )

    masked = mask_spans(masked, dyn_spans)

    if expand_ranges:
        spans: list[tuple[int, int]] = []
        for start, end, span in parse_range_refs_with_spans(masked):
            spans.append(span)
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
                _merge_into(acc, k, EdgeProvenance(causes=frozenset({DependencyCause.static_range})))
        masked = mask_spans(masked, spans)

    for ref, span in parse_cell_refs_with_spans(masked):
        sh = ref.sheet if ref.sheet is not None else current_sheet
        k = format_key(sh, f"{ref.column}{ref.row}")
        _merge_into(acc, k, _prov_with_direct_span(span, span_target=span_target))

    for m in _NAME_TOKEN_RE.finditer(masked):
        token = m.group(1)
        resolved = named_ranges.get(token)
        if resolved is not None:
            sh, a1 = resolved
            k = format_key(sh, a1)
            span = m.span()
            _merge_into(acc, k, _prov_with_direct_span(span, span_target=span_target))
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
                        EdgeProvenance(causes=frozenset({DependencyCause.static_range})),
                    )
            continue
        if token in defined_names:
            raise ValueError(f"Unsupported defined name referenced in formula: {token}")

    return acc


def _call_kind_at_span(formula: str, span: tuple[int, int]) -> str:
    """Return 'OFFSET' or 'INDIRECT' for the dynamic call covering span (cached path)."""
    calls = _find_function_calls_with_spans(formula, {"OFFSET", "INDIRECT"})
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
    dynamic_expansion_cache: dict[tuple[str, str, str], tuple[set[str], set[str], set[str]]] | None = None,
) -> dict[str, EdgeProvenance]:
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
        span_target="formula",
        dynamic_expansion_cache=dynamic_expansion_cache,
    )
    if not normalized or normalized == formula_str:
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
        span_target="normalized",
        dynamic_expansion_cache=dynamic_expansion_cache,
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
        causes = r.causes | n.causes
        out[k] = EdgeProvenance(
            causes=causes,
            direct_sites_formula=r.direct_sites_formula,
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
    dynamic_expansion_cache: dict[tuple[str, str, str], tuple[set[str], set[str], set[str]]] | None = None,
) -> dict[str, EdgeProvenance]:
    """
    Build a map from dependency cell key (``format_key``) to merged :class:`EdgeProvenance`
    for one cell's formula, including IF/IFS/CHOOSE/SWITCH branch union semantics.
    """
    if normalizer is None:
        normalizer = FormulaNormalizer(named_ranges, named_range_ranges)
    f = _ensure_leading_equals(formula)

    if_parts = split_top_level_if(f)
    if if_parts is not None:
        cond_s, then_s, else_s = if_parts
        maps = [
            collect_provenance_for_formula(
                _ensure_leading_equals(cond_s),
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
            ),
            collect_provenance_for_formula(
                _ensure_leading_equals(then_s),
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
            ),
        ]
        if else_s:
            maps.append(
                collect_provenance_for_formula(
                    _ensure_leading_equals(else_s),
                    normalized_formula=None,
                    current_sheet=current_sheet,
                    current_a1=current_a1,
                    named_ranges=named_ranges,
                    named_range_ranges=named_range_ranges,
                    defined_names=defined_names,
                    expand_ranges=expand_ranges,
                    max_range_cells=max_range_cells,
                    use_cached_dynamic_refs=use_cached_dynamic_refs,
                    dynamic_refs=dynamic_refs,
                    wb_formulas=wb_formulas,
                    resolve_cached_value=resolve_cached_value,
                    dynamic_expansion_cache=dynamic_expansion_cache,
                )
            )
        return merge_provenance_maps(maps)

    ifs_args = split_top_level_ifs(f)
    if ifs_args is not None and len(ifs_args) >= 2:
        pairs: list[str] = list(ifs_args)
        default_ifs: str | None = None
        if len(pairs) % 2 == 1:
            default_ifs = pairs[-1]
            pairs = pairs[:-1]
        maps: list[dict[str, EdgeProvenance]] = []
        for i in range(0, len(pairs), 2):
            cond_s, val_s = pairs[i], pairs[i + 1]
            maps.append(
                collect_provenance_for_formula(
                    _ensure_leading_equals(cond_s),
                    normalized_formula=None,
                    current_sheet=current_sheet,
                    current_a1=current_a1,
                    named_ranges=named_ranges,
                    named_range_ranges=named_range_ranges,
                    defined_names=defined_names,
                    expand_ranges=expand_ranges,
                    max_range_cells=max_range_cells,
                    use_cached_dynamic_refs=use_cached_dynamic_refs,
                    dynamic_refs=dynamic_refs,
                    wb_formulas=wb_formulas,
                    resolve_cached_value=resolve_cached_value,
                    dynamic_expansion_cache=dynamic_expansion_cache,
                )
            )
            maps.append(
                collect_provenance_for_formula(
                    _ensure_leading_equals(val_s),
                    normalized_formula=None,
                    current_sheet=current_sheet,
                    current_a1=current_a1,
                    named_ranges=named_ranges,
                    named_range_ranges=named_range_ranges,
                    defined_names=defined_names,
                    expand_ranges=expand_ranges,
                    max_range_cells=max_range_cells,
                    use_cached_dynamic_refs=use_cached_dynamic_refs,
                    dynamic_refs=dynamic_refs,
                    wb_formulas=wb_formulas,
                    resolve_cached_value=resolve_cached_value,
                    dynamic_expansion_cache=dynamic_expansion_cache,
                )
            )
        if default_ifs is not None:
            maps.append(
                collect_provenance_for_formula(
                    _ensure_leading_equals(default_ifs),
                    normalized_formula=None,
                    current_sheet=current_sheet,
                    current_a1=current_a1,
                    named_ranges=named_ranges,
                    named_range_ranges=named_range_ranges,
                    defined_names=defined_names,
                    expand_ranges=expand_ranges,
                    max_range_cells=max_range_cells,
                    use_cached_dynamic_refs=use_cached_dynamic_refs,
                    dynamic_refs=dynamic_refs,
                    wb_formulas=wb_formulas,
                    resolve_cached_value=resolve_cached_value,
                    dynamic_expansion_cache=dynamic_expansion_cache,
                )
            )
        return merge_provenance_maps(maps)

    choose_args = split_top_level_choose(f)
    if choose_args is not None and len(choose_args) >= 2:
        maps = [
            collect_provenance_for_formula(
                _ensure_leading_equals(choose_args[0]),
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
            )
        ]
        for choice_s in choose_args[1:]:
            maps.append(
                collect_provenance_for_formula(
                    _ensure_leading_equals(choice_s),
                    normalized_formula=None,
                    current_sheet=current_sheet,
                    current_a1=current_a1,
                    named_ranges=named_ranges,
                    named_range_ranges=named_range_ranges,
                    defined_names=defined_names,
                    expand_ranges=expand_ranges,
                    max_range_cells=max_range_cells,
                    use_cached_dynamic_refs=use_cached_dynamic_refs,
                    dynamic_refs=dynamic_refs,
                    wb_formulas=wb_formulas,
                    resolve_cached_value=resolve_cached_value,
                    dynamic_expansion_cache=dynamic_expansion_cache,
                )
            )
        return merge_provenance_maps(maps)

    switch_args = split_top_level_switch(f)
    if switch_args is not None and len(switch_args) >= 3:
        expr_s = switch_args[0]
        maps = [
            collect_provenance_for_formula(
                _ensure_leading_equals(expr_s),
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
            )
        ]
        pairs = switch_args[1:]
        default_expr: str | None = None
        if len(pairs) % 2 == 1:
            default_expr = pairs[-1]
            pairs = pairs[:-1]
        for i in range(0, len(pairs), 2):
            val_s, res_s = pairs[i], pairs[i + 1]
            for sub in (val_s, res_s):
                maps.append(
                    collect_provenance_for_formula(
                        _ensure_leading_equals(sub),
                        normalized_formula=None,
                        current_sheet=current_sheet,
                        current_a1=current_a1,
                        named_ranges=named_ranges,
                        named_range_ranges=named_range_ranges,
                        defined_names=defined_names,
                        expand_ranges=expand_ranges,
                        max_range_cells=max_range_cells,
                        use_cached_dynamic_refs=use_cached_dynamic_refs,
                        dynamic_refs=dynamic_refs,
                        wb_formulas=wb_formulas,
                        resolve_cached_value=resolve_cached_value,
                        dynamic_expansion_cache=dynamic_expansion_cache,
                    )
                )
        if default_expr is not None:
            maps.append(
                collect_provenance_for_formula(
                    _ensure_leading_equals(default_expr),
                    normalized_formula=None,
                    current_sheet=current_sheet,
                    current_a1=current_a1,
                    named_ranges=named_ranges,
                    named_range_ranges=named_range_ranges,
                    defined_names=defined_names,
                    expand_ranges=expand_ranges,
                    max_range_cells=max_range_cells,
                    use_cached_dynamic_refs=use_cached_dynamic_refs,
                    dynamic_refs=dynamic_refs,
                    wb_formulas=wb_formulas,
                    resolve_cached_value=resolve_cached_value,
                    dynamic_expansion_cache=dynamic_expansion_cache,
                )
            )
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
    )
