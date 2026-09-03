"""Emit an inverted-tree package from a bound workbook graph."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import TYPE_CHECKING

from excel_grapher.exporter.inverted_tree.ast_emit import (
    emit_helper_body,
    emit_rung2_scc,
    emit_rung3_scc,
    python_annotation,
    python_measure_type,
    python_return_annotation,
)
from excel_grapher.exporter.inverted_tree.catalog import BoundSeries, SeriesCatalog, build_catalog
from excel_grapher.exporter.inverted_tree.deps import (
    SeriesDeps,
    all_formula_root_cells,
    assert_subgraph_bound,
    collect_all_deps,
    formula_closure,
    leaf_closure,
    plan_indices,
)
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.exporter.inverted_tree.schedule import (
    IndexSet,
    build_scc_map,
    plan_fused_scc,
    scan_function_name,
    scc_external_params,
)

if TYPE_CHECKING:
    from excel_grapher.grapher.graph import DependencyGraph
    from excel_grapher.series_bindings.types import WorkbookSeriesBindings

_RUNTIME_PATH = Path(__file__).with_name("runtime.py")


def _data_const_name(series_id: str) -> str:
    return series_id.upper()


def _default_name(series_id: str) -> str:
    return f"{series_id.upper()}_DEFAULT"


def _py_literal(value: object) -> str:
    if isinstance(value, bool):
        return "True" if value else "False"
    if isinstance(value, int | float | str):
        return repr(value)
    if isinstance(value, tuple | list):
        parts = ", ".join(_py_literal(item) for item in value)
        if isinstance(value, list):
            return f"[{parts}]"
        if len(value) == 1:
            return f"({parts},)"
        return f"({parts})"
    item = getattr(value, "item", None)
    if callable(item):
        return _py_literal(item())
    raise InvertedTreeExportError(f"cannot emit a Python literal for {type(value).__name__}")


_GROUP_SPACES = str.maketrans("", "", " \u00a0\u202f\u2009\u2007")


def _parse_numeric_text(text: str) -> float | None:
    """Parse a number, including space-grouped thousands (`1 000`)."""
    stripped = text.strip()
    if not stripped:
        return None
    try:
        return float(stripped)
    except ValueError:
        compact = stripped.translate(_GROUP_SPACES)
        if compact == stripped:
            return None
        try:
            return float(compact)
        except ValueError:
            return None


def _cell_value(graph: DependencyGraph, address: str, dtype: str) -> object:
    """Read one cached cell as a catalog value.

    Float series keep non-numeric cached text (`n/a`, `..`, `--`, empty) as
    strings so IMF-style sentinels survive emit. Grouped numeric text such as
    `1 000` becomes a float. Remaining coercion failures name `address`.
    """
    node = graph.get_node(address)
    value = getattr(node, "value", None) if node is not None else None
    if value is None:
        return 0 if dtype in {"int", "integer", "float", "number"} else ""
    try:
        if dtype in {"int", "integer"}:
            if isinstance(value, str):
                parsed = _parse_numeric_text(value)
                if parsed is None:
                    raise ValueError(f"{value!r} is not an integer")
                return int(parsed)
            return int(value)
        if dtype in {"float", "number"}:
            if isinstance(value, bool):
                return float(value)
            if isinstance(value, int | float):
                return float(value)
            if isinstance(value, str):
                parsed = _parse_numeric_text(value)
                return float(parsed) if parsed is not None else value
            return float(value)
        if dtype in {"string", "str"}:
            return str(value)
        if dtype == "bool":
            return bool(value)
    except (TypeError, ValueError) as exc:
        raise InvertedTreeExportError(
            f"cell {address}: cannot read {dtype} value {value!r}"
        ) from exc
    if isinstance(value, int | float | str | bool):
        return value
    return value


def _series_values(series: BoundSeries, graph: DependencyGraph) -> object:
    values = [_cell_value(graph, addr, series.dtype) for addr in series.cells]
    if series.is_scalar:
        return values[0] if values else 0
    return tuple(values)


def emit_data_module(catalog: SeriesCatalog, graph: DependencyGraph) -> str:
    """Emit `data.py` with constant series and workbook-default input arrays."""
    lines = [
        '"""Constant leaves and workbook-default input arrays."""',
        "",
        "from __future__ import annotations",
        "",
    ]
    for series in catalog.constant_series():
        name = _data_const_name(series.series_id)
        value = _series_values(series, graph)
        anno = python_measure_type(series)
        if series.is_scalar:
            lines.append(f"{name}: {anno} = {_py_literal(value)}")
        else:
            lines.append(f"{name}: tuple[{anno}, ...] = {_py_literal(value)}")
        lines.append("")
    for series in catalog.input_series():
        name = _default_name(series.series_id)
        value = _series_values(series, graph)
        anno = python_measure_type(series)
        if series.is_scalar:
            lines.append(f"{name}: {anno} = {_py_literal(value)}")
        else:
            lines.append(f"{name}: tuple[{anno}, ...] = {_py_literal(value)}")
        lines.append("")
    if len(lines) == 4:
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def _scan_return_annotation(scc: tuple[str, ...], catalog: SeriesCatalog) -> str:
    parts = ", ".join(python_return_annotation(catalog.get(sid)) for sid in scc)
    return f"tuple[{parts}]"


def _helper_signature(series: BoundSeries, deps: SeriesDeps, catalog: SeriesCatalog) -> str:
    params: list[str] = []
    for param_id in deps.param_ids:
        dep = catalog.get(param_id)
        params.append(f"{param_id}: {python_annotation(dep)}")
    joined = ", ".join(params)
    ret = python_return_annotation(series)
    return f"def {series.series_id}({joined}) -> {ret}:"


def _scan_signature(
    scc: tuple[str, ...],
    param_ids: tuple[str, ...],
    catalog: SeriesCatalog,
) -> str:
    params = [f"{pid}: {python_annotation(catalog.get(pid))}" for pid in param_ids]
    joined = ", ".join(params)
    ret = _scan_return_annotation(scc, catalog)
    name = scan_function_name(scc)
    return f"def {name}({joined}) -> {ret}:"


def _emit_scc_wrapper(
    series: BoundSeries,
    scc: tuple[str, ...],
    param_ids: tuple[str, ...],
    catalog: SeriesCatalog,
) -> str:
    args = ", ".join(param_ids)
    unpack = ", ".join(scc)
    idx = scc.index(series.series_id)
    params = [f"{pid}: {python_annotation(catalog.get(pid))}" for pid in param_ids]
    joined = ", ".join(params)
    ret = python_return_annotation(series)
    fn = scan_function_name(scc)
    lines = [
        f"def {series.series_id}({joined}) -> {ret}:",
        f'    """First-level helper for bound series `{series.series_id}`."""',
        f"    {unpack} = {fn}({args})",
        f"    return {scc[idx]}",
    ]
    return "\n".join(lines)


def emit_internals_module(
    catalog: SeriesCatalog,
    deps: Mapping[str, SeriesDeps],
    graph: DependencyGraph,
    scc_map: Mapping[str, tuple[str, ...]],
) -> str:
    """Emit `internals.py` with per-series helpers and fused or demand-driven SCCs."""
    used_runtime: set[str] = set()
    functions: list[str] = []
    emitted_scans: set[tuple[str, ...]] = set()
    for series in catalog.formula_series():
        info = deps[series.series_id]
        scc = scc_map.get(series.series_id, (series.series_id,))
        if len(scc) > 1:
            param_ids = scc_external_params(scc, deps, catalog.order)
            if scc not in emitted_scans:
                emitted_scans.add(scc)
                deps_map = dict(deps)
                plan = plan_fused_scc(scc, catalog=catalog, graph=graph)
                if plan is not None:
                    body, used = emit_rung2_scc(scc, catalog=catalog, deps=deps_map, graph=graph)
                    kind = "Fused union-domain evaluation"
                else:
                    body, used = emit_rung3_scc(scc, catalog=catalog, deps=deps_map, graph=graph)
                    kind = "Demand-driven co-evaluation"
                used_runtime |= used
                joined = ", ".join(f"`{sid}`" for sid in scc)
                doc = f'    """{kind} of zipper series {joined}."""'
                functions.append("\n".join([_scan_signature(scc, param_ids, catalog), doc, *body]))
            functions.append(_emit_scc_wrapper(series, scc, param_ids, catalog))
            continue
        body, used = emit_helper_body(series, catalog=catalog, deps=info, graph=graph)
        used_runtime |= used
        if any(catalog.get(p).is_sequence and p in info.aligned_ids for p in info.param_ids):
            used_runtime.add("require_aligned")
        doc = f'    """First-level helper for bound series `{series.series_id}`."""'
        functions.append(
            "\n".join(
                [
                    _helper_signature(series, info, catalog),
                    doc,
                    *body,
                ]
            )
        )
    runtime_names = sorted(used_runtime)
    lines = [
        '"""First-level-dependency internals for the inverted graph."""',
        "",
        "from __future__ import annotations",
        "",
        "from collections.abc import Sequence",
        "",
    ]
    if runtime_names:
        names = ", ".join(runtime_names)
        lines.append(f"from .runtime import {names}")
        lines.append("")
    lines.append("")
    lines.append("\n\n".join(functions))
    lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def _identity_indices(series: BoundSeries) -> tuple[int, ...]:
    return tuple(range(len(series.cells)))


def _take_after_call(
    series_id: str,
    *,
    result_indices: Mapping[str, tuple[int, ...]],
    call_indices: Mapping[str, tuple[int, ...]],
    runtime: set[str],
) -> str | None:
    """Return a `take` assignment if the callee computed more than consumers need."""
    wanted = result_indices.get(series_id)
    computed = call_indices.get(series_id)
    if wanted is None or computed is None or wanted == computed:
        return None
    work = IndexSet.from_indices(wanted).positions_in(IndexSet.from_indices(computed))
    if work.materialize() == tuple(range(len(computed))):
        return None
    runtime.add("take")
    return f"    {series_id} = take({series_id}, {work.to_source()})"


def _aligned_call_arg(
    param_id: str,
    info: SeriesDeps,
    *,
    catalog: SeriesCatalog,
    host_call: Sequence[int],
    local_indices: Mapping[str, tuple[int, ...]],
    runtime: set[str],
) -> str:
    """Return a call-site argument, `take`n to this host walk's aligned window."""
    if param_id not in info.aligned_ids:
        return param_id
    index_map = info.index_maps.get(param_id)
    if index_map is None:
        return param_id
    needed = tuple(index_map[index] for index in host_call)
    producer = catalog.get(param_id)
    current = local_indices.get(param_id, _identity_indices(producer))
    if needed == current:
        return param_id
    try:
        work = IndexSet.from_indices(needed).positions_in(IndexSet.from_indices(current))
    except ValueError as exc:
        raise InvertedTreeExportError(
            f"series {info.host_id!r}: cannot take {param_id!r} at {needed} from {current}"
        ) from exc
    if work.materialize() == tuple(range(len(current))):
        return param_id
    runtime.add("take")
    return f"take({param_id}, {work.to_source()})"


def _group_key(
    output: BoundSeries,
    *,
    catalog: SeriesCatalog,
    deps: dict[str, SeriesDeps],
) -> tuple[str, ...]:
    """Return required input ids so baseline never shares a shocked runner."""
    leaves = leaf_closure(output.series_id, catalog=catalog, deps=deps)
    return tuple(sid for sid in leaves if catalog.get(sid).direction == "input")


def _union_leaves(
    outputs: Sequence[BoundSeries],
    *,
    catalog: SeriesCatalog,
    deps: dict[str, SeriesDeps],
) -> tuple[str, ...]:
    seen: set[str] = set()
    for output in outputs:
        seen.update(leaf_closure(output.series_id, catalog=catalog, deps=deps))
    return tuple(sid for sid in catalog.order if sid in seen)


def _union_formula_ids(
    outputs: Sequence[BoundSeries],
    *,
    catalog: SeriesCatalog,
    deps: dict[str, SeriesDeps],
    scc_map: dict[str, tuple[str, ...]] | None,
) -> tuple[str, ...]:
    seen: set[str] = set()
    ordered: list[str] = []
    for output in outputs:
        for series_id in formula_closure(
            output.series_id, catalog=catalog, deps=deps, scc_map=scc_map
        ):
            if series_id not in seen:
                seen.add(series_id)
                ordered.append(series_id)
    return tuple(ordered)


def _union_plan_indices(
    outputs: Sequence[BoundSeries],
    *,
    catalog: SeriesCatalog,
    deps: dict[str, SeriesDeps],
    scc_map: dict[str, tuple[str, ...]] | None,
) -> tuple[dict[str, tuple[int, ...]], dict[str, tuple[int, ...]]]:
    result: dict[str, tuple[int, ...]] = {}
    call: dict[str, tuple[int, ...]] = {}
    for output in outputs:
        got_result, got_call = plan_indices(output, catalog=catalog, deps=deps, scc_map=scc_map)
        for series_id, indices in got_result.items():
            previous = result.get(series_id)
            result[series_id] = tuple(
                sorted(set(indices) if previous is None else set(previous) | set(indices))
            )
        for series_id, indices in got_call.items():
            previous = call.get(series_id)
            call[series_id] = tuple(
                sorted(set(indices) if previous is None else set(previous) | set(indices))
            )
    return result, call


def _leaf_signature_parts(
    leaves: Sequence[str],
    catalog: SeriesCatalog,
) -> tuple[list[str], list[str], list[str], set[str]]:
    """Return required ids, defaulted ids, parameter lines, and `data.py` imports."""
    required = [sid for sid in leaves if catalog.get(sid).direction == "input"]
    defaulted = [sid for sid in leaves if catalog.get(sid).direction == "constant"]
    params: list[str] = []
    data_imports: set[str] = set()
    for sid in required:
        params.append(f"{sid}: {python_annotation(catalog.get(sid))}")
    for sid in defaulted:
        const_name = _data_const_name(sid)
        data_imports.add(const_name)
        params.append(f"{sid}: {python_annotation(catalog.get(sid))} = {const_name}")
    return required, defaulted, params, data_imports


def _keyword_signature(name: str, params: Sequence[str], returns: str) -> str:
    if params:
        param_block = ",\n    ".join(params)
        return f"def {name}(\n    *,\n    {param_block},\n) -> {returns}:"
    return f"def {name}() -> {returns}:"


def _emit_result_return(output: BoundSeries) -> str:
    if output.is_scalar:
        return f"    return ({output.series_id},)"
    return f"    return tuple({output.series_id})"


def _emit_evaluation_body(
    *,
    leaves: Sequence[str],
    formula_ids: Sequence[str],
    result_indices: Mapping[str, tuple[int, ...]],
    call_indices: Mapping[str, tuple[int, ...]],
    catalog: SeriesCatalog,
    deps: Mapping[str, SeriesDeps],
    scc_map: Mapping[str, tuple[str, ...]] | None,
) -> tuple[list[str], set[str], set[str]]:
    """Emit `require_length` / `take` / `internals.*` lines for one evaluation walk."""
    runtime: set[str] = set()
    body: list[str] = []
    local_indices: dict[str, tuple[int, ...]] = {}
    for sid in leaves:
        series = catalog.get(sid)
        identity = _identity_indices(series)
        local_indices[sid] = identity
        if not series.is_sequence:
            continue
        runtime.add("require_length")
        body.append(f"    require_length({sid}, {len(series.cells)})")
        wanted = result_indices.get(sid)
        if wanted is None:
            continue
        if wanted == identity:
            continue
        runtime.add("take")
        body.append(f"    {sid} = take({sid}, {IndexSet.from_indices(wanted).to_source()})")
        local_indices[sid] = wanted
    locals_bound: set[str] = set(leaves)
    seen_sccs: set[tuple[str, ...]] = set()
    mapping = scc_map or {}
    for series_id in formula_ids:
        scc = mapping.get(series_id, (series_id,))
        if len(scc) > 1:
            if scc in seen_sccs:
                continue
            seen_sccs.add(scc)
            ext = scc_external_params(scc, deps, catalog.order)
            fn = scan_function_name(scc)
            unpack = ", ".join(scc)
            body.append(f"    {unpack} = internals.{fn}({', '.join(ext)})")
            for sid in scc:
                locals_bound.add(sid)
                computed = call_indices.get(sid, _identity_indices(catalog.get(sid)))
                local_indices[sid] = computed
                taken = _take_after_call(
                    sid,
                    result_indices=result_indices,
                    call_indices=call_indices,
                    runtime=runtime,
                )
                if taken is not None:
                    body.append(taken)
                    wanted = result_indices.get(sid)
                    if wanted is not None:
                        local_indices[sid] = wanted
            continue
        info = deps[series_id]
        host_call = call_indices.get(series_id, _identity_indices(catalog.get(series_id)))
        args = ", ".join(
            _aligned_call_arg(
                param_id,
                info,
                catalog=catalog,
                host_call=host_call,
                local_indices=local_indices,
                runtime=runtime,
            )
            for param_id in info.param_ids
        )
        body.append(f"    {series_id} = internals.{series_id}({args})")
        locals_bound.add(series_id)
        computed = call_indices.get(series_id, _identity_indices(catalog.get(series_id)))
        local_indices[series_id] = computed
        taken = _take_after_call(
            series_id,
            result_indices=result_indices,
            call_indices=call_indices,
            runtime=runtime,
        )
        if taken is not None:
            body.append(taken)
            wanted = result_indices.get(series_id)
            if wanted is not None:
                local_indices[series_id] = wanted
    return body, runtime, locals_bound


def emit_orchestrator(
    output: BoundSeries,
    *,
    catalog: SeriesCatalog,
    deps: Mapping[str, SeriesDeps],
    scc_map: Mapping[str, tuple[str, ...]] | None = None,
) -> tuple[str, set[str], set[str]]:
    """Emit one public `compute_*` function.

    Returns the function source, runtime symbols used, and data constants imported.

    Multi-series lag zippers call the co-scan once and unpack members.
    """
    deps_map = dict(deps)
    scc_map_dict = dict(scc_map) if scc_map is not None else None
    leaves = leaf_closure(output.series_id, catalog=catalog, deps=deps_map)
    formula_ids = formula_closure(
        output.series_id, catalog=catalog, deps=deps_map, scc_map=scc_map_dict
    )
    result_indices, call_indices = plan_indices(
        output, catalog=catalog, deps=deps_map, scc_map=scc_map_dict
    )
    _required, _defaulted, params, data_imports = _leaf_signature_parts(leaves, catalog)
    compute_name = output.compute_name or f"compute_{output.series_id}"
    signature = _keyword_signature(compute_name, params, python_return_annotation(output))
    body, runtime, locals_bound = _emit_evaluation_body(
        leaves=leaves,
        formula_ids=formula_ids,
        result_indices=result_indices,
        call_indices=call_indices,
        catalog=catalog,
        deps=deps,
        scc_map=scc_map,
    )
    if output.series_id in locals_bound:
        body.append(_emit_result_return(output))
    else:
        body.append(f"    return internals.{output.series_id}()")
    doc = f'    """Compute `{output.series_id}` from its subgraph leaf closure."""'
    source = "\n".join([signature, doc, *body])
    return source, runtime, data_imports


def _emit_shared_runner(
    outputs: Sequence[BoundSeries],
    *,
    name: str,
    catalog: SeriesCatalog,
    deps: dict[str, SeriesDeps],
    scc_map: dict[str, tuple[str, ...]] | None,
) -> tuple[str, set[str], set[str]]:
    """Emit one private evaluation walk shared by `outputs`."""
    leaves = _union_leaves(outputs, catalog=catalog, deps=deps)
    formula_ids = _union_formula_ids(outputs, catalog=catalog, deps=deps, scc_map=scc_map)
    result_indices, call_indices = _union_plan_indices(
        outputs, catalog=catalog, deps=deps, scc_map=scc_map
    )
    _required, _defaulted, params, data_imports = _leaf_signature_parts(leaves, catalog)
    returns = ", ".join(python_return_annotation(output) for output in outputs)
    signature = _keyword_signature(name, params, f"tuple[{returns}]")
    body, runtime, _bound = _emit_evaluation_body(
        leaves=leaves,
        formula_ids=formula_ids,
        result_indices=result_indices,
        call_indices=call_indices,
        catalog=catalog,
        deps=deps,
        scc_map=scc_map,
    )
    returned = ", ".join(output.series_id for output in outputs)
    body.append(f"    return {returned}")
    joined = ", ".join(f"`{output.series_id}`" for output in outputs)
    doc = f'    """Evaluate the shared formula closure of {joined}."""'
    return "\n".join([signature, doc, *body]), runtime, data_imports


def _emit_thin_orchestrator(
    output: BoundSeries,
    *,
    runner_name: str,
    outputs: Sequence[BoundSeries],
    runner_leaves: Sequence[str],
    catalog: SeriesCatalog,
    deps: dict[str, SeriesDeps],
) -> tuple[str, set[str]]:
    """Emit a `compute_*` that unpacks one slot from a shared runner."""
    leaves = leaf_closure(output.series_id, catalog=catalog, deps=deps)
    _required, _defaulted, params, data_imports = _leaf_signature_parts(leaves, catalog)
    compute_name = output.compute_name or f"compute_{output.series_id}"
    signature = _keyword_signature(compute_name, params, python_return_annotation(output))
    output_leaves = set(leaves)
    call_args: list[str] = []
    for sid in runner_leaves:
        if catalog.get(sid).direction == "input" or sid in output_leaves:
            call_args.append(f"{sid}={sid}")
    unpack = ", ".join(
        member.series_id if member.series_id == output.series_id else "_" for member in outputs
    )
    call = f"{runner_name}({', '.join(call_args)})"
    body = [
        f"    {unpack} = {call}",
        _emit_result_return(output),
    ]
    doc = f'    """Compute `{output.series_id}` from its subgraph leaf closure."""'
    return "\n".join([signature, doc, *body]), data_imports


def emit_api_module(
    catalog: SeriesCatalog,
    deps: Mapping[str, SeriesDeps],
    scc_map: Mapping[str, tuple[str, ...]] | None = None,
) -> str:
    """Emit `api.py` with keyword-only output orchestrators.

    Outputs that share the same required input leaves share one private
    runner; each `compute_*` keeps its own leaf signature and unpacks its
    slot. Baseline and shocked closures stay apart because their required
    inputs differ.
    """
    functions: list[str] = []
    runtime: set[str] = set()
    data_imports: set[str] = set()
    compute_names: list[str] = []
    deps_map = dict(deps)
    scc_map_dict = dict(scc_map) if scc_map is not None else None
    groups: dict[tuple[str, ...], list[BoundSeries]] = {}
    group_order: list[tuple[str, ...]] = []
    for output in catalog.output_series():
        key = _group_key(output, catalog=catalog, deps=deps_map)
        if key not in groups:
            groups[key] = []
            group_order.append(key)
        groups[key].append(output)
    runner_index = 0
    for key in group_order:
        members = groups[key]
        compute_names.extend(
            output.compute_name or f"compute_{output.series_id}" for output in members
        )
        if len(members) == 1:
            source, used_runtime, used_data = emit_orchestrator(
                members[0], catalog=catalog, deps=deps, scc_map=scc_map
            )
            functions.append(source)
            runtime |= used_runtime
            data_imports |= used_data
            continue
        runner_name = f"_run_{runner_index}"
        runner_index += 1
        runner_source, used_runtime, used_data = _emit_shared_runner(
            members,
            name=runner_name,
            catalog=catalog,
            deps=deps_map,
            scc_map=scc_map_dict,
        )
        functions.append(runner_source)
        runtime |= used_runtime
        data_imports |= used_data
        runner_leaves = _union_leaves(members, catalog=catalog, deps=deps_map)
        for output in members:
            source, used_data = _emit_thin_orchestrator(
                output,
                runner_name=runner_name,
                outputs=members,
                runner_leaves=runner_leaves,
                catalog=catalog,
                deps=deps_map,
            )
            functions.append(source)
            data_imports |= used_data
    lines = [
        '"""Output orchestrators for the inverted graph.',
        "",
        "Each `compute_*` function takes the leaf closure of its subgraph.",
        "There is no evaluation context and no input setters.",
        '"""',
        "",
        "from __future__ import annotations",
        "",
        "from collections.abc import Sequence",
        "",
        "from . import internals",
    ]
    if data_imports:
        imported = ", ".join(sorted(data_imports))
        lines.append(f"from .data import {imported}")
    if runtime:
        lines.append(f"from .runtime import {', '.join(sorted(runtime))}")
    lines.append("")
    lines.append("")
    lines.append("\n\n".join(functions))
    lines.append("")
    lines.append("__all__ = [")
    for name in compute_names:
        lines.append(f"    {name!r},")
    lines.append("]")
    lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def emit_init_module(catalog: SeriesCatalog) -> str:
    """Emit package `__init__.py` re-exporting public `compute_*` functions."""
    names = [s.compute_name or f"compute_{s.series_id}" for s in catalog.output_series()]
    if names:
        imported = ", ".join(names)
        import_line = f"from .api import {imported}"
    else:
        import_line = ""
    lines = [
        '"""Inverted-tree mechanical extraction."""',
        "",
        "from __future__ import annotations",
        "",
    ]
    if import_line:
        lines.append(import_line)
        lines.append("")
    lines.append("__all__ = [")
    for name in names:
        lines.append(f"    {name!r},")
    lines.append("]")
    lines.append("")
    return "\n".join(lines)


def generate_inverted_tree_modules(
    graph: DependencyGraph,
    *,
    series_bindings: WorkbookSeriesBindings,
    bindings_workbook: Path | str,
    targets: Sequence[str] | None = None,
) -> dict[str, str]:
    """Generate api/internals/runtime/data modules for the inverted-tree paradigm.

    Args:
        graph: Dependency graph covering the binding closure.
        series_bindings: Bindings catalog (inputs, constants, internals, outputs).
        bindings_workbook: Workbook path used to expand `data_range`s.
        targets: Ignored; outputs come from the bindings catalog.

    Returns:
        Mapping of package filenames to file contents.

    Raises:
        InvertedTreeExportError: A bound series cannot be inverted fail-closed.
    """
    del targets
    catalog = build_catalog(series_bindings, workbook=bindings_workbook, graph=graph)
    if not catalog.output_series():
        raise InvertedTreeExportError("inverted-tree codegen requires at least one output series")
    deps = collect_all_deps(catalog, graph)
    scc_map = build_scc_map(catalog, deps, graph)
    assert_subgraph_bound(
        catalog=catalog,
        graph=graph,
        roots=list(all_formula_root_cells(catalog)),
    )
    runtime_py = _RUNTIME_PATH.read_text(encoding="utf-8")
    return {
        "__init__.py": emit_init_module(catalog),
        "api.py": emit_api_module(catalog, deps, scc_map),
        "data.py": emit_data_module(catalog, graph),
        "runtime.py": runtime_py if runtime_py.endswith("\n") else runtime_py + "\n",
        "internals.py": emit_internals_module(catalog, deps, graph, scc_map),
    }
