"""Emit an inverted-tree package from a bound workbook graph."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from datetime import date, datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any, Literal

from excel_grapher.exporter.inverted_tree.ast_emit import (
    _is_identity_aligned,
    emit_helper_body,
    emit_rung2_scc,
    emit_rung3_scc,
    python_annotation,
    python_data_annotation,
    python_return_annotation,
)
from excel_grapher.exporter.inverted_tree.catalog import BoundSeries, SeriesCatalog, build_catalog
from excel_grapher.exporter.inverted_tree.deps import (
    CatalogEdges,
    DependenceEdge,
    SeriesDeps,
    all_formula_root_cells,
    assert_subgraph_bound,
    collect_all_deps,
    collect_catalog_edges,
    formula_closure,
    leaf_closure,
    plan_indices,
)
from excel_grapher.exporter.inverted_tree.domains import (
    DomainEmitPlan,
    constants_attr_source,
    domain_annotation,
    domain_const_name,
    key_domain_attr_source,
    plan_domain_emission,
    publish_attr_source,
)
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.exporter.inverted_tree.schedule import (
    IndexSet,
    Rung,
    SccPlan,
    build_scc_map,
    indices_to_source,
    plan_scc,
    scan_function_name,
    scc_external_params,
)
from excel_grapher.series_bindings.input_coerce import (
    input_value_map_from_series,
    measure_domain_from_series,
)
from excel_grapher.series_bindings.validate import validate_series_bindings

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
    if isinstance(value, datetime):
        return (
            "datetime("
            f"{value.year}, {value.month}, {value.day}, "
            f"{value.hour}, {value.minute}, {value.second}"
            f"{f', {value.microsecond}' if value.microsecond else ''})"
        )
    if isinstance(value, date):
        return f"datetime({value.year}, {value.month}, {value.day})"
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


def _coerce_cached_value(value: Any, dtype: str, address: str) -> object:
    """Coerce a cached workbook value to a catalog dtype.

    Float series keep non-numeric cached text (`n/a`, `..`, `--`, empty) as
    strings so IMF-style sentinels survive emit. Grouped numeric text such as
    `1 000` becomes a float. Remaining coercion failures name `address`.
    """
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
        if dtype in {"datetime", "date"}:
            if isinstance(value, datetime):
                return value
            if isinstance(value, date):
                return datetime(value.year, value.month, value.day)
            raise ValueError(f"{value!r} is not a datetime")
    except (TypeError, ValueError) as exc:
        raise InvertedTreeExportError(
            f"cell {address}: cannot read {dtype} value {value!r}"
        ) from exc
    if isinstance(value, int | float | str | bool):
        return value
    return value


def _cell_value(graph: DependencyGraph, address: str, dtype: str) -> object:
    """Read one cached cell as a catalog value."""
    node = graph.get_node(address)
    value = getattr(node, "value", None) if node is not None else None
    if value is None:
        return 0 if dtype in {"int", "integer", "float", "number"} else ""
    return _coerce_cached_value(value, dtype, address)


def _series_values(series: BoundSeries, graph: DependencyGraph) -> object:
    values = [_cell_value(graph, addr, series.dtype) for addr in series.cells]
    if series.is_scalar:
        return values[0] if values else 0
    return tuple(values)


def _uses_datetime(catalog: SeriesCatalog, plan: DomainEmitPlan | None = None) -> bool:
    if any(series.python_dtype == "datetime" for series in catalog.series.values()):
        return True
    return bool(plan is not None and plan.uses_datetime)


def _emit_domain_constants(plan: DomainEmitPlan) -> list[str]:
    """Emit one tuple per distinct key domain, then interned subset domains."""
    lines: list[str] = []
    for field, values in plan.field_domains.items():
        name = domain_const_name(field)
        lines.append(f"{name}: {domain_annotation(values)} = {_py_literal(values)}")
        lines.append("")
    for name, values in plan.interned:
        lines.append(f"{name}: {domain_annotation(values)} = {_py_literal(values)}")
        lines.append("")
    return lines


def emit_data_module(
    catalog: SeriesCatalog,
    graph: DependencyGraph,
    *,
    domains: DomainEmitPlan | None = None,
) -> str:
    """Emit `data.py` with key domains, constants, and workbook-default inputs."""
    plan = domains if domains is not None else plan_domain_emission(catalog)
    constants = catalog.constant_series()
    lines = [
        '"""Constant leaves, key domains, and workbook-default input arrays."""',
        "",
        "from __future__ import annotations",
        "",
    ]
    stdlib: list[str] = []
    if constants:
        stdlib.extend(
            [
                "from collections.abc import Iterator",
                "from contextlib import contextmanager",
            ]
        )
    if _uses_datetime(catalog, plan):
        stdlib.append("from datetime import datetime")
    if stdlib:
        lines.extend([*stdlib, ""])
    lines.extend(_emit_domain_constants(plan))
    constant_names: list[str] = []
    for series in constants:
        name = _data_const_name(series.series_id)
        constant_names.append(name)
        value = _series_values(series, graph)
        lines.append(f"{name}: {python_data_annotation(series)} = {_py_literal(value)}")
        lines.append("")
    for series in catalog.input_series():
        name = _default_name(series.series_id)
        value = _series_values(series, graph)
        lines.append(f"{name}: {python_data_annotation(series)} = {_py_literal(value)}")
        lines.append("")
    if constant_names:
        names_literal = ", ".join(repr(name) for name in constant_names)
        lines.extend(
            [
                f"_CONSTANT_NAMES = frozenset({{{names_literal}}})",
                "",
                "",
                "@contextmanager",
                "def overrides(**values: object) -> Iterator[None]:",
                '    """Replace constant attributes for the duration of the `with` block.',
                "",
                "    Args:",
                "        **values: Mapping of this module's constant names to replacements.",
                "            Unknown names raise `AttributeError`.",
                '    """',
                "    unknown = [name for name in values if name not in _CONSTANT_NAMES]",
                "    if unknown:",
                '        joined = ", ".join(repr(name) for name in unknown)',
                '        raise AttributeError("unknown constant(s): " + joined)',
                "    namespace = globals()",
                "    saved = {name: namespace[name] for name in values}",
                "    namespace.update(values)",
                "    try:",
                "        yield",
                "    finally:",
                "        namespace.update(saved)",
                "",
            ]
        )
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


def _effective_rung(choice: SccPlan, force_rung: Literal[2, 3] | None) -> Rung:
    """Resolve the emit rung, applying `force_rung` when it is legal."""
    if force_rung == 3:
        return 3
    if force_rung == 2 and choice.plan is not None:
        return 2 if len(choice.plan.scc) > 1 else 1
    return choice.rung


def _emit_by_rung(
    scc: tuple[str, ...],
    *,
    series: BoundSeries,
    catalog: SeriesCatalog,
    deps: dict[str, SeriesDeps],
    graph: DependencyGraph,
    edges: Sequence[DependenceEdge],
    choice: SccPlan,
    rung: Rung,
) -> tuple[list[str], set[str]]:
    """Emit one SCC body by resolved rung."""
    match rung:
        case 3:
            return emit_rung3_scc(scc, catalog=catalog, deps=deps, graph=graph, edges=edges)
        case 1 | 2:
            return emit_rung2_scc(
                scc,
                catalog=catalog,
                deps=deps,
                graph=graph,
                plan=choice.plan,
                edges=edges,
            )
        case 0:
            return emit_helper_body(
                series, catalog=catalog, deps=deps[series.series_id], graph=graph
            )
        case _:
            raise ValueError(f"unknown rung {rung!r}")


def _key_domain_attrs(
    name: str,
    *,
    series_id: str | None = None,
    scc: tuple[str, ...] | None = None,
    plan: DomainEmitPlan,
    holes: tuple[int, ...] = (),
) -> str:
    if scc is not None:
        source = key_domain_attr_source(
            name, keys=plan.scc_key[scc], domain_expr=plan.scc_expr[scc]
        )
    elif series_id is None:
        raise ValueError("series_id or scc is required")
    else:
        source = key_domain_attr_source(
            name, keys=plan.series_key[series_id], domain_expr=plan.series_expr[series_id]
        )
    return f"{source}\n{publish_attr_source(name, '__holes__', repr(holes))}"


def emit_internals_module(
    catalog: SeriesCatalog,
    deps: Mapping[str, SeriesDeps],
    graph: DependencyGraph,
    scc_map: Mapping[str, tuple[str, ...]],
    catalog_edges: CatalogEdges,
    *,
    force_rung: Literal[2, 3] | None = None,
    domains: DomainEmitPlan | None = None,
) -> str:
    """Emit `internals.py` with per-series helpers and fused or demand-driven SCCs."""
    if force_rung not in (None, 2, 3):
        raise ValueError(f"force_rung must be 2, 3, or None; got {force_rung!r}")
    plan = domains if domains is not None else plan_domain_emission(catalog, scc_map)
    used_runtime: set[str] = set()
    functions: list[str] = []
    emitted_scans: set[tuple[str, ...]] = set()
    edges = catalog_edges.edges
    deps_map = dict(deps)
    needs_data = False
    for series in catalog.formula_series():
        info = deps[series.series_id]
        scc = scc_map.get(series.series_id, (series.series_id,))
        if len(scc) > 1 and scc in emitted_scans:
            continue
        choice = plan_scc(scc, catalog=catalog, graph=graph, edges=edges)
        rung = _effective_rung(choice, force_rung)
        body, used = _emit_by_rung(
            scc,
            series=series,
            catalog=catalog,
            deps=deps_map,
            graph=graph,
            edges=edges,
            choice=choice,
            rung=rung,
        )
        used_runtime |= used
        if len(scc) > 1:
            emitted_scans.add(scc)
            param_ids = scc_external_params(scc, deps, catalog.order)
            kind = "Demand-driven co-evaluation" if rung == 3 else "Fused union-domain evaluation"
            joined = ", ".join(f"`{sid}`" for sid in scc)
            doc = _helper_docstring(f"{kind} of zipper series {joined}.")
            source = "\n".join([_scan_signature(scc, param_ids, catalog), doc, *body])
            source = f"{source}\n{_key_domain_attrs(scan_function_name(scc), scc=scc, plan=plan)}"
            functions.append(source)
            needs_data = needs_data or plan.uses_data_scc(scc)
            continue
        if rung == 0 and any(
            catalog.get(param).is_sequence and param in info.aligned_ids for param in info.param_ids
        ):
            used_runtime.add("require_aligned")
        if rung == 3:
            summary = f"Demand-driven evaluation of series `{series.series_id}`."
        else:
            summary = f"First-level helper for bound series `{series.series_id}`."
        doc = _helper_docstring(summary, series)
        source = "\n".join([_helper_signature(series, info, catalog), doc, *body])
        attrs = _key_domain_attrs(
            series.series_id,
            series_id=series.series_id,
            plan=plan,
            holes=series.hole_indices,
        )
        source = f"{source}\n{attrs}"
        functions.append(source)
        needs_data = needs_data or plan.uses_data(series.series_id)
    runtime_names = sorted(used_runtime)
    lines = [
        '"""First-level-dependency internals for the inverted graph."""',
        "",
        "from __future__ import annotations",
        "",
        "from collections.abc import Sequence",
        "",
    ]
    if _uses_datetime(catalog, plan):
        lines.extend(["from datetime import datetime", ""])
    if needs_data:
        lines.append("from . import data")
        lines.append("")
    if runtime_names:
        names = ", ".join(runtime_names)
        lines.append(f"from .runtime import {names}")
        lines.append("")
    lines.append("")
    lines.append("\n\n".join(functions))
    lines.append("")
    return "\n".join(lines).rstrip() + "\n"


_HOLE_DOC_LABELS = {
    "blank": "blank",
    "off_closure": "not computed",
    "literal": "cached literal",
    "graph_leaf": "cached literal",
}


def _hole_doc_lines(series: BoundSeries) -> list[str]:
    """Return indented docstring lines naming retained hole cells by category."""
    if not series.holes:
        return []
    grouped: dict[str, list[str]] = {}
    for hole in series.holes:
        grouped.setdefault(_HOLE_DOC_LABELS[hole.kind], []).append(f"`{hole.address}`")
    lines = ["", "    Hole cells kept in the bound rectangle:"]
    for label, addresses in grouped.items():
        lines.append(f"    {label}: {', '.join(addresses)}")
    return lines


def _helper_docstring(summary: str, series: BoundSeries | None = None) -> str:
    lines = [
        f'    """{summary}',
        "",
        "    Each sequence argument must be dense over the producer's `__domain__`.",
        "    Holed series are shorter than the public domain.",
    ]
    if series is not None:
        lines.extend(_hole_doc_lines(series))
    lines.append('    """')
    return "\n".join(lines)


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
    leaf_source: Mapping[str, str],
    runtime: set[str],
) -> str:
    """Return a call-site argument, `take`n to this host walk's aligned window."""
    expr = leaf_source.get(param_id, param_id)
    if param_id not in info.aligned_ids:
        return expr
    affine = info.affine_maps.get(param_id)
    index_map = info.index_maps.get(param_id)
    if affine is not None:
        coeff, offset = affine
        needed = tuple(coeff * index + offset for index in host_call)
    elif index_map is not None:
        needed = tuple(index_map[index] for index in host_call)
    else:
        return expr
    producer = catalog.get(param_id)
    current = local_indices.get(param_id, _identity_indices(producer))
    if needed == current:
        return expr
    pos = {value: index for index, value in enumerate(current)}
    try:
        work = tuple(pos[index] for index in needed)
    except KeyError as exc:
        raise InvertedTreeExportError(
            f"series {info.host_id!r}: cannot take {param_id!r} at {needed} from {current}"
        ) from exc
    if work == tuple(range(len(current))):
        return expr
    runtime.add("take")
    return f"take({expr}, {indices_to_source(work)})"


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
) -> tuple[list[str], list[str], list[str]]:
    """Return input ids, constant ids, and input-only parameter lines."""
    required = [sid for sid in leaves if catalog.get(sid).direction == "input"]
    constants = [sid for sid in leaves if catalog.get(sid).direction == "constant"]
    params = [f"{sid}: {python_annotation(catalog.get(sid))}" for sid in required]
    return required, constants, params


def _compute_attrs(
    name: str,
    constants: Sequence[str],
    series: BoundSeries,
    plan: DomainEmitPlan,
) -> str:
    return "\n".join(
        [
            constants_attr_source(name, constants),
            _key_domain_attrs(
                name,
                series_id=series.series_id,
                plan=plan,
                holes=series.hole_indices,
            ),
        ]
    )


def _leaf_source_map(leaves: Sequence[str], catalog: SeriesCatalog) -> dict[str, str]:
    """Map leaf ids to the expression the orchestrator passes to internals."""
    sources: dict[str, str] = {}
    for sid in leaves:
        if catalog.get(sid).direction == "constant":
            sources[sid] = f"data.{_data_const_name(sid)}"
        else:
            sources[sid] = sid
    return sources


def _emit_input_domain_checks(
    leaves: Sequence[str],
    catalog: SeriesCatalog,
) -> tuple[list[str], set[str]]:
    """Emit `require_input_domain` calls for input leaves that declare a domain."""
    lines: list[str] = []
    runtime: set[str] = set()
    for series_id in leaves:
        series = catalog.get(series_id)
        if series.direction != "input":
            continue
        domain = measure_domain_from_series(series.raw)
        if domain is None:
            continue
        runtime.add("require_input_domain")
        lines.append(f"    require_input_domain({series_id}, {domain!r}, series_id={series_id!r})")
    return lines, runtime


def _emit_input_value_maps(
    leaves: Sequence[str],
    catalog: SeriesCatalog,
) -> tuple[list[str], set[str]]:
    """Emit `apply_input_value_map` after domain checks, before evaluation."""
    lines: list[str] = []
    runtime: set[str] = set()
    for series_id in leaves:
        series = catalog.get(series_id)
        if series.direction != "input":
            continue
        mapping = input_value_map_from_series(series.raw)
        if mapping is None:
            continue
        runtime.add("apply_input_value_map")
        lines.append(
            f"    {series_id} = apply_input_value_map("
            f"{series_id}, {mapping!r}, series_id={series_id!r})"
        )
    return lines, runtime


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
    """Emit domain/map guards, length checks, takes, and internals calls."""
    runtime: set[str] = set()
    body, domain_runtime = _emit_input_domain_checks(leaves, catalog)
    runtime |= domain_runtime
    map_lines, map_runtime = _emit_input_value_maps(leaves, catalog)
    body.extend(map_lines)
    runtime |= map_runtime
    local_indices: dict[str, tuple[int, ...]] = {}
    leaf_source = _leaf_source_map(leaves, catalog)
    for sid in leaves:
        series = catalog.get(sid)
        identity = _identity_indices(series)
        local_indices[sid] = identity
        is_constant = series.direction == "constant"
        if not series.is_sequence:
            continue
        if not is_constant:
            runtime.add("require_length")
            body.append(f"    require_length({sid}, {len(series.cells)})")
        wanted = result_indices.get(sid)
        if wanted is None or wanted == identity:
            continue
        runtime.add("take")
        body.append(
            f"    {sid} = take({leaf_source[sid]}, {IndexSet.from_indices(wanted).to_source()})"
        )
        leaf_source[sid] = sid
        local_indices[sid] = wanted
    locals_bound: set[str] = {
        sid for sid in leaves if catalog.get(sid).direction != "constant" or leaf_source[sid] == sid
    }
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
            args = ", ".join(leaf_source.get(param_id, param_id) for param_id in ext)
            body.append(f"    {unpack} = internals.{fn}({args})")
            for sid in scc:
                locals_bound.add(sid)
                leaf_source[sid] = sid
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
                leaf_source=leaf_source,
                runtime=runtime,
            )
            for param_id in info.param_ids
        )
        body.append(f"    {series_id} = internals.{series_id}({args})")
        locals_bound.add(series_id)
        leaf_source[series_id] = series_id
        host_series = catalog.get(series_id)
        host_n = len(host_series.cells)
        # Identity-aligned helpers set `n` from the taken args, so the result
        # window is `call_indices`. A helper with only `n = host_n` (scalar
        # params, irregular gathers, #695) returns the full catalog.
        uses_aligned_n = any(
            catalog.get(param_id).is_sequence and _is_identity_aligned(info, param_id, host_n)
            for param_id in info.param_ids
        )
        computed = (
            call_indices.get(series_id, _identity_indices(host_series))
            if info.is_scan or uses_aligned_n
            else _identity_indices(host_series)
        )
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
    domains: DomainEmitPlan | None = None,
) -> tuple[str, set[str], bool]:
    """Emit one public `compute_*` function.

    Returns the function source, runtime symbols used, and whether the body
    reads `data`.

    Input leaves that declare `input.domain` are checked with
    `require_input_domain` before the evaluation walk. Scalar inputs that
    declare `input.value_map` are rewritten with `apply_input_value_map` after
    that check, still before any expression sees the value. Multi-series lag
    zippers call the co-scan once and unpack members.
    """
    deps_map = dict(deps)
    scc_map_dict = dict(scc_map) if scc_map is not None else None
    plan = domains if domains is not None else plan_domain_emission(catalog, scc_map_dict)
    leaves = leaf_closure(output.series_id, catalog=catalog, deps=deps_map)
    formula_ids = formula_closure(
        output.series_id, catalog=catalog, deps=deps_map, scc_map=scc_map_dict
    )
    result_indices, call_indices = plan_indices(
        output, catalog=catalog, deps=deps_map, scc_map=scc_map_dict
    )
    _required, constants, params = _leaf_signature_parts(leaves, catalog)
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
    doc = f'    """Compute `{output.series_id}` from its input leaf closure."""'
    source = "\n".join([signature, doc, *body])
    source = f"{source}\n{_compute_attrs(compute_name, constants, output, plan)}"
    return source, runtime, bool(constants) or plan.uses_data(output.series_id)


def _emit_shared_runner(
    outputs: Sequence[BoundSeries],
    *,
    name: str,
    catalog: SeriesCatalog,
    deps: dict[str, SeriesDeps],
    scc_map: dict[str, tuple[str, ...]] | None,
) -> tuple[str, set[str], bool]:
    """Emit one private evaluation walk shared by `outputs`."""
    leaves = _union_leaves(outputs, catalog=catalog, deps=deps)
    formula_ids = _union_formula_ids(outputs, catalog=catalog, deps=deps, scc_map=scc_map)
    result_indices, call_indices = _union_plan_indices(
        outputs, catalog=catalog, deps=deps, scc_map=scc_map
    )
    _required, constants, params = _leaf_signature_parts(leaves, catalog)
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
    return "\n".join([signature, doc, *body]), runtime, bool(constants)


def _emit_thin_orchestrator(
    output: BoundSeries,
    *,
    runner_name: str,
    outputs: Sequence[BoundSeries],
    runner_leaves: Sequence[str],
    catalog: SeriesCatalog,
    deps: dict[str, SeriesDeps],
    domains: DomainEmitPlan,
) -> tuple[str, set[str]]:
    """Emit a `compute_*` that unpacks one slot from a shared runner."""
    leaves = leaf_closure(output.series_id, catalog=catalog, deps=deps)
    _required, constants, params = _leaf_signature_parts(leaves, catalog)
    compute_name = output.compute_name or f"compute_{output.series_id}"
    signature = _keyword_signature(compute_name, params, python_return_annotation(output))
    call_args: list[str] = []
    for sid in runner_leaves:
        if catalog.get(sid).direction == "input":
            call_args.append(f"{sid}={sid}")
    unpack = ", ".join(
        member.series_id if member.series_id == output.series_id else "_" for member in outputs
    )
    call = f"{runner_name}({', '.join(call_args)})"
    domain_lines, domain_runtime = _emit_input_domain_checks(leaves, catalog)
    body = [
        *domain_lines,
        f"    {unpack} = {call}",
        _emit_result_return(output),
    ]
    doc = f'    """Compute `{output.series_id}` from its input leaf closure."""'
    source = "\n".join([signature, doc, *body])
    return (
        f"{source}\n{_compute_attrs(compute_name, constants, output, domains)}",
        domain_runtime,
    )


def emit_api_module(
    catalog: SeriesCatalog,
    deps: Mapping[str, SeriesDeps],
    scc_map: Mapping[str, tuple[str, ...]] | None = None,
    *,
    domains: DomainEmitPlan | None = None,
) -> str:
    """Emit `api.py` with keyword-only output orchestrators.

    Outputs that share the same required input leaves share one private
    runner; each `compute_*` keeps its own input-leaf signature and unpacks
    its slot. Baseline and shocked closures stay apart because their required
    inputs differ. Constant leaves are read from `data`. Each `compute_*`
    publishes `__key__` / `__domain__` alongside `__constants__`.
    """
    functions: list[str] = []
    runtime: set[str] = set()
    uses_data = False
    needs_sequence = False
    compute_names: list[str] = []
    deps_map = dict(deps)
    scc_map_dict = dict(scc_map) if scc_map is not None else None
    plan = domains if domains is not None else plan_domain_emission(catalog, scc_map_dict)
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
        if any(not catalog.get(sid).is_scalar for sid in key):
            needs_sequence = True
        if len(members) == 1:
            source, used_runtime, used_data = emit_orchestrator(
                members[0], catalog=catalog, deps=deps, scc_map=scc_map, domains=plan
            )
            functions.append(source)
            runtime |= used_runtime
            uses_data = uses_data or used_data
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
        uses_data = uses_data or used_data
        runner_leaves = _union_leaves(members, catalog=catalog, deps=deps_map)
        for output in members:
            source, used_runtime = _emit_thin_orchestrator(
                output,
                runner_name=runner_name,
                outputs=members,
                runner_leaves=runner_leaves,
                catalog=catalog,
                deps=deps_map,
                domains=plan,
            )
            functions.append(source)
            runtime |= used_runtime
            uses_data = uses_data or plan.uses_data(output.series_id)
    lines = [
        '"""Output orchestrators for the inverted graph.',
        "",
        "Each `compute_*` function takes the input leaf closure of its subgraph.",
        "Constant leaves are read from `data` and listed on `__constants__`.",
        "Key domains are listed on `__key__` / `__domain__`.",
        "There is no evaluation context and no input setters.",
        '"""',
        "",
        "from __future__ import annotations",
        "",
    ]
    stdlib: list[str] = []
    if needs_sequence:
        stdlib.append("from collections.abc import Sequence")
    if _uses_datetime(catalog, plan):
        stdlib.append("from datetime import datetime")
    if stdlib:
        lines.extend([*stdlib, ""])
    if uses_data:
        lines.append("from . import data")
    lines.append("from . import internals")
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
        "from . import data",
        "from .runtime import as_records",
        "",
    ]
    if import_line:
        lines.append(import_line)
        lines.append("")
    lines.append("__all__ = [")
    lines.append(f"    {'as_records'!r},")
    for name in names:
        lines.append(f"    {name!r},")
    lines.append("]")
    lines.append("")
    return "\n".join(lines)


_REFUSE_BINDING_CODES = frozenset(
    {
        "non_leaf_input_overlap",
        "no_formula_override_targets",
    }
)


def _refuse_invalid_bindings(
    graph: DependencyGraph,
    series_bindings: WorkbookSeriesBindings,
    bindings_workbook: Path | str,
) -> None:
    """Fail closed when bindings would emit a formula cell as a plain input.

    Other validator errors (`duplicate_key`, `bind_resolution_failed`, unbound
    ranges) stay with emit's own fail-closed checks so their messages remain
    specific.
    """
    report = validate_series_bindings(graph, series_bindings, workbook=bindings_workbook)
    codes = sorted(
        {
            issue["code"]
            for issue in report["issues"]
            if issue["level"] == "error" and issue["code"] in _REFUSE_BINDING_CODES
        }
    )
    if codes:
        raise InvertedTreeExportError(f"invalid series bindings ({', '.join(codes)})")


def generate_inverted_tree_modules(
    graph: DependencyGraph,
    *,
    series_bindings: WorkbookSeriesBindings,
    bindings_workbook: Path | str,
    targets: Sequence[str] | None = None,
    force_rung: Literal[2, 3] | None = None,
) -> dict[str, str]:
    """Generate api/internals/runtime/data modules for the inverted-tree paradigm.

    Args:
        graph: Dependency graph covering the binding closure.
        series_bindings: Bindings catalog (inputs, constants, internals, outputs).
        bindings_workbook: Workbook path used to expand `data_range`s.
        targets: Ignored; outputs come from the bindings catalog.
        force_rung: Pin every formula SCC to rung 3 (demand-driven), or
            fuse wherever legal (`2`) and fall through to the auto rung
            otherwise. `None` selects the strongest legal rung.

    Returns:
        Mapping of package filenames to file contents.

    Raises:
        InvertedTreeExportError: A bound series cannot be inverted fail-closed,
            or an input range overlaps an on-graph formula cell without
            `input.mode: override`.
    """
    del targets
    _refuse_invalid_bindings(graph, series_bindings, bindings_workbook)
    catalog = build_catalog(series_bindings, workbook=bindings_workbook, graph=graph)
    if not catalog.output_series():
        raise InvertedTreeExportError("inverted-tree codegen requires at least one output series")
    catalog_edges = collect_catalog_edges(catalog, graph)
    deps = collect_all_deps(catalog, graph, catalog_edges=catalog_edges)
    scc_map = build_scc_map(catalog, deps, edges=catalog_edges.edges)
    domains = plan_domain_emission(catalog, scc_map)
    assert_subgraph_bound(
        catalog=catalog,
        graph=graph,
        roots=list(all_formula_root_cells(catalog)),
    )
    runtime_py = _RUNTIME_PATH.read_text(encoding="utf-8")
    return {
        "__init__.py": emit_init_module(catalog),
        "api.py": emit_api_module(catalog, deps, scc_map, domains=domains),
        "data.py": emit_data_module(catalog, graph, domains=domains),
        "runtime.py": runtime_py if runtime_py.endswith("\n") else runtime_py + "\n",
        "internals.py": emit_internals_module(
            catalog,
            deps,
            graph,
            scc_map,
            catalog_edges=catalog_edges,
            force_rung=force_rung,
            domains=domains,
        ),
    }
