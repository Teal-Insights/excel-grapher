"""Emit an inverted-tree package from a bound workbook graph."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass, replace
from datetime import date, datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any, Literal

from excel_grapher.exporter.inverted_tree.catalog import BoundSeries, SeriesCatalog, build_catalog
from excel_grapher.exporter.inverted_tree.deps import (
    SeriesDeps,
    all_formula_root_cells,
    assert_subgraph_bound,
    bind_blank_rects,
    collect_all_deps,
    collect_catalog_edges,
    formula_closure,
    leaf_closure,
    plan_indices,
    reset_blank_rects,
)
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.exporter.inverted_tree.named_emit import emit_named_modules
from excel_grapher.exporter.inverted_tree.schedule import (
    build_scc_map,
    plan_scc,
)
from excel_grapher.grapher.blank_ranges import normalize_blank_range_specs
from excel_grapher.series_bindings.validate import validate_series_bindings

if TYPE_CHECKING:
    from excel_grapher.grapher.graph import DependencyGraph
    from excel_grapher.series_bindings.types import WorkbookSeriesBindings

_RUNTIME_PATH = Path(__file__).with_name("runtime.py")


def _py_literal(value: object) -> str:
    if value is None:
        return "None"
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


_HOLE_DOC_LABELS = {
    "blank": "blank",
    "off_closure": "not computed",
    "literal": "cached literal",
    "graph_leaf": "cached literal",
    "bound_leaf": "bound leaf",
}


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


_MIN_SHARED_FORMULAS = 2


@dataclass(slots=True)
class _SharedSubplan:
    """A formula prefix shared by outputs with different input closures."""

    name: str
    formula_ids: tuple[str, ...]
    consumers: frozenset[str]
    returns: tuple[str, ...]
    param_ids: tuple[str, ...]
    result_indices: dict[str, tuple[int, ...]]
    call_indices: dict[str, tuple[int, ...]]


def _subplan_signature_params(
    formula_ids: Sequence[str],
    *,
    catalog: SeriesCatalog,
    deps: dict[str, SeriesDeps],
) -> tuple[str, ...]:
    """Return input and external-formula params a shared helper must receive."""
    inside = set(formula_ids)
    needed: set[str] = set()
    for series_id in formula_ids:
        info = deps.get(series_id)
        if info is None:
            continue
        for param_id in info.param_ids:
            if param_id in inside:
                continue
            series = catalog.get(param_id)
            if series.direction == "constant":
                continue
            needed.add(param_id)
    return tuple(sid for sid in catalog.order if sid in needed)


def _subplan_frontier(
    formula_ids: Sequence[str],
    *,
    consumers: frozenset[str],
    closures: Mapping[str, tuple[str, ...]],
    deps: dict[str, SeriesDeps],
) -> tuple[str, ...]:
    """Return series the helper must hand back to private tails."""
    inside = set(formula_ids)
    needed: set[str] = set()
    for output_id in consumers:
        for series_id in closures[output_id]:
            if series_id in inside:
                if series_id == output_id:
                    needed.add(series_id)
                continue
            info = deps.get(series_id)
            if info is None:
                continue
            for param_id in info.param_ids:
                if param_id in inside:
                    needed.add(param_id)
    if not needed and formula_ids:
        needed.add(formula_ids[-1])
    return tuple(sid for sid in formula_ids if sid in needed)


def _plan_shared_subplans(
    catalog: SeriesCatalog,
    deps: dict[str, SeriesDeps],
    scc_map: dict[str, tuple[str, ...]] | None,
) -> tuple[_SharedSubplan, ...]:
    """Extract sufficiently large prefixes shared across distinct input closures.

    Outputs that already share a runner (identical required inputs) are one
    emission site. A helper is emitted only when the same formula series appear
    in two or more input-closure groups, so baseline/shocked splits stay apart
    unless they truly share internals. Helpers may interleave with other
    consumer groups; `_emit_evaluation_body` delays each call until every
    formula `param_id` is bound.
    """
    outputs = list(catalog.output_series())
    closures: dict[str, tuple[str, ...]] = {}
    group_keys: dict[str, tuple[str, ...]] = {}
    for output in outputs:
        closures[output.series_id] = formula_closure(
            output.series_id, catalog=catalog, deps=deps, scc_map=scc_map
        )
        group_keys[output.series_id] = _group_key(output, catalog=catalog, deps=deps)
    consumers: dict[str, set[str]] = {}
    for output_id, formula_ids in closures.items():
        for series_id in formula_ids:
            if series_id == output_id:
                continue
            consumers.setdefault(series_id, set()).add(output_id)
    mapping = scc_map or {}
    dropped: set[str] = set()
    seen_units: set[tuple[str, ...]] = set()
    for unit in mapping.values():
        if len(unit) <= 1 or unit in seen_units:
            continue
        seen_units.add(unit)
        member_consumers = [consumers.get(member, set()) for member in unit]
        if any(group != member_consumers[0] for group in member_consumers):
            dropped.update(unit)
    by_consumers: dict[frozenset[str], list[str]] = {}
    for series_id, group in consumers.items():
        if series_id in dropped or len(group) < 2:
            continue
        keys = {group_keys[output_id] for output_id in group}
        if len(keys) < 2:
            continue
        by_consumers.setdefault(frozenset(group), []).append(series_id)
    subplans: list[_SharedSubplan] = []
    index = 0
    for group in sorted(by_consumers, key=lambda item: (-len(item), tuple(sorted(item)))):
        series_ids = set(by_consumers[group])
        sample = min(group)
        ordered = tuple(sid for sid in closures[sample] if sid in series_ids)
        if len(ordered) < _MIN_SHARED_FORMULAS:
            continue
        members = [catalog.get(output_id) for output_id in sorted(group)]
        result_indices, call_indices = _union_plan_indices(
            members, catalog=catalog, deps=deps, scc_map=scc_map
        )
        subplans.append(
            _SharedSubplan(
                name=f"_shared_{index}",
                formula_ids=ordered,
                consumers=group,
                returns=_subplan_frontier(ordered, consumers=group, closures=closures, deps=deps),
                param_ids=_subplan_signature_params(ordered, catalog=catalog, deps=deps),
                result_indices=result_indices,
                call_indices=call_indices,
            )
        )
        index += 1
    return tuple(subplans)


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
    lines.append(f"    {'data'!r},")
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


def plan_inverted_tree(
    graph: DependencyGraph,
    *,
    series_bindings: WorkbookSeriesBindings,
    bindings_workbook: Path | str,
    blank_ranges: Sequence[str] | None = None,
) -> tuple[SeriesCatalog, dict[str, SeriesDeps], dict[str, tuple[str, ...]]]:
    """Build the catalog, dependency classes, and recurrence groups for emission.

    Raises:
        InvertedTreeExportError: The bindings are invalid, name no output, or
            leave a formula cell of the closure unbound.
    """
    _refuse_invalid_bindings(graph, series_bindings, bindings_workbook)
    blank_rects = normalize_blank_range_specs(blank_ranges)
    catalog = build_catalog(
        series_bindings, workbook=bindings_workbook, graph=graph, blank_ranges=blank_ranges
    )
    if not catalog.output_series():
        raise InvertedTreeExportError("inverted-tree codegen requires at least one output series")
    catalog_edges = collect_catalog_edges(catalog, graph, blank_rects=blank_rects)
    deps = collect_all_deps(catalog, graph, catalog_edges=catalog_edges)
    scc_map = build_scc_map(catalog, deps, edges=catalog_edges.edges)
    # Recurrence groups and nested self recurrences read external producers
    # by coordinate; their producers are never compacted to a host window.
    for sid, info in tuple(deps.items()):
        host = catalog.get(sid)
        if len(scc_map.get(sid, (sid,))) > 1 or (info.is_scan and len(host.key_fields) > 1):
            deps[sid] = replace(info, aligned_ids=frozenset(), index_maps={}, affine_maps={})
    # Scheduling legality: a genuine cell-grain circular reference fails closed
    # at export instead of surfacing as a runtime cycle error.
    checked: set[tuple[str, ...]] = set()
    for series in catalog.formula_series():
        scc = scc_map.get(series.series_id, (series.series_id,))
        if scc in checked:
            continue
        checked.add(scc)
        plan_scc(scc, catalog=catalog, graph=graph, edges=catalog_edges.edges)
    assert_subgraph_bound(
        catalog=catalog,
        graph=graph,
        roots=list(all_formula_root_cells(catalog)),
    )
    return catalog, deps, scc_map


def generate_inverted_tree_modules(
    graph: DependencyGraph,
    *,
    series_bindings: WorkbookSeriesBindings,
    bindings_workbook: Path | str,
    force_rung: Literal[2, 3] | None = None,
    blank_ranges: Sequence[str] | None = None,
) -> dict[str, str]:
    """Generate api/internals/runtime/data modules for inverted-tree export.

    Args:
        graph: Dependency graph covering the binding closure.
        series_bindings: Bindings catalog (inputs, constants, internals, outputs).
        bindings_workbook: Workbook path used to expand `data_range`s.
        force_rung: Pin every formula SCC to rung 3 (demand-driven), or
            fuse wherever legal (`2`) and fall through to the auto rung
            otherwise. `None` selects the strongest legal rung.
        blank_ranges: Sheet-qualified rectangles omitted from the graph that
            resolve as empty (`None`) rather than unbound catalog cells.
            Must match the specs passed to `create_dependency_graph` and
            `FormulaEvaluator`.

    Returns:
        Mapping of package filenames to file contents.

    Raises:
        InvertedTreeExportError: A bound series cannot be inverted fail-closed,
            or an input range overlaps an on-graph formula cell without
            `input.mode: override`.
    """
    blank_rects = normalize_blank_range_specs(blank_ranges)
    token = bind_blank_rects(blank_rects)
    try:
        catalog, deps, scc_map = plan_inverted_tree(
            graph,
            series_bindings=series_bindings,
            bindings_workbook=bindings_workbook,
            blank_ranges=blank_ranges,
        )
        runtime_py = _RUNTIME_PATH.read_text(encoding="utf-8")
        return emit_named_modules(
            catalog,
            deps,
            scc_map,
            graph,
            bindings_workbook,
            init_source=emit_init_module(catalog),
            runtime_source=runtime_py if runtime_py.endswith("\n") else runtime_py + "\n",
        )
    finally:
        reset_blank_rects(token)
