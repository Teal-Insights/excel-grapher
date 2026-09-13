"""Emit an inverted-tree package from a bound workbook graph."""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import replace
from datetime import date, datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any, Literal

from excel_grapher.exporter.inverted_tree.catalog import SeriesCatalog, build_catalog
from excel_grapher.exporter.inverted_tree.deps import (
    SeriesDeps,
    all_formula_root_cells,
    assert_subgraph_bound,
    bind_blank_rects,
    collect_all_deps,
    collect_catalog_edges,
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
        force_rung: Accepted for compatibility and ignored. Every series is
            emitted as named coordinate code; recurrences are demand-driven.
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
