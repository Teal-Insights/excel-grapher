"""Emit an inverted-tree package from a bound workbook graph."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import TYPE_CHECKING

from excel_grapher.exporter.inverted_tree.ast_emit import (
    emit_helper_body,
    python_annotation,
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
)
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError

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


def _cell_value(graph: DependencyGraph, address: str, dtype: str) -> object:
    node = graph.get_node(address)
    value = getattr(node, "value", None) if node is not None else None
    if value is None:
        return 0 if dtype in {"int", "float", "number"} else ""
    if dtype in {"int", "integer"}:
        return int(value)
    if dtype in {"float", "number"}:
        return float(value)
    if dtype in {"string", "str"}:
        return str(value)
    if dtype == "bool":
        return bool(value)
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
        if series.is_scalar:
            lines.append(f"{name}: {series.python_dtype} = {_py_literal(value)}")
        else:
            lines.append(f"{name}: tuple[{series.python_dtype}, ...] = {_py_literal(value)}")
        lines.append("")
    for series in catalog.input_series():
        name = _default_name(series.series_id)
        value = _series_values(series, graph)
        if series.is_scalar:
            lines.append(f"{name}: {series.python_dtype} = {_py_literal(value)}")
        else:
            lines.append(f"{name}: tuple[{series.python_dtype}, ...] = {_py_literal(value)}")
        lines.append("")
    if len(lines) == 4:
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def _helper_signature(series: BoundSeries, deps: SeriesDeps, catalog: SeriesCatalog) -> str:
    params: list[str] = []
    for param_id in deps.param_ids:
        dep = catalog.get(param_id)
        params.append(f"{param_id}: {python_annotation(dep)}")
    joined = ", ".join(params)
    ret = python_return_annotation(series)
    return f"def {series.series_id}({joined}) -> {ret}:"


def emit_internals_module(
    catalog: SeriesCatalog,
    deps: Mapping[str, SeriesDeps],
    graph: DependencyGraph,
) -> str:
    """Emit `internals.py` with one helper per bound formula series."""
    used_runtime: set[str] = set()
    functions: list[str] = []
    for series in catalog.formula_series():
        info = deps[series.series_id]
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


def _horizon_inputs(leaf_ids: Sequence[str], catalog: SeriesCatalog) -> list[str]:
    return [
        sid
        for sid in leaf_ids
        if catalog.get(sid).is_time_series and catalog.get(sid).direction == "input"
    ]


def _time_series_leaves(leaf_ids: Sequence[str], catalog: SeriesCatalog) -> list[str]:
    return [sid for sid in leaf_ids if catalog.get(sid).is_time_series]


def emit_orchestrator(
    output: BoundSeries,
    *,
    catalog: SeriesCatalog,
    deps: Mapping[str, SeriesDeps],
) -> tuple[str, set[str], set[str]]:
    """Emit one public `compute_*` function.

    Returns the function source, runtime symbols used, and data constants imported.
    """
    leaves = leaf_closure(output.series_id, catalog=catalog, deps=dict(deps))
    formula_ids = formula_closure(output.series_id, catalog=catalog, deps=dict(deps))
    required = [sid for sid in leaves if catalog.get(sid).direction == "input"]
    defaulted = [sid for sid in leaves if catalog.get(sid).direction == "constant"]
    compute_name = output.compute_name or f"compute_{output.series_id}"
    params: list[str] = []
    data_imports: set[str] = set()
    for sid in required:
        params.append(f"{sid}: {python_annotation(catalog.get(sid))}")
    for sid in defaulted:
        const_name = _data_const_name(sid)
        data_imports.add(const_name)
        params.append(f"{sid}: {python_annotation(catalog.get(sid))} = {const_name}")
    param_block = ",\n    ".join(params)
    if params:
        signature = f"def {compute_name}(\n    *,\n    {param_block},\n) -> tuple[float, ...]:"
    else:
        signature = f"def {compute_name}() -> tuple[float, ...]:"
    time_inputs = _horizon_inputs(leaves, catalog)
    time_series = _time_series_leaves(leaves, catalog)
    runtime: set[str] = set()
    body: list[str] = []
    out_len = len(output.cells)
    if time_inputs:
        runtime.add("require_aligned")
        runtime.add("trim")
        joined = ", ".join(time_inputs)
        body.append(f"    horizon = min(require_aligned({joined}), {out_len})")
        for sid in time_series:
            body.append(f"    {sid} = trim({sid}, horizon)")
    elif time_series:
        runtime.add("trim")
        body.append(f"    horizon = {out_len}")
        for sid in time_series:
            body.append(f"    {sid} = trim({sid}, horizon)")
    locals_bound: set[str] = set(leaves)
    for series_id in formula_ids:
        info = deps[series_id]
        call = f"internals.{series_id}({', '.join(info.param_ids)})"
        body.append(f"    {series_id} = {call}")
        locals_bound.add(series_id)
    if output.series_id in locals_bound:
        if output.is_scalar:
            body.append(f"    return ({output.series_id},)")
        else:
            body.append(f"    return tuple({output.series_id})")
    else:
        body.append(f"    return internals.{output.series_id}()")
    doc = f'    """Compute `{output.series_id}` from its subgraph leaf closure."""'
    source = "\n".join([signature, doc, *body])
    return source, runtime, data_imports


def emit_api_module(
    catalog: SeriesCatalog,
    deps: Mapping[str, SeriesDeps],
) -> str:
    """Emit `api.py` with keyword-only output orchestrators."""
    functions: list[str] = []
    runtime: set[str] = set()
    data_imports: set[str] = set()
    compute_names: list[str] = []
    for output in catalog.output_series():
        source, used_runtime, used_data = emit_orchestrator(output, catalog=catalog, deps=deps)
        functions.append(source)
        runtime |= used_runtime
        data_imports |= used_data
        compute_names.append(output.compute_name or f"compute_{output.series_id}")
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
    catalog = build_catalog(series_bindings, workbook=bindings_workbook)
    if not catalog.output_series():
        raise InvertedTreeExportError("inverted-tree codegen requires at least one output series")
    deps = collect_all_deps(catalog, graph)
    assert_subgraph_bound(
        catalog=catalog,
        graph=graph,
        roots=list(all_formula_root_cells(catalog)),
    )
    runtime_py = _RUNTIME_PATH.read_text(encoding="utf-8")
    return {
        "__init__.py": emit_init_module(catalog),
        "api.py": emit_api_module(catalog, deps),
        "data.py": emit_data_module(catalog, graph),
        "runtime.py": runtime_py if runtime_py.endswith("\n") else runtime_py + "\n",
        "internals.py": emit_internals_module(catalog, deps, graph),
    }
