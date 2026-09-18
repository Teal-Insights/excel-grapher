"""Emit the public named tensor contract: `data`, `internals`, `validation`, and `api`.

Every formula series becomes one inspectable named function in `internals`
whose body is the workbook formula family expressed over semantic
coordinates. Public `compute_*` functions orchestrate those functions in
dependency order, sharing intermediate results. Input dtype coercion, schema,
domain, and value-map checks live in `validation` so `api` stays the
user-facing surface.
Shared `_CONSTANTS_*` aliases live in `data` and are imported by `api`.
There is no private positional calculation path.
"""

from __future__ import annotations

import ast
import hashlib
import json
from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import TYPE_CHECKING, Any, cast

from excel_grapher.exporter.codegen import REPRESENTATION_VERSION
from excel_grapher.exporter.inverted_tree.ast_emit import (
    EmitContext,
    _as_measure_call,
    _named_keys,
    emit_expr,
    python_measure_type,
)
from excel_grapher.exporter.inverted_tree.catalog import (
    BoundSeries,
    SeriesCatalog,
    _cell_refs_support_series_index,
)
from excel_grapher.exporter.inverted_tree.deps import (
    leaf_closure,
    node_formula_ast,
    try_formula_ast,
)
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.exporter.inverted_tree.named_axes import (
    NamedAxes,
    layout_keys_source,
    python_identifier,
)
from excel_grapher.exporter.inverted_tree.schedule import scan_function_name, scc_external_params
from excel_grapher.series_bindings.input_coerce import (
    input_value_map_from_series,
    measure_domain_from_series,
)
from excel_grapher.series_bindings.resolve import _WorkbookValues

if TYPE_CHECKING:
    from excel_grapher.core.address_keys import CanonicalAddress
    from excel_grapher.exporter.inverted_tree.deps import SeriesDeps
    from excel_grapher.grapher.graph import DependencyGraph

_SKIP_INDEX_CALLS = frozenset(
    {
        "lazy_table",
        "view",
        "span",
        "xl_vlookup",
        "xl_hlookup",
        "xl_lookup",
        "xl_xlookup",
        "xl_match",
        "xl_index",
        "xl_typed_range",
        "xl_lookup_cell",
    }
)
_RESERVED_NAMES = frozenset(
    {
        "data",
        "cast",
        "as_measure",
        "int",
        "str",
        "float",
        "tuple",
        "range",
        "reversed",
        "len",
        "bool",
        "date",
        "datetime",
        "Domain",
        "CoordinateReader",
        "XlError",
        "lazy_table",
        "axis_step",
    }
)


def _generated_helper_imports(used: set[str]) -> list[str]:
    """Split generated helper imports across `.excel` and `.runtime`.

    Names defined independently in both modules fail closed. A name runtime
    only re-exports from excel is imported from `.excel`. Missing names also
    fail closed.
    """
    from excel_grapher.exporter.inverted_tree import excel as inverted_excel
    from excel_grapher.exporter.inverted_tree import runtime as inverted_runtime

    excel_ns = vars(inverted_excel)
    runtime_ns = vars(inverted_runtime)
    excel_names: list[str] = []
    runtime_names: list[str] = []
    for name in sorted(used):
        in_excel = name in excel_ns
        in_runtime = name in runtime_ns
        if in_excel and in_runtime:
            if excel_ns[name] is not runtime_ns[name]:
                raise ValueError(f"{name} is defined independently in excel and runtime")
            excel_names.append(name)
            continue
        if in_excel:
            excel_names.append(name)
            continue
        if in_runtime:
            runtime_names.append(name)
            continue
        raise ValueError(f"{name} is not exported by excel or runtime")
    lines: list[str] = []
    if excel_names:
        lines.append(f"from .excel import {', '.join(excel_names)}")
    if runtime_names:
        lines.append(f"from .runtime import {', '.join(runtime_names)}")
    return lines


def named_codegen_fingerprint(catalog: SeriesCatalog) -> str:
    """Identify the representation, authored provenance, and graph projection."""
    entries = []
    for series in catalog.series.values():
        if series.graph_cells is not None and not series.graph_cells:
            continue
        domain = series.tensor_domain
        required = series.required_coordinates
        entries.append(
            {
                "series_id": series.series_id,
                "direction": series.direction,
                "dtype": series.dtype,
                "axis_labels": series.axis_labels,
                "domain": domain.to_dict(),
                "required": [coord for coord in domain if coord in required],
                "provenance": list(series.coordinate_cells.items()),
            }
        )
    payload = json.dumps(
        {"representation": REPRESENTATION_VERSION, "series": entries},
        sort_keys=True,
        separators=(",", ":"),
    )
    return hashlib.sha256(payload.encode()).hexdigest()


# ---------------------------------------------------------------------------
# Naming and annotations
# ---------------------------------------------------------------------------


def _facade(series: BoundSeries) -> str:
    return "".join(part.capitalize() for part in series.series_id.split("_"))


def _value_annotation(series: BoundSeries) -> str:
    """Describe every permitted workbook value, including blanks and errors."""
    types = python_measure_type(series).split(" | ")
    types.append("str")
    if not series.single_valued:
        types.append("None")
    return " | ".join(dict.fromkeys(types))


def _has_public_type(series: BoundSeries) -> bool:
    """True when `api.py` names this tensor in a public signature."""
    return not series.single_valued and series.direction in {"input", "output"}


def _public_alias(series: BoundSeries) -> str | None:
    """Facade name for public tensors, or `None` when it would shadow the binding."""
    if not _has_public_type(series):
        return None
    facade = _facade(series)
    if facade == series.series_id.upper():
        return None
    return facade


def _annotation(series: BoundSeries) -> str:
    if series.single_valued:
        return _value_annotation(series)
    alias = _public_alias(series)
    if alias is not None:
        return f"data.{alias}"
    return f"data.Series[{_value_annotation(series)}]"


def _binding(series: BoundSeries) -> str:
    return f"data.{series.series_id.upper()}"


def _schema_types(series: BoundSeries) -> str:
    """Name of the shared value-type tuple accepted by the series' schema."""
    return f"{series.python_dtype.upper()}_VALUES"


def _value_types(dtype: str) -> str:
    types = (
        ["int", "float", "bool", "str"]
        if dtype == "float"
        else [dtype, "bool", "str"]
        if dtype == "int"
        else [dtype]
    )
    # Workbook blanks remain observations; the value is not domain membership.
    return "(" + ", ".join(dict.fromkeys([*types, "str", "type(None)"])) + ")"


def _coordinate_names(
    series: BoundSeries, reserved: set[str], *, positional: bool = False
) -> dict[str, str]:
    """Return a loop-variable name for every key field of `series`."""
    names: dict[str, str] = {}
    if positional and len(series.key_fields) == 1:
        candidate = "_p"
        while candidate in reserved:
            candidate += "_"
        names[series.key_fields[0]] = candidate
        reserved.add(candidate)
        return names
    for field in series.key_fields:
        candidate = python_identifier(field.lower())
        if candidate.startswith("xl_"):
            candidate = "axis_" + candidate
        while candidate in reserved:
            candidate += "_key"
        names[field] = candidate
        reserved.add(candidate)
    return names


def _is_runtime_labeller(series: BoundSeries) -> bool:
    return series.axis_labels is not None and series.direction != "constant"


def _axis_position_name(axis: Any) -> str:
    return "_p_" + python_identifier(axis.name.lower())


def _axis_var_name(axis: Any) -> str:
    return "_ax_" + python_identifier(axis.name.lower())


def _index_names_for(
    series: BoundSeries, catalog: SeriesCatalog, names: Mapping[str, str]
) -> dict[str, str]:
    """Map runtime axis names to the position variable used in family tests."""
    if _is_runtime_labeller(series):
        return dict(names)
    return {
        axis.name: _axis_position_name(axis)
        for axis in series.tensor_domain.axes
        if catalog.runtime_labeller(axis.name, axis.keys) is not None
    }


def _runtime_axis_prologue(catalog: SeriesCatalog, source: str) -> list[str]:
    """Bind `_ax_*` labeller axes referenced by a scalar formula body."""
    lines: list[str] = []
    seen: set[str] = set()
    for series in catalog.series.values():
        if not _is_runtime_labeller(series) or not series.tensor_domain.axes:
            continue
        axis = series.tensor_domain.axes[0]
        variable = _axis_var_name(axis)
        if variable not in source or series.series_id in seen:
            continue
        seen.add(series.series_id)
        lines.append(f"    {variable} = {series.series_id}.domain.axes[0]")
    return lines


def _bind_kwargs(series: BoundSeries, catalog: SeriesCatalog) -> str:
    parts = [
        f"{axis.name}={labeller.series_id}"
        for axis in series.tensor_domain.axes
        if (labeller := catalog.runtime_labeller(axis.name, axis.keys)) is not None
        and labeller.series_id != series.series_id
    ]
    return ", ".join(parts)


def _required_expr(series: BoundSeries, catalog: SeriesCatalog) -> str:
    if _is_runtime_labeller(series):
        return f"{_binding(series)}_POSITIONS"
    kwargs = _bind_kwargs(series, catalog)
    if kwargs:
        return f"{_binding(series)}.required.bind({kwargs})"
    return f"{_binding(series)}.required"


def _layout_axis(axis: Any, catalog: SeriesCatalog) -> Any:
    """Return the labeller (or interned-master analog) used for position indexes."""
    labeller = catalog.runtime_labeller(axis.name, axis.keys)
    if labeller is None:
        return axis
    return labeller.tensor_domain.axes[0]


def _labelled_axes_map(catalog: SeriesCatalog) -> dict[str, str]:
    """Map each runtime axis name to its labeller series id."""
    mapping: dict[str, str] = {}
    for series in _retained(catalog):
        if not _is_runtime_labeller(series) or series.axis_labels is None:
            continue
        name = series.axis_labels
        existing = mapping.get(name)
        if existing is not None and existing != series.series_id:
            raise InvertedTreeExportError(
                f"axis {name!r}: multiple runtime labellers {existing!r} and {series.series_id!r}"
            )
        mapping[name] = series.series_id
    return mapping


def _deferred_runtime_inputs(catalog: SeriesCatalog) -> frozenset[str]:
    """Inputs keyed on a runtime axis they do not themselves label."""
    return frozenset(
        series.series_id
        for series in _retained(catalog)
        if series.direction == "input"
        and any(
            (labeller := catalog.runtime_labeller(axis.name, axis.keys)) is not None
            and labeller.series_id != series.series_id
            for axis in series.tensor_domain.axes
        )
    )


def _result_type(scc: Sequence[str]) -> str:
    return (
        "".join(part.capitalize() for part in scan_function_name(tuple(scc)).split("_")) + "Result"
    )


# ---------------------------------------------------------------------------
# Formula bodies
# ---------------------------------------------------------------------------


def _outermost_tables(node: ast.AST, local_names: set[str]) -> list[ast.Call]:
    """Tables and views that can be built once per call, outermost first.

    A table that reads the loop coordinate or a recurrence member is built
    where it is used; views nested inside a hoisted table go with it.
    """
    if (
        isinstance(node, ast.Call)
        and isinstance(node.func, ast.Name)
        and node.func.id in {"lazy_table", "view", "xl_typed_range"}
        and not any(
            isinstance(child, ast.Name) and child.id in local_names for child in ast.walk(node)
        )
    ):
        return [node]
    found: list[ast.Call] = []
    for child in ast.iter_child_nodes(node):
        found.extend(_outermost_tables(child, local_names))
    return found


def _hole_expression(
    series: BoundSeries, index: int, ctx: EmitContext, graph: DependencyGraph
) -> str:
    """Return the named expression of a retained non-formula member."""
    from excel_grapher.exporter.inverted_tree.emit import (
        _cell_value,
        _coerce_cached_value,
        _py_literal,
    )

    hole = series.hole_at(index)
    address = series.cells[index]
    if hole is not None and hole.kind in {"blank", "off_closure"}:
        return "None"
    if hole is not None and hole.kind == "bound_leaf":
        if hole.claimant_id is None:
            raise InvertedTreeExportError(
                f"series {series.series_id!r} cell {address}: bound leaf has no claimant"
            )
        claimant = ctx.catalog.get(hole.claimant_id)
        if claimant.is_scalar:
            return claimant.series_id
        claimant_index = claimant.index_of(address)
        if claimant_index is None:
            raise InvertedTreeExportError(
                f"series {series.series_id!r} cell {address}: "
                f"bound leaf is not in claimant {claimant.series_id!r}"
            )
        keys = _named_keys(claimant, claimant.domain[claimant_index], ctx)
        return f"{claimant.series_id}[{', '.join(keys)}]"
    if hole is not None and hole.kind == "graph_leaf":
        node = graph.get_node(address)
        if node is None or node.value is None:
            raise InvertedTreeExportError(
                f"series {series.series_id!r} cell {address}: graph leaf has no cached value"
            )
        return _as_measure_call(_py_literal(_cell_value(graph, address, series.dtype)), series)
    if hole is not None and hole.literal is not None:
        literal = _py_literal(_coerce_cached_value(hole.literal, series.dtype, address))
        return _as_measure_call(literal, series)
    node = graph.get_node(address)
    if node is None or node.value is None:
        return "None"
    return _as_measure_call(_py_literal(node.value), series)


def _statement_member_groups(
    series: BoundSeries,
    cells: Sequence[CanonicalAddress],
    graph: DependencyGraph,
) -> list[list[int]]:
    """On-graph catalog indices that share one statement and one lowered body.

    Off-graph members stay omitted. Uniform statements sample first, interior,
    and last members and replicate when those expressions match.
    """
    if series.single_valued:
        if not cells or graph.get_node(cells[0]) is None:
            return []
        return [[0]]
    statements = series.statements
    if not statements:
        return [[index] for index, cell in enumerate(cells) if graph.get_node(cell) is not None]
    groups: list[list[int]] = []
    for stmt in statements:
        members = [
            index
            for index in range(stmt.start, min(stmt.stop, len(cells)))
            if graph.get_node(cells[index]) is not None
        ]
        if members:
            groups.append(members)
    return groups


def _sample_member_indices(members: Sequence[int]) -> list[int]:
    """First, interior, and last indices used to detect a uniform family."""
    if len(members) <= 3:
        return list(members)
    middle = members[len(members) // 2]
    return list(dict.fromkeys((members[0], middle, members[-1])))


def _direct_series_subscripts(expression: str, catalog: SeriesCatalog) -> tuple[ast.Subscript, ...]:
    """Catalog-series indexes that are not lookup or table arguments."""
    try:
        parsed = ast.parse(expression, mode="eval")
    except SyntaxError:
        return ()
    skip: set[int] = set()
    for node in ast.walk(parsed):
        if (
            isinstance(node, ast.Call)
            and isinstance(node.func, ast.Name)
            and node.func.id in _SKIP_INDEX_CALLS
        ):
            for child in ast.walk(node):
                if child is not node:
                    skip.add(id(child))
    found: list[ast.Subscript] = []
    for node in ast.walk(parsed):
        if id(node) in skip or not isinstance(node, ast.Subscript):
            continue
        if isinstance(node.value, ast.Name) and node.value.id in catalog.series:
            found.append(node)
    return tuple(found)


def _eval_key_expr(
    node: ast.AST,
    env: Mapping[str, object],
    remaps: Mapping[str, Mapping[object, object]],
) -> object | None:
    """Evaluate a subscript key to a constant, or `None` if it is not static."""
    if isinstance(node, ast.Constant):
        return node.value
    if isinstance(node, ast.Name):
        if node.id in env:
            return env[node.id]
        mapping = remaps.get(node.id)
        return mapping
    if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.UAdd, ast.USub)):
        value = _eval_key_expr(node.operand, env, remaps)
        if type(value) is not int:
            return None
        return value if isinstance(node.op, ast.UAdd) else -value
    if isinstance(node, ast.BinOp) and isinstance(node.op, (ast.Add, ast.Sub)):
        left = _eval_key_expr(node.left, env, remaps)
        right = _eval_key_expr(node.right, env, remaps)
        if type(left) is not int or type(right) is not int:
            return None
        return left + right if isinstance(node.op, ast.Add) else left - right
    if isinstance(node, ast.Subscript):
        mapping = _eval_key_expr(node.value, env, remaps)
        key = _eval_key_expr(node.slice, env, remaps)
        if isinstance(mapping, Mapping) and key is not None:
            try:
                return cast(Mapping[Any, object], mapping)[key]
            except KeyError:
                return None
        return None
    return None


def _index_image_on_producer_axes(
    subscript: ast.Subscript,
    catalog: SeriesCatalog,
    env: Mapping[str, object],
    remaps: Mapping[str, Mapping[object, object]],
) -> bool | None:
    """True when evaluated keys sit on the producer axes; `None` if unknown."""
    if not isinstance(subscript.value, ast.Name):
        return None
    owner = catalog.series.get(subscript.value.id)
    if owner is None:
        return None
    slice_node = subscript.slice
    key_nodes = slice_node.elts if isinstance(slice_node, ast.Tuple) else (slice_node,)
    axes = owner.tensor_domain.axes
    if len(key_nodes) != len(axes):
        return None
    for axis, key_node in zip(axes, key_nodes, strict=True):
        key = _eval_key_expr(key_node, env, remaps)
        if key is None:
            return None
        if key not in axis:
            return False
    return True


def _sampled_index_supported(
    subscripts: Sequence[ast.Subscript],
    *,
    catalog: SeriesCatalog,
    graph: DependencyGraph,
    series: BoundSeries,
    index: int,
    cell: CanonicalAddress,
    names: Mapping[str, str],
    remaps: Mapping[str, Mapping[object, object]],
) -> bool:
    """True when `index` may reuse a sampled series-index expression."""
    env = {names[field]: series.domain[index][field] for field in series.key_fields}
    unevaluable = False
    for subscript in subscripts:
        image = _index_image_on_producer_axes(subscript, catalog, env, remaps)
        if image is None:
            unevaluable = True
            break
        if not image:
            return False
    if unevaluable:
        return _cell_refs_support_series_index(catalog, graph, cell)
    return True


def _semantic_body(
    series: BoundSeries,
    catalog: SeriesCatalog,
    deps: SeriesDeps,
    graph: DependencyGraph,
    named_axes: NamedAxes,
    *,
    deferred: bool = False,
    scc_ids: frozenset[str] = frozenset(),
    literal_tables: dict[str, dict[tuple[object, ...], object]] | None = None,
) -> tuple[list[str], set[str]]:
    """Emit the formula family of `series` as semantic coordinate code.

    Returns the indented body lines and the runtime symbols they use. When
    `deferred` is set the body binds a demand-driven `CoordinateReader`
    instead of returning, so recurrence groups can share one evaluation.
    Uniform statements lower first/interior/last samples and, when those
    expressions match, replicate the grouped body; mixed statements still
    emit one body per distinct expression. A sampled series index is not
    copied onto a member whose affine image is missing from the producer
    axis or whose producer cell is unbound.
    """
    reserved = set(catalog.series) | set(_RESERVED_NAMES)
    names = _coordinate_names(series, reserved, positional=_is_runtime_labeller(series))
    index_names = _index_names_for(series, catalog, names)
    layout_axes = tuple(_layout_axis(axis, catalog) for axis in series.tensor_domain.axes)
    groups: dict[str, list[tuple[object, ...]]] = {}
    tables: dict[str, str] = {}
    used: set[str] = {"XlError"}
    recursive = deferred
    key_remaps: dict[str, dict[object, object]] = {}
    # A scalar layout publishes one observation: its first bound cell.
    cells = series.cells[:1] if series.single_valued else series.cells

    def lower_member(index: int, cell: CanonicalAddress) -> str:
        ctx = EmitContext(
            host=series,
            catalog=catalog,
            deps=deps,
            host_index=index,
            host_cell=cell,
            coordinate_vars={} if series.single_valued else names,
            scc_ids=scc_ids | {series.series_id},
            graph=graph,
            named_axes=named_axes,
            key_remaps=key_remaps,
        )
        node = try_formula_ast(graph, cell)
        if node is None:
            if series.hole_at(index) is None and getattr(
                graph.get_node(cell), "has_formula", False
            ):
                node_formula_ast(graph, cell)
            expression = _hole_expression(series, index, ctx, graph)
        else:
            expression = _as_measure_call(emit_expr(node, ctx), series)
        used.add("as_measure")
        used.update(ctx.used_runtime)
        return expression

    def alias_expression(expression: str) -> str:
        parsed = ast.parse(expression, mode="eval")
        if any(
            isinstance(item, ast.Name) and item.id == series.series_id for item in ast.walk(parsed)
        ):
            nonlocal recursive
            recursive = True
        if "lazy_table(" not in expression and "view(" not in expression:
            return expression
        local_names = set(names.values()) | scc_ids | {series.series_id}
        for item in _outermost_tables(parsed, local_names):
            source = ast.get_source_segment(expression, item)
            if source is None:
                continue
            if source not in tables:
                alias = f"_{series.series_id}_table_{len(tables)}"
                while alias in reserved:
                    alias += "_"
                tables[source] = alias
                reserved.add(alias)
        for source, alias in tables.items():
            expression = expression.replace(source, alias)
        return expression

    def record(index: int, expression: str) -> None:
        coord = tuple(series.domain[index][field] for field in series.key_fields)
        groups.setdefault(expression, []).append(coord)

    for members in _statement_member_groups(series, cells, graph):
        sampled = _sample_member_indices(members)
        sample_exprs = {index: lower_member(index, cells[index]) for index in sampled}
        unique = set(sample_exprs.values())
        if len(unique) == 1:
            expression = alias_expression(next(iter(unique)))
            subscripts = _direct_series_subscripts(expression, catalog)
            if subscripts:
                none_expression = _as_measure_call("None", series)
                for index in members:
                    if _sampled_index_supported(
                        subscripts,
                        catalog=catalog,
                        graph=graph,
                        series=series,
                        index=index,
                        cell=cells[index],
                        names=names,
                        remaps=key_remaps,
                    ):
                        record(index, expression)
                    else:
                        record(index, none_expression)
                continue
            for index in members:
                record(index, expression)
            continue
        for index in members:
            expression = sample_exprs.get(index)
            if expression is None:
                expression = lower_member(index, cells[index])
            record(index, alias_expression(expression))
    if literal_tables is not None and not series.single_valued:
        literals = {}
        literal_groups = []
        for expression, coordinates in groups.items():
            parsed = ast.parse(expression, mode="eval").body
            if not (
                isinstance(parsed, ast.Call)
                and isinstance(parsed.func, ast.Name)
                and parsed.func.id == "as_measure"
                and len(parsed.args) == 1
            ):
                continue
            try:
                literal = ast.literal_eval(parsed.args[0])
            except (ValueError, TypeError):
                continue
            literals.update((coord, literal) for coord in coordinates)
            literal_groups.append(expression)
        if len(literal_groups) > 1:
            for expression in literal_groups:
                del groups[expression]
            table = f"_{series.series_id.upper()}_LITERALS"
            if index_names:
                literals = {
                    tuple(
                        axis.keys.index(key) if axis.name in index_names else key
                        for axis, key in zip(layout_axes, coord, strict=True)
                    ): value
                    for coord, value in literals.items()
                }
            literal_tables[table] = literals
            selectors = ", ".join(
                index_names.get(axis.name, names[axis.name]) for axis in series.tensor_domain.axes
            )
            groups[f"as_measure(data.{table}[{selectors},])"] = list(literals)
    table_lines = [f"    {alias} = {source}" for source, alias in tables.items()]
    if series.single_valued:
        if len(groups) != 1:
            raise InvertedTreeExportError(
                f"series {series.series_id!r}: scalar series has no graph formula"
            )
        expression = next(iter(groups))
        scanned = "\n".join((*tables, expression))
        lines = [*_runtime_axis_prologue(catalog, scanned), *table_lines]
        if not deferred:
            lines.extend(
                [
                    "    try:",
                    f"        return {expression}",
                    "    except XlError as error:",
                    "        return error.code",
                ]
            )
            return lines, used
        used.add("CoordinateReader")
        formula = f"{series.series_id}_formula"
        lines.extend(
            [
                f"    def {formula}() -> {_value_annotation(series)}:",
                f"        return {expression}",
                "",
                f"    {series.series_id} = CoordinateReader({series.series_id!r}, Domain.product(), {formula})",
            ]
        )
        return lines, used
    lines = table_lines
    occupied = reserved
    required = _required_expr(series, catalog)

    def temporary(base: str) -> str:
        candidate = base
        while candidate in occupied:
            candidate += "_"
        occupied.add(candidate)
        return candidate

    formula = temporary(f"{series.series_id}_formula" if deferred else "formula")
    parameters = ", ".join(
        f"{names[axis.name]}: {axis.key_type.__name__}"
        if not _is_runtime_labeller(series)
        else f"{names[axis.name]}: int"
        for axis in series.tensor_domain.axes
    )
    prologue: list[str] = []
    if not _is_runtime_labeller(series):
        seen_labellers: set[str] = set()
        for axis in series.tensor_domain.axes:
            labeller = catalog.runtime_labeller(axis.name, axis.keys)
            if labeller is None or labeller.series_id in seen_labellers:
                continue
            seen_labellers.add(labeller.series_id)
            prologue.append(f"    {_axis_var_name(axis)} = {labeller.series_id}.domain.axes[0]")
    lines.extend(prologue)
    lines.append(f"    def {formula}({parameters}) -> {_value_annotation(series)}:")
    for axis in series.tensor_domain.axes:
        if axis.name in index_names and not _is_runtime_labeller(series):
            lines.append(
                f"        {index_names[axis.name]} = {_axis_var_name(axis)}.position({names[axis.name]})"
            )
    for remap_name, mapping in key_remaps.items():
        lines.append(f"        {remap_name} = {mapping!r}")
    all_coordinates = {coord for coordinates in groups.values() for coord in coordinates}

    def condition(coordinates: list[tuple[object, ...]]) -> str:
        synthesized = _family_condition(
            coordinates,
            all_coordinates,
            layout_axes,
            names,
            index_names=index_names,
        )
        if synthesized is not None:
            return synthesized
        selectors = ", ".join(names[axis.name] for axis in series.tensor_domain.axes)
        if index_names:
            selectors = ", ".join(
                index_names.get(axis.name, names[axis.name]) for axis in series.tensor_domain.axes
            )
            indexed = tuple(
                tuple(
                    axis.keys.index(key) if axis.name in index_names else key
                    for axis, key in zip(layout_axes, coord, strict=True)
                )
                for coord in coordinates
            )
            return f"({selectors},) in {indexed!r}"
        return f"({selectors},) in {tuple(coordinates)!r}"

    # Put the largest family last as the unconditional return, avoiding an
    # explicit coordinate list for ordinary forward or backward recurrence periods.
    families = sorted(groups.items(), key=lambda item: len(item[1]))
    for i, (expression, coordinates) in enumerate(families):
        if i < len(families) - 1:
            lines.append(f"        {'if' if i == 0 else 'elif'} {condition(coordinates)}:")
            lines.append(f"            return {expression}")
        else:
            lines.append(f"        return {expression}")
    lines.append("")
    if recursive:
        used.add("CoordinateReader")
        lines.append(
            f"    {series.series_id} = CoordinateReader({series.series_id!r}, {required}, {formula})"
        )
        if not deferred and deps.is_scan and deps.scan_direction == "reversed":
            coordinate = temporary("coordinate")
            lines.extend(
                [
                    f"    for {coordinate} in reversed(tuple({required})):",
                    f"        {series.series_id}[{coordinate}]",
                ]
            )
        if not deferred:
            if _is_runtime_labeller(series):
                used.add("label_axis")
            lines.append(f"    return {_materialize(series, catalog, required)}")
    elif _is_runtime_labeller(series):
        used.add("label_axis")
        used.add("CoordinateReader")
        axis = series.tensor_domain.axes[0]
        lines.extend(
            [
                f"    {series.series_id} = CoordinateReader({series.series_id!r}, {required}, {formula})",
                "    return Series.from_labels(",
                f"        {_binding(series)},",
                f"        label_axis({axis.name!r}, tuple({series.series_id}[p] for p in {required}), {axis.key_type.__name__})",
                "    )",
            ]
        )
    else:
        used.add("evaluate")
        collect = f"{_binding(series)}.collect(evaluate({formula}, {required})"
        if required != f"{_binding(series)}.required":
            collect += f", domain={required}"
        lines.append(f"    return {collect})")
    return lines, used


def _materialize(series: BoundSeries, catalog: SeriesCatalog, required: str | None = None) -> str:
    """Publish a completed demand-driven reader as an immutable tensor."""
    if required is None:
        required = _required_expr(series, catalog)
    if _is_runtime_labeller(series):
        axis = series.tensor_domain.axes[0]
        return (
            f"Series.from_labels({_binding(series)}, "
            f"label_axis({axis.name!r}, tuple({series.series_id}[p] for p in {required}), "
            f"{axis.key_type.__name__}))"
        )
    records = f"(coord, {series.series_id}[coord]) for coord in {required}"
    if required != f"{_binding(series)}.required":
        return f"{_binding(series)}.collect(({records}), domain={required})"
    return f"{_binding(series)}.collect({records})"


# ---------------------------------------------------------------------------
# internals.py
# ---------------------------------------------------------------------------


def _signature(name: str, params: Sequence[BoundSeries], returns: str) -> str:
    joined = ", ".join(f"{series.series_id}: {_annotation(series)}" for series in params)
    return f"def {name}({('*, ' + joined) if joined else ''}) -> {returns}:"


def _schema_checks(params: Sequence[BoundSeries], catalog: SeriesCatalog) -> list[str]:
    """Validate tensor parameters that arrive from outside the generated model.

    Results of other named functions are `Series` instances validated on
    construction, so only inputs and constants are checked again here.
    """
    lines: list[str] = []
    for series in params:
        if series.single_valued or series.direction not in {"input", "constant"}:
            continue
        kwargs = _bind_kwargs(series, catalog)
        schema = f"{_binding(series)}.schema"
        if kwargs:
            lines.append(f"    {schema}.bind({kwargs}).validate({series.series_id})")
        else:
            lines.append(f"    {schema}.validate({series.series_id})")
    return lines


def _publish_line(series: BoundSeries, constants: str | None = None) -> str:
    name = series.series_id.upper()
    arguments = (
        [f"key={series.key_fields!r}", "domain=None"]
        if series.single_valued
        else [f"{_binding(series)}.schema"]
    )
    if constants is not None:
        arguments.append(f"constants={constants}")
    if series.single_valued:
        arguments.append(f"cells=data.{name}_CELLS")
    else:
        arguments.append(f"cells={_binding(series)}.cells")
    return f"@publish({', '.join(arguments)})"


def emit_named_internals(
    catalog: SeriesCatalog,
    deps: Mapping[str, SeriesDeps],
    scc_map: Mapping[str, tuple[str, ...]],
    graph: DependencyGraph,
    named_axes: NamedAxes,
    literal_tables: dict[str, dict[tuple[object, ...], object]],
) -> str:
    """Emit one named calculation function per formula series or recurrence group."""
    functions: list[str] = []
    used: set[str] = {"publish"}
    emitted_groups: set[tuple[str, ...]] = set()
    for series in _retained_formula_series(catalog):
        scc = scc_map.get(series.series_id, (series.series_id,))
        if len(scc) > 1:
            if scc in emitted_groups:
                continue
            emitted_groups.add(scc)
            source, group_used = _emit_recurrence_group(
                scc, catalog, deps, graph, named_axes, literal_tables
            )
            functions.append(source)
            used |= group_used
            continue
        info = deps[series.series_id]
        params = [catalog.get(sid) for sid in info.param_ids]
        body, body_used = _semantic_body(
            series, catalog, info, graph, named_axes, literal_tables=literal_tables
        )
        used |= body_used
        functions.append(
            "\n".join(
                [
                    _publish_line(series),
                    _signature(series.series_id, params, _annotation(series)),
                    f'    """Compute `{series.series_id}` using authored coordinate identities."""',
                    *_schema_checks(params, catalog),
                    *body,
                ]
            )
        )
    tensor_names = ["Domain", "Series"]
    if "label_axis" in used:
        tensor_names.append("label_axis")
        used.remove("label_axis")
    lines = [
        '"""Named calculation functions for every bound formula series."""',
        "from __future__ import annotations",
        "from dataclasses import dataclass",
        "from datetime import datetime",
        "from typing import cast",
        "from . import data",
        f"from .tensor import {', '.join(tensor_names)}",
        *_generated_helper_imports(used),
        "",
        "\n\n".join(functions),
        "",
    ]
    return "\n".join(lines)


def _family_condition(
    coordinates: Sequence[tuple[object, ...]],
    universe: set[tuple[object, ...]],
    axes: Sequence[Any],
    names: Mapping[str, str],
    index_names: Mapping[str, str] | None = None,
) -> str | None:
    """Describe a formula family by its axis relations rather than a coordinate list.

    A family that is a product of per-axis selections becomes a conjunction of
    equality, run, or membership tests; a family on one diagonal of two
    integer axes adds the difference between them. Irregular families keep
    their explicit coordinate list.
    """
    members = set(coordinates)
    tests: list[str] = []
    selected: list[set[object]] = []
    positions = index_names or {}
    for index, axis in enumerate(axes):
        present = [key for key in axis.keys if any(coord[index] == key for coord in universe)]
        keys = [key for key in present if any(coord[index] == key for coord in members)]
        selected.append(set(keys))
        name = positions.get(axis.name, names[axis.name])
        emitted = [axis.keys.index(key) for key in keys] if axis.name in positions else list(keys)
        if len(keys) == len(present):
            continue
        if len(keys) == 1:
            tests.append(f"{name} == {emitted[0]!r}")
            continue
        first, last = present.index(keys[0]), present.index(keys[-1])
        contiguous = present[first : last + 1] == keys
        ascending = axis.name in positions or (
            axis.key_type is int and present == sorted(cast(Sequence[int], present))
        )
        if contiguous and ascending:
            if last == len(present) - 1:
                tests.append(f"{name} >= {emitted[0]!r}")
            elif first == 0:
                tests.append(f"{name} <= {emitted[-1]!r}")
            else:
                tests.append(f"{emitted[0]!r} <= {name} <= {emitted[-1]!r}")
            continue
        tests.append(f"{name} in {tuple(emitted)!r}")
    product = {
        coord
        for coord in universe
        if all(coord[index] in keys for index, keys in enumerate(selected))
    }
    if product == members:
        return " and ".join(tests) if tests else None
    relation = _diagonal_condition(
        members, universe, product, tests, axes, names, index_names=positions
    )
    if relation is not None:
        return relation
    return _union_condition(members, universe, axes, names, index_names=positions)


def _diagonal_condition(
    members: set[tuple[object, ...]],
    universe: set[tuple[object, ...]],
    product: set[tuple[object, ...]],
    tests: Sequence[str],
    axes: Sequence[Any],
    names: Mapping[str, str],
    index_names: Mapping[str, str] | None = None,
) -> str | None:
    """A family on one diagonal band of two integer axes."""
    positions = index_names or {}
    integer_axes = [
        index for index, axis in enumerate(axes) if axis.key_type is int or axis.name in positions
    ]
    for position, left in enumerate(integer_axes):
        for right in integer_axes[position + 1 :]:
            use_positions = axes[left].name in positions or axes[right].name in positions

            def difference(
                coord: tuple[object, ...],
                left: int = left,
                right: int = right,
                use_positions: bool = use_positions,
            ) -> int:
                if use_positions:
                    return axes[right].keys.index(coord[right]) - axes[left].keys.index(coord[left])
                return cast(int, coord[right]) - cast(int, coord[left])

            differences = sorted({difference(coord) for coord in members})
            low, high = differences[0], differences[-1]
            if differences != list(range(low, high + 1)):
                continue
            left_name = positions.get(axes[left].name, names[axes[left].name])
            right_name = positions.get(axes[right].name, names[axes[right].name])
            for candidates, prefix in ((universe, []), (product, tests)):
                if {coord for coord in candidates if low <= difference(coord) <= high} != members:
                    continue
                bounds = {difference(coord) for coord in candidates}
                if low == high == 0:
                    relation = f"{left_name} == {right_name}"
                elif low == high:
                    relation = f"{right_name} - {left_name} == {low}"
                elif high == max(bounds):
                    relation = f"{right_name} - {left_name} >= {low}"
                elif low == min(bounds):
                    relation = f"{right_name} - {left_name} <= {high}"
                else:
                    relation = f"{low} <= {right_name} - {left_name} <= {high}"
                return " and ".join([*prefix, relation])
    return None


def _union_condition(
    members: set[tuple[object, ...]],
    universe: set[tuple[object, ...]],
    axes: Sequence[Any],
    names: Mapping[str, str],
    index_names: Mapping[str, str] | None = None,
) -> str | None:
    """A family that splits along one axis into a few expressible sub-families."""
    positions = index_names or {}
    for index, axis in enumerate(axes):
        keys = [key for key in axis.keys if any(coord[index] == key for coord in members)]
        if not 2 <= len(keys) <= 8:
            continue
        parts: list[str] = []
        name = positions.get(axis.name, names[axis.name])
        for key in keys:
            group = [coord for coord in members if coord[index] == key]
            restricted = {coord for coord in universe if coord[index] == key}
            emitted = axis.keys.index(key) if axis.name in positions else key
            selector = f"{name} == {emitted!r}"
            if set(group) == restricted:
                parts.append(selector)
                continue
            inner = _family_condition(group, restricted, axes, names, index_names=positions)
            if inner is None:
                break
            parts.append(f"({selector} and {inner})")
        else:
            return " or ".join(parts)
    return None


def _emit_recurrence_group(
    scc: tuple[str, ...],
    catalog: SeriesCatalog,
    deps: Mapping[str, SeriesDeps],
    graph: DependencyGraph,
    named_axes: NamedAxes,
    literal_tables: dict[str, dict[tuple[object, ...], object]],
) -> tuple[str, set[str]]:
    """Emit a mutually recursive series group as demand-driven readers."""
    params = [catalog.get(sid) for sid in scc_external_params(scc, deps, catalog.order)]
    result_type = _result_type(scc)
    used: set[str] = set()
    lines = [
        "@dataclass(frozen=True, slots=True)",
        f"class {result_type}:",
        '    """Complete named results of one recurrence group evaluation."""',
    ]
    for sid in scc:
        lines.append(f"    {sid}: {_annotation(catalog.get(sid))}")
    joined = ", ".join(f"`{sid}`" for sid in scc)
    lines.extend(
        [
            "",
            _signature(scan_function_name(scc), params, result_type),
            f'    """Evaluate the recurrence group {joined} and publish complete tensors."""',
            *_schema_checks(params, catalog),
        ]
    )
    members = frozenset(scc)
    for sid in scc:
        body, body_used = _semantic_body(
            catalog.get(sid),
            catalog,
            deps[sid],
            graph,
            named_axes,
            deferred=True,
            scc_ids=members,
            literal_tables=literal_tables,
        )
        used |= body_used
        lines.extend(body)
    lines.append(f"    return {result_type}(")
    for sid in scc:
        series = catalog.get(sid)
        value = f"{sid}[()]" if series.single_valued else _materialize(series, catalog)
        lines.append(f"        {sid}={value},")
    lines.append("    )")
    return "\n".join(lines), used


# ---------------------------------------------------------------------------
# api.py
# ---------------------------------------------------------------------------


def _argument_source(series_id: str, catalog: SeriesCatalog) -> str:
    series = catalog.get(series_id)
    if series.direction == "constant":
        return f"data.{series_id.upper()}"
    return f"self.{series_id}"


# ---------------------------------------------------------------------------
# validation.py
# ---------------------------------------------------------------------------


def _python_literal(value: object) -> str:
    """Render a Python literal using double quotes."""
    if isinstance(value, str):
        return json.dumps(value)
    if isinstance(value, frozenset):
        inner = ", ".join(_python_literal(item) for item in sorted(value, key=repr))
        return f"frozenset({{{inner}}})"
    if isinstance(value, dict):
        items = ", ".join(
            f"{_python_literal(key)}: {_python_literal(item)}" for key, item in value.items()
        )
        return "{" + items + "}"
    if isinstance(value, tuple):
        inner = ", ".join(_python_literal(item) for item in value)
        return f"({inner},)" if len(value) == 1 else f"({inner})"
    if isinstance(value, list):
        return "[" + ", ".join(_python_literal(item) for item in value) + "]"
    return repr(value)


def _check_signature(series_id: str, annotation: str, extra: Sequence[str] = ()) -> str:
    """Write a check signature, wrapping only when the one-liner exceeds 100 columns."""
    name = f"_check_{series_id}"
    extras = ", ".join(extra)
    arguments = f"{series_id}: {annotation}" + (f", *, {extras}" if extras else "")
    one_line = f"def {name}({arguments}) -> {annotation}:"
    if len(one_line) <= 100:
        return one_line
    if extras:
        return f"def {name}(\n    {series_id}: {annotation},\n    *,\n    {extras},\n) -> {annotation}:"
    return f"def {name}(\n    {series_id}: {annotation},\n) -> {annotation}:"


def _binding_dtype(series: BoundSeries) -> str:
    """Return the setter dtype used to coerce a public compute argument."""
    raw = series.dtype
    return {"str": "string", "integer": "int", "date": "datetime"}.get(raw, raw)


def _input_check(
    series: BoundSeries, catalog: SeriesCatalog
) -> tuple[list[str], set[str], list[str]]:
    """Coerce dtype, validate schema and domain, then apply the value map."""
    lines: list[str] = []
    used: set[str] = set()
    series_id = series.series_id
    quoted_id = _python_literal(series_id)
    runtime_axes = [
        (axis, labeller)
        for axis in series.tensor_domain.axes
        if (labeller := catalog.runtime_labeller(axis.name, axis.keys)) is not None
        and labeller.series_id != series.series_id
    ]
    bind = ", ".join(f"{axis.name}={axis.name}" for axis, _labeller in runtime_axes)
    extras = [f"{axis.name}: {_annotation(labeller)}" for axis, labeller in runtime_axes]
    if not series.single_valued:
        schema = f"{_binding(series)}.schema"
        if bind:
            lines.append(f"    {schema}.bind({bind}).validate({series_id})")
        else:
            lines.append(f"    {schema}.validate({series_id})")
    used.add("coerce_input_measure")
    lines.append(
        f"    {series_id} = coerce_input_measure("
        f"{series_id}, dtype={_python_literal(_binding_dtype(series))}, series_id={quoted_id})"
    )
    domain = measure_domain_from_series(series.raw)
    if domain is not None:
        used.add("require_input_domain")
        domain_literal = _python_literal(domain)
        if series.single_valued:
            lines.append(
                f"    require_input_domain({series_id}, {domain_literal}, series_id={quoted_id})"
            )
        else:
            # Restrict validation to this extraction's required domain;
            # wider source tensors retain their off-graph observations.
            required = (
                f"{_binding(series)}.required.bind({bind})"
                if bind
                else f"{_binding(series)}.required"
            )
            lines.extend(
                [
                    f"    for coordinate in {required}:",
                    f"        require_input_domain({series_id}[coordinate], {domain_literal}, series_id={quoted_id} + repr(coordinate))",
                ]
            )
    mapping = input_value_map_from_series(series.raw)
    if mapping is not None:
        used.add("apply_input_value_map")
        lines.append(
            f"    return apply_input_value_map({series_id}, {_python_literal(mapping)}, series_id={quoted_id})"
        )
    else:
        lines.append(f"    return {series_id}")
    return lines, used, extras


def _input_check_functions(catalog: SeriesCatalog) -> tuple[list[str], list[str], set[str]]:
    """Collect non-trivial input check sources and the runtime helpers they use."""
    checks: list[str] = []
    checked: list[str] = []
    used: set[str] = set()
    for series in _retained(catalog):
        if series.direction != "input":
            continue
        body, check_used, extras = _input_check(series, catalog)
        if len(body) == 1 and body[0] == f"    return {series.series_id}":
            continue
        used |= check_used
        checked.append(series.series_id)
        annotation = _annotation(series)
        checks.append(
            "\n".join(
                [
                    _check_signature(series.series_id, annotation, extras),
                    f'    """Validate `{series.series_id}` before the model reads it."""',
                    *body,
                ]
            )
        )
    return checks, checked, used


def emit_named_validation(catalog: SeriesCatalog) -> str:
    """Emit input schema, dtype, domain, and value-map checks for `Model` construction."""
    checks, checked, used = _input_check_functions(catalog)
    lines = [
        '"""Input schema, domain, and value-map checks for bound Model arguments."""',
        "",
        "from __future__ import annotations",
        "",
    ]
    stdlib: list[str] = []
    local: list[str] = []
    if any("datetime" in _annotation(catalog.get(sid)) for sid in checked):
        stdlib.append("from datetime import datetime")
    if any(not catalog.get(sid).single_valued for sid in checked):
        local.append("from . import data")
    local.extend(_generated_helper_imports(used))
    extra = [*stdlib, *(["", *local] if stdlib and local else local)]
    if extra:
        lines.extend(extra)
        lines.append("")
        lines.append("")
    if checks:
        lines.append("\n\n\n".join(checks))
        lines.extend(["", "", "CHECKS = {"])
        lines.extend(f"    {_python_literal(sid)}: _check_{sid}," for sid in checked)
        lines.append("}")
    else:
        lines.append("CHECKS = {}")
    lines.append("")
    return "\n".join(lines)


def _model_attribute(
    series: BoundSeries, deps: Mapping[str, SeriesDeps], catalog: SeriesCatalog
) -> list[str]:
    series_id = series.series_id
    args = ", ".join(f"{sid}={_argument_source(sid, catalog)}" for sid in deps[series_id].param_ids)
    return [
        "    @cached_property",
        f"    def {series_id}(self) -> {_annotation(series)}:",
        f"        return internals.{series_id}({args})",
    ]


def _model_recurrence_group(
    scc: tuple[str, ...], deps: Mapping[str, SeriesDeps], catalog: SeriesCatalog
) -> list[str]:
    name = scan_function_name(scc)
    params = scc_external_params(scc, deps, catalog.order)
    args = ", ".join(f"{sid}={_argument_source(sid, catalog)}" for sid in params)
    lines = [
        "    @cached_property",
        f"    def _{name}(self) -> internals.{_result_type(scc)}:",
        f"        return internals.{name}({args})",
    ]
    for sid in scc:
        lines.extend(
            [
                "",
                "    @cached_property",
                f"    def {sid}(self) -> {_annotation(catalog.get(sid))}:",
                f"        return self._{name}.{sid}",
            ]
        )
    return lines


def _model_init(catalog: SeriesCatalog) -> list[str]:
    """Bind static inputs, evaluate labellers, then check runtime-axis inputs."""
    deferred_ids = _deferred_runtime_inputs(catalog)
    deferred = tuple(
        series.series_id for series in _retained(catalog) if series.series_id in deferred_ids
    )
    lines = [
        "",
        "    def __init__(self, **inputs: object) -> None:",
        "        for name, value in inputs.items():",
    ]
    if deferred:
        names = ", ".join(repr(name) for name in deferred)
        lines.append(f"            if name in {{{names}}}:")
        lines.append("                continue")
    lines.extend(
        [
            "            check = validation.CHECKS.get(name)",
            "            setattr(self, name, value if check is None else check(value))",
        ]
    )
    if not deferred:
        return lines
    labellers = tuple(
        series.series_id for series in _retained(catalog) if _is_runtime_labeller(series)
    )
    quoted = ", ".join(repr(name) for name in labellers)
    lines.append(f"        for name in ({quoted},):")
    lines.append("            getattr(self, name)")
    for series_id in deferred:
        series = catalog.get(series_id)
        extras = []
        for axis in series.tensor_domain.axes:
            labeller = catalog.runtime_labeller(axis.name, axis.keys)
            if labeller is not None and labeller.series_id != series.series_id:
                extras.append(f"{axis.name}=self.{labeller.series_id}")
        extra = f", {', '.join(extras)}" if extras else ""
        lines.extend(
            [
                f"        if {series_id!r} in inputs:",
                f"            value = inputs[{series_id!r}]",
                f"            check = validation.CHECKS.get({series_id!r})",
                "            setattr(",
                "                self,",
                f"                {series_id!r},",
                f"                value if check is None else check(value{extra}),",
                "            )",
            ]
        )
    return lines


def _model_cells_method() -> list[str]:
    """Bind provenance templates over this model's evaluated labellers."""
    return [
        "",
        "    def cells(self, series_id: str) -> object:",
        '        """Authored worksheet cells of `series_id` over this model\'s labels."""',
        "        binding = getattr(data, series_id.upper())",
        "        cells = getattr(binding, 'cells', None)",
        "        if cells is None:",
        "            cells = getattr(data, f'{series_id.upper()}_CELLS')",
        "        bind = getattr(cells, 'bind', None)",
        "        if bind is None:",
        "            return cells",
        "        return bind(",
        "            **{",
        "                axis: getattr(self, labeller)",
        "                for axis, labeller in data.LABELLED_AXES.items()",
        "            }",
        "        )",
    ]


def _key_note(
    output: BoundSeries,
    catalog: SeriesCatalog,
    deps: Mapping[str, SeriesDeps],
) -> list[str]:
    """Docstring lines naming the inputs that determine runtime axis keys."""
    notes: list[str] = []
    for axis in output.tensor_domain.axes:
        labeller = catalog.runtime_labeller(axis.name, axis.keys)
        if labeller is None:
            continue
        inputs = [
            sid
            for sid in leaf_closure(labeller.series_id, catalog=catalog, deps=dict(deps))
            if catalog.get(sid).direction == "input"
        ]
        if not inputs:
            continue
        listed = ", ".join(f"`{sid}`" for sid in inputs)
        notes.append(f"    Keys along `{axis.name}` are determined by {listed}.")
    return notes


def emit_named_api(
    catalog: SeriesCatalog,
    deps: Mapping[str, SeriesDeps],
    scc_map: Mapping[str, tuple[str, ...]],
    constant_sets: Mapping[frozenset[str], str],
) -> str:
    """Emit the memoized `Model` and the public `compute_*` functions over it."""
    inputs = [s for s in _retained(catalog) if s.direction == "input"]
    model = [
        "class Model:",
        '    """Formula series of the workbook, evaluated on demand from bound inputs.',
        "",
        "    Each attribute evaluates its named formula once per model. Only the",
        "    inputs bound at construction are available, so a public function",
        "    supplies exactly the leaves of its output.",
        '    """',
        "",
    ]
    for series in inputs:
        model.append(f"    {series.series_id}: {_annotation(series)}")
    model.extend(_model_init(catalog))
    if _labelled_axes_map(catalog):
        model.extend(_model_cells_method())
    emitted_groups: set[tuple[str, ...]] = set()
    for series in _retained_formula_series(catalog):
        scc = scc_map.get(series.series_id, (series.series_id,))
        model.append("")
        if len(scc) > 1:
            if scc in emitted_groups:
                model.pop()
                continue
            emitted_groups.add(scc)
            model.extend(_model_recurrence_group(scc, deps, catalog))
            continue
        model.extend(_model_attribute(series, deps, catalog))
    functions: list[str] = []
    compute_names: list[str] = []
    for output in catalog.output_series():
        leaves = leaf_closure(output.series_id, catalog=catalog, deps=dict(deps))
        constants = frozenset(sid for sid in leaves if catalog.get(sid).direction == "constant")
        source, name = _public_function(output, leaves, catalog, constant_sets[constants], deps)
        functions.append(source)
        compute_names.append(name)
    aliases = list(constant_sets.values())
    lines = [
        '"""Generated functions accepting and returning named-coordinate values."""',
        "from __future__ import annotations",
        "from datetime import datetime",
        "from functools import cached_property",
        "from . import data, internals, validation",
        *([_constants_import(aliases)] if aliases else []),
        "from .runtime import publish",
        "",
        "",
        "\n".join(model),
        "",
        "\n\n".join(functions),
        "",
        "__all__ = [",
        "    'Model',",
        *(f"    {name!r}," for name in compute_names),
        "]",
        "",
    ]
    return "\n".join(lines)


def _output_constant_sets(
    catalog: SeriesCatalog, deps: Mapping[str, SeriesDeps]
) -> tuple[dict[frozenset[str], str], list[str]]:
    """Alias each public output's constant-leaf set, sharing subset unions."""
    constant_sets: dict[frozenset[str], str] = {}
    unique: list[frozenset[str]] = []
    for output in catalog.output_series():
        leaves = leaf_closure(output.series_id, catalog=catalog, deps=dict(deps))
        constants = frozenset(sid for sid in leaves if catalog.get(sid).direction == "constant")
        if constants not in constant_sets:
            constant_sets[constants] = f"_CONSTANTS_{len(constant_sets)}"
            unique.append(constants)
    names = sorted({name for group in unique for name in group})
    atoms: dict[tuple[int, ...], set[str]] = {}
    for name in names:
        key = tuple(i for i, group in enumerate(unique) if name in group)
        if key:
            atoms.setdefault(key, set()).add(name)
    family = set(unique)
    known: dict[frozenset[str], str] = {}
    lines: list[str] = []
    group_index = 0
    for atom_names in atoms.values():
        atom = frozenset(atom_names)
        users = sum(atom <= group for group in unique)
        if users >= 2 and atom not in family:
            alias = f"_CONSTANTS_GROUP_{group_index}"
            group_index += 1
            known[atom] = alias
            lines.append(f"{alias} = {_constant_literal(atom)}")
    for constants in unique:
        alias = constant_sets[constants]
        lines.append(f"{alias} = {_constant_set_source(constants, known)}")
        known[constants] = alias
    return constant_sets, lines


def _constants_import(aliases: Sequence[str]) -> str:
    """Import shared `_CONSTANTS_*` aliases from `data`, wrapping at 100 columns."""
    one_line = f"from .data import {', '.join(aliases)}"
    if len(one_line) <= 100:
        return one_line
    return "from .data import (\n    " + ",\n    ".join(aliases) + ",\n)"


def _constant_literal(names: frozenset[str]) -> str:
    return "frozenset({" + ", ".join(repr(name) for name in sorted(names)) + "})"


def _constant_set_source(constants: frozenset[str], known: Mapping[frozenset[str], str]) -> str:
    """Write a constant set as the largest known subset plus its extra members."""
    bases = [base for base in known if base and base < constants]
    if not bases:
        return _constant_literal(constants)
    base = max(bases, key=len)
    extra = constants - base
    if not extra:
        return known[base]
    return f"{known[base]} | {_constant_literal(extra)}"


def _public_function(
    output: BoundSeries,
    leaves: Sequence[str],
    catalog: SeriesCatalog,
    constants: str,
    deps: Mapping[str, SeriesDeps] | None = None,
) -> tuple[str, str]:
    inputs = [catalog.get(sid) for sid in leaves if catalog.get(sid).direction == "input"]
    name = output.compute_name or f"compute_{output.series_id}"
    summary = f"Compute `{output.series_id}` using authored coordinate identities."
    notes = _key_note(output, catalog, deps) if deps is not None else []
    docstring = [f'    """{summary}', "", *notes, '    """'] if notes else [f'    """{summary}"""']
    source = "\n".join(
        [
            _publish_line(output, constants),
            _signature(name, inputs, _annotation(output)),
            *docstring,
            f"    return Model(**locals()).{output.series_id}",
        ]
    )
    return source, name


# ---------------------------------------------------------------------------
# data.py
# ---------------------------------------------------------------------------


def _retained(catalog: SeriesCatalog) -> list[BoundSeries]:
    """Series with at least one authored cell inside the extracted graph."""
    return [
        series
        for series in catalog.series.values()
        if series.graph_cells is None or series.graph_cells
    ]


def _retained_formula_series(catalog: SeriesCatalog) -> list[BoundSeries]:
    return [series for series in _retained(catalog) if series.is_formula_series]


def _read_defaults(catalog: SeriesCatalog, workbook: Path | str) -> dict[str, dict[str, object]]:
    """Read authored workbook values for every input and constant series."""
    from excel_grapher.exporter.inverted_tree.emit import _coerce_cached_value

    defaults: dict[str, dict[str, object]] = {}
    leaves = [series for series in _retained(catalog) if series.direction in {"input", "constant"}]
    with _WorkbookValues(workbook) as reader:
        reader.prefetch(
            cell for series in leaves for cell in (series.authored_cells or series.cells)
        )
        for series in leaves:
            values: dict[str, object] = {}
            for cell in series.authored_cells or series.cells:
                value = reader.read(cell)
                values[cell] = (
                    None if value is None else _coerce_cached_value(value, series.dtype, cell)
                )
            defaults[series.series_id] = values
    return defaults


def _provenance_source(
    series: BoundSeries,
    named_axes: NamedAxes,
    domain_source: str | None = None,
    catalog: SeriesCatalog | None = None,
) -> str:
    """Describe authored cells as a worksheet rectangle when they form one.

    Falls back to an explicit coordinate-to-cell dictionary for irregular
    layouts. A rectangle with a few relocated cells keeps those as explicit
    exceptions.
    """
    cells = series.coordinate_cells
    literal = repr(dict(cells))
    if not cells or series.single_valued:
        return literal
    rectangle = _rectangle_source(series, cells, named_axes, catalog)
    if rectangle is not None:
        return rectangle
    domain = domain_source or f"{series.series_id.upper()}_DOMAIN"
    grid = _grid_source(series, dict(cells), domain)
    if grid is not None and len(grid) < len(literal):
        return grid
    return literal


def _rectangle_source(
    series: BoundSeries,
    cells: Mapping[tuple[Any, ...], CanonicalAddress],
    named_axes: NamedAxes,
    catalog: SeriesCatalog | None = None,
) -> str | None:
    """Describe a dense rectangle over one or two axes, or `None`."""
    from excel_grapher.core.address_keys import format_cell_key
    from excel_grapher.exporter.export_runtime.provenance import column_letter
    from excel_grapher.exporter.inverted_tree.catalog import _dense_rect

    rect = _dense_rect(list(cells.values()))
    if rect is None:
        return None
    sheet, first_row, first_col, last_row, last_col = rect
    height, width = last_row - first_row + 1, last_col - first_col + 1
    axes = series.tensor_domain.axes

    def layout_arg(axis: Any) -> str:
        if catalog is None:
            return layout_keys_source(axis, named_axes)
        return _axis_source_for_domain(axis, named_axes, catalog)

    def matches(layout: Mapping[tuple[Any, ...], str]) -> dict[tuple[Any, ...], str]:
        return {coord: address for coord, address in cells.items() if layout.get(coord) != address}

    if len(axes) == 1:
        axis = axes[0]
        if height == 1 and len(axis.keys) == width:
            layout = {
                (key,): format_cell_key(sheet, column_letter(first_col + i), first_row)
                for i, key in enumerate(axis.keys)
            }
            if not matches(layout):
                return (
                    f"row_cells({sheet!r}, {first_row}, {column_letter(first_col)!r}, "
                    f"{layout_arg(axis)})"
                )
        if width == 1 and len(axis.keys) == height:
            layout = {
                (key,): format_cell_key(sheet, column_letter(first_col), first_row + i)
                for i, key in enumerate(axis.keys)
            }
            if not matches(layout):
                return (
                    f"column_cells({sheet!r}, {column_letter(first_col)!r}, {first_row}, "
                    f"{layout_arg(axis)})"
                )
        return None
    if len(axes) == 2:
        for row_position, col_position in ((0, 1), (1, 0)):
            row_axis, col_axis = axes[row_position], axes[col_position]
            if len(row_axis.keys) != height or len(col_axis.keys) != width:
                continue
            layout = {}
            for r, row_key in enumerate(row_axis.keys):
                for c, col_key in enumerate(col_axis.keys):
                    coord = (row_key, col_key) if row_position == 0 else (col_key, row_key)
                    layout[coord] = format_cell_key(
                        sheet, column_letter(first_col + c), first_row + r
                    )
            exceptions = matches(layout)
            if len(exceptions) * 4 > len(cells):
                continue
            arguments = [
                repr(sheet),
                str(first_row),
                repr(column_letter(first_col)),
                layout_arg(row_axis),
                layout_arg(col_axis),
            ]
            if row_position == 1:
                arguments.append("cols_first=True")
            if exceptions:
                arguments.append(f"exceptions={exceptions!r}")
            return f"block_cells({', '.join(arguments)})"
    return None


def _grid_source(
    series: BoundSeries, cells: dict[tuple[Any, ...], str], domain_source: str
) -> str | None:
    """Describe cells whose sheet, row, and column each follow a group of key fields."""
    from itertools import product

    from excel_grapher.core.address_keys import parse_cell_coords

    fields = series.key_fields
    if not fields:
        return None
    parsed = {coord: parse_cell_coords(address) for coord, address in cells.items()}
    best: str | None = None
    for assignment in product(range(3), repeat=len(fields)):
        groups: list[list[int]] = [[], [], []]
        for position, group in enumerate(assignment):
            groups[group].append(position)
        mappings: list[dict[Any, Any]] = [{}, {}, {}]
        exceptions: dict[Any, str] = {}
        for coord, located in parsed.items():
            consistent = True
            for group, value in enumerate(located):
                key = tuple(coord[position] for position in groups[group])
                if mappings[group].setdefault(key, value) != value:
                    consistent = False
            if not consistent:
                exceptions[coord] = cells[coord]
        if len(exceptions) * 4 > len(cells):
            continue
        source = _render_grid(series, fields, groups, mappings, exceptions, domain_source)
        if best is None or len(source) < len(best):
            best = source
    return best


def _render_grid(
    series: BoundSeries,
    fields: Sequence[str],
    groups: Sequence[Sequence[int]],
    mappings: Sequence[Mapping[Any, Any]],
    exceptions: Mapping[Any, str],
    domain_source: str,
) -> str:
    from excel_grapher.exporter.export_runtime.provenance import column_letter

    def render(group: int, value: Any) -> str:
        positions = groups[group]
        mapping = mappings[group]
        if not positions:
            return repr(value(mapping[()]))
        names = tuple(fields[position] for position in positions)
        entries = {
            (key[0] if len(positions) == 1 else key): value(target)
            for key, target in mapping.items()
        }
        return f"({names!r}, {entries!r})"

    arguments = [
        render(0, str),
        domain_source,
        f"rows={render(1, int)}",
        f"cols={render(2, column_letter)}",
    ]
    if exceptions:
        arguments.append(f"exceptions={dict(exceptions)!r}")
    return f"grid_cells({', '.join(arguments)})"


def _product_axis_source(axis: Any, named_axes: NamedAxes) -> str:
    """Name the emitted axis, or build a subset `Axis` from `span` / keys."""
    emitted = named_axes.emitted(axis)
    if axis.keys == emitted.keys:
        return named_axes.constant(axis)
    return f"Axis({axis.name!r}, {layout_keys_source(axis, named_axes)}, {axis.key_type.__name__})"


def _coordinates_source(
    coordinates: Sequence[tuple[Any, ...]],
    axes: Sequence[Any],
    named_axes: NamedAxes,
    catalog: SeriesCatalog | None = None,
) -> str:
    """List sparse coordinates as runs along the last axis when that is shorter."""
    runtime = catalog is not None and any(
        catalog.runtime_labeller(axis.name, axis.keys) is not None for axis in axes
    )
    if runtime:
        positions = tuple(
            tuple(axis.keys.index(key) for axis, key in zip(axes, coord, strict=True))
            for coord in coordinates
        )
        return repr(positions)
    literal = repr(tuple(coordinates))
    if not coordinates or not axes:
        return literal
    last = named_axes.emitted(axes[-1])
    keys = last.keys
    runs: list[tuple[tuple[Any, ...], Any, Any]] = []
    for coordinate in coordinates:
        prefix, key = coordinate[:-1], coordinate[-1]
        if runs and runs[-1][0] == prefix:
            previous = keys.index(runs[-1][2])
            if previous + 1 < len(keys) and keys[previous + 1] == key:
                runs[-1] = (prefix, runs[-1][1], key)
                continue
        runs.append((prefix, key, key))
    source = f"coordinate_runs({named_axes.constant(axes[-1])}, {tuple(runs)!r})"
    return source if len(source) < len(literal) else literal


def _axis_source_for_domain(axis: Any, named_axes: NamedAxes, catalog: SeriesCatalog) -> str:
    """Name an interned axis, or an `AxisTemplate` for a runtime-labelled axis."""
    labeller = catalog.runtime_labeller(axis.name, axis.keys)
    if labeller is None:
        return _product_axis_source(axis, named_axes)
    emitted = named_axes.emitted(axis)
    if axis.keys == emitted.keys:
        return named_axes.constant(axis)
    source = labeller.tensor_domain.axes[0].keys
    extra = f", source={source!r}" if source != axis.keys else ""
    return (
        f"AxisTemplate({axis.name!r}, {axis.key_type.__name__}, "
        f"size={len(axis.keys)}, labeller={labeller.series_id!r}, snapshot={axis.keys!r}{extra})"
    )


def _domain_source(series: BoundSeries, named_axes: NamedAxes, catalog: SeriesCatalog) -> str:
    domain = series.tensor_domain
    runtime = any(
        catalog.runtime_labeller(axis.name, axis.keys) is not None for axis in domain.axes
    )
    if domain.coordinates is None:
        axes = ", ".join(_axis_source_for_domain(axis, named_axes, catalog) for axis in domain.axes)
        ctor = "DomainTemplate" if runtime else "Domain"
        return f"{ctor}.product({axes})"
    if runtime:
        axes = ", ".join(_axis_source_for_domain(axis, named_axes, catalog) for axis in domain.axes)
        coordinates = _coordinates_source(tuple(domain), domain.axes, named_axes, catalog)
        return f"DomainTemplate.explicit(axes=({axes},), coordinates={coordinates})"
    axes = ", ".join(named_axes.constant(axis) for axis in domain.axes)
    emitted = tuple(named_axes.emitted(axis) for axis in domain.axes)
    coordinates = _coordinates_source(tuple(domain), emitted, named_axes)
    return f"Domain.explicit(axes=({axes},), coordinates={coordinates})"


def _uses_grid_cells(series: BoundSeries, named_axes: NamedAxes) -> bool:
    """True when provenance is a `grid_cells` expression that names a domain."""
    cells = series.coordinate_cells
    if not cells or series.single_valued:
        return False
    if _rectangle_source(series, cells, named_axes) is not None:
        return False
    literal = repr(dict(cells))
    grid = _grid_source(series, dict(cells), "DOMAIN")
    return grid is not None and len(grid) < len(literal)


def _format_define_series(
    name: str,
    series_id: str,
    domain_source: str,
    values_source: str | None,
    cells_source: str,
    value_types: str,
    required_source: str | None,
    annotation: str,
) -> str:
    arguments = [repr(series_id), domain_source]
    if values_source is not None:
        arguments.append(values_source)
    arguments.append(f"cells={cells_source}")
    arguments.append(f"value_types={value_types}")
    if required_source is not None:
        arguments.append(f"required={required_source}")
    assignment = f"{name}: {annotation} = define_series({', '.join(arguments)})"
    if len(assignment) <= 100:
        return assignment
    body = ",\n".join(f"    {argument}" for argument in arguments)
    return f"{name}: {annotation} = define_series(\n{body},\n)"


def emit_named_data(
    catalog: SeriesCatalog,
    workbook: Path | str,
    named_axes: NamedAxes,
    literal_tables: Mapping[str, Mapping[tuple[object, ...], object]],
    constant_lines: Sequence[str] = (),
) -> str:
    """Emit shared axes, bound series, provenance, and workbook defaults."""
    from excel_grapher.exporter.inverted_tree.emit import _py_literal

    retained = _retained(catalog)
    defaults = _read_defaults(catalog, workbook)
    labelled = _labelled_axes_map(catalog)
    tensor_names = [
        "Axis",
        "Domain",
        "Series",
        "SeriesSpec",
        "coordinate_runs",
        "define_series",
    ]
    if labelled:
        tensor_names = [
            "Axis",
            "AxisTemplate",
            "Domain",
            "DomainTemplate",
            "Series",
            "SeriesSpec",
            "coordinate_runs",
            "define_series",
        ]
    lines = [
        '"""Authored domains, validated tensor types, and workbook defaults."""',
        "from __future__ import annotations",
        "from collections.abc import Iterator",
        "from contextlib import contextmanager",
        "from datetime import datetime",
        "from .provenance import block_cells, column_cells, grid_cells, row_cells",
        "from .runtime import span",
        f"from .tensor import {', '.join(tensor_names)}",
        f"CODEGEN_SCHEMA_VERSION = {REPRESENTATION_VERSION!r}",
        f"CODEGEN_FINGERPRINT = {named_codegen_fingerprint(catalog)!r}",
        "",
    ]
    for constant, axis in named_axes.items():
        labeller = catalog.runtime_labeller(axis.name, axis.keys)
        if labeller is None:
            lines.append(
                f"{constant} = Axis({axis.name!r}, {axis.keys!r}, {axis.key_type.__name__})"
            )
            continue
        source = labeller.tensor_domain.axes[0].keys
        extra = f", source={source!r}" if source != axis.keys else ""
        lines.append(
            f"{constant} = AxisTemplate({axis.name!r}, {axis.key_type.__name__}, "
            f"size={len(axis.keys)}, labeller={labeller.series_id!r}, snapshot={axis.keys!r}{extra})"
        )
    if labelled:
        lines.append(f"LABELLED_AXES = {labelled!r}")
    lines.append("")
    dtypes = sorted({series.python_dtype for series in retained if not series.single_valued})
    lines.extend(f"{dtype.upper()}_VALUES = {_value_types(dtype)}" for dtype in dtypes)
    domain_owner: dict[str | tuple[str, str], str] = {}
    constant_series: list[BoundSeries] = []
    for series in retained:
        name = series.series_id.upper()
        if series.direction == "constant":
            constant_series.append(series)
        if series.single_valued:
            lines.append(f"{name}_CELLS = {_provenance_source(series, named_axes)}")
            if series.direction in {"input", "constant"}:
                constant = name + ("_DEFAULT" if series.direction == "input" else "")
                value = next(iter(defaults[series.series_id].values()), None)
                lines.append(f"{constant} = {_py_literal(value)}")
            continue
        domain = series.tensor_domain
        constructed = _domain_source(series, named_axes, catalog)
        required_coords = series.required_coordinates
        required = tuple(coord for coord in domain if coord in required_coords)
        required_differs = len(required) != len(domain)
        runtime = any(
            catalog.runtime_labeller(axis.name, axis.keys) is not None for axis in domain.axes
        )
        uses_grid = _uses_grid_cells(series, named_axes)
        if runtime:
            owner = domain_owner.get(("rt", domain.fingerprint))
            if owner is None:
                domain_owner[("rt", domain.fingerprint)] = name
                domain_source = f"{name}_DOMAIN"
                lines.append(f"{name}_DOMAIN = {constructed}")
            else:
                domain_source = f"{owner}_DOMAIN"
        else:
            owner = domain_owner.get(domain.fingerprint)
            if owner is None:
                domain_owner[domain.fingerprint] = name
                if uses_grid or required_differs or constructed.startswith("Domain.explicit"):
                    domain_source = f"{name}_DOMAIN"
                    lines.append(f"{name}_DOMAIN = {constructed}")
                else:
                    domain_source = constructed
            else:
                domain_source = f"{owner}.domain"
        cells_source = _provenance_source(series, named_axes, domain_source, catalog)
        required_source = None
        if required_differs:
            ctor = "DomainTemplate" if runtime else "Domain"
            required_source = (
                f"{ctor}.explicit(axes={domain_source}.axes, "
                f"coordinates={_coordinates_source(required, domain.axes, named_axes, catalog)})"
            )
        values_source = None
        define_domain = domain_source
        if runtime and series.direction in {"input", "constant"}:
            snapshot = ", ".join(
                f"Axis({axis.name!r}, {axis.keys!r}, {axis.key_type.__name__})"
                for axis in domain.axes
            )
            define_domain = f"Domain.product({snapshot})"
            if required_source is None:
                required_source = domain_source
        if series.direction in {"input", "constant"}:
            cells = series.coordinate_cells
            values = tuple(defaults[series.series_id][cells[coord]] for coord in domain)
            values_source = _py_literal(values)
        if _is_runtime_labeller(series):
            axis = domain.axes[0]
            lines.append(
                f"{name}_POSITIONS = Domain.product(Axis({axis.name!r}, "
                f"{tuple(range(len(axis.keys)))!r}, int))"
            )
        annotation = (
            f"Series[{_value_annotation(series)}]"
            if values_source is not None
            else f"SeriesSpec[{_value_annotation(series)}]"
        )
        lines.append(
            _format_define_series(
                name,
                series.series_id,
                define_domain,
                values_source,
                cells_source,
                _schema_types(series),
                required_source,
                annotation,
            )
        )
        alias = _public_alias(series)
        if alias is not None:
            lines.append(f"{alias} = Series[{_value_annotation(series)}]")
        if series.direction == "input":
            lines.append(f"{name}_DEFAULT = {name}")
    for table, values in literal_tables.items():
        lines.append(f"{table} = {dict(values)!r}")
    names = tuple(series.series_id.upper() for series in constant_series)
    schemas = ", ".join(
        f"{series.series_id.upper()!r}: {series.series_id.upper()}.schema"
        for series in constant_series
        if not series.single_valued
    )
    lines.extend(
        [
            "",
            f"_CONSTANT_NAMES = frozenset({names!r})",
            "_CONSTANT_SCHEMAS = {" + schemas + "}",
            *([""] + list(constant_lines) if constant_lines else []),
            "",
            "",
            "@contextmanager",
            "def overrides(**values: object) -> Iterator[None]:",
            '    """Temporarily replace constants after validating their named schemas."""',
            "    unknown = values.keys() - _CONSTANT_NAMES",
            "    if unknown:",
            "        raise AttributeError(f'unknown constants: {sorted(unknown)}')",
            "    namespace = globals()",
            "    for name, value in values.items():",
            "        if name in _CONSTANT_SCHEMAS:",
            "            _CONSTANT_SCHEMAS[name].validate(value)",
            "    previous = {name: namespace[name] for name in values}",
            "    namespace.update(values)",
            "    try:",
            "        yield",
            "    finally:",
            "        namespace.update(previous)",
            "",
        ]
    )
    if not any("span(" in line for line in lines):
        lines.remove("from .runtime import span")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Package assembly
# ---------------------------------------------------------------------------


def emit_named_modules(
    catalog: SeriesCatalog,
    deps: Mapping[str, SeriesDeps],
    scc_map: Mapping[str, tuple[str, ...]],
    graph: DependencyGraph,
    workbook: Path | str,
    *,
    init_source: str,
    runtime_source: str,
    excel_source: str,
) -> dict[str, str]:
    """Assemble the standalone named package."""
    named_axes = NamedAxes.plan(
        axis
        for series in _retained(catalog)
        if not series.single_valued
        for axis in series.tensor_domain.axes
    )
    literal_tables: dict[str, dict[tuple[object, ...], object]] = {}
    internals = emit_named_internals(catalog, deps, scc_map, graph, named_axes, literal_tables)
    validation = emit_named_validation(catalog)
    constant_sets, constant_lines = _output_constant_sets(catalog, deps)
    api = emit_named_api(catalog, deps, scc_map, constant_sets)
    data = emit_named_data(catalog, workbook, named_axes, literal_tables, constant_lines)
    from excel_grapher.exporter.inverted_tree.standalone import build_runtime_modules

    export_runtime = Path(__file__).parents[1] / "export_runtime"
    tensor_source = (export_runtime / "tensor.py").read_text(encoding="utf-8")
    provenance_source = (export_runtime / "provenance.py").read_text(encoding="utf-8")
    return {
        "__init__.py": init_source
        + "\nfrom .tensor import Axis, Domain, Series, Tensor, TensorSchema\n"
        + "__all__ += ['Axis', 'Domain', 'Series', 'Tensor', 'TensorSchema']\n",
        "api.py": api,
        "validation.py": validation,
        "internals.py": internals,
        "data.py": data,
        "tensor.py": tensor_source,
        "provenance.py": provenance_source,
        **build_runtime_modules(runtime_source, excel_source),
    }


# ---------------------------------------------------------------------------
# Diagnostics
# ---------------------------------------------------------------------------


def inventory_named_emission(
    catalog: SeriesCatalog,
    deps: Mapping[str, SeriesDeps],
    scc_map: Mapping[str, tuple[str, ...]],
    graph: DependencyGraph,
) -> list[dict[str, object]]:
    """Attempt every formula body and report the series that cannot be lowered.

    Each entry names the series, its recurrence group when it has one, and
    the export error. An empty list means the whole model lowers to named
    computation.
    """
    named_axes = NamedAxes.plan(
        axis
        for series in _retained(catalog)
        if not series.single_valued
        for axis in series.tensor_domain.axes
    )
    failures: list[dict[str, object]] = []
    for series in _retained_formula_series(catalog):
        scc = scc_map.get(series.series_id, (series.series_id,))
        try:
            _semantic_body(
                series,
                catalog,
                deps[series.series_id],
                graph,
                named_axes,
                deferred=len(scc) > 1,
                scc_ids=frozenset(scc) if len(scc) > 1 else frozenset(),
                literal_tables={},
            )
        except InvertedTreeExportError as exc:
            failures.append(
                {
                    "series_id": series.series_id,
                    "group": list(scc) if len(scc) > 1 else None,
                    "layout": series.layout,
                    "cells": len(series.cells),
                    "error": str(exc),
                }
            )
    return failures
