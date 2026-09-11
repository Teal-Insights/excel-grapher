"""Emit the public named tensor contract: `data`, `internals`, `validation`, and `api`.

Every formula series becomes one inspectable named function in `internals`
whose body is the workbook formula family expressed over semantic
coordinates. Public `compute_*` functions orchestrate those functions in
dependency order, sharing intermediate results. Input schema, domain, and
value-map checks live in `validation` so `api` stays the user-facing surface.
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
from excel_grapher.exporter.inverted_tree.deps import (
    leaf_closure,
    node_formula_ast,
    try_formula_ast,
)
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.exporter.inverted_tree.named_axes import NamedAxes, python_identifier
from excel_grapher.exporter.inverted_tree.schedule import scan_function_name, scc_external_params
from excel_grapher.series_bindings.input_coerce import (
    input_value_map_from_series,
    measure_domain_from_series,
)
from excel_grapher.series_bindings.resolve import _WorkbookValues

if TYPE_CHECKING:
    from excel_grapher.core.address_keys import CanonicalAddress
    from excel_grapher.exporter.inverted_tree.catalog import BoundSeries, SeriesCatalog
    from excel_grapher.exporter.inverted_tree.deps import SeriesDeps
    from excel_grapher.grapher.graph import DependencyGraph

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


def _annotation(series: BoundSeries) -> str:
    return _value_annotation(series) if series.single_valued else f"data.{_facade(series)}"


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


def _coordinate_names(series: BoundSeries, reserved: set[str]) -> dict[str, str]:
    """Return a loop-variable name for every key field of `series`."""
    names: dict[str, str] = {}
    for field in series.key_fields:
        candidate = python_identifier(field.lower())
        if candidate.startswith("xl_"):
            candidate = "axis_" + candidate
        while candidate in reserved:
            candidate += "_key"
        names[field] = candidate
        reserved.add(candidate)
    return names


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
        and node.func.id in {"lazy_table", "view"}
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
    """
    reserved = set(catalog.series) | set(_RESERVED_NAMES)
    names = _coordinate_names(series, reserved)
    groups: dict[str, list[tuple[object, ...]]] = {}
    tables: dict[str, str] = {}
    used: set[str] = {"XlError"}
    recursive = deferred
    # A scalar layout publishes one observation: its first bound cell.
    cells = series.cells[:1] if series.single_valued else series.cells
    for index, cell in enumerate(cells):
        if graph.get_node(cell) is None:
            # Retain authored off-graph metadata, but do not synthesize a
            # value or request a formula outside the extracted graph.
            continue
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
        used |= ctx.used_runtime
        parsed = ast.parse(expression, mode="eval")
        if any(
            isinstance(item, ast.Name) and item.id == series.series_id for item in ast.walk(parsed)
        ):
            recursive = True
        if "lazy_table(" in expression or "view(" in expression:
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
        coord = tuple(series.domain[index][field] for field in series.key_fields)
        groups.setdefault(expression, []).append(coord)
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
            literal_tables[table] = literals
            selectors = ", ".join(names.values()) + ","
            groups[f"as_measure(data.{table}[{selectors}])"] = list(literals)
    lines = [f"    {alias} = {source}" for source, alias in tables.items()]
    if series.single_valued:
        if len(groups) != 1:
            raise InvertedTreeExportError(
                f"series {series.series_id!r}: scalar series has no graph formula"
            )
        expression = next(iter(groups))
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
    name = series.series_id.upper()
    occupied = reserved

    def temporary(base: str) -> str:
        candidate = base
        while candidate in occupied:
            candidate += "_"
        occupied.add(candidate)
        return candidate

    formula = temporary(f"{series.series_id}_formula" if deferred else "formula")
    parameters = ", ".join(
        f"{names[axis.name]}: {axis.key_type.__name__}" for axis in series.tensor_domain.axes
    )
    lines.append(f"    def {formula}({parameters}) -> {_value_annotation(series)}:")
    all_coordinates = {coord for coordinates in groups.values() for coord in coordinates}

    def condition(coordinates: list[tuple[object, ...]]) -> str:
        synthesized = _family_condition(
            coordinates, all_coordinates, series.tensor_domain.axes, names
        )
        if synthesized is not None:
            return synthesized
        selectors = ", ".join(names[axis.name] for axis in series.tensor_domain.axes)
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
            f"    {series.series_id} = CoordinateReader({series.series_id!r}, data.{name}_REQUIRED, {formula})"
        )
        if not deferred and deps.is_scan and deps.scan_direction == "reversed":
            coordinate = temporary("coordinate")
            lines.extend(
                [
                    f"    for {coordinate} in reversed(tuple(data.{name}_REQUIRED)):",
                    f"        {series.series_id}[{coordinate}]",
                ]
            )
        if not deferred:
            lines.append(f"    return {_materialize(series)}")
    else:
        used.add("evaluate")
        lines.append(
            f"    return {_annotation(series)}.collect(evaluate({formula}, data.{name}_REQUIRED))"
        )
    return lines, used


def _materialize(series: BoundSeries) -> str:
    """Publish a completed demand-driven reader as an immutable tensor."""
    name = series.series_id.upper()
    return (
        f"{_annotation(series)}.collect("
        f"(coord, {series.series_id}[coord]) for coord in data.{name}_REQUIRED)"
    )


# ---------------------------------------------------------------------------
# internals.py
# ---------------------------------------------------------------------------


def _signature(name: str, params: Sequence[BoundSeries], returns: str) -> str:
    joined = ", ".join(f"{series.series_id}: {_annotation(series)}" for series in params)
    return f"def {name}({('*, ' + joined) if joined else ''}) -> {returns}:"


def _schema_checks(params: Sequence[BoundSeries]) -> list[str]:
    """Validate tensor parameters that arrive from outside the generated model.

    Results of other named functions are `Series` instances validated on
    construction, so only inputs and constants are checked again here.
    """
    return [
        f"    data.{series.series_id.upper()}_SCHEMA.validate({series.series_id})"
        for series in params
        if not series.single_valued and series.direction in {"input", "constant"}
    ]


def _publish_line(series: BoundSeries, constants: str | None = None) -> str:
    name = series.series_id.upper()
    arguments = (
        [f"key={series.key_fields!r}", "domain=None"]
        if series.single_valued
        else [f"data.{name}_SCHEMA"]
    )
    if constants is not None:
        arguments.append(f"constants={constants}")
    arguments.append(f"cells=data.{name}_CELLS")
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
                    *_schema_checks(params),
                    *body,
                ]
            )
        )
    lines = [
        '"""Named calculation functions for every bound formula series."""',
        "from __future__ import annotations",
        "from dataclasses import dataclass",
        "from datetime import datetime",
        "from typing import cast",
        "from . import data",
        "from .tensor import Domain",
        f"from .runtime import {', '.join(sorted(used))}",
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
    for index, axis in enumerate(axes):
        present = [key for key in axis.keys if any(coord[index] == key for coord in universe)]
        keys = [key for key in present if any(coord[index] == key for coord in members)]
        selected.append(set(keys))
        name = names[axis.name]
        if len(keys) == len(present):
            continue
        if len(keys) == 1:
            tests.append(f"{name} == {keys[0]!r}")
            continue
        first, last = present.index(keys[0]), present.index(keys[-1])
        contiguous = present[first : last + 1] == keys
        ascending = axis.key_type is int and present == sorted(cast(Sequence[int], present))
        if contiguous and ascending:
            if last == len(present) - 1:
                tests.append(f"{name} >= {keys[0]!r}")
            elif first == 0:
                tests.append(f"{name} <= {keys[-1]!r}")
            else:
                tests.append(f"{keys[0]!r} <= {name} <= {keys[-1]!r}")
            continue
        tests.append(f"{name} in {tuple(keys)!r}")
    product = {
        coord
        for coord in universe
        if all(coord[index] in keys for index, keys in enumerate(selected))
    }
    if product == members:
        return " and ".join(tests) if tests else None
    relation = _diagonal_condition(members, universe, product, tests, axes, names)
    if relation is not None:
        return relation
    return _union_condition(members, universe, axes, names)


def _diagonal_condition(
    members: set[tuple[object, ...]],
    universe: set[tuple[object, ...]],
    product: set[tuple[object, ...]],
    tests: Sequence[str],
    axes: Sequence[Any],
    names: Mapping[str, str],
) -> str | None:
    """A family on one diagonal band of two integer axes."""
    integer_axes = [index for index, axis in enumerate(axes) if axis.key_type is int]
    for position, left in enumerate(integer_axes):
        for right in integer_axes[position + 1 :]:

            def difference(coord: tuple[object, ...], left: int = left, right: int = right) -> int:
                return cast(int, coord[right]) - cast(int, coord[left])

            differences = sorted({difference(coord) for coord in members})
            low, high = differences[0], differences[-1]
            if differences != list(range(low, high + 1)):
                continue
            left_name, right_name = names[axes[left].name], names[axes[right].name]
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
) -> str | None:
    """A family that splits along one axis into a few expressible sub-families."""
    for index, axis in enumerate(axes):
        keys = [key for key in axis.keys if any(coord[index] == key for coord in members)]
        if not 2 <= len(keys) <= 4:
            continue
        parts: list[str] = []
        for key in keys:
            group = [coord for coord in members if coord[index] == key]
            restricted = {coord for coord in universe if coord[index] == key}
            selector = f"{names[axis.name]} == {key!r}"
            if set(group) == restricted:
                parts.append(selector)
                continue
            inner = _family_condition(group, restricted, axes, names)
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
            *_schema_checks(params),
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
        value = f"{sid}[()]" if series.single_valued else _materialize(series)
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


def _check_signature(series_id: str, annotation: str) -> str:
    """Write a check signature, wrapping only when the one-liner exceeds 100 columns."""
    name = f"_check_{series_id}"
    one_line = f"def {name}({series_id}: {annotation}) -> {annotation}:"
    if len(one_line) <= 100:
        return one_line
    return f"def {name}(\n    {series_id}: {annotation},\n) -> {annotation}:"


def _input_check(series: BoundSeries) -> tuple[list[str], set[str]]:
    """Validate one input's schema and declared domain, then apply its value map."""
    lines: list[str] = []
    used: set[str] = set()
    series_id = series.series_id
    quoted_id = _python_literal(series_id)
    if not series.single_valued:
        lines.append(f"    data.{series_id.upper()}_SCHEMA.validate({series_id})")
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
            lines.extend(
                [
                    f"    for coordinate in data.{series_id.upper()}_REQUIRED:",
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
    return lines, used


def _input_check_functions(catalog: SeriesCatalog) -> tuple[list[str], list[str], set[str]]:
    """Collect non-trivial input check sources and the runtime helpers they use."""
    checks: list[str] = []
    checked: list[str] = []
    used: set[str] = set()
    for series in _retained(catalog):
        if series.direction != "input":
            continue
        body, check_used = _input_check(series)
        if len(body) == 1 and body[0] == f"    return {series.series_id}":
            continue
        used |= check_used
        checked.append(series.series_id)
        annotation = _annotation(series)
        checks.append(
            "\n".join(
                [
                    _check_signature(series.series_id, annotation),
                    f'    """Validate `{series.series_id}` before the model reads it."""',
                    *body,
                ]
            )
        )
    return checks, checked, used


def emit_named_validation(catalog: SeriesCatalog) -> str:
    """Emit input schema, domain, and value-map checks for `Model` construction."""
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
    if used:
        local.append(f"from .runtime import {', '.join(sorted(used))}")
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
    model.extend(
        [
            "",
            "    def __init__(self, **inputs: object) -> None:",
            "        for name, value in inputs.items():",
            "            check = validation.CHECKS.get(name)",
            "            setattr(self, name, value if check is None else check(value))",
        ]
    )
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
        source, name = _public_function(output, leaves, catalog, constant_sets[constants])
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
    lines: list[str] = []
    for output in catalog.output_series():
        leaves = leaf_closure(output.series_id, catalog=catalog, deps=dict(deps))
        constants = frozenset(sid for sid in leaves if catalog.get(sid).direction == "constant")
        if constants not in constant_sets:
            alias = f"_CONSTANTS_{len(constant_sets)}"
            lines.append(f"{alias} = {_constant_set_source(constants, constant_sets)}")
            constant_sets[constants] = alias
    return constant_sets, lines


def _constants_import(aliases: Sequence[str]) -> str:
    """Import shared `_CONSTANTS_*` aliases from `data`, wrapping at 100 columns."""
    one_line = f"from .data import {', '.join(aliases)}"
    if len(one_line) <= 100:
        return one_line
    return "from .data import (\n    " + ",\n    ".join(aliases) + ",\n)"


def _constant_set_source(constants: frozenset[str], known: Mapping[frozenset[str], str]) -> str:
    """Write a constant set as the largest known subset plus its extra members."""

    def literal(names: frozenset[str]) -> str:
        return "frozenset({" + ", ".join(repr(name) for name in sorted(names)) + "})"

    bases = [base for base in known if base and base < constants]
    if not bases:
        return literal(constants)
    base = max(bases, key=len)
    return f"{known[base]} | {literal(constants - base)}"


def _public_function(
    output: BoundSeries,
    leaves: Sequence[str],
    catalog: SeriesCatalog,
    constants: str,
) -> tuple[str, str]:
    inputs = [catalog.get(sid) for sid in leaves if catalog.get(sid).direction == "input"]
    name = output.compute_name or f"compute_{output.series_id}"
    source = "\n".join(
        [
            _publish_line(output, constants),
            _signature(name, inputs, _annotation(output)),
            f'    """Compute `{output.series_id}` using authored coordinate identities."""',
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


def _provenance_source(series: BoundSeries, named_axes: NamedAxes) -> str:
    """Describe authored cells as a worksheet rectangle when they form one.

    Falls back to an explicit coordinate-to-cell dictionary for irregular
    layouts. A rectangle with a few relocated cells keeps those as explicit
    exceptions.
    """
    cells = series.coordinate_cells
    literal = repr(dict(cells))
    if not cells or series.single_valued:
        return literal
    rectangle = _rectangle_source(series, cells, named_axes)
    if rectangle is not None:
        return rectangle
    grid = _grid_source(series, dict(cells))
    if grid is not None and len(grid) < len(literal):
        return grid
    return literal


def _rectangle_source(
    series: BoundSeries, cells: Mapping[tuple[Any, ...], CanonicalAddress], named_axes: NamedAxes
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
                    f"{named_axes.constant(axis)})"
                )
        if width == 1 and len(axis.keys) == height:
            layout = {
                (key,): format_cell_key(sheet, column_letter(first_col), first_row + i)
                for i, key in enumerate(axis.keys)
            }
            if not matches(layout):
                return (
                    f"column_cells({sheet!r}, {column_letter(first_col)!r}, {first_row}, "
                    f"{named_axes.constant(axis)})"
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
                named_axes.constant(row_axis),
                named_axes.constant(col_axis),
            ]
            if row_position == 1:
                arguments.append("cols_first=True")
            if exceptions:
                arguments.append(f"exceptions={exceptions!r}")
            return f"block_cells({', '.join(arguments)})"
    return None


def _grid_source(series: BoundSeries, cells: dict[tuple[Any, ...], str]) -> str | None:
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
        source = _render_grid(series, fields, groups, mappings, exceptions)
        if best is None or len(source) < len(best):
            best = source
    return best


def _render_grid(
    series: BoundSeries,
    fields: Sequence[str],
    groups: Sequence[Sequence[int]],
    mappings: Sequence[Mapping[Any, Any]],
    exceptions: Mapping[Any, str],
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
        f"{series.series_id.upper()}_DOMAIN",
        f"rows={render(1, int)}",
        f"cols={render(2, column_letter)}",
    ]
    if exceptions:
        arguments.append(f"exceptions={dict(exceptions)!r}")
    return f"grid_cells({', '.join(arguments)})"


def _coordinates_source(
    coordinates: Sequence[tuple[Any, ...]], axes: Sequence[Any], named_axes: NamedAxes
) -> str:
    """List sparse coordinates as runs along the last axis when that is shorter."""
    literal = repr(tuple(coordinates))
    if not coordinates or not axes:
        return literal
    keys = axes[-1].keys
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


def _domain_source(series: BoundSeries, named_axes: NamedAxes) -> str:
    domain = series.tensor_domain
    axes = ", ".join(named_axes.constant(axis) for axis in domain.axes)
    if domain.coordinates is None:
        return f"Domain.product({axes})"
    coordinates = _coordinates_source(domain.coordinates, domain.axes, named_axes)
    return f"Domain.explicit(axes=({axes},), coordinates={coordinates})"


def emit_named_data(
    catalog: SeriesCatalog,
    workbook: Path | str,
    named_axes: NamedAxes,
    literal_tables: Mapping[str, Mapping[tuple[object, ...], object]],
    constant_lines: Sequence[str] = (),
) -> str:
    """Emit shared axes, domains, schemas, facades, provenance, and defaults."""
    from excel_grapher.exporter.inverted_tree.emit import _py_literal

    retained = _retained(catalog)
    defaults = _read_defaults(catalog, workbook)
    lines = [
        '"""Authored domains, validated tensor types, and workbook defaults."""',
        "from __future__ import annotations",
        "from collections.abc import Iterator",
        "from contextlib import contextmanager",
        "from datetime import datetime",
        "from .provenance import block_cells, column_cells, grid_cells, row_cells",
        "from .tensor import Axis, Domain, Series, TensorSchema, coordinate_runs",
        f"CODEGEN_SCHEMA_VERSION = {REPRESENTATION_VERSION!r}",
        f"CODEGEN_FINGERPRINT = {named_codegen_fingerprint(catalog)!r}",
        "",
    ]
    for constant, axis in named_axes.items():
        lines.append(f"{constant} = Axis({axis.name!r}, {axis.keys!r}, {axis.key_type.__name__})")
    lines.append("")
    dtypes = sorted({series.python_dtype for series in retained if not series.single_valued})
    lines.extend(f"{dtype.upper()}_VALUES = {_value_types(dtype)}" for dtype in dtypes)
    domain_names: dict[str, str] = {}
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
        domain_source = domain_names.get(domain.fingerprint)
        if domain_source is None:
            domain_source = _domain_source(series, named_axes)
            domain_names[domain.fingerprint] = f"{name}_DOMAIN"
        required = tuple(coord for coord in domain if coord in series.required_coordinates)
        required_source = (
            f"{name}_DOMAIN"
            if len(required) == len(domain)
            else f"Domain.explicit(axes={name}_DOMAIN.axes, "
            f"coordinates={_coordinates_source(required, domain.axes, named_axes)})"
        )
        keyed_by = ", ".join(axis.name for axis in domain.axes)
        lines.extend(
            [
                f"{name}_DOMAIN = {domain_source}",
                f"{name}_REQUIRED = {required_source}",
                f"{name}_CELLS = {_provenance_source(series, named_axes)}",
                f"{name}_SCHEMA = TensorSchema({series.series_id!r}, {name}_REQUIRED, {_schema_types(series)})",
                f"class {_facade(series)}(Series[{_value_annotation(series)}]):",
                f'    """`{series.series_id}` by {keyed_by}."""',
                f"    schema = {name}_SCHEMA",
                "",
            ]
        )
        if series.direction in {"input", "constant"}:
            constant = name + ("_DEFAULT" if series.direction == "input" else "")
            cells = series.coordinate_cells
            values = tuple(defaults[series.series_id][cells[coord]] for coord in domain)
            lines.append(f"{constant} = {_facade(series)}({name}_DOMAIN, {_py_literal(values)})")
    for table, values in literal_tables.items():
        lines.append(f"{table} = {dict(values)!r}")
    names = tuple(series.series_id.upper() for series in constant_series)
    schemas = ", ".join(
        f"{series.series_id.upper()!r}: {series.series_id.upper()}_SCHEMA"
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
            "    for name, value in values.items():",
            "        if name in _CONSTANT_SCHEMAS:",
            "            _CONSTANT_SCHEMAS[name].validate(value)",
            "    namespace = globals()",
            "    previous = {name: namespace[name] for name in values}",
            "    namespace.update(values)",
            "    try:",
            "        yield",
            "    finally:",
            "        namespace.update(previous)",
            "",
        ]
    )
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
        + "\nfrom .tensor import Axis, Domain, Tensor, TensorSchema\n"
        + "__all__ += ['Axis', 'Domain', 'Tensor', 'TensorSchema']\n",
        "api.py": api,
        "validation.py": validation,
        "internals.py": internals,
        "data.py": data,
        "tensor.py": tensor_source,
        "provenance.py": provenance_source,
        **build_runtime_modules(runtime_source),
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
