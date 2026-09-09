"""Emit the named tensor contract over private scheduled calculation kernels."""

from __future__ import annotations

import ast
import hashlib
import json
import keyword
from collections.abc import Mapping
from dataclasses import fields, is_dataclass
from pathlib import Path
from typing import TYPE_CHECKING

from excel_grapher.core.formula_ast import FunctionCallNode
from excel_grapher.exporter.codegen import REPRESENTATION_VERSION
from excel_grapher.exporter.inverted_tree.ast_emit import (
    EmitContext,
    _as_measure_call,
    emit_expr,
    python_measure_type,
)
from excel_grapher.exporter.inverted_tree.deps import leaf_closure, try_formula_ast
from excel_grapher.series_bindings.input_coerce import measure_domain_from_series
from excel_grapher.series_bindings.resolve import _WorkbookValues

if TYPE_CHECKING:
    from excel_grapher.exporter.inverted_tree.catalog import BoundSeries, SeriesCatalog
    from excel_grapher.exporter.inverted_tree.deps import SeriesDeps
    from excel_grapher.grapher.graph import DependencyGraph


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


def _scalar_expression(node: object) -> bool:
    """Whether the existing scalar emitter has no positional range operations."""
    if isinstance(node, FunctionCallNode) and node.name.upper() in {
        "OFFSET",
        "INDIRECT",
    }:
        return False
    if is_dataclass(node) and not isinstance(node, type):
        return all(_scalar_expression(getattr(node, field.name)) for field in fields(node))
    if isinstance(node, tuple | list):
        return all(_scalar_expression(item) for item in node)
    return True


def _semantic_body(
    series: BoundSeries,
    catalog: SeriesCatalog,
    deps: SeriesDeps,
    graph: DependencyGraph,
    *,
    deferred: bool = False,
    scc_ids: frozenset[str] = frozenset(),
    literal_tables: dict[str, dict[tuple[object, ...], object]] | None = None,
) -> list[str] | None:
    """Emit proven scalar formula families as semantic coordinate loops."""
    reserved = set(catalog.series) | {
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
    }
    names: dict[str, str] = {}
    for field in series.key_fields:
        candidate = field.lower()
        if not candidate.isidentifier():
            return None
        if candidate.startswith("xl_"):
            candidate = "axis_" + candidate
        while keyword.iskeyword(candidate) or candidate in reserved:
            candidate += "_key"
        names[field] = candidate
        reserved.add(candidate)
    groups: dict[str, list[tuple[object, ...]]] = {}
    tables: dict[str, str] = {}
    used: set[str] = set()
    recursive = deferred
    for index, cell in enumerate(series.cells):
        node = try_formula_ast(graph, cell)
        if node is None:
            value_node = graph.get_node(cell)
            if value_node is None:
                # Retain authored off-graph metadata, but do not synthesize a
                # value or request a formula outside the extracted graph.
                continue
            from excel_grapher.exporter.inverted_tree.emit import _py_literal

            literal = _py_literal(value_node.value)
            expression = literal if value_node.value is None else _as_measure_call(literal, series)
            used.add("as_measure")
            coord = tuple(series.domain[index][field] for field in series.key_fields)
            groups.setdefault(expression, []).append(coord)
            continue
        if not _scalar_expression(node):
            return None
        ctx = EmitContext(
            host=series,
            catalog=catalog,
            deps=deps,
            host_index=index,
            host_cell=cell,
            index_var=None,
            prior_var=None,
            coordinate_vars=names,
            scc_ids=scc_ids | {series.series_id},
            graph=graph,
        )
        expression = _as_measure_call(emit_expr(node, ctx), series)
        used.add("as_measure")
        if series.series_id + "[" in expression:
            recursive = True
        if "lazy_table(" in expression:
            parsed = ast.parse(expression, mode="eval")
            for item in ast.walk(parsed):
                if not (
                    isinstance(item, ast.Call)
                    and isinstance(item.func, ast.Name)
                    and item.func.id == "lazy_table"
                ):
                    continue
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
        used |= ctx.used_runtime
    if literal_tables is not None and series.layout != "scalar":
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
    if series.layout == "scalar":
        if len(groups) != 1:
            return None
        expression = next(iter(groups))
        return [
            f"    from .runtime import {', '.join(sorted(used | {'XlError'}))}",
            *(f"    {alias} = {source}" for source, alias in tables.items()),
            "    try:",
            f"        return {expression}",
            "    except XlError as error:",
            "        return error.code",
        ]
    name = series.series_id.upper()
    occupied = reserved

    def temporary(base: str) -> str:
        candidate = base
        while candidate in occupied:
            candidate += "_"
        occupied.add(candidate)
        return candidate

    coordinate = temporary("_coordinate")
    records = temporary("_records")
    value = temporary("_value")
    error = temporary("_error")
    compute = temporary(
        f"_compute_{series.series_id}_coordinate" if deferred else "_compute_coordinate"
    )
    if recursive:
        used.add("CoordinateReader")
    lines = [
        f"    from .runtime import {', '.join(sorted(used | {'XlError'}))}",
        *(f"    {alias} = {source}" for source, alias in tables.items()),
    ]
    if recursive:
        lines.append(
            f"    def {compute}({coordinate}: tuple[str | int, ...]) -> {_value_annotation(series)}:"
        )
    else:
        lines.extend([f"    {records} = []", f"    for {coordinate} in data.{name}_REQUIRED:"])
    for i, axis in enumerate(series.tensor_domain.axes):
        lines.append(
            f"        {names[axis.name]} = cast({axis.key_type.__name__}, {coordinate}[{i}])"
        )
    lines.append("        try:")
    all_coordinates = {coord for coordinates in groups.values() for coord in coordinates}

    def condition(coordinates: list[tuple[object, ...]]) -> str:
        members = set(coordinates)
        for index, field in enumerate(series.key_fields):
            keys = {coord[index] for coord in coordinates}
            if len(keys) == 1:
                key = next(iter(keys))
                if {coord for coord in all_coordinates if coord[index] == key} == members:
                    return f"{names[field]} == {key!r}"
        return f"{coordinate} in {tuple(coordinates)!r}"

    # Put the largest family in the final else branch, avoiding an explicit
    # coordinate list for ordinary forward or backward recurrence periods.
    for i, (expression, coordinates) in enumerate(
        sorted(groups.items(), key=lambda item: len(item[1]))
    ):
        indent = "            "
        if len(groups) > 1:
            lines.append(
                "            else:"
                if i == len(groups) - 1
                else f"            {'if' if i == 0 else 'elif'} {condition(coordinates)}:"
            )
            indent += "    "
        lines.append(f"{indent}{value} = {expression}")
    lines.extend(
        [
            f"        except XlError as {error}:",
            f"            {value} = {error}.code",
        ]
    )
    if recursive:
        lines.extend(
            [
                f"        return {value}",
                f"    {series.series_id} = CoordinateReader({series.series_id!r}, data.{name}_REQUIRED, {compute})",
            ]
        )
        if not deferred and deps.is_scan and deps.scan_direction == "reversed":
            lines.extend(
                [
                    f"    for {coordinate} in reversed(tuple(data.{name}_REQUIRED)):",
                    f"        {series.series_id}[{coordinate}]",
                ]
            )
        lines.append(
            f"    return {_annotation(series)}.from_records(domain=data.{name}_REQUIRED, records=((coord, {series.series_id}[coord]) for coord in data.{name}_REQUIRED))"
        )
    else:
        lines.extend(
            [
                f"        {records}.append(({coordinate}, {value}))",
                f"    return {_annotation(series)}.from_records(domain=data.{name}_REQUIRED, records={records})",
            ]
        )
    return lines[:-1] if deferred else lines


def _facade(series: BoundSeries) -> str:
    return "".join(part.capitalize() for part in series.series_id.split("_"))


def _annotation(series: BoundSeries) -> str:
    inner = _value_annotation(series)
    return inner if series.layout == "scalar" else f"data.{_facade(series)}[{inner}]"


def _value_annotation(series: BoundSeries) -> str:
    """Describe every permitted workbook value, including blanks and errors."""
    types = python_measure_type(series).split(" | ")
    types.append("str")
    if series.layout != "scalar":
        types.append("None")
    return " | ".join(dict.fromkeys(types))


def _domain_source(series: BoundSeries, *, required: bool = False) -> str:
    domain = series.tensor_domain
    axes = ", ".join(
        f"Axis({axis.name!r}, {axis.keys!r}, {axis.key_type.__name__})" for axis in domain.axes
    )
    if required:
        required_coordinates = series.required_coordinates
        coordinates = tuple(coord for coord in domain if coord in required_coordinates)
        return (
            f"Domain.explicit(axes=({axes},), coordinates={coordinates!r})"
            if axes
            else f"Domain.explicit(axes=(), coordinates={coordinates!r})"
        )
    if domain.coordinates is None:
        return f"Domain.product({axes})"
    return f"Domain.explicit(axes=({axes},), coordinates={domain.coordinates!r})"


def _catalog_order(series: BoundSeries) -> tuple[tuple[object, ...], ...]:
    return tuple(tuple(point[name] for name in series.key_fields) for point in series.domain)


def _schema_types(series: BoundSeries) -> str:
    dtype = series.python_dtype
    types = (
        ["int", "float", "bool", "str"]
        if dtype == "float"
        else [dtype, "bool", "str"]
        if dtype == "int"
        else [dtype]
    )
    # Workbook blanks remain observations; the value is not domain membership.
    return "(" + ", ".join(dict.fromkeys([*types, "str", "type(None)"])) + ",)"


def _fused_helpers(
    functions: Mapping[str, ast.FunctionDef],
    catalog: SeriesCatalog,
    deps: Mapping[str, SeriesDeps],
    graph: DependencyGraph,
    literal_tables: dict[str, dict[tuple[object, ...], object]],
) -> list[str]:
    """Publish scheduled multi-series results as frozen named tensor fields."""
    lines = ["from dataclasses import dataclass", ""]
    for function in functions.values():
        if function.name in catalog.series or function.name.startswith("_"):
            continue
        returns = [node for node in function.body if isinstance(node, ast.Return)]
        if not returns or not isinstance(returns[-1].value, ast.Tuple):
            continue
        output_ids = []
        for item in returns[-1].value.elts:
            while (
                isinstance(item, ast.Call)
                and isinstance(item.func, ast.Name)
                and item.func.id in {"tuple", "reversed"}
                and len(item.args) == 1
            ):
                item = item.args[0]
            if not isinstance(item, ast.Name) or item.id not in catalog.series:
                break
            output_ids.append(item.id)
        else:
            params = [*function.args.args, *function.args.kwonlyargs]
            if any(param.arg not in catalog.series for param in params):
                continue
            result_type = "".join(part.capitalize() for part in function.name.split("_")) + "Result"
            lines.extend(
                [
                    "@dataclass(frozen=True, slots=True)",
                    f"class {result_type}:",
                    '    """Complete named results from one dependency evaluation."""',
                ]
            )
            for sid in output_ids:
                lines.append(f"    {sid}: {_annotation(catalog.get(sid))}")
            signature = ", ".join(
                f"{param.arg}: {_annotation(catalog.get(param.arg))}" for param in params
            )
            lines.extend(
                [
                    "",
                    f"def {function.name}({('*, ' + signature) if signature else ''}) -> {result_type}:",
                    '    """Evaluate the recurrence and publish complete tensors."""',
                ]
            )
            arguments = []
            for param in params:
                sid = param.arg
                series = catalog.get(sid)
                source = sid
                if series.layout != "scalar":
                    lines.append(f"    data.{sid.upper()}_SCHEMA.validate({sid})")
                    source = (
                        f"{sid}[data._{sid.upper()}_ORDER[0]]"
                        if series.is_scalar
                        else f"CoordinateBuffer({sid}, data._{sid.upper()}_ORDER)"
                    )
                arguments.append(f"{sid}={source}")
            semantic_bodies = [
                _semantic_body(
                    catalog.get(sid),
                    catalog,
                    deps[sid],
                    graph,
                    deferred=True,
                    scc_ids=frozenset(output_ids),
                    literal_tables=literal_tables,
                )
                if catalog.get(sid).layout != "scalar"
                else None
                for sid in output_ids
            ]
            semantic = all(body is not None for body in semantic_bodies)
            if semantic:
                for body in semantic_bodies:
                    assert body is not None
                    lines.extend(body)
            else:
                lines.append(f"    _raw_results = _kernels.{function.name}({', '.join(arguments)})")
            lines.append(f"    return {result_type}(")
            for index, sid in enumerate(output_ids):
                series = catalog.get(sid)
                value = f"_raw_results[{index}]"
                if series.layout != "scalar":
                    name = sid.upper()
                    records = (
                        f"((coord, {sid}[coord]) for coord in data.{name}_REQUIRED)"
                        if semantic
                        else f"((coord, value) for coord, value in zip(data._{name}_ORDER, {value}, strict=True) if coord in data._{name}_REQUIRED_MEMBERS)"
                    )
                    value = f"{_annotation(series)}.from_records(domain=data.{name}_REQUIRED, records={records})"
                lines.append(f"        {sid}={value},")
            lines.extend(["    )", ""])
    return lines


def emit_named_modules(
    modules: dict[str, str],
    catalog: SeriesCatalog,
    deps: Mapping[str, SeriesDeps],
    graph: DependencyGraph,
    workbook: Path | str,
) -> dict[str, str]:
    """Replace the public flat contract with schema-enforcing tensor functions.

    Flat buffers are confined to private scheduled kernels. Their ordering is
    explicitly recorded and never inferred from the tensor's canonical order.
    """
    from excel_grapher.exporter.inverted_tree.emit import _coerce_cached_value, _py_literal

    defaults = {}
    retained = [
        series
        for series in catalog.series.values()
        if series.graph_cells is None or series.graph_cells
    ]
    with _WorkbookValues(workbook) as reader:
        leaves = [series for series in retained if series.direction in {"input", "constant"}]
        reader.prefetch(
            cell for series in leaves for cell in (series.authored_cells or series.cells)
        )
        for series in leaves:
            values = []
            for cell in series.authored_cells or series.cells:
                value = reader.read(cell)
                values.append(
                    None if value is None else _coerce_cached_value(value, series.dtype, cell)
                )
            defaults[series.series_id] = values
    data_lines = [
        '"""Authored domains, validated tensor types, and workbook defaults."""',
        "from __future__ import annotations",
        "from datetime import datetime",
        "from typing import TypeVar",
        "from .tensor import Axis, Domain, Tensor, TensorSchema",
        "from . import _data",
        "T = TypeVar('T')",
        f"CODEGEN_SCHEMA_VERSION = {REPRESENTATION_VERSION!r}",
        f"CODEGEN_FINGERPRINT = {named_codegen_fingerprint(catalog)!r}",
        "",
    ]
    domain_names: dict[str, str] = {}
    for series in retained:
        data_lines.append(f"{series.series_id.upper()}_CELLS = {dict(series.coordinate_cells)!r}")
        if series.layout == "scalar":
            if series.direction in {"input", "constant"}:
                name = series.series_id.upper() + (
                    "_DEFAULT" if series.direction == "input" else ""
                )
                data_lines.append(f"{name} = {_py_literal(defaults[series.series_id][0])}")
            continue
        name = series.series_id.upper()
        domain = series.tensor_domain
        domain_source = domain_names.get(domain.fingerprint)
        if domain_source is None:
            domain_source = _domain_source(series)
            domain_names[domain.fingerprint] = f"{name}_DOMAIN"
        required_coordinates = series.required_coordinates
        required = tuple(coord for coord in domain if coord in required_coordinates)
        required_source = (
            f"{name}_DOMAIN"
            if len(required) == len(domain)
            else f"Domain.explicit(axes={name}_DOMAIN.axes, coordinates={required!r})"
        )
        catalog_order = _catalog_order(series)
        authored_order = tuple(series.coordinate_cells)
        order_source = (
            f"tuple({name}_CELLS)"
            if catalog_order == authored_order
            else f"tuple(coord for coord in {name}_CELLS if coord in _{name}_REQUIRED_MEMBERS)"
            if catalog_order
            == tuple(coord for coord in authored_order if coord in required_coordinates)
            else repr(catalog_order)
        )
        data_lines.extend(
            [
                f"{name}_DOMAIN = {domain_source}",
                f"{name}_REQUIRED = {required_source}",
                f"_{name}_REQUIRED_MEMBERS = frozenset({name}_REQUIRED)",
                f"{name}_SCHEMA = TensorSchema({series.series_id!r}, {name}_REQUIRED, {_schema_types(series)})",
                f"_{name}_ORDER = {order_source}",
                f"class {_facade(series)}(Tensor[T]):",
                '    """Validated authored series with complete coordinate indexing."""',
                "    __slots__ = ()",
                "    def __post_init__(self) -> None:",
                "        super().__post_init__()",
                f"        {name}_SCHEMA.validate(self)",
            ]
        )
        key_types = [axis.key_type.__name__ for axis in domain.axes]
        key_type = key_types[0] if len(key_types) == 1 else "tuple[" + ", ".join(key_types) + "]"
        data_lines.extend(
            [
                f"    def __getitem__(self, key: {key_type}) -> T:",
                "        return super().__getitem__(key)",
                "",
            ]
        )
        if series.direction in {"input", "constant"}:
            default = name + ("_DEFAULT" if series.direction == "input" else "")
            # Records filter structural blanks while retaining declared observations.
            data_lines.append(
                f"{default} = {_facade(series)}[{_value_annotation(series)}].from_legacy(domain={name}_DOMAIN, values={_py_literal(tuple(defaults[series.series_id]))}, coordinate_order=tuple({name}_CELLS))"
            )

    api_lines = [
        '"""Generated functions accepting and returning named-coordinate values."""',
        "from __future__ import annotations",
        "from datetime import datetime",
        "from . import data, _data, _kernel",
        "from typing import cast",
        "from .tensor import Tensor",
        "from ._tensor_lowering import CoordinateBuffer",
        "from .runtime import publish, XlError",
        "",
    ]
    constant_series = [series for series in retained if series.direction == "constant"]
    if constant_series:
        data_lines.extend(
            [
                "from contextlib import contextmanager",
                "from collections.abc import Iterator",
                f"_CONSTANT_NAMES = frozenset({tuple(series.series_id.upper() for series in constant_series)!r})",
                "_CONSTANT_SCHEMAS = {"
                + ", ".join(
                    f"{series.series_id.upper()!r}: {series.series_id.upper()}_SCHEMA"
                    for series in constant_series
                    if series.layout != "scalar"
                )
                + "}",
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
            ]
        )
    internal_lines = [*api_lines[:-1], "from . import _kernels", ""]
    helper_functions = {
        node.name: node
        for node in ast.parse(modules["internals.py"]).body
        if isinstance(node, ast.FunctionDef)
    }
    public_lines = api_lines
    literal_tables: dict[str, dict[tuple[object, ...], object]] = {}
    jobs = [(series, False) for series in catalog.output_series()] + [
        (series, True)
        for series in catalog.formula_series()
        if series.series_id in helper_functions
    ]
    for output, internal in jobs:
        api_lines = internal_lines if internal else public_lines
        leaves = leaf_closure(output.series_id, catalog=catalog, deps=dict(deps))
        inputs = (
            [
                catalog.get(arg.arg)
                for arg in [
                    *helper_functions[output.series_id].args.args,
                    *helper_functions[output.series_id].args.kwonlyargs,
                ]
            ]
            if internal
            else [catalog.get(sid) for sid in leaves if catalog.get(sid).direction == "input"]
        )
        constants = (
            []
            if internal
            else [catalog.get(sid) for sid in leaves if catalog.get(sid).direction == "constant"]
        )
        function = (
            output.series_id if internal else output.compute_name or f"compute_{output.series_id}"
        )
        params = ", ".join(f"{series.series_id}: {_annotation(series)}" for series in inputs)
        api_lines.append(
            f"@publish(key={output.key_fields!r}, domain={'None' if output.layout == 'scalar' else 'data.' + output.series_id.upper() + '_REQUIRED'}, constants={tuple(s.series_id for s in constants)!r}, cells=data.{output.series_id.upper()}_CELLS)"
        )
        api_lines.append(
            f"def {function}({('*, ' + params) if params else ''}) -> {_annotation(output)}:"
        )
        api_lines.append(
            f'    """Compute `{output.series_id}` using authored coordinate identities."""'
        )
        for series in [*inputs, *constants]:
            if series.layout != "scalar":
                variable = (
                    f"data.{series.series_id.upper()}"
                    if series.direction == "constant" and not internal
                    else series.series_id
                )
                api_lines.append(f"    data.{series.series_id.upper()}_SCHEMA.validate({variable})")
        if not internal:
            for series in inputs:
                domain = measure_domain_from_series(series.raw)
                if domain is None:
                    continue
                sid = series.series_id
                api_lines.append("    from .runtime import require_input_domain")
                if series.layout == "scalar":
                    api_lines.append(
                        f"    require_input_domain({sid}, {domain!r}, series_id={sid!r})"
                    )
                else:
                    # Restrict validation to this extraction's required domain;
                    # wider source tensors retain their off-graph observations.
                    coord = "_input_coordinate"
                    while coord in catalog.series:
                        coord += "_"
                    api_lines.extend(
                        [
                            f"    for {coord} in data.{sid.upper()}_REQUIRED:",
                            f"        require_input_domain({sid}[{coord}], {domain!r}, series_id={sid!r} + repr({coord}))",
                        ]
                    )
        args = []
        semantic = (
            _semantic_body(
                output, catalog, deps[output.series_id], graph, literal_tables=literal_tables
            )
            if internal
            else None
        )
        if semantic is not None:
            api_lines.extend([*semantic, ""])
            continue
        for series in inputs:
            sid = series.series_id
            source = (
                sid
                if series.layout == "scalar"
                else f"{sid}[data._{sid.upper()}_ORDER[0]]"
                if series.is_scalar
                else f"CoordinateBuffer({sid}, data._{sid.upper()}_ORDER)"
            )
            args.append(f"{sid}={source}")
        constant_args = []
        for series in constants:
            name = series.series_id.upper()
            source = (
                f"data.{name}"
                if series.layout == "scalar"
                else f"data.{name}[data._{name}_ORDER[0]]"
                if series.is_scalar
                else f"CoordinateBuffer(data.{name}, data._{name}_ORDER)"
            )
            constant_args.append(f"{name}={source}")
        indent = "    "
        if constant_args:
            api_lines.append(f"    with _data.overrides({', '.join(constant_args)}):")
            indent += "    "
        call = f"result = {'_kernels' if internal else '_kernel'}.{function}({', '.join(args)})"
        if output.is_scalar:
            api_lines.extend(
                [
                    f"{indent}try:",
                    f"{indent}    {call}",
                    f"{indent}except XlError as error:",
                    f"{indent}    result = error.code",
                ]
            )
        else:
            api_lines.append(f"{indent}{call}")
        if output.layout == "scalar":
            api_lines.append("    return result[0] if isinstance(result, tuple) else result")
        else:
            name = output.series_id.upper()
            if output.is_scalar:
                api_lines.append("    result = result if isinstance(result, tuple) else (result,)")
            api_lines.append(
                f"    return {_annotation(output)}.from_records(domain=data.{name}_REQUIRED, records=((coord, value) for coord, value in zip(data._{name}_ORDER, result, strict=True) if coord in data._{name}_REQUIRED_MEMBERS))"
            )
        api_lines.append("")
    internal_lines.extend(_fused_helpers(helper_functions, catalog, deps, graph, literal_tables))
    data_lines.extend(f"{name} = {values!r}" for name, values in literal_tables.items())
    result = {
        "__init__.py": modules["__init__.py"]
        + "\nfrom .tensor import Axis, Domain, Tensor, TensorSchema\n"
        + "__all__ += ['Axis', 'Domain', 'Tensor', 'TensorSchema']\n",
        "api.py": "\n".join(public_lines) + "\n",
        "internals.py": "\n".join(internal_lines) + "\n",
        "data.py": "\n".join(data_lines) + "\n",
        "tensor.py": Path(__file__)
        .parents[1]
        .joinpath("export_runtime", "tensor.py")
        .read_text(encoding="utf-8"),
        "_tensor_lowering.py": Path(__file__)
        .parents[1]
        .joinpath("export_runtime", "tensor_lowering.py")
        .read_text(encoding="utf-8"),
        "_kernel.py": modules["api.py"]
        .replace("from . import data", "from . import _data as data")
        .replace("from . import internals", "from . import _kernels as internals"),
        "_kernels.py": modules["internals.py"]
        .replace("from . import data", "from . import _data as data")
        .replace("from .data import", "from ._data import"),
        "_data.py": modules["data.py"],
        "runtime.py": modules["runtime.py"].replace(
            "from excel_grapher.exporter.export_runtime.tensor import Domain, Tensor",
            "from .tensor import Domain, Tensor",
        ),
    }
    return result
