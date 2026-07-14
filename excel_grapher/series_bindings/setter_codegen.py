"""Generate `set_*` functions that apply Records to graph input leaves."""

from __future__ import annotations

import ast
import warnings
from collections.abc import Iterable, Mapping, Sequence
from pathlib import Path
from typing import TYPE_CHECKING, Any

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.codegen_literals import (
    emit_setter_type_alias_lines,
    py_scalar_literal,
    resolutions_include_datetime,
)
from excel_grapher.series_bindings.docstrings import (
    emit_docstring_literal,
    resolve_series_function_docstring,
)
from excel_grapher.series_bindings.normalize import effective_dimension_id, has_input_direction
from excel_grapher.series_bindings.resolve import (
    _lookup_concept_dtype,
    resolve_series_bindings,
    warn_series_resolution_issues,
)
from excel_grapher.series_bindings.types import (
    SeriesResolution,
    WorkbookSeriesBindings,
)

if TYPE_CHECKING:
    from excel_grapher.series_bindings.docstring_renderers import SeriesDocstringRendererSpec
    from excel_grapher.series_bindings.docstrings import SeriesBindingDocstringCallbackSpec


def _key_tuple_literal(key_fields: list[str], key: Mapping[str, object]) -> str:
    pairs = ", ".join(
        f"({repr(f)}, {py_scalar_literal(key[f], datetime_ref='datetime')})" for f in key_fields
    )
    return f"({pairs},)" if len(key_fields) == 1 else f"({pairs})"


def _measure_concept(series: dict[str, Any]) -> str:
    measure = (series.get("structure") or {}).get("measure") or {}
    return str(measure.get("concept") or "OBS_VALUE")


def _measure_dtype_for_codegen(
    series: dict[str, Any],
    concept_scheme: dict[str, Any] | None = None,
) -> str | None:
    """Return the measure dtype to enforce in generated setters, if any.

    Precedence matches resolve/docstrings: explicit `structure.measure.dtype`, then
    `concept_scheme` for the measure concept, then a non-`auto` measure `bind.read`.
    """
    measure = (series.get("structure") or {}).get("measure") or {}
    concept_name = str(measure.get("concept") or "OBS_VALUE")
    inferred = _lookup_concept_dtype(concept_scheme, series, concept_name)
    if inferred is not None:
        return inferred
    raw_bind = measure.get("bind")
    bind: dict[str, Any] = raw_bind if isinstance(raw_bind, dict) else {}
    read = bind.get("read")
    if read is not None and read != "auto":
        return str(read)
    return None


def _accepts_scalar_shorthand(
    series: dict[str, Any],
    resolved: SeriesResolution,
) -> bool:
    """True when the setter may accept a bare measure value instead of record(s)."""
    if series.get("layout") != "scalar":
        return False
    if series.get("key"):
        return False
    return len(resolved["leaves"]) == 1


def _allowed_record_fields(
    series: dict[str, Any],
    *,
    allow_address: bool,
    requires_address: bool,
) -> set[str]:
    fields = {_measure_concept(series)}
    fields.update(str(c) for c in (series.get("key") or []))
    for dim in (series.get("structure") or {}).get("dimensions") or []:
        if isinstance(dim, dict) and dim.get("include_in_record", True):
            fields.add(effective_dimension_id(dim))
    for attr in (series.get("structure") or {}).get("attributes") or []:
        if isinstance(attr, dict) and attr.get("include_in_record", False):
            fields.add(effective_dimension_id(attr))
    fields.update(str(c) for c in (series.get("series_context") or {}))
    if allow_address or requires_address:
        fields.update({"address", "cell_address"})
    return {f for f in fields if f}


def _allowed_fields_literal(allowed: Sequence[str]) -> str:
    """Emit a deterministic frozenset literal for generated setter calls."""
    inner = ", ".join(repr(field) for field in allowed)
    return f"frozenset({{{inner}}})"


def _emit_python_module_body(path: Path) -> list[str]:
    """Emit top-level definitions from a module without imports or ``__all__``."""
    source = path.read_text(encoding="utf-8")
    tree = ast.parse(source)
    lines_out: list[str] = []
    for node in tree.body:
        if isinstance(node, (ast.Import, ast.ImportFrom)):
            continue
        if isinstance(node, ast.Assign) and any(
            isinstance(target, ast.Name) and target.id == "__all__" for target in node.targets
        ):
            continue
        segment = ast.get_source_segment(source, node)
        if segment:
            lines_out.extend(segment.splitlines())
            lines_out.append("")
    return lines_out


def emit_input_coerce_helpers() -> list[str]:
    """Emit inlined ``coerce_scalar`` and ``coerce_setter_input`` for generated modules."""
    package_dir = Path(__file__).resolve().parent
    lines = [
        "# --- Setter input coercion (inlined from series_bindings) ---",
        "from collections.abc import Iterable, Mapping, Sequence",
        "from datetime import date, datetime, timedelta",
        "from typing import Any, Literal, TypeGuard, cast",
        "",
        'Layout = Literal["scalar", "series", "matrix"]',
        'EmptyMeasure = Literal["skip", "write", "error"]',
        "SetterInput = object",
        "Scalar = object",
        "Record = dict[str, object]",
        "Records = list[Record]",
        "",
    ]
    lines.extend(_emit_python_module_body(package_dir / "coerce.py"))
    lines.extend(_emit_python_module_body(package_dir / "input_coerce.py"))
    return lines


SERIES_HELPERS_STDLIB_IMPORTS: tuple[str, ...] = (
    "from collections.abc import Iterable, Mapping, Sequence",
    "from datetime import date, datetime, timedelta",
    "from typing import TYPE_CHECKING, Any, Literal, TypeAlias, TypeGuard, cast",
)
"""Standard-library imports required by `emit_series_helpers_definitions`."""


def emit_series_helpers_definitions() -> list[str]:
    """Emit consolidated type aliases, coercion, and record-apply helpers.

    The returned block is the body of a dedicated helper module for the multi-module
    export. It defines the setter type aliases, the inlined `coerce_scalar` /
    `coerce_setter_input` machinery, and the shared `_apply_series_records` helper.
    Callers must provide the imports in `SERIES_HELPERS_STDLIB_IMPORTS` plus a
    `coerce_inputs_dict` binding (e.g. `from .runtime import coerce_inputs_dict`).

    Returns:
        Source lines defining the helpers, without any import statements.
    """
    package_dir = Path(__file__).resolve().parent
    lines = [
        'Layout: TypeAlias = Literal["scalar", "series", "matrix"]',
        'EmptyMeasure: TypeAlias = Literal["skip", "write", "error"]',
        "SetterInput: TypeAlias = object",
        "Scalar: TypeAlias = str | int | float | bool | datetime | None",
        "Record: TypeAlias = dict[str, object]",
        "Records: TypeAlias = list[Record]",
        "",
        "if TYPE_CHECKING:",
        "    import pandas as pd",
        "    import polars as pl",
        "",
        "    DataFrameInput: TypeAlias = pd.DataFrame | pl.DataFrame",
        "else:",
        "    DataFrameInput: TypeAlias = object",
        "",
        "SeriesInput: TypeAlias = Records | Record | Sequence[Scalar] | DataFrameInput",
        "",
    ]
    lines.extend(_emit_python_module_body(package_dir / "coerce.py"))
    lines.extend(_emit_python_module_body(package_dir / "input_coerce.py"))
    lines.extend(emit_setter_helpers())
    return lines


def _key_dtypes_for_codegen(series: dict[str, Any], key_fields: list[str]) -> dict[str, str]:
    dtypes: dict[str, str] = {}
    for dim in (series.get("structure") or {}).get("dimensions") or []:
        if not isinstance(dim, dict):
            continue
        field_name = effective_dimension_id(dim)
        if field_name not in key_fields:
            continue
        bind = dim.get("bind") if isinstance(dim.get("bind"), dict) else {}
        if "read" in bind:
            dtypes[field_name] = str(bind["read"])
        elif dim.get("dtype") is not None:
            dtypes[field_name] = str(dim["dtype"])
    return dtypes


def _canonical_key_order(
    resolved: SeriesResolution,
    key_fields: list[str],
) -> tuple[object, ...] | None:
    if len(key_fields) != 1 or resolved["requires_address"]:
        return None
    key_field = key_fields[0]
    leaves = sorted(resolved["leaves"], key=lambda leaf: leaf["key"][key_field])
    return tuple(leaf["key"][key_field] for leaf in leaves)


def _key_order_constant_name(series_id: str) -> str:
    return f"_KEY_ORDER_{series_id.upper()}"


def _emit_key_order_constant(
    resolved: SeriesResolution,
    key_fields: list[str],
) -> list[str]:
    key_order = _canonical_key_order(resolved, key_fields)
    if key_order is None:
        return []
    name = _key_order_constant_name(resolved["series_id"])
    inner = ", ".join(py_scalar_literal(value, datetime_ref="datetime") for value in key_order)
    return [f"{name} = ({inner},)" if len(key_order) == 1 else f"{name} = ({inner})", ""]


def _coerce_setter_input_call(
    series: dict[str, Any],
    resolved: SeriesResolution,
    *,
    key_fields: list[str],
    measure_concept: str,
    strict_kwarg: str,
    empty_measure_kwarg: str | None,
    requires_address: bool,
    concept_scheme: dict[str, Any] | None = None,
    emit_requires_address: bool = True,
) -> str:
    layout = str(series.get("layout") or "series")
    key_order = _canonical_key_order(resolved, key_fields)
    key_order_expr = (
        _key_order_constant_name(resolved["series_id"]) if key_order is not None else "None"
    )
    key_dtypes = _key_dtypes_for_codegen(series, key_fields)
    measure_dtype = _measure_dtype_for_codegen(series, concept_scheme=concept_scheme)
    parts = [
        "coerce_setter_input(",
        "            records,",
        f"            layout={layout!r},",
        f"            key_fields={tuple(key_fields)!r},",
        f"            measure_field={measure_concept!r},",
        f"            key_order={key_order_expr},",
        f"            strict={strict_kwarg},",
    ]
    if empty_measure_kwarg is not None:
        parts.append(f"            empty_measure={empty_measure_kwarg},")
    if emit_requires_address:
        parts.append(f"            requires_address={requires_address!r},")
    if key_dtypes:
        parts.append(f"            key_dtypes={key_dtypes!r},")
    if measure_dtype is not None:
        parts.append(f"            measure_dtype={measure_dtype!r},")
    parts.append("        )")
    return "\n".join(parts)


def emit_setter_helpers() -> list[str]:
    """Emit shared private helpers used by generated `set_*` functions."""
    return [
        "def _apply_series_records(",
        "    ctx,",
        "    records,",
        "    *,",
        "    key_fields,",
        "    allowed_fields,",
        "    measure_field,",
        "    leaf_index,",
        "    strict,",
        "    fn_name,",
        "    allow_address=False,",
        "    requires_address=False,",
        ") -> None:",
        "    updates: dict[str, object] = {}",
        "    first_record_by_address: dict[str, int] = {}",
        "    for index, record in enumerate(records):",
        "        if strict:",
        "            unknown = set(record) - allowed_fields",
        "            if unknown:",
        '                raise ValueError(f"record[{index}]: unknown fields {sorted(unknown)!r}")',
        "        if measure_field not in record:",
        '            raise ValueError(f"record[{index}]: missing required field {measure_field!r}")',
        "        address = None",
        "        key_tuple = None",
        "        if allow_address or requires_address:",
        '            address = record.get("address") or record.get("cell_address")',
        "        if requires_address and address is None:",
        "            raise ValueError(",
        '                f"record[{index}]: address required for {fn_name} (duplicate keys in binding)"',
        "            )",
        "        if address is None:",
        "            if not requires_address:",
        "                missing = [field for field in key_fields if field not in record]",
        "                if missing:",
        '                    raise ValueError(f"record[{index}]: missing key fields {missing!r}")',
        "                key_tuple = tuple((field, record[field]) for field in key_fields)",
        "                address = leaf_index.get(key_tuple)",
        "                if address is None:",
        "                    raise ValueError(",
        '                        f"record[{index}]: no leaf matches key {dict(key_tuple)!r}"',
        "                    )",
        "        elif not requires_address and all(field in record for field in key_fields):",
        "            key_tuple = tuple((field, record[field]) for field in key_fields)",
        "        assert address is not None",
        "        if address in updates:",
        "            prior = first_record_by_address[address]",
        "            if key_tuple is not None:",
        '                detail = f"duplicate key {dict(key_tuple)!r} matches record[{prior}]"',
        "            else:",
        '                detail = f"duplicate cell {address!r} matches record[{prior}]"',
        '            raise ValueError(f"record[{index}]: {detail}")',
        "        first_record_by_address[address] = index",
        "        updates[address] = record[measure_field]",
        "    if updates:",
        "        ctx.set_inputs(coerce_inputs_dict(updates))",
        "",
    ]


def emit_setter_function(
    series: dict[str, Any],
    resolved: SeriesResolution,
    *,
    graph: DependencyGraph | None = None,
    workbook: Path | str | None = None,
    bindings: WorkbookSeriesBindings | None = None,
    series_docstring_callback: SeriesBindingDocstringCallbackSpec | None = None,
    docstring_renderer: SeriesDocstringRendererSpec = "google",
) -> list[str]:
    """Emit Python source lines for one series binding setter."""
    if not resolved["leaves"]:
        raise ValueError(f"Cannot codegen setter for {resolved['series_id']!r}: no resolved leaves")
    if not resolved["ok"] and not resolved["requires_address"]:
        raise ValueError(f"Cannot codegen setter for {resolved['series_id']!r}: resolution failed")

    input_block = series.get("input") or {}
    setter = input_block.get("setter") or series.get("setter") or {}
    fn_name = str(setter.get("name", f"set_{resolved['series_id']}"))
    strict = bool(setter.get("strict", True))
    allow_address = bool(setter.get("allow_address", False))
    requires_address = bool(resolved["requires_address"])
    key_fields = [str(c) for c in (series.get("key") or [])]
    measure_concept = _measure_concept(series)
    allowed = sorted(
        _allowed_record_fields(
            series,
            allow_address=allow_address,
            requires_address=requires_address,
        )
    )
    index_name = f"_LEAF_INDEX_{resolved['series_id'].upper()}"

    lines: list[str] = []
    if not requires_address:
        lines.append(f"{index_name} = {{")
        for leaf in resolved["leaves"]:
            key_tuple_src = _key_tuple_literal(key_fields, leaf["key"])
            lines.append(f"    {key_tuple_src}: {repr(leaf['address'])},")
        lines.append("}")
        lines.append("")
    lines.extend(_emit_key_order_constant(resolved, key_fields))
    scalar_shorthand = _accepts_scalar_shorthand(series, resolved)
    input_type = "Records | Record | Scalar" if scalar_shorthand else "SeriesInput"
    lines.append(f"def {fn_name}(")
    lines.append("    ctx: EvalContext,")
    lines.append(f"    records: {input_type},")
    lines.append("    *,")
    lines.append(f"    strict: bool = {strict!r},")
    if not scalar_shorthand:
        lines.append('    empty_measure: EmptyMeasure = "write",')
    lines.append(") -> None:")
    if series_docstring_callback is not None and (
        graph is None or workbook is None or bindings is None
    ):
        raise ValueError("series_docstring_callback requires graph, workbook, and bindings context")
    if graph is not None and workbook is not None and bindings is not None:
        doc = resolve_series_function_docstring(
            graph=graph,
            workbook=workbook,
            bindings=bindings,
            series=series,
            resolution=resolved,
            function_kind="setter",
            function_name=fn_name,
            callback_spec=series_docstring_callback,
            docstring_renderer=docstring_renderer,
        )
    else:
        doc = (
            series.get("notes")
            or series.get("sdmx_notes")
            or f"Apply records for {resolved['series_id']}."
        )
    if doc is not None:
        lines.extend(emit_docstring_literal(doc))
    leaf_index_arg = "{}" if requires_address else index_name
    concept_scheme = bindings.get("concept_scheme") if bindings is not None else None
    if not isinstance(concept_scheme, dict):
        concept_scheme = None
    coerced_records = _coerce_setter_input_call(
        series,
        resolved,
        key_fields=key_fields,
        measure_concept=measure_concept,
        strict_kwarg="strict",
        empty_measure_kwarg=None if scalar_shorthand else "empty_measure",
        requires_address=requires_address,
        concept_scheme=concept_scheme,
        emit_requires_address=not scalar_shorthand,
    )
    lines.append("    _apply_series_records(")
    lines.append("        ctx,")
    lines.append(f"        {coerced_records},")
    lines.append(f"        key_fields={tuple(key_fields)!r},")
    lines.append(f"        allowed_fields={_allowed_fields_literal(allowed)},")
    lines.append(f"        measure_field={measure_concept!r},")
    lines.append(f"        leaf_index={leaf_index_arg},")
    lines.append("        strict=strict,")
    lines.append(f"        fn_name={fn_name!r},")
    lines.append(f"        allow_address={allow_address!r},")
    lines.append(f"        requires_address={requires_address!r},")
    lines.append("    )")
    lines.append("")
    return lines


def emit_setters_block(
    graph: DependencyGraph,
    workbook: Path | str,
    bindings: WorkbookSeriesBindings,
    *,
    export_addresses: Iterable[str] | None = None,
    include_type_aliases: bool = True,
    include_helpers: bool = True,
    series_docstring_callback: SeriesBindingDocstringCallbackSpec | None = None,
    docstring_renderer: SeriesDocstringRendererSpec = "google",
) -> list[str]:
    """Emit all series setter functions for a validated binding manifest.

    When `include_helpers` is true the coercion/record-apply helpers and type aliases
    are inlined alongside the setters (single-file export). When false only the setter
    functions are emitted and callers must supply the helpers separately, e.g. via a
    dedicated `_api_helpers` module in the multi-module export.
    """
    report = resolve_series_bindings(
        graph,
        bindings,
        workbook=workbook,
        direction="input",
        export_addresses=export_addresses,
    )
    lines: list[str] = ["# --- Series binding setters (Records API) ---", ""]
    include_datetime = resolutions_include_datetime(report["series"])
    if include_helpers:
        lines.extend(emit_input_coerce_helpers())
        if include_type_aliases:
            lines.extend(emit_setter_type_alias_lines(include_datetime=include_datetime))
        lines.extend(emit_setter_helpers())
    elif include_datetime:
        lines.extend(["from datetime import datetime", ""])
    by_id = {
        s["id"]: s
        for s in bindings.get("series", [])
        if isinstance(s, dict) and has_input_direction(s)
    }
    failed: list[str] = []
    for resolved in report["series"]:
        if not resolved["ok"] and not resolved["requires_address"]:
            failed.append(resolved["series_id"])
            continue
        if not resolved["leaves"]:
            warnings.warn(
                f"No resolved input cells for series {resolved['series_id']!r}; skipping setter emission",
                UserWarning,
                stacklevel=2,
            )
            continue
        warn_series_resolution_issues(resolved)
        series = by_id.get(resolved["series_id"])
        if series is None:
            continue
        lines.extend(
            emit_setter_function(
                series,
                resolved,
                graph=graph,
                workbook=workbook,
                bindings=bindings,
                series_docstring_callback=series_docstring_callback,
                docstring_renderer=docstring_renderer,
            )
        )
    if failed:
        raise ValueError(f"Cannot codegen setters: resolution failed for {failed!r}")
    return lines


def generate_setters_module(
    graph: DependencyGraph,
    workbook: Path | str,
    bindings: WorkbookSeriesBindings,
) -> str:
    """Generate a standalone module fragment with setters (requires EvalContext imports)."""
    header = [
        "from excel_grapher.runtime.cache import EvalContext, coerce_inputs_dict",
        "",
    ]
    return "\n".join(header + emit_setters_block(graph, workbook, bindings))
