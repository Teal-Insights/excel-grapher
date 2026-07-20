"""Generate `set_*` / `read_*` functions that write and read graph input leaves."""

from __future__ import annotations

import ast
import re
import warnings
from collections.abc import Iterable, Mapping, Sequence
from pathlib import Path
from typing import TYPE_CHECKING, Any

from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.series_bindings.codegen_literals import (
    emit_setter_type_alias_lines,
    py_scalar_literal,
    python_annotation_for_dtype,
    resolutions_include_datetime,
    setter_input_annotation,
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
    ResolutionReport,
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


def _leaf_index_name(series_id: str) -> str:
    return f"_LEAF_INDEX_{series_id.upper()}"


def emit_leaf_index_lines(series: dict[str, Any], resolved: SeriesResolution) -> list[str]:
    """Emit the `_LEAF_INDEX_<ID>` constant shared by a setter/reader dual."""
    if resolved["requires_address"]:
        return []
    key_fields = [str(c) for c in (series.get("key") or [])]
    index_name = _leaf_index_name(resolved["series_id"])
    lines = [f"{index_name} = {{"]
    for leaf in resolved["leaves"]:
        key_tuple_src = _key_tuple_literal(key_fields, leaf["key"])
        lines.append(f"    {key_tuple_src}: {repr(leaf['address'])},")
    lines.append("}")
    lines.append("")
    return lines


def dimension_id_to_param_name(field: str) -> str:
    """Convert an SDMX / effective dimension id to a Python keyword parameter name.

    Examples:
        `TIME_PERIOD` -> `time_period`
        `ref_area` -> `ref_area`
    """
    slug = re.sub(r"[^a-zA-Z0-9]+", "_", field).strip("_").lower()
    if not slug:
        slug = "key"
    if slug[0].isdigit():
        slug = f"dim_{slug}"
    return slug


def _reader_function_name(series: dict[str, Any], resolved: SeriesResolution) -> str:
    input_block = series.get("input") or {}
    reader = input_block.get("reader")
    if isinstance(reader, dict) and reader.get("name"):
        return str(reader["name"])
    return f"read_{resolved['series_id']}"


def _should_emit_reader(resolved: SeriesResolution) -> bool:
    """Emit a reader whenever the setter has at least one resolved leaf."""
    return bool(resolved["leaves"])


def _should_emit_reader_range(series: dict[str, Any], resolved: SeriesResolution) -> bool:
    """Emit a range reader when the series data_range is multi-cell / non-scalar."""
    if not _should_emit_reader(resolved):
        return False
    if series.get("layout") == "scalar":
        return False
    return len(resolved["leaves"]) > 1 or ":" in str(series.get("data_range") or "")


def _reader_addresses_name(series_id: str) -> str:
    return f"_READER_ADDRESSES_{series_id.upper()}"


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
    include_leaf_index: bool = True,
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
    index_name = _leaf_index_name(resolved["series_id"])

    lines: list[str] = []
    if include_leaf_index and not requires_address:
        lines.extend(emit_leaf_index_lines(series, resolved))
    lines.extend(_emit_key_order_constant(resolved, key_fields))
    scalar_shorthand = _accepts_scalar_shorthand(series, resolved)
    concept_scheme = bindings.get("concept_scheme") if bindings is not None else None
    if not isinstance(concept_scheme, dict):
        concept_scheme = None
    measure_dtype = _measure_dtype_for_codegen(series, concept_scheme=concept_scheme)
    input_type = setter_input_annotation(
        layout=str(series.get("layout") or "series"),
        measure_dtype=measure_dtype,
        scalar_shorthand=scalar_shorthand,
    )
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


def emit_reader_function(
    series: dict[str, Any],
    resolved: SeriesResolution,
    *,
    graph: DependencyGraph | None = None,
    workbook: Path | str | None = None,
    bindings: WorkbookSeriesBindings | None = None,
    series_docstring_callback: SeriesBindingDocstringCallbackSpec | None = None,
    docstring_renderer: SeriesDocstringRendererSpec = "google",
) -> list[str]:
    """Emit a `read_*` dual that resolves domain keys via `_LEAF_INDEX_*`.

    For uniquely keyed series, expects the matching `_LEAF_INDEX_<ID>` constant to
    already be present (emitted by `emit_leaf_index_lines` / `emit_readers_block`, or
    co-emitted by `emit_setter_function` in single-file mode). For duplicate-key
    bindings (`requires_address`), emits an address-keyed reader validated against the
    resolved leaf addresses — the inverse of the address-required setter path.
    """
    if not _should_emit_reader(resolved):
        return []

    key_fields = [str(c) for c in (series.get("key") or [])]
    fn_name = _reader_function_name(series, resolved)
    index_name = _leaf_index_name(resolved["series_id"])
    # Readers live in `_readers.py` (modular) or alongside the runtime embed
    # (single-file), both of which expose `CellValue`.
    return_annotation = "CellValue"
    key_dtypes = _key_dtypes_for_codegen(series, key_fields)
    requires_address = bool(resolved["requires_address"])

    lines: list[str] = []
    if requires_address:
        addresses_name = _reader_addresses_name(resolved["series_id"])
        addr_inner = ", ".join(repr(leaf["address"]) for leaf in resolved["leaves"])
        lines.append(f"{addresses_name} = frozenset({{{addr_inner}}})")
        lines.append("")
        lines.append(f"def {fn_name}(ctx: EvalContext, *, address: str) -> {return_annotation}:")
    elif key_fields:
        lines.append(f"def {fn_name}(")
        lines.append("    ctx: EvalContext,")
        lines.append("    *,")
        for field in key_fields:
            param = dimension_id_to_param_name(field)
            annotation = python_annotation_for_dtype(key_dtypes.get(field)) or "object"
            lines.append(f"    {param}: {annotation},")
        lines.append(f") -> {return_annotation}:")
    else:
        lines.append(f"def {fn_name}(ctx: EvalContext) -> {return_annotation}:")

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
            function_kind="reader",
            function_name=fn_name,
            callback_spec=series_docstring_callback,
            docstring_renderer=docstring_renderer,
        )
    else:
        doc = (
            series.get("notes")
            or series.get("sdmx_notes")
            or f"Read the value for {resolved['series_id']}."
        )
    if doc is not None:
        lines.extend(emit_docstring_literal(doc))

    if requires_address:
        addresses_name = _reader_addresses_name(resolved["series_id"])
        lines.append(f"    if address not in {addresses_name}:")
        lines.append(
            "        raise ValueError("
            'f"address {address!r} is not a leaf of '
            f"{resolved['series_id']!r}"
            '")'
        )
        lines.append("    return xl_cell(ctx, address)")
    elif key_fields:
        param_pairs = ", ".join(
            f"({repr(field)}, {dimension_id_to_param_name(field)})" for field in key_fields
        )
        key_tuple_expr = f"({param_pairs},)" if len(key_fields) == 1 else f"({param_pairs})"
        lines.append(f"    key_tuple = {key_tuple_expr}")
        lines.append(f"    address = {index_name}.get(key_tuple)")
        lines.append("    if address is None:")
        lines.append('        raise ValueError(f"no leaf matches key {dict(key_tuple)!r}")')
        lines.append("    return xl_cell(ctx, address)")
    else:
        # Keyless scalar: leaf index maps () -> address.
        single_address = resolved["leaves"][0]["address"]
        lines.append(f"    return xl_cell(ctx, {single_address!r})")
    lines.append("")
    return lines


def emit_reader_range_function(
    series: dict[str, Any],
    resolved: SeriesResolution,
    *,
    graph: DependencyGraph | None = None,
    workbook: Path | str | None = None,
    bindings: WorkbookSeriesBindings | None = None,
    series_docstring_callback: SeriesBindingDocstringCallbackSpec | None = None,
    docstring_renderer: SeriesDocstringRendererSpec = "google",
) -> list[str]:
    """Emit `read_<id>_range` for a binding-aligned multi-cell `data_range`.

    Returns an empty list for scalar / single-leaf series without a multi-cell
    `data_range`.
    """
    if not _should_emit_reader_range(series, resolved):
        return []

    data_range = series.get("data_range")
    if not isinstance(data_range, str) or not data_range:
        return []

    reader_name = _reader_function_name(series, resolved)
    fn_name = f"{reader_name}_range"
    lines: list[str] = [f"def {fn_name}(ctx: EvalContext) -> CellValue:"]
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
            function_kind="reader",
            function_name=fn_name,
            callback_spec=series_docstring_callback,
            docstring_renderer=docstring_renderer,
        )
    else:
        doc = f"Read the binding-aligned range for {resolved['series_id']}."
    if doc is not None:
        lines.extend(emit_docstring_literal(doc))
    lines.append(f"    return xl_range(ctx, {data_range!r})")
    lines.append("")
    return lines


def _iter_input_resolutions(
    graph: DependencyGraph,
    workbook: Path | str,
    bindings: WorkbookSeriesBindings,
    *,
    export_addresses: Iterable[str] | None = None,
) -> tuple[list[tuple[dict[str, Any], SeriesResolution]], list[str], ResolutionReport]:
    """Resolve input series and return (pairs, failed ids, report)."""
    report = resolve_series_bindings(
        graph,
        bindings,
        workbook=workbook,
        direction="input",
        export_addresses=export_addresses,
    )
    by_id = {
        s["id"]: s
        for s in bindings.get("series", [])
        if isinstance(s, dict) and has_input_direction(s)
    }
    pairs: list[tuple[dict[str, Any], SeriesResolution]] = []
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
        pairs.append((series, resolved))
    return pairs, failed, report


def emit_readers_block(
    graph: DependencyGraph,
    workbook: Path | str,
    bindings: WorkbookSeriesBindings,
    *,
    export_addresses: Iterable[str] | None = None,
    series_docstring_callback: SeriesBindingDocstringCallbackSpec | None = None,
    docstring_renderer: SeriesDocstringRendererSpec = "google",
    include_leaf_indexes: bool = True,
) -> list[str]:
    """Emit `_LEAF_INDEX_*` maps and `read_*` / `read_*_range` functions.

    Used by the modular export's private `_readers.py` module so formula bodies
    can call readers without importing the public `api` module.
    """
    pairs, failed, report = _iter_input_resolutions(
        graph,
        workbook,
        bindings,
        export_addresses=export_addresses,
    )
    lines: list[str] = ["# --- Series binding readers ---", ""]
    concept_scheme = bindings.get("concept_scheme")
    if not isinstance(concept_scheme, dict):
        concept_scheme = None
    include_datetime = resolutions_include_datetime(report["series"]) or any(
        _measure_dtype_for_codegen(series, concept_scheme=concept_scheme) == "datetime"
        for series in bindings.get("series", [])
        if isinstance(series, dict) and has_input_direction(series)
    )
    if include_datetime:
        lines.extend(["from datetime import datetime", ""])
    for series, resolved in pairs:
        if include_leaf_indexes:
            lines.extend(emit_leaf_index_lines(series, resolved))
        lines.extend(
            emit_reader_function(
                series,
                resolved,
                graph=graph,
                workbook=workbook,
                bindings=bindings,
                series_docstring_callback=series_docstring_callback,
                docstring_renderer=docstring_renderer,
            )
        )
        lines.extend(
            emit_reader_range_function(
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
        raise ValueError(f"Cannot codegen readers: resolution failed for {failed!r}")
    return lines


def emit_setters_block(
    graph: DependencyGraph,
    workbook: Path | str,
    bindings: WorkbookSeriesBindings,
    *,
    export_addresses: Iterable[str] | None = None,
    include_type_aliases: bool = True,
    include_helpers: bool = True,
    include_readers: bool = True,
    include_leaf_indexes: bool = True,
    series_docstring_callback: SeriesBindingDocstringCallbackSpec | None = None,
    docstring_renderer: SeriesDocstringRendererSpec = "google",
) -> list[str]:
    """Emit all series setter and reader functions for a validated binding manifest.

    When `include_helpers` is true the coercion/record-apply helpers and type aliases
    are inlined alongside the setters (single-file export). When false only the setter
    and reader functions are emitted and callers must supply the helpers separately, e.g. via a
    dedicated `_api_helpers` module in the multi-module export.

    When `include_readers` / `include_leaf_indexes` are false, readers and leaf maps
    are omitted so a separate `_readers` module can own them.
    """
    pairs, failed, report = _iter_input_resolutions(
        graph,
        workbook,
        bindings,
        export_addresses=export_addresses,
    )
    lines: list[str] = ["# --- Series binding setters (Records API) ---", ""]
    concept_scheme = bindings.get("concept_scheme")
    if not isinstance(concept_scheme, dict):
        concept_scheme = None
    include_datetime = resolutions_include_datetime(report["series"]) or any(
        _measure_dtype_for_codegen(series, concept_scheme=concept_scheme) == "datetime"
        for series in bindings.get("series", [])
        if isinstance(series, dict) and has_input_direction(series)
    )
    if include_helpers:
        lines.extend(emit_input_coerce_helpers())
        if include_type_aliases:
            lines.extend(emit_setter_type_alias_lines(include_datetime=include_datetime))
        lines.extend(emit_setter_helpers())
    elif include_datetime:
        lines.extend(["from datetime import datetime", ""])
    for series, resolved in pairs:
        lines.extend(
            emit_setter_function(
                series,
                resolved,
                graph=graph,
                workbook=workbook,
                bindings=bindings,
                series_docstring_callback=series_docstring_callback,
                docstring_renderer=docstring_renderer,
                include_leaf_index=include_leaf_indexes,
            )
        )
        if include_readers:
            lines.extend(
                emit_reader_function(
                    series,
                    resolved,
                    graph=graph,
                    workbook=workbook,
                    bindings=bindings,
                    series_docstring_callback=series_docstring_callback,
                    docstring_renderer=docstring_renderer,
                )
            )
            lines.extend(
                emit_reader_range_function(
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
