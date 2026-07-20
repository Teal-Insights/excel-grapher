"""Pluggable docstring renderers for generated series-binding API functions."""

from __future__ import annotations

from collections.abc import Mapping
from typing import Any, Literal, Protocol, TypeAlias, cast, runtime_checkable

from excel_grapher.series_bindings.codegen_literals import setter_input_annotation
from excel_grapher.series_bindings.docstrings import (
    FieldContract,
    SeriesBindingDocstringContract,
    SeriesFunctionDoc,
)

SeriesDocstringRendererName = Literal["plain", "rst", "google", "numpy"]


@runtime_checkable
class SeriesDocstringRenderer(Protocol):
    def render(
        self,
        contract: SeriesBindingDocstringContract,
        doc: SeriesFunctionDoc,
        *,
        series: Mapping[str, Any] | None = None,
    ) -> str: ...


@runtime_checkable
class SeriesDocstringRenderCallable(Protocol):
    def __call__(
        self,
        contract: SeriesBindingDocstringContract,
        doc: SeriesFunctionDoc,
        *,
        series: Mapping[str, Any] | None = None,
    ) -> str: ...


SeriesDocstringRendererSpec: TypeAlias = (
    SeriesDocstringRendererName | SeriesDocstringRenderer | SeriesDocstringRenderCallable
)


def _field_description_line(
    doc: SeriesFunctionDoc,
    field_name: str,
    field_contract: FieldContract,
) -> str:
    field_doc = doc.field_descriptions.get(field_name)
    description = field_doc.description if field_doc is not None else field_contract.concept_name
    expected_value = field_contract.expected_value
    if expected_value is None:
        return f"{field_name}: {description}"
    return f'{field_name}: {description} If supplied, expected value: "{expected_value}".'


def _indent_lines(lines: list[str], prefix: str) -> list[str]:
    """Indent non-empty lines, leaving blank lines free of trailing whitespace."""
    return [f"{prefix}{line}" if line else "" for line in lines]


def _render_record(record: Mapping[str, object]) -> str:
    fields = ", ".join(f"{field_name!r}: {value!r}" for field_name, value in record.items())
    return f"{{{fields}}}"


def _measure_field(contract: SeriesBindingDocstringContract) -> str | None:
    """Return the measure field name (the last required field), if any."""
    if contract.required_fields:
        return contract.required_fields[-1]
    return None


def _is_single_key(contract: SeriesBindingDocstringContract) -> bool:
    """True when the binding has exactly one key field (one key plus the measure)."""
    return len(contract.required_fields) == 2


def _setter_input_type_name(contract: SeriesBindingDocstringContract) -> str:
    """Return the input type text advertised in generated setter docstrings."""
    scalar_shorthand = contract.layout == "scalar" and len(contract.required_fields) <= 1
    return setter_input_annotation(
        layout=contract.layout,
        measure_dtype=contract.value_type,
        scalar_shorthand=scalar_shorthand,
    )


def _setter_input_description(contract: SeriesBindingDocstringContract) -> str:
    """Describe the input shapes accepted for the binding's layout."""
    if contract.layout == "scalar":
        return "A bare scalar value, a single record dict, or a list of records."
    if contract.layout == "series" and _is_single_key(contract):
        return (
            "A list of records, a single record dict, a tidy pandas/polars "
            "DataFrame, or a 1-D iterable of measure values in key order."
        )
    return "A list of records, a single record dict, or a tidy pandas/polars DataFrame."


def _setter_supports_empty_measure(contract: SeriesBindingDocstringContract) -> bool:
    return contract.function_kind == "setter" and contract.layout != "scalar"


def _empty_measure_kwarg_description() -> str:
    return (
        "How to treat rows with missing measure values (`None` or float NaN after "
        'DataFrame coercion). "write" (default) passes values through; "skip" drops '
        'them; "error" raises. Empty key fields always raise.'
    )


def _render_records_example(contract: SeriesBindingDocstringContract) -> list[str]:
    lines = [f"{contract.function_name}(ctx, ["]
    for record in contract.example_records:
        lines.append(f"    {_render_record(record)},")
    lines.append("])")
    return lines


def _render_scalar_example(contract: SeriesBindingDocstringContract) -> list[str]:
    measure = _measure_field(contract)
    value = contract.example_records[0].get(measure) if contract.example_records else None
    return [f"{contract.function_name}(ctx, {value!r})"]


def _render_positional_example(contract: SeriesBindingDocstringContract) -> list[str]:
    """Render a positional measure example only when full key arity is known."""
    values = contract.positional_example_values
    if not values:
        return []
    return [f"{contract.function_name}(ctx, {list(values)!r})"]


def _render_example_call(contract: SeriesBindingDocstringContract) -> list[str]:
    if contract.function_kind != "setter":
        return [f"{contract.function_name}(ctx=ctx)"]
    if contract.layout == "scalar":
        return _render_scalar_example(contract)
    blocks = [_render_records_example(contract)]
    if contract.layout == "series" and _is_single_key(contract):
        positional = _render_positional_example(contract)
        if positional:
            blocks.append(positional)
    lines: list[str] = []
    for index, block in enumerate(blocks):
        if index:
            lines.append("")
        lines.extend(block)
    return lines


def _required_and_optional_fields(
    contract: SeriesBindingDocstringContract,
) -> tuple[list[str], list[str]]:
    required = [name for name, field in contract.fields.items() if field.required]
    optional = [name for name, field in contract.fields.items() if not field.required]
    return required, optional


def _source_binding_lines(
    contract: SeriesBindingDocstringContract,
    series: Mapping[str, Any] | None,
) -> list[str]:
    data_range = series.get("data_range", contract.data_range) if series else contract.data_range
    layout = series.get("layout", contract.layout) if series else contract.layout
    value_type = contract.value_type if contract.value_type is not None else "unspecified"
    return [
        f"Workbook range: {data_range}",
        f"Layout: {layout}",
        f"Value type: {value_type}",
    ]


def _append_intro(lines: list[str], doc: SeriesFunctionDoc) -> None:
    lines.append(doc.summary)
    lines.append("")
    if doc.purpose:
        lines.append(doc.purpose)
    if doc.record_matching:
        lines.append(doc.record_matching)
    if doc.purpose or doc.record_matching:
        lines.append("")


def _append_field_bullets(
    lines: list[str],
    *,
    doc: SeriesFunctionDoc,
    contract: SeriesBindingDocstringContract,
    required: list[str],
    optional: list[str],
    indent: str = "",
) -> None:
    lines.append(f"{indent}Required record fields:")
    for name in required:
        lines.append(f"{indent}    - {_field_description_line(doc, name, contract.fields[name])}")
    if optional:
        lines.append(f"{indent}Optional record fields:")
        for name in optional:
            lines.append(
                f"{indent}    - {_field_description_line(doc, name, contract.fields[name])}"
            )


class PlainSeriesDocstringRenderer:
    def render(
        self,
        contract: SeriesBindingDocstringContract,
        doc: SeriesFunctionDoc,
        *,
        series: Mapping[str, Any] | None = None,
    ) -> str:
        required, optional = _required_and_optional_fields(contract)
        example_lines = _render_example_call(contract)

        lines: list[str] = []
        _append_intro(lines, doc)
        if contract.function_kind == "setter":
            lines.append("Accepted input:")
            lines.append(f"    {_setter_input_description(contract)}")
            lines.append("")
            if _setter_supports_empty_measure(contract):
                lines.append("Keyword arguments:")
                lines.append(f"    empty_measure: {_empty_measure_kwarg_description()}")
                lines.append("")
        lines.append("Required record fields:")
        lines.extend(
            [
                f"    {_field_description_line(doc, name, contract.fields[name])}"
                for name in required
            ]
        )
        if optional:
            lines.append("")
            lines.append("Optional record fields:")
            lines.extend(
                [
                    f"    {_field_description_line(doc, name, contract.fields[name])}"
                    for name in optional
                ]
            )
        lines.extend(
            [
                "",
                "Source binding:",
                *[f"    {line}" for line in _source_binding_lines(contract, series)],
                "",
                "Example:",
                *_indent_lines(example_lines, "    "),
            ]
        )
        return "\n".join(lines)


class RstSeriesDocstringRenderer:
    def render(
        self,
        contract: SeriesBindingDocstringContract,
        doc: SeriesFunctionDoc,
        *,
        series: Mapping[str, Any] | None = None,
    ) -> str:
        required, optional = _required_and_optional_fields(contract)
        example_lines = _render_example_call(contract)

        lines = [doc.summary, ""]
        if doc.purpose:
            lines.extend(["Purpose", "-------", doc.purpose, ""])
        if doc.record_matching:
            lines.extend(["Record matching", "---------------", doc.record_matching, ""])
        if contract.function_kind == "setter":
            lines.extend(
                [
                    "Accepted input",
                    "--------------",
                    _setter_input_description(contract),
                    "",
                ]
            )
            if _setter_supports_empty_measure(contract):
                lines.extend(
                    [
                        "Keyword arguments",
                        "-------------------",
                        f"empty_measure: {_empty_measure_kwarg_description()}",
                        "",
                    ]
                )
        lines.append("Required record fields")
        lines.append("----------------------")
        for name in required:
            field_line = _field_description_line(doc, name, contract.fields[name])
            field_name, _, description = field_line.partition(": ")
            lines.append(f":{field_name}: {description}")
        if optional:
            lines.extend(["", "Optional record fields", "----------------------"])
            for name in optional:
                field_line = _field_description_line(doc, name, contract.fields[name])
                field_name, _, description = field_line.partition(": ")
                lines.append(f":{field_name}: {description}")
        lines.extend(["", "Source binding", "--------------"])
        lines.extend(_source_binding_lines(contract, series))
        lines.extend(["", "Example", "-------"])
        lines.extend(example_lines)
        return "\n".join(lines)


class GoogleSeriesDocstringRenderer:
    def render(
        self,
        contract: SeriesBindingDocstringContract,
        doc: SeriesFunctionDoc,
        *,
        series: Mapping[str, Any] | None = None,
    ) -> str:
        required, optional = _required_and_optional_fields(contract)
        example_lines = _render_example_call(contract)

        lines: list[str] = []
        _append_intro(lines, doc)
        lines.append("Args:")
        if contract.function_kind == "setter":
            lines.append(
                f"    records ({_setter_input_type_name(contract)}): "
                f"{_setter_input_description(contract)}"
            )
            if _setter_supports_empty_measure(contract):
                lines.append(
                    f"    empty_measure (EmptyMeasure): {_empty_measure_kwarg_description()}"
                )
            _append_field_bullets(
                lines,
                doc=doc,
                contract=contract,
                required=required,
                optional=optional,
                indent="        ",
            )
        elif contract.function_kind == "reader":
            lines.append("    ctx (EvalContext): Evaluation context.")
        else:
            lines.append("    ctx (EvalContext | None): Existing evaluation context, if available.")
            lines.append(
                "    inputs (dict[str, object] | None): Optional input map when ctx is omitted."
            )
        lines.append("")
        lines.append("Returns:")
        if contract.function_kind == "setter":
            lines.append("    None: Applies the input updates to ctx.")
        elif contract.function_kind == "reader":
            lines.append("    CellValue: Value read from the bound cell or range.")
        else:
            lines.append("    Records: Computed output records.")
            _append_field_bullets(
                lines,
                doc=doc,
                contract=contract,
                required=required,
                optional=optional,
                indent="        ",
            )
        lines.extend(["", "Source binding:"])
        for line in _source_binding_lines(contract, series):
            lines.append(f"    {line}")
        lines.extend(["", "Examples:"])
        lines.extend(_indent_lines(example_lines, "    "))
        return "\n".join(lines)


class NumpySeriesDocstringRenderer:
    def render(
        self,
        contract: SeriesBindingDocstringContract,
        doc: SeriesFunctionDoc,
        *,
        series: Mapping[str, Any] | None = None,
    ) -> str:
        required, optional = _required_and_optional_fields(contract)
        example_lines = _render_example_call(contract)

        lines = [doc.summary, ""]
        if doc.purpose:
            lines.extend([doc.purpose, ""])
        if doc.record_matching:
            lines.extend([doc.record_matching, ""])
        lines.extend(["Parameters", "----------"])
        if contract.function_kind == "setter":
            lines.append(f"records : {_setter_input_type_name(contract)}")
            lines.append(f"    {_setter_input_description(contract)}")
            if _setter_supports_empty_measure(contract):
                lines.append('empty_measure : EmptyMeasure, default "write"')
                lines.append(f"    {_empty_measure_kwarg_description()}")
            _append_field_bullets(
                lines,
                doc=doc,
                contract=contract,
                required=required,
                optional=optional,
                indent="    ",
            )
        elif contract.function_kind == "reader":
            lines.append("ctx : EvalContext")
            lines.append("    Evaluation context.")
        else:
            lines.append("ctx : EvalContext | None, optional")
            lines.append("    Existing evaluation context.")
            lines.append("inputs : dict[str, object] | None, optional")
            lines.append("    Optional input map when ctx is omitted.")
        lines.extend(["", "Returns", "-------"])
        if contract.function_kind == "setter":
            lines.append("None")
            lines.append("    Applies the input updates to ctx.")
        elif contract.function_kind == "reader":
            lines.append("CellValue")
            lines.append("    Value read from the bound cell or range.")
        else:
            lines.append("Records")
            lines.append("    Computed output records.")
            _append_field_bullets(
                lines,
                doc=doc,
                contract=contract,
                required=required,
                optional=optional,
                indent="    ",
            )
        lines.extend(["", "Source binding", "--------------"])
        lines.extend(_source_binding_lines(contract, series))
        lines.extend(["", "Examples", "--------"])
        lines.extend(example_lines)
        return "\n".join(lines)


class _CallableSeriesDocstringRendererAdapter:
    def __init__(self, func: SeriesDocstringRenderCallable) -> None:
        self._func = func

    def render(
        self,
        contract: SeriesBindingDocstringContract,
        doc: SeriesFunctionDoc,
        *,
        series: Mapping[str, Any] | None = None,
    ) -> str:
        return self._func(contract, doc, series=series)


_BUILTIN_RENDERERS: dict[SeriesDocstringRendererName, SeriesDocstringRenderer] = {
    "plain": PlainSeriesDocstringRenderer(),
    "rst": RstSeriesDocstringRenderer(),
    "google": GoogleSeriesDocstringRenderer(),
    "numpy": NumpySeriesDocstringRenderer(),
}


def resolve_series_docstring_renderer(
    renderer: SeriesDocstringRendererSpec,
) -> SeriesDocstringRenderer:
    if isinstance(renderer, str):
        if renderer not in _BUILTIN_RENDERERS:
            known = ", ".join(sorted(_BUILTIN_RENDERERS))
            raise ValueError(f"Unknown series docstring renderer: {renderer!r}. Known: {known}")
        renderer_name = cast(SeriesDocstringRendererName, renderer)
        return _BUILTIN_RENDERERS[renderer_name]
    if isinstance(renderer, SeriesDocstringRenderer):
        return renderer
    if callable(renderer):
        return _CallableSeriesDocstringRendererAdapter(renderer)
    raise TypeError(
        "Custom renderer must be a renderer object with .render(...) or a callable "
        "(contract, doc, *, series=None) -> str."
    )


__all__ = [
    "GoogleSeriesDocstringRenderer",
    "NumpySeriesDocstringRenderer",
    "PlainSeriesDocstringRenderer",
    "RstSeriesDocstringRenderer",
    "SeriesDocstringRenderCallable",
    "SeriesDocstringRenderer",
    "SeriesDocstringRendererName",
    "SeriesDocstringRendererSpec",
    "resolve_series_docstring_renderer",
]
