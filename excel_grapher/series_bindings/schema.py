from __future__ import annotations

import json
import re
from collections.abc import Sequence
from functools import lru_cache
from importlib.resources import files
from typing import Any, cast

from jsonschema import Draft202012Validator
from jsonschema.exceptions import ValidationError

from excel_grapher.series_bindings.normalize import normalize_bindings_document
from excel_grapher.series_bindings.types import WorkbookSeriesBindings

_REQUIRED_PROPERTY_RE = re.compile(r"'([^']+)' is a required property")
_UNEXPECTED_PROPERTY_RE = re.compile(r"\('([^']+)' was unexpected\)")
_FIELD_HINTS: dict[str, str] = {
    "key": "Add a key list naming the dimension ids that identify each record.",
    "structure": "Add a structure block describing measure, dimensions, and attributes.",
    "data_range": "Add the Excel range this series covers (for example Sheet1!B3:Q3).",
    "layout": "Optional layout intent: scalar, series, or matrix.",
    "input": "Add an input block with a setter for editable series.",
    "output": "Add an output block with a compute for derived series.",
    "internal": "Add an internal block for non-I/O formula-cell key triangulation.",
}


class SeriesBindingsSchemaError(ValueError):
    """Raised when a binding manifest fails JSON Schema validation."""


def _infer_series_sheets(document: dict[str, Any]) -> None:
    """Fill `sheet` on each series when omitted and `data_range` is sheet-qualified."""
    from excel_grapher.core.address_keys import parse_address
    from excel_grapher.grapher.target_expansion import split_range_target_on_colon

    series_list = document.get("series")
    if not isinstance(series_list, list):
        return
    for series in series_list:
        if not isinstance(series, dict) or "sheet" in series:
            continue
        data_range = series.get("data_range")
        if not isinstance(data_range, str) or "!" not in data_range:
            continue
        split = split_range_target_on_colon(data_range)
        start = split[0] if split is not None else data_range
        sheet, _ = parse_address(start)
        series["sheet"] = sheet


@lru_cache(maxsize=1)
def _schema_validator() -> Any:
    schema_text = (
        files("excel_grapher.series_bindings")
        .joinpath("series_binding.schema.json")
        .read_text(encoding="utf-8")
    )
    schema = json.loads(schema_text)
    Draft202012Validator.check_schema(schema)
    return Draft202012Validator(schema)


def _series_id_at_index(document: dict[str, Any], index: int) -> str | None:
    series_list = document.get("series")
    if not isinstance(series_list, list) or not (0 <= index < len(series_list)):
        return None
    entry = series_list[index]
    if isinstance(entry, dict) and entry.get("id") is not None:
        return str(entry["id"])
    return None


def _format_schema_location(path: Sequence[str | int], *, document: dict[str, Any]) -> str:
    parts = list(path)
    if not parts:
        return "document root"

    segments: list[str] = []
    index = 0
    while index < len(parts):
        part = parts[index]
        if part == "series" and index + 1 < len(parts):
            next_part = parts[index + 1]
            if isinstance(next_part, int):
                series_id = _series_id_at_index(document, next_part)
                if series_id is not None:
                    segments.append(f'series[{next_part}] "{series_id}"')
                else:
                    segments.append(f"series[{next_part}]")
                index += 2
                continue
        if isinstance(part, int):
            segments.append(f"[{part}]")
        elif segments:
            segments.append(f".{part}")
        else:
            segments.append(str(part))
        index += 1
    return "".join(segments)


def _humanize_schema_message(error: ValidationError) -> str:
    if error.validator == "required":
        match = _REQUIRED_PROPERTY_RE.match(error.message)
        if match is not None:
            field = match.group(1)
            hint = _FIELD_HINTS.get(field)
            if hint is not None:
                return f"missing required field `{field}` — {hint}"
            return f"missing required field `{field}`"
    if error.validator == "additionalProperties":
        match = _UNEXPECTED_PROPERTY_RE.search(error.message)
        if match is not None:
            return f"unknown field `{match.group(1)}`"
    return error.message


def format_schema_error(error: ValidationError, document: dict[str, Any]) -> str:
    """Format one JSON Schema validation error for humans."""
    location = _format_schema_location(error.absolute_path, document=document)
    detail = _humanize_schema_message(error)
    return f"{location}: {detail}"


def validate_bindings_document(document: dict[str, Any]) -> WorkbookSeriesBindings:
    """Validate `document` against the series binding JSON Schema.

    Returns the document unchanged on success (typed for callers).
    Raises `SeriesBindingsSchemaError` on failure.
    """
    _infer_series_sheets(document)
    document = normalize_bindings_document(document)
    validator = _schema_validator()
    errors = sorted(validator.iter_errors(document), key=lambda e: list(e.absolute_path))
    if errors:
        raise SeriesBindingsSchemaError(format_schema_error(errors[0], document))
    return cast(WorkbookSeriesBindings, document)


def format_schema_errors(document: dict[str, Any]) -> list[str]:
    """Return human-readable schema error strings (for tests and CLI)."""
    _infer_series_sheets(document)
    document = normalize_bindings_document(document)
    validator = _schema_validator()
    return [
        format_schema_error(error, document)
        for error in sorted(validator.iter_errors(document), key=lambda e: list(e.absolute_path))
    ]
