from __future__ import annotations

import json
from functools import lru_cache
from importlib.resources import files
from typing import Any, cast

from jsonschema import Draft202012Validator

from excel_grapher.series_bindings.types import WorkbookSeriesBindings


class SeriesBindingsSchemaError(ValueError):
    """Raised when a binding manifest fails JSON Schema validation."""


def _infer_series_sheets(document: dict[str, Any]) -> None:
    """Fill ``sheet`` on each series when omitted and ``data_range`` is sheet-qualified."""
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


def validate_bindings_document(document: dict[str, Any]) -> WorkbookSeriesBindings:
    """Validate ``document`` against the series binding JSON Schema.

    Returns the document unchanged on success (typed for callers).
    Raises :class:`SeriesBindingsSchemaError` on failure.
    """
    _infer_series_sheets(document)
    validator = _schema_validator()
    errors = sorted(validator.iter_errors(document), key=lambda e: list(e.absolute_path))
    if errors:
        first = errors[0]
        path = ".".join(str(p) for p in first.absolute_path) or "<root>"
        raise SeriesBindingsSchemaError(f"{path}: {first.message}")
    return cast(WorkbookSeriesBindings, document)


def format_schema_errors(document: dict[str, Any]) -> list[str]:
    """Return human-readable schema error strings (for tests and CLI)."""
    _infer_series_sheets(document)
    validator = _schema_validator()
    messages: list[str] = []
    for error in sorted(validator.iter_errors(document), key=lambda e: list(e.absolute_path)):
        path = ".".join(str(p) for p in error.absolute_path) or "<root>"
        messages.append(f"{path}: {error.message}")
    return messages
