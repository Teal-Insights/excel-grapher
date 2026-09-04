"""Coerce setter caller input into canonical record lists for series bindings."""

from __future__ import annotations

from collections.abc import Iterable, Mapping, Sequence
from typing import Any, TypeGuard, cast

from excel_grapher.series_bindings.coerce import coerce_scalar, validate_binding_scalar
from excel_grapher.series_bindings.setter_input_types import EmptyMeasure, Layout, SetterInput
from excel_grapher.series_bindings.types import Record, Records

__all__ = [
    "EmptyMeasure",
    "Layout",
    "coerce_setter_input",
    "measure_domain_from_series",
    "require_input_domain",
]


def _is_mapping(value: object) -> TypeGuard[Mapping[str, object]]:
    return isinstance(value, Mapping) and not isinstance(value, (str, bytes, bytearray))


def _is_pandas_dataframe(data: object) -> bool:
    cls = type(data)
    module = cls.__module__
    return cls.__name__ == "DataFrame" and (module == "pandas" or module.startswith("pandas."))


def _is_polars_dataframe(data: object) -> bool:
    cls = type(data)
    return cls.__module__.startswith("polars.") and cls.__name__ == "DataFrame"


def _is_tabular_dataframe(data: object) -> bool:
    return _is_pandas_dataframe(data) or _is_polars_dataframe(data)


def _import_pandas() -> Any:
    try:
        import pandas as pd
    except ImportError as exc:
        raise ImportError(
            "DataFrame input requires pandas; install it or pass records / a 1D iterable"
        ) from exc
    return pd


def _import_polars() -> Any:
    try:
        import polars as pl
    except ImportError as exc:
        raise ImportError(
            "DataFrame input requires polars; install it or pass records / a 1D iterable"
        ) from exc
    return pl


def _coerce_scalar_records(
    data: object,
    measure_field: str,
) -> Records:
    """Normalize scalar-layout setter input to a record list."""
    if isinstance(data, list):
        return cast(Records, data)
    if _is_mapping(data):
        return [dict(data)]
    return [{measure_field: data}]


def _is_records_list(data: Sequence[object]) -> TypeGuard[Records]:
    if not data:
        return True
    return all(_is_mapping(item) for item in data)


def _coerce_key_value(
    field: str,
    raw: object,
    key_dtypes: Mapping[str, str] | None,
) -> object:
    if key_dtypes is None:
        return raw
    read_as = key_dtypes.get(field)
    if read_as is None:
        return raw
    return coerce_scalar(raw, read_as)


def _dataframe_column_names(data: object) -> list[str]:
    columns = getattr(data, "columns", None)
    if columns is None:
        raise TypeError(f"unsupported DataFrame-like input: {type(data)!r}")
    return [str(column) for column in columns]


def _is_missing_value(value: object) -> bool:
    """Return whether a normalized cell value counts as missing for empty-measure policy.

    Treats ``None`` and float NaN as missing. Tabular inputs are normalized via
    pandas/polars record conversion before this check runs.
    """
    return value is None or (isinstance(value, float) and value != value)


def _validate_nonempty_key_fields(
    records: Records,
    *,
    key_fields: tuple[str, ...],
) -> None:
    for index, record in enumerate(records):
        for field in key_fields:
            if field in record and _is_missing_value(record[field]):
                raise ValueError(f"record[{index}]: empty key field {field!r}")


def _validate_dataframe_columns(
    column_names: list[str],
    *,
    key_fields: tuple[str, ...],
    measure_field: str,
    strict: bool,
) -> None:
    required = set(key_fields) | {measure_field}
    present = set(column_names)
    missing = sorted(required - present)
    if missing:
        msg = f"missing required column(s): {missing!r}"
        extra = sorted(present - set(key_fields) - {measure_field})
        if measure_field in missing and extra:
            tidy_columns = ", ".join([*key_fields, measure_field])
            msg += (
                f"; input looks wide (extra columns {extra!r}) — "
                f"melt or stack to tidy with columns {tidy_columns!r}"
            )
        raise ValueError(msg)
    if strict:
        unknown = sorted(present - required)
        if unknown:
            raise ValueError(f"unknown columns {unknown!r}")


def _apply_empty_measure(
    records: Records,
    *,
    key_fields: tuple[str, ...],
    measure_field: str,
    empty_measure: EmptyMeasure,
) -> Records:
    """Apply empty key/measure policy after input normalization."""
    _validate_nonempty_key_fields(records, key_fields=key_fields)
    if empty_measure == "write":
        return records

    kept: list[dict[str, object]] = []
    for index, record in enumerate(records):
        if measure_field not in record:
            if empty_measure == "error":
                raise ValueError(f"record[{index}]: missing required field {measure_field!r}")
            continue
        if _is_missing_value(record[measure_field]):
            if empty_measure == "error":
                raise ValueError(f"record[{index}]: empty measure field {measure_field!r}")
            continue
        kept.append(record)
    return kept


def _row_dicts_from_dataframe(data: object) -> list[Record]:
    if _is_pandas_dataframe(data):
        pd = _import_pandas()
        if not isinstance(data, pd.DataFrame):
            raise ImportError(
                "DataFrame input requires pandas; install it or pass records / a 1D iterable"
            )
        return data.to_dict(orient="records")
    if _is_polars_dataframe(data):
        pl = _import_polars()
        if not isinstance(data, pl.DataFrame):
            raise ImportError(
                "DataFrame input requires polars; install it or pass records / a 1D iterable"
            )
        return data.to_dicts()
    raise TypeError(f"unsupported DataFrame-like input: {type(data)!r}")


def _apply_key_dtypes(
    records: Records,
    *,
    key_fields: tuple[str, ...],
    key_dtypes: Mapping[str, str] | None,
) -> Records:
    """Coerce key field values on each record using binding read modes."""
    if not key_dtypes:
        return records
    coerced: list[dict[str, object]] = []
    for record in records:
        updated = dict(record)
        for field in key_fields:
            if field in updated:
                updated[field] = _coerce_key_value(field, updated[field], key_dtypes)
        coerced.append(updated)
    return coerced


def _apply_measure_dtype(
    records: Records,
    *,
    measure_field: str,
    measure_dtype: str | None,
) -> Records:
    """Validate and coerce measure values against the binding measure dtype."""
    if measure_dtype is None:
        return records
    validated: list[dict[str, object]] = []
    for index, record in enumerate(records):
        if measure_field not in record:
            validated.append(record)
            continue
        raw = record[measure_field]
        try:
            value = validate_binding_scalar(raw, measure_dtype)
        except TypeError as exc:
            raise TypeError(
                f"record[{index}]: {measure_field} must be {measure_dtype}, "
                f"got {type(raw).__name__}: {raw!r}"
            ) from exc
        except ValueError as exc:
            raise ValueError(f"record[{index}]: {measure_field}: {exc}") from exc
        if value is raw:
            validated.append(record)
            continue
        updated = dict(record)
        updated[measure_field] = value
        validated.append(updated)
    return validated


def _format_measure_domain(domain: Mapping[str, Any]) -> str:
    """Render a measure domain for error messages."""
    if "enum" in domain:
        values = domain["enum"]
        rendered = ", ".join(repr(value) for value in sorted(values, key=repr))
        return f"{{{rendered}}}"
    if "between" in domain:
        bounds = domain["between"]
        return f"between(min={bounds.get('min')!r}, max={bounds.get('max')!r})"
    if "real_between" in domain:
        bounds = domain["real_between"]
        return f"real_between(min={bounds.get('min')!r}, max={bounds.get('max')!r})"
    return repr(dict(domain))


def _in_closed_bounds(value: int | float, bounds: Mapping[str, Any]) -> bool:
    """Return whether `value` lies in an inclusive min/max interval."""
    lo = bounds.get("min")
    hi = bounds.get("max")
    return (lo is None or value >= lo) and (hi is None or value <= hi)


def _value_in_measure_domain(value: object, domain: Mapping[str, Any]) -> bool:
    """Return whether `value` is inside a measure domain declaration."""
    if "enum" in domain:
        return value in domain["enum"]
    if "between" in domain:
        if isinstance(value, bool) or not isinstance(value, int):
            return False
        return _in_closed_bounds(value, domain["between"])
    if "real_between" in domain:
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            return False
        return _in_closed_bounds(value, domain["real_between"])
    return True


def _is_measure_sequence(value: object) -> TypeGuard[Sequence[object]]:
    return isinstance(value, Sequence) and not isinstance(value, (str, bytes, bytearray))


def _reject_out_of_domain(
    value: object,
    domain: Mapping[str, Any],
    *,
    label: str,
) -> None:
    """Raise `ValueError` when a non-null `value` is outside `domain`."""
    if value is None:
        return
    if not _value_in_measure_domain(value, domain):
        raise ValueError(
            f"{label} out of domain: {value!r} not in {_format_measure_domain(domain)}"
        )


def require_input_domain(
    value: object,
    domain: Mapping[str, Any],
    *,
    series_id: str,
) -> None:
    """Reject a scalar or sequence argument outside `input.domain`.

    Args:
        value: One measure, or a catalog-order sequence of measures.
        domain: Normalized `enum` / `between` / `real_between` declaration.
        series_id: Binding series id used in the error message.

    Raises:
        ValueError: When any non-`None` member is outside `domain`.
    """
    if _is_measure_sequence(value):
        for index, member in enumerate(value):
            _reject_out_of_domain(member, domain, label=f"{series_id}[{index}]")
        return
    _reject_out_of_domain(value, domain, label=series_id)


def measure_domain_from_series(series: Mapping[str, Any]) -> dict[str, Any] | None:
    """Normalize `input.domain` for codegen and runtime checks."""
    input_block = series.get("input")
    if not isinstance(input_block, dict):
        return None
    domain = input_block.get("domain")
    if not isinstance(domain, dict):
        return None
    if "enum" in domain:
        values = domain["enum"]
        if not isinstance(values, (list, tuple, set, frozenset)):
            return None
        return {"enum": frozenset(values)}
    if "between" in domain:
        bounds = domain["between"]
        if not isinstance(bounds, dict):
            return None
        return {"between": dict(bounds)}
    if "real_between" in domain:
        bounds = domain["real_between"]
        if not isinstance(bounds, dict):
            return None
        return {"real_between": dict(bounds)}
    return None


def _apply_measure_domain(
    records: Records,
    *,
    measure_field: str,
    measure_domain: Mapping[str, Any] | None,
) -> Records:
    """Reject measure values outside an optional `input.domain` declaration."""
    if measure_domain is None:
        return records
    for index, record in enumerate(records):
        if measure_field not in record:
            continue
        _reject_out_of_domain(
            record[measure_field],
            measure_domain,
            label=f"record[{index}]: {measure_field}",
        )
    return records


def _coerce_dataframe_records(
    data: object,
    *,
    key_fields: tuple[str, ...],
    measure_field: str,
    strict: bool,
) -> Records:
    column_names = _dataframe_column_names(data)
    _validate_dataframe_columns(
        column_names,
        key_fields=key_fields,
        measure_field=measure_field,
        strict=strict,
    )
    records: list[dict[str, object]] = []
    for row in _row_dicts_from_dataframe(data):
        record: dict[str, object] = {field: row[field] for field in key_fields}
        record[measure_field] = row[measure_field]
        records.append(record)
    return records


def _non_scalar_input_hint(
    *,
    layout: Layout,
    key_fields: tuple[str, ...],
    measure_field: str,
) -> str:
    if layout == "matrix":
        columns = ", ".join([*key_fields, measure_field])
        return f"pass records or a tidy DataFrame with columns {columns!r}"
    return "pass records, a 1D iterable of measure values, or a tidy DataFrame"


def _unsupported_non_scalar_input_type_error(
    data: object,
    *,
    layout: Layout,
    key_fields: tuple[str, ...],
    measure_field: str,
) -> TypeError:
    hint = _non_scalar_input_hint(
        layout=layout,
        key_fields=key_fields,
        measure_field=measure_field,
    )
    return TypeError(f"unsupported {layout} setter input type {type(data)!r}; {hint}")


def _coerce_positional_records(
    data: Iterable[object],
    *,
    layout: Layout,
    key_fields: tuple[str, ...],
    measure_field: str,
    key_order: tuple[object, ...],
) -> Records:
    if len(key_fields) != 1:
        if layout == "matrix":
            raise ValueError(
                "positional measure values are not supported for matrix setters; "
                "pass records or a tidy DataFrame"
            )
        raise ValueError(
            "positional measure values require a single-key series binding; "
            f"got key_fields={list(key_fields)!r}"
        )
    values = list(data)
    if len(values) != len(key_order):
        raise ValueError(
            f"expected {len(key_order)} values for positional input, got {len(values)}"
        )
    key_field = key_fields[0]
    return [
        {key_field: key, measure_field: value} for key, value in zip(key_order, values, strict=True)
    ]


def _coerce_non_scalar_records(
    data: SetterInput,
    *,
    layout: Layout,
    key_fields: tuple[str, ...],
    measure_field: str,
    key_order: tuple[object, ...] | None,
    strict: bool,
) -> Records:
    """Normalize series/matrix setter input to records before key coercion."""
    if _is_tabular_dataframe(data):
        return _coerce_dataframe_records(
            data,
            key_fields=key_fields,
            measure_field=measure_field,
            strict=strict,
        )

    if _is_mapping(data):
        return [dict(data)]

    if isinstance(data, list):
        if _is_records_list(data):
            return data
        if key_order is None:
            raise ValueError("positional input requires key_order")
        return _coerce_positional_records(
            data,
            layout=layout,
            key_fields=key_fields,
            measure_field=measure_field,
            key_order=key_order,
        )

    if isinstance(data, (str, bytes, bytearray)):
        raise _unsupported_non_scalar_input_type_error(
            data,
            layout=layout,
            key_fields=key_fields,
            measure_field=measure_field,
        )

    if isinstance(data, Iterable):
        if key_order is None:
            raise ValueError("positional input requires key_order")
        return _coerce_positional_records(
            data,
            layout=layout,
            key_fields=key_fields,
            measure_field=measure_field,
            key_order=key_order,
        )

    raise _unsupported_non_scalar_input_type_error(
        data,
        layout=layout,
        key_fields=key_fields,
        measure_field=measure_field,
    )


def coerce_setter_input(
    data: SetterInput,
    *,
    layout: Layout,
    key_fields: tuple[str, ...],
    measure_field: str,
    key_order: tuple[object, ...] | None,
    strict: bool,
    key_dtypes: Mapping[str, str] | None = None,
    measure_dtype: str | None = None,
    measure_domain: Mapping[str, Any] | None = None,
    empty_measure: EmptyMeasure = "write",
    requires_address: bool = False,
) -> Records:
    """Normalize caller input into records for ``_apply_series_records``.

    Args:
        data: Scalar value, record(s), 1D measure values, or tidy DataFrame.
        layout: Binding layout (`scalar`, `series`, or `matrix`).
        key_fields: Key column names from the binding manifest.
        measure_field: Measure concept name (e.g. `OBS_VALUE`).
        key_order: Canonical key values for positional measure iterables.
        strict: When true, reject unknown DataFrame columns.
        key_dtypes: Optional read modes per key field applied to all input shapes.
        measure_dtype: Optional binding dtype enforced for `measure_field` values.
        measure_domain: Optional `input.domain` (`enum` / `between` / `real_between`)
            enforced after dtype coercion.
        empty_measure: How to treat rows with missing/NaN measure values.
        requires_address: When true, reject DataFrame input (records must carry addresses).

    Returns:
        List of record dicts ready for leaf resolution.

    Raises:
        ImportError: When a DataFrame-like value is passed but pandas/polars is missing.
        TypeError: When the input shape is unsupported for the layout, or a measure
            value does not match `measure_dtype`.
        ValueError: When columns, keys, positional lengths, or domains are invalid.
    """
    if layout == "scalar":
        if _is_tabular_dataframe(data):
            raise TypeError("scalar setters do not accept DataFrame input")
        records = _coerce_scalar_records(data, measure_field)
        records = _apply_measure_dtype(
            records,
            measure_field=measure_field,
            measure_dtype=measure_dtype,
        )
        return _apply_measure_domain(
            records,
            measure_field=measure_field,
            measure_domain=measure_domain,
        )

    if requires_address and _is_tabular_dataframe(data):
        raise TypeError(
            "DataFrame input is not supported when the binding requires address "
            "disambiguation; pass records with 'address' or 'cell_address'"
        )

    records = _coerce_non_scalar_records(
        data,
        layout=layout,
        key_fields=key_fields,
        measure_field=measure_field,
        key_order=key_order,
        strict=strict,
    )
    records = _apply_key_dtypes(
        records,
        key_fields=key_fields,
        key_dtypes=key_dtypes,
    )
    records = _apply_measure_dtype(
        records,
        measure_field=measure_field,
        measure_dtype=measure_dtype,
    )
    records = _apply_measure_domain(
        records,
        measure_field=measure_field,
        measure_domain=measure_domain,
    )
    return _apply_empty_measure(
        records,
        key_fields=key_fields,
        measure_field=measure_field,
        empty_measure=empty_measure,
    )
