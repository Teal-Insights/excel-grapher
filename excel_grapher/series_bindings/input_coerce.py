"""Coerce setter caller input into canonical record lists for series bindings."""

from __future__ import annotations

from collections.abc import Iterable, Mapping, Sequence
from typing import Any, TypeGuard, cast

from excel_grapher.series_bindings.coerce import coerce_scalar
from excel_grapher.series_bindings.records_types import Record, Records
from excel_grapher.series_bindings.setter_input_types import Layout, SetterInput

__all__ = ["Layout", "coerce_setter_input"]


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
        raise ValueError(f"missing required column(s): {missing!r}")
    if strict:
        unknown = sorted(present - required)
        if unknown:
            raise ValueError(f"unknown columns {unknown!r}")


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

    Returns:
        List of record dicts ready for leaf resolution.

    Raises:
        ImportError: When a DataFrame-like value is passed but pandas/polars is missing.
        TypeError: When the input shape is unsupported for the layout.
        ValueError: When columns, keys, or positional lengths are invalid.
    """
    if layout == "scalar":
        if _is_tabular_dataframe(data):
            raise TypeError("scalar setters do not accept DataFrame input")
        return _coerce_scalar_records(data, measure_field)

    records = _coerce_non_scalar_records(
        data,
        layout=layout,
        key_fields=key_fields,
        measure_field=measure_field,
        key_order=key_order,
        strict=strict,
    )
    return _apply_key_dtypes(
        records,
        key_fields=key_fields,
        key_dtypes=key_dtypes,
    )
