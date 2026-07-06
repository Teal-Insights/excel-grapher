"""Unit tests for setter input coercion (records, positional, DataFrame)."""

from __future__ import annotations

from datetime import datetime
from typing import Any, Literal, TypedDict, cast

import pytest

from excel_grapher.series_bindings.input_coerce import coerce_setter_input
from excel_grapher.series_bindings.types import Records


class _SeriesCoerceKwargs(TypedDict):
    layout: Literal["series", "matrix"]
    key_fields: tuple[str, ...]
    measure_field: str
    key_order: tuple[int, ...] | None
    strict: bool


_SERIES_KWARGS: _SeriesCoerceKwargs = {
    "layout": "series",
    "key_fields": ("TIME_PERIOD",),
    "measure_field": "OBS_VALUE",
    "key_order": (1, 2, 3, 4, 5),
    "strict": True,
}


def test_records_pass_through_unchanged() -> None:
    records: Records = [
        {"TIME_PERIOD": 4, "OBS_VALUE": 7.5},
        {"TIME_PERIOD": 5, "OBS_VALUE": 8.0},
    ]
    result = coerce_setter_input(records, **_SERIES_KWARGS)
    assert result is records


def test_positional_series_builds_records() -> None:
    result = coerce_setter_input([-2.0, -1.0, 0.0, 7.5, 8.0], **_SERIES_KWARGS)
    assert result == [
        {"TIME_PERIOD": 1, "OBS_VALUE": -2.0},
        {"TIME_PERIOD": 2, "OBS_VALUE": -1.0},
        {"TIME_PERIOD": 3, "OBS_VALUE": 0.0},
        {"TIME_PERIOD": 4, "OBS_VALUE": 7.5},
        {"TIME_PERIOD": 5, "OBS_VALUE": 8.0},
    ]


def test_positional_wrong_length_raises() -> None:
    with pytest.raises(ValueError, match="expected 5 values"):
        coerce_setter_input([1.0, 2.0], **_SERIES_KWARGS)


def test_positional_multi_key_series_raises() -> None:
    with pytest.raises(ValueError, match="single-key"):
        coerce_setter_input(
            [1.0, 2.0],
            layout="series",
            key_fields=("REF_AREA", "TIME_PERIOD"),
            measure_field="OBS_VALUE",
            key_order=((1, 2),),
            strict=True,
        )


def test_positional_requires_key_order() -> None:
    with pytest.raises(ValueError, match="key_order"):
        coerce_setter_input(
            [1.0],
            layout="series",
            key_fields=("TIME_PERIOD",),
            measure_field="OBS_VALUE",
            key_order=None,
            strict=True,
        )


def test_series_single_dict_wrapped_as_record() -> None:
    result = coerce_setter_input(
        {"TIME_PERIOD": 4, "OBS_VALUE": 7.5},
        **_SERIES_KWARGS,
    )
    assert result == [{"TIME_PERIOD": 4, "OBS_VALUE": 7.5}]


def test_tidy_pandas_dataframe_matrix_layout() -> None:
    pd = pytest.importorskip("pandas")
    df = pd.DataFrame(
        {
            "INDICATOR": ["GDP growth", "Debt"],
            "TIME_PERIOD": [2025, 2026],
            "OBS_VALUE": [9.9, 44.4],
        }
    )
    result = coerce_setter_input(
        df,
        layout="matrix",
        key_fields=("INDICATOR", "TIME_PERIOD"),
        measure_field="OBS_VALUE",
        key_order=None,
        strict=True,
    )
    assert result == [
        {"INDICATOR": "GDP growth", "TIME_PERIOD": 2025, "OBS_VALUE": 9.9},
        {"INDICATOR": "Debt", "TIME_PERIOD": 2026, "OBS_VALUE": 44.4},
    ]


def test_tidy_pandas_dataframe() -> None:
    pd = pytest.importorskip("pandas")
    df = pd.DataFrame({"TIME_PERIOD": [4, 5], "OBS_VALUE": [7.5, 8.0]})
    result = coerce_setter_input(df, **_SERIES_KWARGS)
    assert result == [
        {"TIME_PERIOD": 4, "OBS_VALUE": 7.5},
        {"TIME_PERIOD": 5, "OBS_VALUE": 8.0},
    ]


def test_tidy_pandas_dataframe_partial_rows_ok() -> None:
    pd = pytest.importorskip("pandas")
    df = pd.DataFrame({"TIME_PERIOD": [4], "OBS_VALUE": [7.5]})
    result = coerce_setter_input(df, **_SERIES_KWARGS)
    assert result == [{"TIME_PERIOD": 4, "OBS_VALUE": 7.5}]


def test_tidy_dataframe_extra_column_strict_raises() -> None:
    pd = pytest.importorskip("pandas")
    df = pd.DataFrame(
        {"TIME_PERIOD": [4], "OBS_VALUE": [7.5], "notes": ["draft"]},
    )
    with pytest.raises(ValueError, match="unknown columns"):
        coerce_setter_input(df, **_SERIES_KWARGS)


def test_tidy_dataframe_extra_column_non_strict_ok() -> None:
    pd = pytest.importorskip("pandas")
    df = pd.DataFrame(
        {"TIME_PERIOD": [4], "OBS_VALUE": [7.5], "notes": ["draft"]},
    )
    result = coerce_setter_input(
        df,
        layout=_SERIES_KWARGS["layout"],
        key_fields=_SERIES_KWARGS["key_fields"],
        measure_field=_SERIES_KWARGS["measure_field"],
        key_order=_SERIES_KWARGS["key_order"],
        strict=False,
    )
    assert result == [{"TIME_PERIOD": 4, "OBS_VALUE": 7.5}]


def test_missing_measure_column_raises() -> None:
    pd = pytest.importorskip("pandas")
    df = pd.DataFrame({"TIME_PERIOD": [4]})
    with pytest.raises(ValueError, match="OBS_VALUE"):
        coerce_setter_input(df, **_SERIES_KWARGS)


def test_key_coercion_int_from_float_records() -> None:
    records: Records = [
        {"TIME_PERIOD": 4.0, "OBS_VALUE": 7.5},
        {"TIME_PERIOD": 5.0, "OBS_VALUE": 8.0},
    ]
    result = coerce_setter_input(
        records,
        key_dtypes={"TIME_PERIOD": "int"},
        **_SERIES_KWARGS,
    )
    assert result == [
        {"TIME_PERIOD": 4, "OBS_VALUE": 7.5},
        {"TIME_PERIOD": 5, "OBS_VALUE": 8.0},
    ]
    assert result is not records


def test_key_coercion_int_from_float() -> None:
    pd = pytest.importorskip("pandas")
    df = pd.DataFrame({"TIME_PERIOD": [4.0, 5.0], "OBS_VALUE": [7.5, 8.0]})
    result = coerce_setter_input(
        df,
        key_dtypes={"TIME_PERIOD": "int"},
        **_SERIES_KWARGS,
    )
    assert result == [
        {"TIME_PERIOD": 4, "OBS_VALUE": 7.5},
        {"TIME_PERIOD": 5, "OBS_VALUE": 8.0},
    ]


def test_tidy_polars_dataframe() -> None:
    pl = pytest.importorskip("polars")
    df = pl.DataFrame({"TIME_PERIOD": [4, 5], "OBS_VALUE": [7.5, 8.0]})
    result = coerce_setter_input(df, **_SERIES_KWARGS)
    assert result == [
        {"TIME_PERIOD": 4, "OBS_VALUE": 7.5},
        {"TIME_PERIOD": 5, "OBS_VALUE": 8.0},
    ]


def test_dataframe_like_without_pandas_raises() -> None:
    class _FakePandasDataFrame:
        __module__ = "pandas.core.frame"

        @property
        def columns(self) -> list[str]:
            return ["TIME_PERIOD", "OBS_VALUE"]

    fake = _FakePandasDataFrame()
    type(fake).__name__ = "DataFrame"  # type: ignore[misc]

    with pytest.raises(ImportError, match="pandas"):
        coerce_setter_input(cast(Any, fake), **_SERIES_KWARGS)


def test_scalar_layout_bare_value() -> None:
    result = coerce_setter_input(
        "France",
        layout="scalar",
        key_fields=(),
        measure_field="OBS_VALUE",
        key_order=None,
        strict=True,
    )
    assert result == [{"OBS_VALUE": "France"}]


def test_scalar_layout_dict_record() -> None:
    result = coerce_setter_input(
        {"OBS_VALUE": "France"},
        layout="scalar",
        key_fields=(),
        measure_field="OBS_VALUE",
        key_order=None,
        strict=True,
    )
    assert result == [{"OBS_VALUE": "France"}]


def test_scalar_layout_list_passes_through() -> None:
    records: Records = [{"OBS_VALUE": "France"}]
    result = coerce_setter_input(
        records,
        layout="scalar",
        key_fields=(),
        measure_field="OBS_VALUE",
        key_order=None,
        strict=True,
    )
    assert result is records


def test_dataframe_on_scalar_layout_raises() -> None:
    pd = pytest.importorskip("pandas")
    df = pd.DataFrame({"OBS_VALUE": [1.0]})
    with pytest.raises(TypeError, match="scalar"):
        coerce_setter_input(
            df,
            layout="scalar",
            key_fields=(),
            measure_field="OBS_VALUE",
            key_order=None,
            strict=True,
        )


def test_unsupported_type_raises() -> None:
    with pytest.raises(TypeError, match=r"unsupported series setter input type"):
        coerce_setter_input(42, **_SERIES_KWARGS)


def test_unsupported_type_matrix_layout() -> None:
    with pytest.raises(TypeError, match=r"unsupported matrix setter input type"):
        coerce_setter_input(
            42,
            layout="matrix",
            key_fields=("INDICATOR", "TIME_PERIOD"),
            measure_field="OBS_VALUE",
            key_order=None,
            strict=True,
        )
    with pytest.raises(TypeError, match="INDICATOR, TIME_PERIOD, OBS_VALUE"):
        coerce_setter_input(
            42,
            layout="matrix",
            key_fields=("INDICATOR", "TIME_PERIOD"),
            measure_field="OBS_VALUE",
            key_order=None,
            strict=True,
        )


def test_positional_multi_key_matrix_raises() -> None:
    with pytest.raises(ValueError, match="matrix setters"):
        coerce_setter_input(
            [1.0, 2.0],
            layout="matrix",
            key_fields=("INDICATOR", "TIME_PERIOD"),
            measure_field="OBS_VALUE",
            key_order=((1, 2),),
            strict=True,
        )


def test_positional_rejects_string() -> None:
    with pytest.raises(TypeError, match=r"unsupported series setter input type"):
        coerce_setter_input("not-a-record", **_SERIES_KWARGS)


def test_key_coercion_datetime() -> None:
    pd = pytest.importorskip("pandas")
    dt = datetime(2020, 1, 15)
    df = pd.DataFrame({"TIME_PERIOD": ["2020-01-15"], "OBS_VALUE": [1.0]})
    result = coerce_setter_input(
        df,
        key_dtypes={"TIME_PERIOD": "datetime"},
        layout="series",
        key_fields=("TIME_PERIOD",),
        measure_field="OBS_VALUE",
        key_order=None,
        strict=True,
    )
    assert result == [{"TIME_PERIOD": dt, "OBS_VALUE": 1.0}]


_MATRIX_KWARGS: _SeriesCoerceKwargs = {
    "layout": "matrix",
    "key_fields": ("INDICATOR", "TIME_PERIOD"),
    "measure_field": "OBS_VALUE",
    "key_order": None,
    "strict": True,
}


def test_tidy_polars_dataframe_matrix_layout() -> None:
    pl = pytest.importorskip("polars")
    df = pl.DataFrame(
        {
            "INDICATOR": ["GDP growth", "Debt"],
            "TIME_PERIOD": [2025, 2026],
            "OBS_VALUE": [9.9, 44.4],
        }
    )
    result = coerce_setter_input(df, **_MATRIX_KWARGS)
    assert result == [
        {"INDICATOR": "GDP growth", "TIME_PERIOD": 2025, "OBS_VALUE": 9.9},
        {"INDICATOR": "Debt", "TIME_PERIOD": 2026, "OBS_VALUE": 44.4},
    ]


def test_empty_dataframe_matrix_layout_is_noop() -> None:
    pd = pytest.importorskip("pandas")
    df = pd.DataFrame({"INDICATOR": [], "TIME_PERIOD": [], "OBS_VALUE": []})
    result = coerce_setter_input(df, **_MATRIX_KWARGS)
    assert result == []


def test_wide_dataframe_matrix_layout_hint() -> None:
    pd = pytest.importorskip("pandas")
    df = pd.DataFrame(
        {
            "INDICATOR": ["GDP growth"],
            2024: [1.5],
            2025: [1.8],
        }
    )
    with pytest.raises(ValueError, match="looks wide"):
        coerce_setter_input(df, **_MATRIX_KWARGS)


def test_empty_measure_write_passes_none_through() -> None:
    records: Records = [{"TIME_PERIOD": 4, "OBS_VALUE": None}]
    result = coerce_setter_input(records, **_SERIES_KWARGS, empty_measure="write")
    assert result == [{"TIME_PERIOD": 4, "OBS_VALUE": None}]


def test_empty_measure_skip_drops_row() -> None:
    pd = pytest.importorskip("pandas")
    df = pd.DataFrame({"TIME_PERIOD": [4, 5], "OBS_VALUE": [7.5, float("nan")]})
    result = coerce_setter_input(df, **_SERIES_KWARGS, empty_measure="skip")
    assert result == [{"TIME_PERIOD": 4, "OBS_VALUE": 7.5}]


def test_empty_measure_error_raises() -> None:
    records: Records = [{"TIME_PERIOD": 4, "OBS_VALUE": None}]
    with pytest.raises(ValueError, match="empty measure field"):
        coerce_setter_input(records, **_SERIES_KWARGS, empty_measure="error")


def test_empty_key_always_errors() -> None:
    pd = pytest.importorskip("pandas")
    df = pd.DataFrame({"TIME_PERIOD": [None], "OBS_VALUE": [7.5]})
    with pytest.raises(ValueError, match="empty key field"):
        coerce_setter_input(df, **_SERIES_KWARGS, empty_measure="skip")


def test_requires_address_dataframe_rejected() -> None:
    pd = pytest.importorskip("pandas")
    df = pd.DataFrame({"TIME_PERIOD": [4], "OBS_VALUE": [7.5]})
    with pytest.raises(TypeError, match="requires address"):
        coerce_setter_input(
            df,
            **_SERIES_KWARGS,
            requires_address=True,
        )
