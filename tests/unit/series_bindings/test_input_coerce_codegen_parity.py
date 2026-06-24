"""Parity between library ``coerce_setter_input`` and codegen-emitted helpers."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any, Literal, TypedDict, cast

import pytest
import xlsxwriter

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.runtime.cache import EvalContext, coerce_inputs_dict
from excel_grapher.series_bindings import (
    expand_data_range,
    load_series_bindings,
    resolve_series_binding,
)
from excel_grapher.series_bindings.input_coerce import coerce_setter_input
from excel_grapher.series_bindings.setter_codegen import (
    _canonical_key_order,
    _key_dtypes_for_codegen,
    emit_input_coerce_helpers,
    emit_setter_function,
    emit_setter_helpers,
)

FIXTURES = Path(__file__).resolve().parents[2] / "fixtures" / "series_bindings"


class _SeriesCoerceKwargs(TypedDict, total=False):
    layout: Literal["scalar", "series"]
    key_fields: tuple[str, ...]
    measure_field: str
    key_order: tuple[object, ...] | None
    strict: bool
    key_dtypes: dict[str, str]


def _write_borvelia_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Inputs")
    ws.write("A2", "Borvelia")
    ws.write("A5", "Primary balance (% of GDP)")
    for col, year in enumerate([1, 2, 3, 4, 5], start=5):
        ws.write(0, col, year)
        ws.write_number(4, col, float(year - 3))
    wb.close()


def _emitted_coerce_setter_input() -> Callable[..., list[dict[str, object]]]:
    namespace: dict[str, object] = {}
    exec("\n".join(emit_input_coerce_helpers()), namespace)
    return cast(Callable[..., list[dict[str, object]]], namespace["coerce_setter_input"])


def _borvelia_series_kwargs(tmp_path: Path) -> _SeriesCoerceKwargs:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, expand_data_range("Inputs!F5:J5"), load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series)
    key_fields = [str(c) for c in (series.get("key") or [])]
    key_dtypes = _key_dtypes_for_codegen(series, key_fields)
    kwargs: _SeriesCoerceKwargs = {
        "layout": "series",
        "key_fields": tuple(key_fields),
        "measure_field": "OBS_VALUE",
        "key_order": _canonical_key_order(resolved, key_fields),
        "strict": True,
    }
    if key_dtypes:
        kwargs["key_dtypes"] = key_dtypes
    return kwargs


@pytest.mark.parametrize(
    ("data",),
    [
        ([{"TIME_PERIOD": 4, "OBS_VALUE": 7.5}, {"TIME_PERIOD": 5, "OBS_VALUE": 8.0}],),
        ({"TIME_PERIOD": 4, "OBS_VALUE": 7.5},),
        ([-2.0, -1.0, 0.0, 7.5, 8.0],),
    ],
)
def test_emitted_coerce_matches_library_for_series_inputs(
    tmp_path: Path,
    data: object,
) -> None:
    emitted = _emitted_coerce_setter_input()
    kwargs = _borvelia_series_kwargs(tmp_path)
    assert coerce_setter_input(data, **kwargs) == emitted(data, **kwargs)


def test_emitted_coerce_matches_library_for_pandas_dataframe(tmp_path: Path) -> None:
    pd = pytest.importorskip("pandas")
    emitted = _emitted_coerce_setter_input()
    kwargs = _borvelia_series_kwargs(tmp_path)
    df = pd.DataFrame({"TIME_PERIOD": [4, 5], "OBS_VALUE": [7.5, 8.0]})
    assert coerce_setter_input(df, **kwargs) == emitted(df, **kwargs)


def test_generated_setter_matches_library_coercion_for_positional_input(tmp_path: Path) -> None:
    wb_path = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(wb_path)
    graph = create_dependency_graph(wb_path, expand_data_range("Inputs!F5:J5"), load_values=True)
    bindings = load_series_bindings(FIXTURES / "borvelia_primary_balance.yaml")
    series = bindings["series"][0]
    resolved = resolve_series_binding(graph, wb_path, series)
    lines = emit_setter_function(series, resolved)
    namespace: dict[str, object] = {
        "EvalContext": EvalContext,
        "coerce_inputs_dict": coerce_inputs_dict,
    }
    exec("\n".join(emit_input_coerce_helpers() + emit_setter_helpers() + lines), namespace)
    setter = cast(Callable[[EvalContext, object], None], namespace["set_borvelia_primary_balance"])

    key_fields = [str(c) for c in (series.get("key") or [])]
    key_dtypes = _key_dtypes_for_codegen(series, key_fields)
    kwargs: dict[str, Any] = {
        "layout": "series",
        "key_fields": tuple(key_fields),
        "measure_field": "OBS_VALUE",
        "key_order": _canonical_key_order(resolved, key_fields),
        "strict": True,
    }
    if key_dtypes:
        kwargs["key_dtypes"] = key_dtypes
    values = [-2.0, -1.0, 0.0, 7.5, 8.0]
    records = coerce_setter_input(values, **kwargs)

    ctx_library = EvalContext(inputs=coerce_inputs_dict({}), resolver=lambda _a: None)
    ctx_generated = EvalContext(inputs=coerce_inputs_dict({}), resolver=lambda _a: None)
    setter(ctx_generated, values)
    apply_records = cast(
        Callable[..., None],
        namespace["_apply_series_records"],
    )
    apply_records(
        ctx_library,
        records,
        key_fields=kwargs["key_fields"],
        allowed_fields=frozenset({"OBS_VALUE", "TIME_PERIOD"}),
        measure_field="OBS_VALUE",
        leaf_index=namespace["_LEAF_INDEX_BORVELIA_PRIMARY_BALANCE"],
        strict=True,
        fn_name="set_borvelia_primary_balance",
        allow_address=False,
        requires_address=False,
    )
    assert ctx_generated.inputs == ctx_library.inputs
