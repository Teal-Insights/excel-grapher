"""Lazy worksheet views over named coordinates replace per-cell callbacks."""

from __future__ import annotations

from excel_grapher.exporter.export_runtime.tensor import Axis, Domain, Tensor
from excel_grapher.exporter.inverted_tree import runtime

YEARS = Axis("TIME_PERIOD", (2020, 2021, 2022, 2023), int)
COUNTRIES = Axis("COUNTRY", ("France", "Kenya"), str)


def _series() -> Tensor[float]:
    return Tensor(Domain.product(YEARS), (1.0, 2.0, 3.0, 4.0))


def _matrix() -> Tensor[float]:
    return Tensor(Domain.product(COUNTRIES, YEARS), tuple(float(i) for i in range(8)))


def test_span_returns_inclusive_keys_in_axis_order() -> None:
    assert runtime.span(YEARS, 2021, 2023) == (2021, 2022, 2023)
    assert runtime.span(YEARS, 2022, 2022) == (2022,)


def test_span_outside_axis_is_a_reference_error() -> None:
    try:
        runtime.span(YEARS, 2019, 2021)
    except runtime.XlError as error:
        assert error.code == "#REF!"
    else:
        raise AssertionError("expected #REF!")


def test_row_view_exposes_a_one_row_range() -> None:
    row = runtime.view(_series(), cols=runtime.span(YEARS, 2021, 2023))
    assert row.shape == (1, 3)
    assert runtime.xl_sum(row) == 9.0
    assert runtime.xl_index(row, None, 2) == 3.0


def test_column_view_resolves_only_selected_cells() -> None:
    reads: list[tuple[int, ...]] = []

    class Recording:
        def __getitem__(self, key: tuple[int, ...]) -> float:
            reads.append(key)
            return float(key[0])

    column = runtime.view(Recording(), rows=YEARS.keys)
    assert column.shape == (4, 1)
    assert runtime.xl_index(column, 3, 1) == 2022.0
    assert reads == [(2022,)]


def test_block_view_composes_row_and_column_keys() -> None:
    block = runtime.view(_matrix(), rows=COUNTRIES.keys, cols=runtime.span(YEARS, 2022, 2023))
    assert block.shape == (2, 2)
    assert runtime.xl_index(block, 2, 1) == 6.0
    transposed = runtime.view(
        Tensor(Domain.product(YEARS, COUNTRIES), tuple(float(i) for i in range(8))),
        rows=COUNTRIES.keys,
        cols=YEARS.keys,
        cols_first=True,
    )
    assert runtime.xl_index(transposed, 2, 1) == 1.0


def test_stored_error_in_aggregated_view_propagates() -> None:
    values = Tensor(Domain.product(YEARS), (1.0, "#N/A", 3.0, 4.0))
    try:
        runtime.xl_sum(runtime.view(values, cols=YEARS.keys))
    except runtime.XlError as error:
        assert error.code == "#N/A"
    else:
        raise AssertionError("expected the stored error to propagate")
