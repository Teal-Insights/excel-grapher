from dataclasses import FrozenInstanceError
from datetime import date, datetime

import pytest

from excel_grapher.exporter.export_runtime.tensor import (
    Axis,
    AxisError,
    CoordinateError,
    Domain,
    DomainError,
    SchemaError,
    Tensor,
    TensorSchema,
)


def years(*keys: int) -> Axis:
    return Axis("year", keys=keys, key_type=int)


def test_product_label_access_selection_and_immutability() -> None:
    domain = Domain.product(
        Axis("scenario", keys=("base", "shock"), key_type=str), years(2025, 2026)
    )
    tensor = Tensor.from_nested(domain=domain, values=((0, None), ("#DIV/0!", 4)))
    assert tensor["base", 2025] == 0
    assert tensor.sel(scenario="base")[2026] is None
    assert tensor.isel(scenario=1, year=0) == "#DIV/0!"
    assert tensor.sel(scenario="shock", year=2026) == 4
    with pytest.raises(CoordinateError):
        tensor[0, 0]
    with pytest.raises(CoordinateError):
        tensor["base"]
    with pytest.raises(FrozenInstanceError):
        tensor.domain = Domain.product()  # ty: ignore[invalid-assignment]


def test_sparse_domain_order_holes_and_exact_records() -> None:
    domain = Domain.explicit(axes=(years(2026, 2025, 2027),), coordinates=((2027,), (2026,)))
    tensor = Tensor.from_records(domain=domain, records=(((2027,), None), ((2026,), 0)))
    assert list(tensor.items()) == [((2026,), 0), ((2027,), None)]
    with pytest.raises(CoordinateError, match="2025"):
        tensor[2025]
    for records in [
        (((2026,), 1),),
        (((2026,), 1), ((2027,), 2), ((2025,), 3)),
        (((2026,), 1), ((2026,), 2)),
    ]:
        with pytest.raises(DomainError):
            Tensor.from_records(domain=domain, records=records)


def test_arbitrary_rank_and_scalar_product() -> None:
    axes = tuple(
        Axis(name, keys=(0, 1), key_type=int)
        for name in ("scenario", "instrument", "holder", "year")
    )
    domain = Domain.product(*axes)
    tensor = Tensor.from_records(domain=domain, records=((coord, sum(coord)) for coord in domain))
    assert tensor[1, 0, 1, 1] == 3
    assert Tensor.from_nested(domain=Domain.product(), values=42)[()] == 42


@pytest.mark.parametrize(
    "name,keys,key_type",
    [
        ("", (1,), int),
        ("year", (1, 1), int),
        ("year", (True,), int),
        ("year", (1.0,), int),
        ("year", ([1],), str),
    ],
)
def test_invalid_axes(name, keys, key_type) -> None:
    with pytest.raises(AxisError):
        Axis(name, keys=keys, key_type=key_type)


def test_invalid_domains_and_shapes() -> None:
    with pytest.raises(DomainError):
        Domain.product(years(2025), years(2026))
    for coordinates in [((2024,),), ((True,),), ((2025, 0),), ((2025,), (2025,))]:
        with pytest.raises(DomainError):
            Domain.explicit(axes=(years(2025),), coordinates=coordinates)
    with pytest.raises(DomainError):
        Tensor.from_nested(domain=Domain.product(years(2025, 2026)), values=(1,))
    tensor = Tensor.from_nested(domain=Domain.product(years(2025)), values=(1,))
    with pytest.raises(CoordinateError):
        tensor.sel(unknown=2025)


def test_required_schema_rejects_same_size_different_labels() -> None:
    schema = TensorSchema(
        "gdp", Domain.product(years(2025, 2026)), value_types=(int, float, str, type(None))
    )
    tensor = Tensor.from_nested(domain=Domain.product(years(2026, 2027)), values=(1, 2))
    with pytest.raises(SchemaError, match="gdp.*2025"):
        schema.validate(tensor)
    wider = Tensor.from_nested(domain=Domain.product(years(2024, 2025, 2026)), values=(0, 1, 2))
    schema.validate(wider)
    with pytest.raises(SchemaError):
        schema.validate(wider, exact=True)


def test_sparse_serialization_and_explicit_legacy_order() -> None:
    domain = Domain.explicit(axes=(years(2025, 2026),), coordinates=((2026,), (2025,)))
    tensor = Tensor.from_legacy(
        domain=domain, values=("#N/A", None), coordinate_order=((2026,), (2025,))
    )
    assert tensor[2026] == "#N/A"
    assert tensor.to_legacy(coordinate_order=((2026,), (2025,))) == ("#N/A", None)
    assert Tensor.from_json(tensor.to_json()) == tensor
    assert (
        domain.fingerprint
        == Domain.explicit(axes=domain.axes, coordinates=reversed(tuple(domain))).fingerprint
    )


def test_year_scenario_facades_preserve_ragged_paths() -> None:
    from excel_grapher.exporter.export_runtime.tensor import ScenarioSeries, YearSeries

    baseline = YearSeries(years=(2025, 2026, 2027), values=(100, 105, 110))
    shock = YearSeries(years=(2026, 2027), values=(102, 106))
    paths = ScenarioSeries.from_paths({"base": baseline, "shock": shock})
    assert paths["shock", 2026] == 102
    assert isinstance(paths["shock"], YearSeries)
    assert paths["shock"].years == (2026, 2027)
    with pytest.raises(CoordinateError):
        paths["shock", 2025]
    with pytest.raises(AxisError):
        YearSeries(years=(True,), values=(1,))


def test_facade_constructors_and_serialization_enforce_schema() -> None:
    from excel_grapher.exporter.export_runtime.tensor import ScenarioSeries, YearSeries

    series = YearSeries.from_nested(domain=Domain.product(years(2025)), values=(3,))
    assert isinstance(series, YearSeries)
    assert series[2025] == 3
    assert YearSeries.from_json(series.to_json()) == series
    with pytest.raises(SchemaError):
        YearSeries.from_nested(domain=Domain.product(Axis("creditor", ("bank",), str)), values=(3,))
    with pytest.raises(SchemaError):
        ScenarioSeries.from_nested(
            domain=Domain.product(Axis("scenario", (1,), int), years(2025)), values=((3,),)
        )


def test_schema_validates_its_own_contract() -> None:
    with pytest.raises(SchemaError):
        TensorSchema("gdp", Domain.product(years()), value_types=(float,))
    with pytest.raises(SchemaError):
        TensorSchema("", Domain.product(years(2025)), value_types=(float,))


def test_large_sparse_domain_does_not_expand_product() -> None:
    axes = tuple(
        Axis(name, tuple(range(1000)), int) for name in ("instrument", "holder", "vintage", "year")
    )
    domain = Domain.explicit(axes=axes, coordinates=((999, 999, 999, 999), (0, 0, 0, 0)))
    assert len(domain) == 2
    assert tuple(domain) == ((0, 0, 0, 0), (999, 999, 999, 999))


def test_workbook_date_values_serialize_without_coercion() -> None:
    values = (date(2025, 1, 1), datetime(2026, 1, 1, 12, 30))
    tensor = Tensor.from_nested(domain=Domain.product(years(2025, 2026)), values=values)
    restored = Tensor.from_json(tensor.to_json())
    assert restored == tensor
    assert type(restored[2025]) is date
    assert type(restored[2026]) is datetime


def test_domain_positions_are_shared_and_tensors_hold_only_values() -> None:
    import tracemalloc

    from excel_grapher.exporter.export_runtime.tensor import Axis, Domain, Tensor

    years = Axis("year", tuple(range(2000, 2100)), int)
    countries = Axis("country", tuple(f"c{i}" for i in range(100)), str)
    product = Domain.product(countries, years)
    assert product.position(("c1", 2000)) == 100
    assert product.position(("c99", 2099)) == 9999
    sparse = Domain.explicit(axes=(countries, years), coordinates=[("c3", 2001), ("c1", 2005)])
    assert sparse.position(("c1", 2005)) == 0
    assert sparse.position(("c3", 2001)) == 1
    values = tuple(float(i) for i in range(len(product)))
    first = Tensor(product, values)
    tracemalloc.start()
    second = Tensor(product, values)
    _current, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    assert second[("c1", 2000)] == 100.0
    assert first.domain is second.domain
    assert peak < 40 * len(product), peak
