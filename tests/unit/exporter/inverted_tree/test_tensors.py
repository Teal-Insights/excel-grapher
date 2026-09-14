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
    relabel_input,
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


def test_relabel_replaces_axis_keys_without_moving_values() -> None:
    scenario = Axis("scenario", ("base", "shock"), str)
    year = years(2024, 2025)
    tensor = Tensor.from_nested(domain=Domain.product(scenario, year), values=((1, 2), (3, 4)))
    labels = Tensor.from_nested(domain=Domain.product(year), values=(2026, 2027))

    relabelled = tensor.relabel(year=labels)

    assert list(relabelled.items()) == [
        (("base", 2026), 1),
        (("base", 2027), 2),
        (("shock", 2026), 3),
        (("shock", 2027), 4),
    ]
    assert tensor.domain.axes[1].keys == (2024, 2025)


def test_relabel_rejects_invalid_labels() -> None:
    tensor = Tensor.from_nested(domain=Domain.product(years(2024, 2025)), values=(1, 2))
    for values, match in [((2026, 2026), "duplicate.*2026"), ((2026, "2027"), "year.*int")]:
        labels = Tensor.from_nested(domain=tensor.domain, values=values)
        with pytest.raises(AxisError, match=match):
            tensor.relabel(year=labels)


def test_relabel_input_maps_public_labels_to_snapshot_keys() -> None:
    labels = Tensor.from_nested(domain=Domain.product(years(2024, 2025)), values=(2026, 2027))
    public = Tensor.from_nested(domain=Domain.product(years(2026, 2027)), values=(10, 20))

    internal = relabel_input(public, labels, "year", series_id="investment")

    assert list(internal.items()) == [((2024,), 10), ((2025,), 20)]
    stale = Tensor.from_nested(domain=Domain.product(years(2024, 2025)), values=(10, 20))
    with pytest.raises(SchemaError, match="investment.*2024.*2026.*2027"):
        relabel_input(stale, labels, "year", series_id="investment")

    reversed_public = Tensor.from_nested(domain=Domain.product(years(2027, 2026)), values=(20, 10))
    assert list(relabel_input(reversed_public, labels, "year", series_id="investment").items()) == [
        ((2024,), 10),
        ((2025,), 20),
    ]


def test_relabel_input_rejects_duplicate_runtime_labels() -> None:
    labels = Tensor.from_nested(domain=Domain.product(years(2024, 2025)), values=(2026, 2026))
    public = Tensor.from_nested(domain=Domain.product(years(2026)), values=(10,))

    with pytest.raises(AxisError, match="duplicate.*2026"):
        relabel_input(public, labels, "year", series_id="investment")


def test_schema_can_validate_structure_before_runtime_coordinates_exist() -> None:
    schema = TensorSchema("investment", Domain.product(years(2024, 2025)), value_types=(int,))
    shifted = Tensor.from_nested(domain=Domain.product(years(2026, 2027)), values=(10, 20))

    schema.validate_structure(shifted)
    with pytest.raises(SchemaError, match="missing required coordinate"):
        schema.validate(shifted)


def test_sparse_domain_order_holes_and_exact_records() -> None:
    domain = Domain.explicit(axes=(years(2026, 2025, 2027),), coordinates=((2027,), (2026,)))
    tensor = Tensor.from_records(domain=domain, records=(((2027,), None), ((2026,), 0)))
    assert list(tensor.items()) == [((2026,), 0), ((2027,), None)]
    assert tensor[2025] is None
    with pytest.raises(CoordinateError, match="2024"):
        tensor[2024]
    for records in [
        (((2026,), 1),),
        (((2026,), 1), ((2027,), 2), ((2025,), 3)),
        (((2026,), 1), ((2026,), 2)),
    ]:
        with pytest.raises(DomainError):
            Tensor.from_records(domain=domain, records=records)


def test_sparse_get_blank_vs_invalid_key() -> None:
    domain = Domain.explicit(axes=(years(2026, 2025, 2027),), coordinates=((2027,), (2026,)))
    assert domain.sparse_get((2026,)) == domain.position((2026,))
    assert domain.sparse_get((2027,)) == domain.position((2027,))
    assert domain.sparse_get((2025,)) is None
    with pytest.raises(CoordinateError, match="2024"):
        domain.sparse_get((2024,))
    with pytest.raises(CoordinateError, match="absent from domain"):
        domain.position((2025,))


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


def test_sparse_serialization_round_trip() -> None:
    domain = Domain.explicit(axes=(years(2025, 2026),), coordinates=((2026,), (2025,)))
    tensor = Tensor.from_records(domain=domain, records=(((2026,), "#N/A"), ((2025,), None)))
    assert tensor[2026] == "#N/A"
    assert Tensor.from_json(tensor.to_json()) == tensor
    assert (
        domain.fingerprint
        == Domain.explicit(axes=domain.axes, coordinates=reversed(tuple(domain))).fingerprint
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
