"""Runtime primitives for label-keyed axes."""

from __future__ import annotations

from typing import Any, cast

import pytest

from excel_grapher.exporter.export_runtime.provenance import row_cells
from excel_grapher.exporter.export_runtime.tensor import (
    Axis,
    AxisError,
    AxisTemplate,
    CoordinateError,
    Domain,
    DomainTemplate,
    SchemaError,
    SchemaTemplate,
    Series,
    SeriesSpec,
    Tensor,
    label_axis,
)
from excel_grapher.exporter.inverted_tree.runtime import as_records, publish


def test_axis_position_returns_the_key_index() -> None:
    axis = Axis("TIME_PERIOD", (2024, 2025, 2026), int)
    assert axis.position(2024) == 0
    assert axis.position(2026) == 2
    with pytest.raises(CoordinateError, match="TIME_PERIOD"):
        axis.position(2023)


def test_label_axis_builds_an_axis_and_names_collisions() -> None:
    axis = label_axis("TIME_PERIOD", (2026, 2027), int)
    assert axis == Axis("TIME_PERIOD", (2026, 2027), int)
    with pytest.raises(AxisError, match="TIME_PERIOD.*duplicate label 2026"):
        label_axis("TIME_PERIOD", (2026, 2026), int)
    with pytest.raises(AxisError, match="TIME_PERIOD.*label 2026.5 must be int"):
        label_axis("TIME_PERIOD", (2026.5, 2027), int)


def test_series_from_labels_is_the_identity_map() -> None:
    axis = label_axis("TIME_PERIOD", (2026, 2027), int)
    spec = SeriesSpec[object](
        schema=SchemaTemplate(
            "year_labels",
            DomainTemplate.product(
                AxisTemplate("TIME_PERIOD", int, size=2, labeller="year_labels", snapshot=(1, 2))
            ),
            (int, str, type(None)),
        ),
        domain=DomainTemplate.product(
            AxisTemplate("TIME_PERIOD", int, size=2, labeller="year_labels", snapshot=(1, 2))
        ),
        cells={},
    )
    labels = spec.from_labels(axis)
    assert isinstance(labels, Series)
    assert tuple(labels.domain.axes[0].keys) == (2026, 2027)
    assert list(labels.items()) == [((2026,), 2026), ((2027,), 2027)]


def test_axis_template_bind_checks_name_type_and_size() -> None:
    template = AxisTemplate("TIME_PERIOD", int, size=2, labeller="labels", snapshot=(1, 2))
    axis = Axis("TIME_PERIOD", (2026, 2027), int)
    assert template.bind(axis) is axis
    with pytest.raises(AxisError, match="TIME_PERIOD"):
        template.bind(Axis("SCENARIO", (2026, 2027), int))
    with pytest.raises(AxisError, match="TIME_PERIOD"):
        template.bind(Axis("TIME_PERIOD", ("2026", "2027"), str))
    with pytest.raises(AxisError, match="TIME_PERIOD"):
        template.bind(Axis("TIME_PERIOD", (2026,), int))


def test_axis_template_bind_accepts_already_sliced_axis() -> None:
    """A subset template binds a full labeller or a tensor that is already sliced."""
    template = AxisTemplate(
        "TIME_PERIOD",
        int,
        size=3,
        labeller="year_labels",
        snapshot=(2026, 2027, 2028),
        source=(2024, 2025, 2026, 2027, 2028),
    )
    labeller = Axis("TIME_PERIOD", (2034, 2035, 2036, 2037, 2038), int)
    assert template.bind(labeller) == Axis("TIME_PERIOD", (2036, 2037, 2038), int)
    sliced = Axis("TIME_PERIOD", (2036, 2037, 2038), int)
    assert template.bind(sliced) is sliced
    with pytest.raises(AxisError, match="expected 3 keys"):
        template.bind(Axis("TIME_PERIOD", (2036, 2037), int))
    with pytest.raises(AxisError, match="expected 3 keys"):
        template.bind(Axis("TIME_PERIOD", (2034, 2035, 2036, 2037), int))


def test_schema_template_validate_accepts_sliced_tensor_own_axes() -> None:
    schema = SchemaTemplate(
        "growth",
        DomainTemplate.product(
            AxisTemplate(
                "TIME_PERIOD",
                int,
                size=3,
                labeller="year_labels",
                snapshot=(2026, 2027, 2028),
                source=(2024, 2025, 2026, 2027, 2028),
            )
        ),
        (int, float, str, type(None)),
    )
    tensor = Tensor.from_nested(
        domain=Domain.product(Axis("TIME_PERIOD", (2026, 2027, 2028), int)),
        values=(1.0, 2.0, 3.0),
    )
    schema.validate(tensor)
    labels = Tensor.from_nested(
        domain=Domain.product(Axis("TIME_PERIOD", (2024, 2025, 2026, 2027, 2028), int)),
        values=(2024, 2025, 2026, 2027, 2028),
    )
    schema.bind(TIME_PERIOD=labels).validate(tensor)


def test_domain_and_schema_templates_bind_labellers_by_axis_name() -> None:
    years = AxisTemplate("TIME_PERIOD", int, size=2, labeller="year_labels", snapshot=(1, 2))
    domain = DomainTemplate.product(years)
    labels = Tensor.from_nested(
        domain=Domain.product(Axis("TIME_PERIOD", (2026, 2027), int)), values=(2026, 2027)
    )
    bound = domain.bind(TIME_PERIOD=labels)
    assert bound == Domain.product(Axis("TIME_PERIOD", (2026, 2027), int))
    schema = SchemaTemplate("src", domain, (int, float, str, type(None)))
    tensor = Tensor.from_nested(domain=bound, values=(3, 4))
    schema.bind(TIME_PERIOD=labels).validate(tensor)
    schema.validate(tensor)
    with pytest.raises(
        SchemaError, match="unknown labels.*1.*accepted labels are \\(2026, 2027\\)"
    ):
        schema.bind(TIME_PERIOD=labels).validate(
            Tensor.from_nested(
                domain=Domain.product(Axis("TIME_PERIOD", (1, 2), int)), values=(3, 4)
            )
        )


def test_schema_template_validate_uses_the_tensor_own_axes() -> None:
    schema = SchemaTemplate(
        "out",
        DomainTemplate.product(
            AxisTemplate("TIME_PERIOD", int, size=2, labeller="labels", snapshot=(1, 2))
        ),
        (int, str, type(None)),
    )
    tensor = Tensor.from_nested(
        domain=Domain.product(Axis("TIME_PERIOD", (2030, 2031), int)), values=(6, 8)
    )
    schema.validate(tensor)
    too_short = Tensor.from_nested(
        domain=Domain.product(Axis("TIME_PERIOD", (2030,), int)), values=(6,)
    )
    with pytest.raises(SchemaError, match="expected 2 keys"):
        schema.validate(too_short)


def test_provenance_template_bind_rewrites_row_cells_over_runtime_keys() -> None:
    template = row_cells(
        "Engine",
        10,
        "C",
        AxisTemplate("TIME_PERIOD", int, size=3, labeller="labels", snapshot=(1, 2, 3)),
    )
    labels = Tensor.from_nested(
        domain=Domain.product(Axis("TIME_PERIOD", (2026, 2027, 2028), int)),
        values=(2026, 2027, 2028),
    )
    cells = template.bind(TIME_PERIOD=labels)
    assert dict(cells) == {
        (2026,): "Engine!C10",
        (2027,): "Engine!D10",
        (2028,): "Engine!E10",
    }


def test_publish_and_as_records_accept_a_schema_template() -> None:
    schema = SchemaTemplate(
        "out",
        DomainTemplate.product(
            AxisTemplate("TIME_PERIOD", int, size=2, labeller="labels", snapshot=(1, 2))
        ),
        (int, str, type(None)),
    )

    @publish(schema)
    def compute_out() -> Tensor[int]:
        return Tensor.from_nested(
            domain=Domain.product(Axis("TIME_PERIOD", (2030, 2031), int)), values=(6, 8)
        )

    result = compute_out()
    published = cast(Any, compute_out)
    assert published.__key__ == ("TIME_PERIOD",)
    assert isinstance(published.__domain__, DomainTemplate)
    assert as_records(compute_out, result) == [
        {"TIME_PERIOD": 2030, "OBS_VALUE": 6},
        {"TIME_PERIOD": 2031, "OBS_VALUE": 8},
    ]
