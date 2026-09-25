"""Labeller lookup is an index, not a scan of every series (#1000).

`labeller_for` runs once per coordinate during emit, including when the
workbook has no runtime labeller. The catalog indexes series that declare
`axis_labels` when it is built and reuses those covered-key sets.
"""

from __future__ import annotations

from collections.abc import Sequence

import pytest

from excel_grapher.exporter.inverted_tree.catalog import (
    BoundSeries,
    KeyPoint,
    SeriesCatalog,
    Statement,
)
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from tests.unit.exporter.inverted_tree.helpers import make_catalog


class _CountingDict(dict[str, BoundSeries]):
    """Dict that counts full scans of its contents."""

    def __init__(self, *args: object, **kwargs: object) -> None:
        super().__init__(*args, **kwargs)  # type: ignore[arg-type]
        self.scans = 0

    def values(self):  # type: ignore[override]
        self.scans += 1
        return super().values()

    def items(self):  # type: ignore[override]
        self.scans += 1
        return super().items()

    def __iter__(self):
        self.scans += 1
        return super().__iter__()

    def __getitem__(self, key: str) -> BoundSeries:
        self.scans += 1
        return super().__getitem__(key)


def _series(
    series_id: str,
    values: Sequence[int | str],
    *,
    key: str | None = "TIME_PERIOD",
    direction: str = "input",
    axis_labels: str | None = None,
) -> BoundSeries:
    if key is None:
        cells = (f"{series_id}!A1",)
        domain = (KeyPoint(()),)
        key_fields: tuple[str, ...] = ()
    else:
        cells = tuple(f"{series_id}!A{index + 1}" for index in range(len(values)))
        domain = tuple(KeyPoint(((key, value),)) for value in values)
        key_fields = (key,)
    return BoundSeries(
        series_id=series_id,
        layout="series" if key_fields else "scalar",
        direction=direction,  # type: ignore[arg-type]
        cells=cells,
        key_fields=key_fields,
        dtype="int" if key_fields else "float",
        compute_name=None,
        raw={},
        domain=domain,
        statements=(Statement(series_id, series_id, None, 0, len(cells), cells, domain),),
        axis_labels=axis_labels,
    )


def _catalog(*series: BoundSeries) -> tuple[SeriesCatalog, _CountingDict]:
    mapping = _CountingDict((item.series_id, item) for item in series)
    catalog = make_catalog(
        mapping,
        tuple(item.series_id for item in series),
        {cell: item.series_id for item in series for cell in item.cells},
    )
    return catalog, mapping


def test_labeller_for_matches_the_covering_series() -> None:
    history = _series("history", [2020, 2021], direction="internal", axis_labels="TIME_PERIOD")
    projection = _series(
        "projection", [2022, 2023], direction="internal", axis_labels="TIME_PERIOD"
    )
    areas = _series(
        "areas", ["US", "CA"], key="REF_AREA", direction="internal", axis_labels="REF_AREA"
    )
    unlabeled = _series("plain", [2020, 2021, 2022, 2023], direction="output")
    catalog, _scans = _catalog(_series("extra", [1]), history, projection, areas, unlabeled)

    assert catalog.labeller_for("TIME_PERIOD", (2020, 2021)) is history
    assert catalog.labeller_for("TIME_PERIOD", (2020,)) is history
    assert catalog.labeller_for("TIME_PERIOD", (2020, 2020)) is history
    assert catalog.labeller_for("TIME_PERIOD", ("2020", "2021")) is None
    assert catalog.labeller_for("TIME_PERIOD", (2022, 2023)) is projection
    assert catalog.labeller_for("TIME_PERIOD", (2021, 2022)) is None
    assert catalog.labeller_for("TIME_PERIOD", (2020, 2021, 2022, 2023)) is None
    assert catalog.labeller_for("REF_AREA", ("US",)) is areas
    assert catalog.labeller_for("REF_AREA", ("US", "CA", "MX")) is None
    assert catalog.labeller_for("COUNTRY", (2020, 2021)) is None
    assert catalog.runtime_labeller("TIME_PERIOD", (2020, 2021)) is history


def test_labeller_for_errors_when_two_labellers_cover_the_same_keys() -> None:
    first = _series("first", [2020, 2021, 2022], direction="internal", axis_labels="TIME_PERIOD")
    second = _series(
        "second", [2019, 2020, 2021, 2022], direction="internal", axis_labels="TIME_PERIOD"
    )
    catalog, _scans = _catalog(first, second)
    with pytest.raises(InvertedTreeExportError, match="multiple labellers"):
        catalog.labeller_for("TIME_PERIOD", (2020, 2021))


def test_runtime_labeller_ignores_constant_labellers() -> None:
    years = _series("years", [2020, 2021], direction="constant", axis_labels="TIME_PERIOD")
    catalog, _scans = _catalog(years)
    assert catalog.labeller_for("TIME_PERIOD", (2020, 2021)) is years
    assert catalog.runtime_labeller("TIME_PERIOD", (2020, 2021)) is None


def test_labeller_for_does_not_rescan_the_catalog(monkeypatch: pytest.MonkeyPatch) -> None:
    years = _series("years", [2020, 2021, 2022], direction="internal", axis_labels="TIME_PERIOD")
    pinned = _series("pinned", [], key=None, axis_labels="TIME_PERIOD")
    fillers = [_series(f"extra_{index}", [index]) for index in range(400)]
    catalog, scans = _catalog(*fillers, pinned, years)
    indexed = catalog._labellers_by_axis["TIME_PERIOD"]
    assert [series.series_id for series, _covered in indexed] == ["years"]
    assert indexed[0][1] == frozenset({2020, 2021, 2022})
    assert set(catalog._labellers_by_axis) == {"TIME_PERIOD"}
    before = scans.scans
    accesses = {"n": 0}
    original = BoundSeries.tensor_domain.fget
    assert original is not None

    def counting(self: BoundSeries):
        accesses["n"] += 1
        return original(self)

    monkeypatch.setattr(BoundSeries, "tensor_domain", property(counting))
    for _ in range(50):
        assert catalog.labeller_for("TIME_PERIOD", (2020, 2021)) is years
        assert catalog.labeller_for("MISSING", (2020,)) is None
        assert catalog.runtime_labeller("TIME_PERIOD", (2022,)) is years
    assert scans.scans == before
    assert accesses["n"] == 0
