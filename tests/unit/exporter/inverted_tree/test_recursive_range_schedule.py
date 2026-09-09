"""Recursive range reads require catalog-instance scheduling."""

from pathlib import Path

import pytest

from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)


@pytest.mark.parametrize("window", ["B2:C2", "$B$2:$C$2"])
@pytest.mark.parametrize("second", ["=2", "=B2+1"])
def test_recursive_sum_selects_available_catalog_instances(
    tmp_path: Path, window: str, second: str
) -> None:
    moving = window == "B2:C2"
    workbook = write_workbook(
        tmp_path / "recursive.xlsx",
        {
            "Data": {
                "B1": 2020,
                "C1": 2021,
                "D1": 2022,
                "E1": 2023,
                "B2": "=1",
                "C2": second,
                "D2": f"=SUM({window})",
                "E2": "=SUM(C2:D2)" if moving else f"=SUM({window})",
            }
        },
    )
    bindings = bindings_document(
        series_entry(
            "totals",
            "Data!B2:E2",
            layout="series",
            direction="output",
            header_row=1,
            compute_name="compute_totals",
        )
    )
    modules = generate_inverted(workbook, bindings)
    package = load_package(modules, tmp_path, name=f"recursive_{moving}")
    assert dict(package.compute_totals().items()) == {
        (2020,): 1.0,
        (2021,): 2.0,
        (2022,): 3.0,
        (2023,): 5.0 if moving else 3.0,
    }
