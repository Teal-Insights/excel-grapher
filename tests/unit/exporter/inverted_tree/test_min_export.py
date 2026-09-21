"""MIN export uses shared aggregate semantics."""

from pathlib import Path

import pytest

from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    invoke_public_compute,
    load_package,
    series_entry,
    write_workbook,
)


@pytest.mark.parametrize("formula", ["=MIN(A1,A2)", "=MIN(A1:A2)"])
def test_min_export(tmp_path: Path, formula: str) -> None:
    workbook = write_workbook(
        tmp_path / "minimum.xlsx", {"Engine": {"A1": 2, "A2": 0, "B1": formula}}
    )
    document = bindings_document(
        series_entry("first", "Engine!A1"),
        series_entry("second", "Engine!A2"),
        series_entry("result", "Engine!B1", direction="output"),
    )
    package = load_package(generate_inverted(workbook, document), tmp_path)
    assert (
        invoke_public_compute(package, package.compute_result, dict(first=2.0, second=0.0)) == 0.0
    )
    assert (
        invoke_public_compute(package, package.compute_result, dict(first=-3.0, second=1.0)) == -3.0
    )
    assert (
        invoke_public_compute(package, package.compute_result, dict(first="#DIV/0!", second=1.0))
        == "#DIV/0!"
    )
