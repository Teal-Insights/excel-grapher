"""MIN export uses shared aggregate semantics."""

from pathlib import Path
from typing import Literal

import pytest

from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)


@pytest.mark.parametrize("force_rung", [None, 3])
@pytest.mark.parametrize("formula", ["=MIN(A1,A2)", "=MIN(A1:A2)"])
def test_min_export(tmp_path: Path, force_rung: Literal[3] | None, formula: str) -> None:
    workbook = write_workbook(
        tmp_path / "minimum.xlsx", {"Engine": {"A1": 2, "A2": 0, "B1": formula}}
    )
    document = bindings_document(
        series_entry("first", "Engine!A1"),
        series_entry("second", "Engine!A2"),
        series_entry("result", "Engine!B1", direction="output"),
    )
    package = load_package(generate_inverted(workbook, document, force_rung=force_rung), tmp_path)
    assert package.api.compute_result(first=2.0, second=0.0) == 0.0
    assert package.api.compute_result(first=-3.0, second=1.0) == -3.0
    assert package.api.compute_result(first="#DIV/0!", second=1.0) == "#DIV/0!"
