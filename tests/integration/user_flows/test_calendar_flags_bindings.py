"""Integration tests for bool flag + datetime header series bindings."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import (
    expand_data_range,
    load_series_bindings,
    resolve_series_binding,
    validate_series_bindings,
)
from tests.integration.user_flows.utils import write_calendar_flags_workbook
from tests.paths import SERIES_BINDINGS_FIXTURES as FIXTURES


def test_calendar_flags_workbook_validates_and_resolves(tmp_path: Path) -> None:
    wb_path = tmp_path / "calendar_flags.xlsx"
    write_calendar_flags_workbook(wb_path)
    bindings = load_series_bindings(FIXTURES / "calendar_flags.yaml")
    series = bindings["series"][0]
    targets = expand_data_range(series["data_range"], workbook=wb_path)
    graph = create_dependency_graph(wb_path, targets, load_values=True)

    report = validate_series_bindings(graph, bindings, workbook=wb_path)
    assert report["ok"] is True

    resolved = resolve_series_binding(
        graph,
        wb_path,
        series,
        concept_scheme=bindings.get("concept_scheme"),
        direction="input",
    )
    assert resolved["ok"] is True
    assert resolved["requires_address"] is False
    keys = {(leaf["key"]["IS_ACTIVE"], leaf["key"]["TIME_PERIOD"]) for leaf in resolved["leaves"]}
    assert keys == {
        (True, datetime(2024, 1, 1)),
        (True, datetime(2024, 2, 1)),
        (True, datetime(2024, 3, 1)),
        (False, datetime(2024, 1, 1)),
        (False, datetime(2024, 2, 1)),
        (False, datetime(2024, 3, 1)),
    }
