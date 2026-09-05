"""View-level groups sequence single-file `generate()` setters (`list_groups`).

Package exports (`generate_modules`) omit `list_groups`; groups remain a
`generate()` presentation concern.
"""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Any, cast

import pytest

from excel_grapher.exporter import CodeGenerator
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import expand_data_range, validate_bindings_document
from tests.integration.user_flows.utils import write_series_bindings_workbook


def _row_series(series_id: str, row: int, groups: list[dict[str, Any]] | None) -> dict[str, Any]:
    entry: dict[str, Any] = {
        "id": series_id,
        "sheet": "Sheet1",
        "data_range": f"Sheet1!F{row}:J{row}",
        "layout": "series",
        "setter": {"name": f"set_{series_id}"},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "float",
                "bind": {"kind": "data_cell", "read": "float"},
            },
            "dimensions": [
                {
                    "concept": "TIME_PERIOD",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
                }
            ],
        },
        "key": ["TIME_PERIOD"],
    }
    if groups is not None:
        entry["groups"] = groups
    return entry


GROUPED_DOCUMENT: dict[str, Any] = {
    "schema_version": "1.5.0",
    "workbook": "series_bindings.xlsx",
    "series": [
        _row_series("primary_balance", 5, [{"path": ["Fiscal"], "order": 1}]),
        _row_series("gdp_growth", 3, [{"path": ["Macro", "Growth"]}]),
        _row_series("interest_rate", 4, None),
    ],
}


@pytest.fixture
def workbook(tmp_path: Path) -> Path:
    path = tmp_path / "series_bindings.xlsx"
    write_series_bindings_workbook(path)
    return path


def _generate(workbook: Path, document: dict[str, Any] = GROUPED_DOCUMENT) -> str:
    bindings = validate_bindings_document(deepcopy(document))
    targets: list[str] = []
    for series in bindings["series"]:
        targets.extend(expand_data_range(series["data_range"], workbook=workbook))
    graph = create_dependency_graph(workbook, targets, load_values=True)
    with CodeGenerator(graph) as gen:
        return gen.generate(
            targets,
            series_bindings=bindings,
            bindings_workbook=workbook,
        )


def test_grouped_export_sequences_definitions(workbook: Path) -> None:
    code = _generate(workbook)
    assert (
        code.index("def set_primary_balance(")
        < code.index("def set_gdp_growth(")
        < code.index("def set_interest_rate(")
    )


def test_grouped_export_emits_list_groups_discovery(workbook: Path) -> None:
    code = _generate(workbook)
    assert "def list_groups(" in code
    namespace: dict[str, Any] = {}
    exec(code, namespace)
    groups = cast(dict[str, Any], namespace["list_groups"]())
    assert [g["label"] for g in groups["groups"]] == ["Fiscal", "Macro"]


def test_grouped_export_preserves_setter_semantics(workbook: Path) -> None:
    namespace: dict[str, Any] = {}
    exec(_generate(workbook), namespace)
    ctx = namespace["make_context"]()
    namespace["set_primary_balance"](ctx, [{"TIME_PERIOD": 4, "OBS_VALUE": 7.5}])
    assert ctx.inputs["Sheet1!I5"] == 7.5


def test_ungrouped_bindings_export_omits_list_groups(workbook: Path) -> None:
    document = deepcopy(GROUPED_DOCUMENT)
    for series in document["series"]:
        series.pop("groups", None)
    code = _generate(workbook, document)
    assert "def list_groups(" not in code
    assert (
        code.index("def set_primary_balance(")
        < code.index("def set_gdp_growth(")
        < code.index("def set_interest_rate(")
    )
