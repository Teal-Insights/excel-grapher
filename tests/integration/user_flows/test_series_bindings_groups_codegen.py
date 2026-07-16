"""Integration: view-level groups sequence the exported Records API without changing semantics."""

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


# Declaration order deliberately interleaves groups so grouped export must resequence:
# primary_balance (Fiscal) is declared first but Macro/Growth groups appear first
# in the group tree only if first-appearance ordering is honored.
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


def _generate_modules(workbook: Path) -> dict[str, str]:
    bindings = validate_bindings_document(deepcopy(GROUPED_DOCUMENT))
    targets: list[str] = []
    for series in bindings["series"]:
        targets.extend(expand_data_range(series["data_range"], workbook=workbook))
    graph = create_dependency_graph(workbook, targets, load_values=True)
    with CodeGenerator(graph) as gen:
        return gen.generate_modules(
            targets,
            series_bindings=bindings,
            bindings_workbook=workbook,
        )


def test_grouped_export_sequences_definitions_and_all(workbook: Path) -> None:
    modules = _generate_modules(workbook)
    api = modules["api.py"]

    # Grouped bindings first (Fiscal, then Macro/Growth), ungrouped trail.
    assert (
        api.index("def set_primary_balance(")
        < api.index("def set_gdp_growth(")
        < api.index("def set_interest_rate(")
    )

    init = modules["__init__.py"]
    all_line = next(line for line in init.splitlines() if line.startswith("__all__"))
    assert all_line.index("set_primary_balance") < all_line.index("set_gdp_growth")
    assert all_line.index("set_gdp_growth") < all_line.index("set_interest_rate")


def test_grouped_export_emits_list_groups_discovery(workbook: Path) -> None:
    modules = _generate_modules(workbook)

    assert "groups.json" not in modules
    assert "def list_groups(" in modules["api.py"]
    assert "list_groups" in modules["__init__.py"]


def test_grouped_export_preserves_setter_semantics(workbook: Path, tmp_path: Path) -> None:
    modules = _generate_modules(workbook)
    pkg_dir = tmp_path / "grouped_pkg"
    pkg_dir.mkdir()
    for filename, content in modules.items():
        (pkg_dir / filename).write_text(content, encoding="utf-8")

    import importlib
    import sys

    sys.path.insert(0, str(tmp_path))
    try:
        pkg = importlib.import_module("grouped_pkg")
        ctx = pkg.make_context()
        pkg.set_primary_balance(ctx, [{"TIME_PERIOD": 4, "OBS_VALUE": 7.5}])
        assert ctx.inputs["Sheet1!I5"] == 7.5

        groups = cast(dict[str, Any], pkg.list_groups())
        assert [g["label"] for g in groups["groups"]] == ["Fiscal", "Macro"]
    finally:
        sys.path.remove(str(tmp_path))
        for name in list(sys.modules):
            if name == "grouped_pkg" or name.startswith("grouped_pkg."):
                del sys.modules[name]


def test_ungrouped_bindings_export_is_unchanged(workbook: Path) -> None:
    document = deepcopy(GROUPED_DOCUMENT)
    for series in document["series"]:
        series.pop("groups", None)
    bindings = validate_bindings_document(document)
    targets: list[str] = []
    for series in bindings["series"]:
        targets.extend(expand_data_range(series["data_range"], workbook=workbook))
    graph = create_dependency_graph(workbook, targets, load_values=True)
    with CodeGenerator(graph) as gen:
        modules = gen.generate_modules(
            targets,
            series_bindings=bindings,
            bindings_workbook=workbook,
        )

    assert "groups.json" not in modules
    assert "list_groups" not in modules["api.py"]
    # Flat export keeps declaration order for definitions.
    api = modules["api.py"]
    assert (
        api.index("def set_primary_balance(")
        < api.index("def set_gdp_growth(")
        < api.index("def set_interest_rate(")
    )
