"""Generated `excel.py` and `runtime.py` are layered, not duplicates (#843)."""

from __future__ import annotations

import ast
from pathlib import Path

from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    series_entry,
    write_workbook,
)

_NAMED_AXIS_HELPERS = frozenset(
    {"publish", "view", "span", "take", "CoordinateReader", "as_records"}
)


def _add_modules(tmp_path: Path) -> dict[str, str]:
    workbook = write_workbook(
        tmp_path / "standalone_layers.xlsx",
        {"S": {"A1": 1.0, "B1": 2.0, "C1": "=A1+B1"}},
    )
    return generate_inverted(
        workbook,
        bindings_document(
            series_entry("left", "S!A1"),
            series_entry("right", "S!B1"),
            series_entry("total", "S!C1", direction="output"),
        ),
    )


def _top_level_names(source: str) -> set[str]:
    tree = ast.parse(source)
    return {
        node.name
        for node in tree.body
        if isinstance(node, ast.FunctionDef | ast.AsyncFunctionDef | ast.ClassDef)
    }


def test_generated_excel_and_runtime_are_distinct_layers(tmp_path: Path) -> None:
    modules = _add_modules(tmp_path)
    excel = modules["excel.py"]
    runtime = modules["runtime.py"]
    internals = modules["internals.py"]
    excel_names = _top_level_names(excel)
    runtime_names = _top_level_names(runtime)

    assert excel != runtime
    assert "excel_grapher" not in excel
    assert "excel_grapher" not in runtime
    assert "from .excel import" in runtime
    assert "from .runtime import" not in excel
    assert "from .runtime import" in internals
    assert "from .excel import" not in internals

    assert runtime_names >= _NAMED_AXIS_HELPERS
    assert excel_names.isdisjoint(_NAMED_AXIS_HELPERS)
    assert "xl_add" in runtime_names
    assert "xl_add" in excel_names
    assert "Range" in excel_names
    assert "xl_add as _core_add" in runtime
    assert "_adapt_core(_core_add(" in runtime
