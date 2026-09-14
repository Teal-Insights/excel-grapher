"""Generated `excel.py` holds Excel ops; `runtime.py` holds named-axis primitives."""

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
    {"publish", "view", "span", "take", "CoordinateReader", "as_records", "lazy_table"}
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


def _imported_modules(source: str) -> set[str]:
    tree = ast.parse(source)
    modules: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.ImportFrom) and node.module:
            modules.add(node.module)
        elif isinstance(node, ast.Import):
            modules.update(alias.name for alias in node.names)
    return modules


def test_generated_excel_owns_operators_runtime_owns_named_axis(tmp_path: Path) -> None:
    modules = _add_modules(tmp_path)
    excel = modules["excel.py"]
    runtime = modules["runtime.py"]
    internals = modules["internals.py"]
    excel_names = _top_level_names(excel)
    runtime_names = _top_level_names(runtime)

    assert "xl_add" in excel_names
    assert "xl_add" not in runtime_names
    assert excel.count("def xl_add") == 1
    assert "def xl_add" not in runtime
    assert runtime_names >= _NAMED_AXIS_HELPERS
    assert excel_names.isdisjoint(_NAMED_AXIS_HELPERS)

    assert "excel_grapher" not in _imported_modules(excel)
    assert "excel_grapher" not in _imported_modules(runtime)
    assert "from .excel import" in runtime
    assert "from .runtime import" not in excel
    assert "from .excel import" in internals
    assert "from .runtime import" in internals
    excel_import = next(
        line for line in internals.splitlines() if line.startswith("from .excel import")
    )
    runtime_import = next(
        line for line in internals.splitlines() if line.startswith("from .runtime import")
    )
    assert "xl_add" in excel_import
    assert "publish" in runtime_import
    assert "xl_add" not in runtime_import
    assert "publish" not in excel_import
    assert "def _core_add" in excel
    seen_body = False
    for node in ast.parse(excel).body:
        is_import = isinstance(node, ast.Import | ast.ImportFrom)
        if isinstance(node, ast.Expr) and isinstance(node.value, ast.Constant):
            continue
        if is_import:
            assert not seen_body
            continue
        seen_body = True
