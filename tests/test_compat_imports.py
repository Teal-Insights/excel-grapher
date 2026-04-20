"""Regression guard for the ``excel_grapher.evaluator.{codegen,export_runtime}`` compat shims.

Issue #101 relocated the code generator and export runtime under
``excel_grapher.exporter``. Thin re-exports remain at the old evaluator paths so
that existing consumers keep working. These tests pin the shim contract: every
name listed here must resolve to the same object at both the old and new
locations.
"""

from __future__ import annotations

import importlib

import pytest

# (old_module, new_module, symbol)
_SHIM_SYMBOLS: list[tuple[str, str, str]] = [
    # codegen
    ("excel_grapher.evaluator.codegen", "excel_grapher.exporter.codegen", "CodeGenerator"),
    ("excel_grapher.evaluator.codegen", "excel_grapher.exporter.codegen", "GraphLike"),
    ("excel_grapher.evaluator.codegen", "excel_grapher.exporter.codegen", "GraphNode"),
    ("excel_grapher.evaluator.codegen", "excel_grapher.exporter.codegen", "GenerationParts"),
    # export_runtime.cache
    (
        "excel_grapher.evaluator.export_runtime.cache",
        "excel_grapher.exporter.export_runtime.cache",
        "CircularReferenceWarning",
    ),
    (
        "excel_grapher.evaluator.export_runtime.cache",
        "excel_grapher.exporter.export_runtime.cache",
        "EvalContext",
    ),
    (
        "excel_grapher.evaluator.export_runtime.cache",
        "excel_grapher.exporter.export_runtime.cache",
        "xl_cell",
    ),
    # export_runtime.core
    (
        "excel_grapher.evaluator.export_runtime.core",
        "excel_grapher.exporter.export_runtime.core",
        "CellValue",
    ),
    (
        "excel_grapher.evaluator.export_runtime.core",
        "excel_grapher.exporter.export_runtime.core",
        "XlError",
    ),
    (
        "excel_grapher.evaluator.export_runtime.core",
        "excel_grapher.exporter.export_runtime.core",
        "_format_general_number",
    ),
    # export_runtime.embed
    (
        "excel_grapher.evaluator.export_runtime.embed",
        "excel_grapher.exporter.export_runtime.embed",
        "emit_runtime",
    ),
    # export_runtime submodules (one canary symbol each)
    (
        "excel_grapher.evaluator.export_runtime.math",
        "excel_grapher.exporter.export_runtime.math",
        "xl_sum",
    ),
    (
        "excel_grapher.evaluator.export_runtime.text",
        "excel_grapher.exporter.export_runtime.text",
        "xl_concatenate",
    ),
    (
        "excel_grapher.evaluator.export_runtime.info",
        "excel_grapher.exporter.export_runtime.info",
        "xl_isnumber",
    ),
    (
        "excel_grapher.evaluator.export_runtime.logic",
        "excel_grapher.exporter.export_runtime.logic",
        "xl_and",
    ),
    (
        "excel_grapher.evaluator.export_runtime.lookup",
        "excel_grapher.exporter.export_runtime.lookup",
        "xl_vlookup",
    ),
    (
        "excel_grapher.evaluator.export_runtime.operators",
        "excel_grapher.exporter.export_runtime.operators",
        "xl_add",
    ),
    (
        "excel_grapher.evaluator.export_runtime.reference",
        "excel_grapher.exporter.export_runtime.reference",
        "xl_row",
    ),
    (
        "excel_grapher.evaluator.export_runtime.offset_runtime",
        "excel_grapher.exporter.export_runtime.offset_runtime",
        "xl_offset",
    ),
]


@pytest.mark.parametrize(("old_mod", "new_mod", "symbol"), _SHIM_SYMBOLS)
def test_shim_exports_same_object(old_mod: str, new_mod: str, symbol: str) -> None:
    old = getattr(importlib.import_module(old_mod), symbol)
    new = getattr(importlib.import_module(new_mod), symbol)
    assert old is new, f"{old_mod}.{symbol} diverged from {new_mod}.{symbol}"


def test_exporter_public_api() -> None:
    exporter = importlib.import_module("excel_grapher.exporter")
    canonical = importlib.import_module("excel_grapher.exporter.codegen")
    assert exporter.CodeGenerator is canonical.CodeGenerator
    # export_runtime is exposed as a subpackage attribute for `exporter.export_runtime.*`.
    assert exporter.export_runtime is importlib.import_module(
        "excel_grapher.exporter.export_runtime"
    )


def test_grapher_does_not_import_evaluator_or_exporter() -> None:
    """#101 boundary rule: grapher must not depend on evaluator or exporter."""
    import pkgutil

    import excel_grapher.grapher as grapher_pkg

    offenders: list[tuple[str, str]] = []
    for mod_info in pkgutil.walk_packages(grapher_pkg.__path__, prefix="excel_grapher.grapher."):
        mod = importlib.import_module(mod_info.name)
        for name, value in vars(mod).items():
            if not hasattr(value, "__module__"):
                continue
            origin = getattr(value, "__module__", "") or ""
            if origin.startswith("excel_grapher.evaluator") or origin.startswith(
                "excel_grapher.exporter"
            ):
                offenders.append((mod_info.name, f"{name} <- {origin}"))
    assert not offenders, f"grapher leaked imports from evaluator/exporter: {offenders}"
