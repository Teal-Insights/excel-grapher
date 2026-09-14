"""Shared `emit_runtime` module lists for exporter tests."""

from __future__ import annotations

from pathlib import Path

from excel_grapher.exporter.embed import _core_modules, emit_runtime

_PACKAGE_ROOT = Path(__file__).resolve().parents[3]
_EXPORT_RUNTIME = _PACKAGE_ROOT / "excel_grapher" / "exporter" / "export_runtime"

_EXPORT_RUNTIME_MODULES: tuple[tuple[str, Path], ...] = (
    ("export_runtime.errors", _EXPORT_RUNTIME / "errors.py"),
    ("export_runtime.ranges", _EXPORT_RUNTIME / "ranges.py"),
    ("export_runtime.values", _EXPORT_RUNTIME / "values.py"),
    ("export_runtime.lookup", _EXPORT_RUNTIME / "lookup.py"),
    ("export_runtime.operators", _EXPORT_RUNTIME / "operators.py"),
    ("export_runtime.aggregates", _EXPORT_RUNTIME / "aggregates.py"),
    ("export_runtime.math", _EXPORT_RUNTIME / "math.py"),
    ("export_runtime.text", _EXPORT_RUNTIME / "text.py"),
    ("export_runtime.logic", _EXPORT_RUNTIME / "logic.py"),
    ("export_runtime.reference", _EXPORT_RUNTIME / "reference.py"),
    ("export_runtime.offset", _EXPORT_RUNTIME / "offset.py"),
    ("export_runtime.info", _EXPORT_RUNTIME / "info.py"),
    ("export_runtime.error_funcs", _EXPORT_RUNTIME / "error_funcs.py"),
)

_MODULES: tuple[tuple[str, Path], ...] = (
    *_core_modules(include_operators_fastpath=True),
    *_EXPORT_RUNTIME_MODULES,
)


def emit_export_runtime(required_symbols: set[str]) -> str:
    """Emit core plus export-runtime symbols for wrapper-boundary tests."""
    return emit_runtime(required_symbols, modules=_MODULES)
