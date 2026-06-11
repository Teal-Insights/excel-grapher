"""Public `excel_grapher` package exports remain importable and discoverable (integration).

Smoke-checks the advertised API surface so packaging, re-exports, and optional
extras do not silently drop symbols users rely on in notebooks and apps.
"""

from __future__ import annotations

from types import ModuleType


def test_public_api_imports() -> None:
    """Basic smoke test that the library imports and exposes the expected public API."""
    import excel_grapher as eg

    assert isinstance(eg.exporter, ModuleType)
    assert isinstance(eg.grapher, ModuleType)
    assert isinstance(eg.series_bindings, ModuleType)
    assert "exporter" in eg.__all__
    assert "grapher" in eg.__all__
    assert "series_bindings" in eg.__all__
    assert eg.create_dependency_graph is not None
    assert eg.CodeGenerator is not None
    assert eg.DependencyGraph is not None
    assert eg.Node is not None
    assert eg.to_graphviz is not None
    assert eg.to_mermaid is not None
    assert eg.to_networkx is not None
    assert eg.validate_graph is not None
    assert eg.FromWorkbook is not None
    assert eg.GreaterThanCell is not None
    assert eg.NotEqualCell is not None
    assert eg.RealBetween is not None
    assert eg.RealIntervalDomain is not None
