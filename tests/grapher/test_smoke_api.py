from __future__ import annotations


def test_public_api_imports() -> None:
    """
    Basic smoke test that the library imports and exposes the expected public API.
    """
    import excel_grapher as eg

    assert eg.create_dependency_graph is not None
    assert eg.DependencyGraph is not None
    assert eg.Node is not None
    assert eg.ValueType is not None
    assert eg.to_graphviz is not None
    assert eg.to_mermaid is not None
    assert eg.to_networkx is not None
    assert eg.validate_graph is not None
    assert eg.FromWorkbook is not None
    assert eg.GreaterThanCell is not None
    assert eg.NotEqualCell is not None

