"""Tests for optional setters module generation."""

from __future__ import annotations

import sys

from excel_grapher import DependencyGraph, Node
from excel_grapher.core.address_keys import parse_address
from excel_grapher.exporter.codegen import CodeGenerator
from excel_grapher.exporter.input_groups import SetterGenerationOptions
from tests.integration.utils.generated_package import import_generated_package, purge_module_cache
from tests.integration.utils.parity_harness import records_to_address_dict


def _make_node(address: str, formula: str | None, value: object) -> Node:
    sheet, coord = parse_address(address)
    col = "".join(c for c in coord if c.isalpha())
    row = int("".join(c for c in coord if c.isdigit()))
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=formula,
        value=value,
        is_leaf=formula is None,
    )


def test_generate_modules_without_setters_unchanged_file_set(tmp_path) -> None:
    graph = DependencyGraph()
    graph.add_node(_make_node("S!A1", None, 1.0))
    files = CodeGenerator(graph).generate_modules(["S!A1"])
    assert "exported/setters.py" not in files
    assert set(files.keys()) == {
        "exported/__init__.py",
        "exported/constants.py",
        "exported/entrypoint.py",
        "exported/inputs.py",
        "exported/internals.py",
        "exported/runtime.py",
    }


def test_generate_modules_with_setters_emits_setters_py(tmp_path) -> None:
    graph = DependencyGraph()
    graph.add_node(_make_node("S!A1", None, 10.0))
    graph.add_node(_make_node("S!B1", "=S!A1*2", None))
    files = CodeGenerator(graph).generate_modules(
        ["S!B1"],
        setters=SetterGenerationOptions(),
    )
    assert "exported/setters.py" in files
    assert "def set_" in files["exported/setters.py"]
    assert "setters" in files["exported/__init__.py"]


def test_generated_compute_all_returns_records(tmp_path) -> None:
    graph = DependencyGraph()
    graph.add_node(_make_node("S!A1", None, 10.0))
    graph.add_node(_make_node("S!B1", "=S!A1*2", None))
    package_name = "exported_setters_records"
    files = CodeGenerator(graph).generate_modules(["S!B1"], package_name=package_name)
    pkg = import_generated_package(tmp_path, files, package_name=package_name)
    try:
        records = pkg.compute_all()
        assert isinstance(records, list)
        by_addr = records_to_address_dict(records)
        assert by_addr["S!B1"] == 20.0
    finally:
        sys.path.remove(str(tmp_path))
        purge_module_cache(package_name)
