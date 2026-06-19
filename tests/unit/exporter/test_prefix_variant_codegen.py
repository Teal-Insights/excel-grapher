"""Unit tests for codegen export with ``_xlfn.`` and ``_xludf.`` formulas."""

from __future__ import annotations

import re

import pytest

from excel_grapher import DependencyGraph, Node
from excel_grapher.core.address_keys import parse_address
from excel_grapher.exporter.codegen import CodeGenerator

_XLUDF_RUNTIME_PATTERN = re.compile(r"xl__xludf_[a-z0-9_]+")


def _make_node(address: str, formula: str | None, value: object) -> Node:
    sheet, coord = parse_address(address)
    column = "".join(character for character in coord if character.isalpha())
    row = int("".join(character for character in coord if character.isdigit()))
    return Node(
        sheet=sheet,
        column=column,
        row=row,
        formula=formula,
        normalized_formula=formula,
        value=value,
        is_leaf=formula is None,
    )


def _make_graph(*nodes: Node) -> DependencyGraph:
    graph = DependencyGraph()
    for node in nodes:
        graph.add_node(node)
    return graph


def _lookup_graph(*, target_formula: str) -> DependencyGraph:
    return _make_graph(
        _make_node("S!A1", None, 1),
        _make_node("S!A2", None, 2),
        _make_node("S!A3", None, 3),
        _make_node("S!B1", None, "a"),
        _make_node("S!B2", None, "b"),
        _make_node("S!B3", None, "c"),
        _make_node("S!C1", target_formula, None),
    )


@pytest.mark.parametrize(
    ("formula", "expected_runtime"),
    [
        ("=_xludf.XLOOKUP(2,S!A1:S!A3,S!B1:S!B3)", "xl__xlfn_xlookup"),
        ("=_xlfn.XLOOKUP(2,S!A1:S!A3,S!B1:S!B3)", "xl__xlfn_xlookup"),
        ('=_xludf.NUMBERVALUE("1,234.56", ".", ",")', "xl__xlfn_numbervalue"),
    ],
)
def test_generate_modules_maps_prefixed_builtins_to_runtime_symbols(
    formula: str,
    expected_runtime: str,
) -> None:
    files = CodeGenerator(_lookup_graph(target_formula=formula)).generate_modules(["S!C1"])
    combined = "\n".join(files.values())
    assert _XLUDF_RUNTIME_PATTERN.search(combined) is None
    assert expected_runtime in combined


def test_generate_modules_nested_xludf_ifna_xlookup_uses_runtime_symbols() -> None:
    formula = '=_xludf.IFNA(_xludf.XLOOKUP(1,S!A1:S!A1,S!A1:S!A1),"x")'
    files = CodeGenerator(_lookup_graph(target_formula=formula)).generate_modules(["S!C1"])
    internals = files["internals.py"]
    assert _XLUDF_RUNTIME_PATTERN.search(internals) is None
    assert "xl_ifna" in internals or "xl__xlfn_xlookup" in internals


def test_generate_single_file_never_emits_xludf_runtime_symbols() -> None:
    formula = '=_xludf.IFNA(_xludf.XLOOKUP(1,S!A1:S!A1,S!B1:S!B1),"x")'
    code = CodeGenerator(_lookup_graph(target_formula=formula)).generate(["S!C1"])
    assert _XLUDF_RUNTIME_PATTERN.search(code) is None
    assert "xl__xlfn_xlookup" in code or "xl_ifna" in code
