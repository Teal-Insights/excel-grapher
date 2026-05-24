"""Cross-workflow tests for Records contract parity."""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

from excel_grapher import DependencyGraph, Node
from excel_grapher.core.address_keys import parse_address
from excel_grapher.exporter.codegen import CodeGenerator
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


def test_single_file_and_modular_records_equivalent() -> None:
    graph = DependencyGraph()
    graph.add_node(_make_node("S!A1", None, 3.0))
    graph.add_node(_make_node("S!B1", "=S!A1*2", None))
    targets = ["S!B1"]

    single_code = CodeGenerator(graph).generate(targets)
    single_ns: dict = {}
    exec(single_code, single_ns)
    single_records = records_to_address_dict(single_ns["compute_all"]())

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        package_name = "exported_records_equiv"
        files = CodeGenerator(graph).generate_modules(targets, package_name=package_name)
        pkg = import_generated_package(tmp_path, files, package_name=package_name)
        try:
            modular_records = records_to_address_dict(pkg.compute_all())
        finally:
            sys.path.remove(str(tmp_path))
            purge_module_cache(package_name)

    assert single_records == modular_records == {"S!B1": 6.0}


def test_rectangular_entrypoint_records_row_major() -> None:
    graph = DependencyGraph()
    for addr, val in [("S!A1", 1.0), ("S!B1", 2.0), ("S!A2", 3.0), ("S!B2", 4.0)]:
        graph.add_node(_make_node(addr, None, val))
    graph.add_node(_make_node("S!C1", "=S!A1+S!B1+S!A2+S!B2", None))

    code = CodeGenerator(graph).generate(
        [],
        entrypoints={"block": ["S!A1:B2"]},
    )
    ns: dict = {}
    exec(code, ns)
    records = ns["compute_block"]()
    addresses = [rec["address"] for rec in records if "address" in rec]
    assert addresses == ["S!A1", "S!B1", "S!A2", "S!B2"]
