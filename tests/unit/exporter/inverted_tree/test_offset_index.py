"""OFFSET(INDEX(...)) lowers to a lookup of the destination series (#729)."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.core.address_keys import as_canonical
from excel_grapher.core.formula_ast import FunctionCallNode, parse
from excel_grapher.exporter.inverted_tree import InvertedTreeExportError
from excel_grapher.exporter.inverted_tree.deps import (
    ast_literal_int,
    collect_all_dependence_edges,
    offset_index_destination,
)
from tests.unit.exporter.inverted_tree.helpers import (
    all_param_names,
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    series_entry,
    write_workbook,
)


def test_offset_index_destination_applies_literal_column_shift() -> None:
    ast = parse("=OFFSET(INDEX(Lookup!C4:C5,1,1),0,-1)")
    assert isinstance(ast, FunctionCallNode)
    dest = offset_index_destination(ast, as_canonical("Engine!A1"))
    assert dest == (as_canonical("Lookup!B4"), as_canonical("Lookup!B5"))
    minus = ast.args[2]
    assert ast_literal_int(minus) == -1


def _mcve_workbook(tmp_path: Path) -> Path:
    """Issue 729 MCVE: OFFSET's reference is INDEX, not a cell or range token."""
    return write_workbook(
        tmp_path / "offset_index_mcve.xlsx",
        {
            "Lookup": {"B4": 111.0, "C4": "Afghanistan"},
            "Engine": {"A1": "=OFFSET(INDEX(Lookup!C4:C4,1,1),0,-1)"},
            "Outputs": {"A1": "=Engine!A1"},
        },
    )


def _mcve_bindings() -> dict:
    return bindings_document(
        series_entry("code", "Lookup!B4", layout="scalar", direction="input"),
        series_entry("offset_index", "Engine!A1", layout="scalar", direction="internal"),
        series_entry("result", "Outputs!A1", layout="scalar", direction="output"),
    )


def _country_workbook(
    tmp_path: Path,
    *,
    engine_formulas: dict[str, str],
    defined_names: dict[str, str] | None = None,
) -> Path:
    return write_workbook(
        tmp_path / "offset_index_countries.xlsx",
        {
            "Lookup": {
                "A4": "AF",
                "B4": 111.0,
                "C4": "Afghanistan",
                "A5": "BR",
                "B5": 222.0,
                "C5": "Brazil",
            },
            "Engine": {
                "A1": "AF",
                "A2": "BR",
                **engine_formulas,
            },
            "Outputs": {
                "A1": "AF",
                "A2": "BR",
                "B1": "=Engine!B1",
                "B2": "=Engine!B2",
            },
        },
        defined_names=defined_names,
    )


def _country_bindings(*, engine_id: str = "imported") -> dict:
    return bindings_document(
        series_entry(
            "codes",
            "Lookup!B4:B5",
            layout="series",
            direction="input",
            label_column="A",
            key_concept="COUNTRY",
            key_read="string",
        ),
        series_entry(
            "names",
            "Lookup!C4:C5",
            layout="series",
            direction="constant",
            dtype="string",
            label_column="A",
            key_concept="COUNTRY",
            key_read="string",
        ),
        series_entry(
            engine_id,
            "Engine!B1:B2",
            layout="series",
            direction="internal",
            label_column="A",
            key_concept="COUNTRY",
            key_read="string",
        ),
        series_entry(
            "result",
            "Outputs!B1:B2",
            layout="series",
            direction="output",
            label_column="A",
            key_concept="COUNTRY",
            key_read="string",
        ),
    )


def test_offset_index_mcve_emits_destination_lookup(tmp_path: Path) -> None:
    workbook = _mcve_workbook(tmp_path)
    catalog, _deps, graph = inverted_graph_parts(workbook, _mcve_bindings())
    assert sorted(graph.leaf_keys()) == ["Lookup!B4"]
    assert sorted(graph.formula_keys()) == ["Engine!A1", "Outputs!A1"]
    edges = collect_all_dependence_edges(catalog, graph)
    producers = {edge.producer_id for edge in edges if edge.consumer_id == "offset_index"}
    assert producers == {"code"}

    modules = generate_inverted(workbook, _mcve_bindings())
    internals = modules["internals.py"]
    assert "xl_offset" not in internals
    assert "FunctionCallNode" not in internals
    pkg = load_package(modules, tmp_path, name="offset_index_mcve")
    assert "code" in all_param_names(pkg.internals.offset_index)
    assert pkg.internals.offset_index(111.0) == pytest.approx(111.0)
    assert pkg.compute_result(code=111.0) == pytest.approx((111.0,))


def test_offset_index_steps_onto_adjacent_column_series(tmp_path: Path) -> None:
    workbook = _country_workbook(
        tmp_path,
        engine_formulas={
            "B1": "=OFFSET(INDEX(Lookup!C4:C5,1,1),0,-1)",
            "B2": "=OFFSET(INDEX(Lookup!C4:C5,2,1),0,-1)",
        },
    )
    modules = generate_inverted(workbook, _country_bindings())
    pkg = load_package(modules, tmp_path, name="offset_index_codes")
    params = all_param_names(pkg.internals.imported)
    assert "codes" in params
    assert "names" not in params
    assert pkg.compute_result(codes=(111.0, 222.0)) == pytest.approx((111.0, 222.0))


def test_offset_index_zero_offset_stays_on_index_column(tmp_path: Path) -> None:
    workbook = _country_workbook(
        tmp_path,
        engine_formulas={
            "B1": "=OFFSET(INDEX(Lookup!C4:C5,1,1),0,0)",
            "B2": "=OFFSET(INDEX(Lookup!C4:C5,2,1),0,0)",
        },
    )
    modules = generate_inverted(workbook, _country_bindings(engine_id="imported_names"))
    pkg = load_package(modules, tmp_path, name="offset_index_names")
    params = all_param_names(pkg.internals.imported_names)
    assert "names" in params
    assert "codes" not in params
    assert pkg.compute_result() == ("Afghanistan", "Brazil")


def test_offset_index_named_range_array(tmp_path: Path) -> None:
    workbook = _country_workbook(
        tmp_path,
        engine_formulas={
            "B1": "=OFFSET(INDEX(Country_list,1,1),0,-1)",
            "B2": "=OFFSET(INDEX(Country_list,2,1),0,-1)",
        },
        defined_names={"Country_list": "Lookup!$C$4:$C$5"},
    )
    pkg = load_package(
        generate_inverted(workbook, _country_bindings()),
        tmp_path,
        name="offset_index_named",
    )
    assert pkg.compute_result(codes=(111.0, 222.0)) == pytest.approx((111.0, 222.0))


def test_offset_index_unbound_destination_fail_closed(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "offset_index_unbound.xlsx",
        {
            "Lookup": {"C4": "Afghanistan"},
            "Engine": {"A1": "=OFFSET(INDEX(Lookup!C4:C4,1,1),0,-1)"},
            "Outputs": {"A1": "=Engine!A1"},
        },
    )
    document = bindings_document(
        series_entry("name", "Lookup!C4", layout="scalar", direction="input", dtype="string"),
        series_entry("offset_index", "Engine!A1", layout="scalar", direction="internal"),
        series_entry("result", "Outputs!A1", layout="scalar", direction="output"),
    )
    with pytest.raises(InvertedTreeExportError, match="not a bound series|not in any bound series"):
        generate_inverted(workbook, document)
