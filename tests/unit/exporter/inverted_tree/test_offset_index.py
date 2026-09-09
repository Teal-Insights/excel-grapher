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
    index_call_is_ref,
    offset_index_destination,
)
from excel_grapher.grapher.dynamic_refs import DynamicRefConfig
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
    assert pkg.internals.offset_index(code=111.0) == pytest.approx(111.0)
    assert pkg.compute_result(code=111.0) == pytest.approx(111.0)


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
    result = pkg.compute_result(
        codes=pkg.data.Codes.from_nested(domain=pkg.data.CODES_DOMAIN, values=(111.0, 222.0))
    )
    assert (result["AF"], result["BR"]) == pytest.approx((111.0, 222.0))


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
    result = pkg.compute_result()
    assert (result["AF"], result["BR"]) == ("Afghanistan", "Brazil")


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
    result = pkg.compute_result(
        codes=pkg.data.Codes.from_nested(domain=pkg.data.CODES_DOMAIN, values=(111.0, 222.0))
    )
    assert (result["AF"], result["BR"]) == pytest.approx((111.0, 222.0))


def _constraint_refs() -> DynamicRefConfig:
    return DynamicRefConfig.from_constraints({}, {})


def _row_select_sheets(*, formula: str) -> dict[str, dict[str, object]]:
    return {
        "Data": {
            "A1": 1,
            "A2": 2,
            "B1": 10.0,
            "B2": 20.0,
            "C1": formula,
            "C2": formula,
        }
    }


def _row_select_bindings(*, include_labels: bool = False) -> dict:
    series = []
    if include_labels:
        series.append(
            series_entry(
                "labels",
                "Data!A1:A2",
                layout="series",
                direction="constant",
                label_column="A",
            )
        )
    series.extend(
        [
            series_entry(
                "values",
                "Data!B1:B2",
                layout="series",
                direction="constant",
                label_column="A",
            ),
            series_entry(
                "selected",
                "Data!C1:C2",
                layout="series",
                direction="output",
                label_column="A",
            ),
        ]
    )
    return bindings_document(*series)


def test_index_call_is_ref_when_row_is_past_the_array() -> None:
    ast = parse("=INDEX(Data!$B$1:$B$2,ROW()+2,1)")
    assert isinstance(ast, FunctionCallNode)
    assert index_call_is_ref(ast, as_canonical("Data!C1")) is True
    in_bounds = parse("=INDEX(Data!$B$1:$B$2,ROW(),1)")
    assert isinstance(in_bounds, FunctionCallNode)
    assert index_call_is_ref(in_bounds, as_canonical("Data!C1")) is False


def test_offset_index_zero_offset_constraint_extraction_uses_row_selector(
    tmp_path: Path,
) -> None:
    """Issue 778: whole-array OFFSET edges still emit each host's INDEX pick."""
    workbook = write_workbook(
        tmp_path / "offset_index_row.xlsx",
        _row_select_sheets(formula="=OFFSET(INDEX($B$1:$B$2,ROW(),1),0,0)"),
    )
    _catalog, _deps, graph = inverted_graph_parts(
        workbook, _row_select_bindings(), dynamic_refs=_constraint_refs()
    )
    assert sorted(graph.get_dependencies("Data!C1")) == ["Data!B1", "Data!B2"]

    modules = generate_inverted(workbook, _row_select_bindings(), dynamic_refs=_constraint_refs())
    internals = modules["internals.py"]
    assert "INDIRECT edge sets" not in internals
    pkg = load_package(modules, tmp_path, name="offset_index_row_sel")
    result = pkg.compute_selected()
    assert (result[1], result[2]) == pytest.approx((10.0, 20.0))


def test_offset_index_constraint_shift_uses_index_row_not_graph_edges(
    tmp_path: Path,
) -> None:
    workbook = write_workbook(
        tmp_path / "offset_index_row_shift.xlsx",
        _row_select_sheets(formula="=OFFSET(INDEX($B$1:$B$2,ROW(),1),0,-1)"),
    )
    pkg = load_package(
        generate_inverted(
            workbook,
            _row_select_bindings(include_labels=True),
            dynamic_refs=_constraint_refs(),
        ),
        tmp_path,
        name="offset_index_row_shift",
    )
    result = pkg.compute_selected()
    assert (result[1], result[2]) == pytest.approx((1.0, 2.0))


def test_offset_index_provably_oob_emits_ref_under_constraint_extraction(
    tmp_path: Path,
) -> None:
    """Issue 778: out-of-bounds INDEX is `#REF!`, not a missing-edge classify."""
    workbook = write_workbook(
        tmp_path / "offset_index_oob.xlsx",
        _row_select_sheets(formula="=OFFSET(INDEX($B$1:$B$2,ROW()+2,1),0,-1)"),
    )
    modules = generate_inverted(
        workbook,
        _row_select_bindings(include_labels=True),
        dynamic_refs=_constraint_refs(),
    )
    assert "xl_raise('#REF!')" in modules["internals.py"]
    pkg = load_package(modules, tmp_path, name="offset_index_oob")
    result = pkg.compute_selected()
    assert (result[1], result[2]) == ("#REF!", "#REF!")


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
