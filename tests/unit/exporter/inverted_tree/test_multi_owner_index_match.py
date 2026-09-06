"""MATCH/INDEX ranges that span multiple bound series and declared blanks (#710).

Excel treats lookup vectors and INDEX rectangles as positional arrays. Binding
boundaries and declared blanks must not drop, shift, or replace those positions.
"""

from __future__ import annotations

from collections.abc import Mapping
from pathlib import Path
from typing import Any

import pytest

from excel_grapher.core.address_keys import as_canonical
from excel_grapher.core.types import XlError
from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree.deps import resolve_positional_range
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.grapher.blank_ranges import normalize_blank_range_specs
from excel_grapher.series_bindings import validate_bindings_document, validate_series_bindings
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    call_compute,
    generate_inverted,
    input_kwargs,
    inverted_graph_parts,
    load_package,
    series_entry,
    write_workbook,
)

_BLANK = ("Engine!A6", "Engine!B5:C10", "Engine!D6:E7")
_MEASURE = {
    "concept": "OBS_VALUE",
    "dtype": "float",
    "bind": {"kind": "data_cell", "read": "float"},
}


def _multi_owner_sheets() -> dict[str, dict[str, object]]:
    """Compact form of the `input5_vintage_terms` INDEX/MATCH geometry.

    Row vector ``A5:A10`` has three owners plus a declared interior blank.
    Column vector ``A5:E5`` splits a note, declared blanks, and year headers.
    Two output formulas share one shape so generated iteration cannot drift.
    """
    return {
        "Engine": {
            "A5": '="Note"',
            "A7": '="Title"',
            "A8": "Bond",
            "A9": "Loan",
            "A10": "Equity",
            "D5": 2020,
            "E5": 2021,
            "D8": 10.0,
            "E8": 11.0,
            "D9": 20.0,
            "E9": 21.0,
            "D10": 30.0,
            "E10": 31.0,
            "A228": "Loan",
            "B230": 2020,
            "B231": 2021,
            "E230": ("=INDEX($A$5:$E$10,MATCH(A$228,$A$5:$A$10,0),MATCH(B230,$A$5:$E$5,0))"),
            "E231": ("=INDEX($A$5:$E$10,MATCH(A$228,$A$5:$A$10,0),MATCH(B231,$A$5:$E$5,0))"),
        }
    }


def _multi_owner_bindings() -> dict[str, Any]:
    return bindings_document(
        series_entry("note", "Engine!A5", layout="scalar", direction="internal", dtype="string"),
        series_entry("title", "Engine!A7", layout="scalar", direction="internal", dtype="string"),
        series_entry(
            "labels",
            "Engine!A8:A10",
            layout="series",
            direction="constant",
            dtype="string",
            label_column="A",
            key_concept="COUNTRY",
            key_read="string",
        ),
        series_entry(
            "years",
            "Engine!D5:E5",
            layout="series",
            direction="constant",
            dtype="int",
            header_row=5,
        ),
        {
            "id": "terms",
            "sheet": "Engine",
            "data_range": "Engine!D8:E10",
            "layout": "matrix",
            "constant": {},
            "structure": {
                "measure": _MEASURE,
                "dimensions": [
                    {
                        "id": "COUNTRY",
                        "concept": "COUNTRY",
                        "role": "key",
                        "scope": "cell",
                        "bind": {
                            "kind": "row_label",
                            "label_column": "A",
                            "read": "string",
                        },
                    },
                    {
                        "id": "TIME_PERIOD",
                        "concept": "TIME_PERIOD",
                        "role": "key",
                        "scope": "cell",
                        "bind": {
                            "kind": "column_header",
                            "header_row": 5,
                            "read": "int",
                        },
                    },
                ],
            },
            "key": ["COUNTRY", "TIME_PERIOD"],
        },
        series_entry(
            "row_key",
            "Engine!A228",
            layout="scalar",
            direction="input",
            dtype="string",
        ),
        series_entry(
            "col_key",
            "Engine!B230:B231",
            layout="series",
            direction="input",
            dtype="int",
            label_column="B",
            key_concept="TIME_PERIOD",
            key_read="int",
        ),
        series_entry(
            "picked",
            "Engine!E230:E231",
            layout="series",
            direction="output",
            label_column="B",
            key_concept="TIME_PERIOD",
            key_read="int",
        ),
    )


def _norm_measure(value: object) -> object:
    if isinstance(value, XlError):
        return value.value
    return value


def _export_picked(
    tmp_path: Path,
    *,
    stem: str,
    document: dict[str, Any] | None = None,
    sheets: Mapping[str, Mapping[str, object]] | None = None,
    inputs: Mapping[str, object] | None = None,
) -> tuple[tuple[object, ...], tuple[object, ...], dict[str, str]]:
    bound = document or _multi_owner_bindings()
    workbook = write_workbook(tmp_path / f"{stem}.xlsx", sheets or _multi_owner_sheets())
    catalog, _deps, graph = inverted_graph_parts(workbook, bound, blank_ranges=_BLANK)
    modules = generate_inverted(workbook, bound, blank_ranges=_BLANK)
    pkg = load_package(modules, tmp_path, name=stem)
    cells = ("Engine!E230", "Engine!E231")
    with FormulaEvaluator(graph, blank_ranges=_BLANK) as ev:
        expected = ev.evaluate(list(cells))
    kwargs = input_kwargs(catalog, graph)
    if inputs is not None:
        kwargs.update(inputs)
    got = call_compute(pkg, "picked", kwargs)
    assert isinstance(got, tuple)
    want = tuple(_norm_measure(expected[cell]) for cell in cells)
    return tuple(_norm_measure(value) for value in got), want, modules


def test_resolve_positional_range_keeps_declared_blanks(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "positional.xlsx", _multi_owner_sheets())
    catalog, _deps, _graph = inverted_graph_parts(
        workbook, _multi_owner_bindings(), blank_ranges=_BLANK
    )
    cells, missing = resolve_positional_range(
        [
            as_canonical("Engine!A5"),
            as_canonical("Engine!A6"),
            as_canonical("Engine!A7"),
            as_canonical("Engine!A8"),
        ],
        catalog,
        normalize_blank_range_specs(_BLANK),
    )
    assert missing == ()
    assert [cell.blank for cell in cells] == [False, True, False, False]
    assert [cell.series_id for cell in cells] == ["note", None, "title", "labels"]
    assert cells[3].catalog_index == 0


def test_coverage_passes_before_export(tmp_path: Path) -> None:
    """Ownership coverage is green; the gap is export, not a missing binding."""
    workbook = write_workbook(tmp_path / "coverage.xlsx", _multi_owner_sheets())
    bindings = validate_bindings_document(_multi_owner_bindings())
    graph = create_dependency_graph(
        workbook,
        ["Engine!E230", "Engine!E231"],
        load_values=True,
        use_cached_dynamic_refs=True,
        blank_ranges=_BLANK,
    )
    report = validate_series_bindings(graph, bindings, workbook=workbook)
    assert report["ok"] is True


def test_export_matches_evaluator_across_binding_boundaries(tmp_path: Path) -> None:
    got, want, modules = _export_picked(tmp_path, stem="mo_idx")
    assert got == pytest.approx(want)
    assert got == pytest.approx((20.0, 21.0))
    internals = modules["internals.py"]
    assert "None" in internals
    assert "xl_match(" in internals


def test_leading_and_interior_blanks_keep_match_positions(tmp_path: Path) -> None:
    """``Loan`` is Excel position 5; dropping ``A6`` would match at 4."""
    got, want, _modules = _export_picked(tmp_path, stem="mo_pos")
    assert got == pytest.approx(want)
    assert got == pytest.approx((20.0, 21.0))


def test_missing_lookup_is_na(tmp_path: Path) -> None:
    got, _want, _modules = _export_picked(
        tmp_path,
        stem="mo_na",
        inputs={"row_key": "Missing"},
    )
    assert got == ("#N/A", "#N/A")


def test_index_into_declared_blank_is_blank(tmp_path: Path) -> None:
    got, _want, _modules = _export_picked(
        tmp_path,
        stem="mo_blank",
        inputs={"row_key": "Title"},
    )
    assert got == (None, None)


def test_match_slice_does_not_use_the_rest_of_the_series(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "match_slice.xlsx",
        {
            "Engine": {
                "A1": "A",
                "A2": "B",
                "A3": "C",
                "B1": '=MATCH("C",$A$1:$A$2,0)',
                "B2": '=MATCH("B",$A$2:$A$3,0)',
            }
        },
    )
    document = bindings_document(
        series_entry(
            "labels",
            "Engine!A1:A3",
            layout="series",
            direction="constant",
            dtype="string",
            label_column="A",
            key_concept="COUNTRY",
            key_read="string",
        ),
        series_entry(
            "picked",
            "Engine!B1:B2",
            layout="series",
            direction="output",
            dtype="int",
            label_column="A",
            key_concept="COUNTRY",
            key_read="string",
        ),
    )
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="match_slice")
    with FormulaEvaluator(graph) as ev:
        expected = ev.evaluate(["Engine!B1", "Engine!B2"])
    got = call_compute(pkg, "picked", input_kwargs(catalog, graph))
    assert isinstance(got, tuple)
    want = tuple(_norm_measure(expected[cell]) for cell in ("Engine!B1", "Engine!B2"))
    assert tuple(_norm_measure(value) for value in got) == want
    assert want[0] == "#N/A"
    assert want[1] == 1


def test_invalid_index_coordinates_use_shared_ref_semantics(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "index_oob.xlsx",
        {
            "Engine": {
                "A1": 10.0,
                "A2": 20.0,
                "B1": 1.0,
                "B2": 2.0,
                "C1": "=INDEX($A$1:$B$2,3,1)",
                "C2": "=INDEX($A$1:$B$2,1,3)",
            }
        },
    )
    document = bindings_document(
        series_entry("corner", "Engine!A1", layout="scalar", direction="constant"),
        series_entry("a2", "Engine!A2", layout="scalar", direction="constant"),
        series_entry("b1", "Engine!B1", layout="scalar", direction="constant"),
        series_entry("b2", "Engine!B2", layout="scalar", direction="constant"),
        series_entry(
            "picked",
            "Engine!C1:C2",
            layout="series",
            direction="output",
            label_column="A",
            key_concept="TIME_PERIOD",
            key_read="int",
        ),
    )
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="index_oob")
    with FormulaEvaluator(graph) as ev:
        expected = ev.evaluate(["Engine!C1", "Engine!C2"])
    got = call_compute(pkg, "picked", input_kwargs(catalog, graph))
    assert isinstance(got, tuple)
    want = tuple(_norm_measure(expected[cell]) for cell in ("Engine!C1", "Engine!C2"))
    assert tuple(_norm_measure(value) for value in got) == want
    assert want == ("#REF!", "#REF!")


def test_unbound_lookup_cell_names_host_reference_and_address(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "unbound.xlsx", _multi_owner_sheets())
    document = _multi_owner_bindings()
    document["series"] = [series for series in document["series"] if series["id"] != "terms"]
    with pytest.raises(InvertedTreeExportError) as exc:
        generate_inverted(workbook, document, blank_ranges=_BLANK)
    message = str(exc.value)
    assert "Engine!E230" in message
    assert "Engine!A5:E10" in message or "unbound cells" in message
    assert "Engine!D8" in message


def test_unbound_match_vector_cell_fails_closed(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "unbound_match.xlsx",
        {
            "Engine": {
                "A1": "A",
                "A2": "B",
                "A3": "C",
                "B1": '=MATCH("B",$A$1:$A$3,0)',
            }
        },
    )
    document = bindings_document(
        series_entry("first", "Engine!A1", layout="scalar", direction="constant", dtype="string"),
        series_entry("third", "Engine!A3", layout="scalar", direction="constant", dtype="string"),
        series_entry("picked", "Engine!B1", layout="scalar", direction="output", dtype="int"),
    )
    with pytest.raises(InvertedTreeExportError) as exc:
        generate_inverted(workbook, document)
    message = str(exc.value)
    assert "Engine!B1" in message
    assert "Engine!A2" in message
