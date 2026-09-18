"""INDEX(named_range,,1) keeps the worksheet country-code column (#891).

LIC-DSF looks up market access with:

    INDEX(MktFin, MATCH(country_code, INDEX(MktFin,,1), 0), COLUMNS(MktFin))

`MktFin` is a headered table whose first worksheet column is country codes
(including the header cell). The inverted-tree reconstruction must INDEX that
column — not a bound-series measure column after a row-label is stripped —
and must keep Excel cell types so `INDEX(...) = 1` is numeric, not `'1' = 1`.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Literal

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.grapher import DynamicRefConfig, create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    assert_package_matches_evaluator,
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)

_HEADERS = [
    "Country code",
    "Country",
    "Eurobond",
    "PRGT market access",
    "Market access (for stress tests and market fin. Module)",
]
_ROWS = (
    (4, "Afghanistan", "No", "No", 0),
    (50, "Benin", "No", "No", 0),
    (652, "Ghana", "Yes", "Yes", 1),
    (668, "Togo", "No", "No", 0),
)


def _mcve_workbook(tmp_path: Path) -> Path:
    cells: dict[str, object] = {
        "C6": "On",
        "C7": "Ghana",
        "C8": 652,
        "C11": (
            '=IF($C$6="On",IFERROR(IF(INDEX(MktFin,MATCH(C8,INDEX(MktFin,,1),0),'
            'COLUMNS(MktFin))=1,"Yes","No"),"No"),"No")'
        ),
    }
    trigger: dict[str, object] = {}
    for col, header in enumerate(_HEADERS, start=1):
        letter = "ABCDE"[col - 1]
        trigger[f"{letter}3"] = col
        trigger[f"{letter}4"] = header
    for offset, row in enumerate(_ROWS):
        for col, value in enumerate(row, start=1):
            trigger[f"{'ABCDE'[col - 1]}{5 + offset}"] = value
    chart = {
        "I21": 1,
        "I19": '=IF(Input!C11="Yes",1,0)',
        "A50": "=I19",
        "A254": "=A50",
        "D261": 67.17,
        "D254": 76.13,
        "D255": 74.12,
        "D251": '=IFERROR(IF($I$21=$A50,D261," ")," ")',
        "D242": "=IF($I$21=$A254,D254,D255)",
    }
    return write_workbook(
        tmp_path / "mkt_fin.xlsx",
        {"Trigger": trigger, "Input": cells, "Chart": chart},
        defined_names={"MktFin": "Trigger!$A$4:$E$8"},
    )


def _measure(dtype: str) -> dict[str, Any]:
    return {
        "concept": "OBS_VALUE",
        "dtype": dtype,
        "bind": {"kind": "data_cell", "read": dtype if dtype != "number" else "float"},
    }


def _mcve_bindings() -> dict[str, Any]:
    document = bindings_document(
        {
            "id": "mkt_fin_headers",
            "sheet": "Trigger",
            "data_range": "Trigger!A4:E4",
            "layout": "series",
            "constant": {},
            "structure": {
                "measure": _measure("string"),
                "dimensions": [
                    {
                        "id": "POSITION",
                        "concept": "INDICATOR",
                        "role": "key",
                        "scope": "cell",
                        "bind": {
                            "kind": "column_header",
                            "header_row": 3,
                            "read": "int",
                        },
                    }
                ],
            },
            "key": ["POSITION"],
        },
        {
            "id": "mkt_fin",
            "sheet": "Trigger",
            "data_range": "Trigger!A5:E8",
            "layout": "matrix",
            "constant": {},
            "structure": {
                "measure": _measure("string"),
                "dimensions": [
                    {
                        "id": "COUNTRY",
                        "concept": "COUNTRY",
                        "role": "key",
                        "scope": "cell",
                        "bind": {
                            "kind": "row_label",
                            "label_column": "B",
                            "read": "string",
                        },
                    },
                    {
                        "id": "INDICATOR",
                        "concept": "INDICATOR",
                        "role": "key",
                        "scope": "cell",
                        "bind": {
                            "kind": "column_header",
                            "header_row": 4,
                            "read": "string",
                        },
                    },
                ],
            },
            "key": ["COUNTRY", "INDICATOR"],
        },
        series_entry("marker", "Chart!I21", layout="scalar", direction="constant", dtype="int"),
        series_entry("c4_value", "Chart!D261", layout="scalar", direction="constant"),
        series_entry("b2_market", "Chart!D254", layout="scalar", direction="constant"),
        series_entry("b2_non_market", "Chart!D255", layout="scalar", direction="constant"),
        series_entry("flag", "Chart!I19", layout="scalar", direction="internal"),
        series_entry("flag_a50", "Chart!A50", layout="scalar", direction="internal"),
        series_entry("flag_a254", "Chart!A254", layout="scalar", direction="internal"),
        series_entry(
            "enabled",
            "Input!C6",
            layout="scalar",
            direction="input",
            dtype="string",
            domain={"enum": ["On", "Off"]},
        ),
        series_entry(
            "country_code",
            "Input!C8",
            layout="scalar",
            direction="input",
            dtype="int",
        ),
        series_entry(
            "yes_no",
            "Input!C11",
            layout="scalar",
            direction="output",
            dtype="string",
            compute_name="compute_yes_no",
        ),
        series_entry(
            "chart",
            "Chart!D251",
            layout="scalar",
            direction="output",
            compute_name="compute_chart",
        ),
        series_entry(
            "b2",
            "Chart!D242",
            layout="scalar",
            direction="output",
            compute_name="compute_b2",
        ),
        schema_version="1.16.0",
    )
    document["concept_scheme"]["concepts"].append({"id": "INDICATOR", "dtype": "string"})
    return document


def _mcve_dynamic_refs() -> DynamicRefConfig:
    countries = Literal[4, 50, 652, 668]
    names = Literal["Afghanistan", "Benin", "Ghana", "Togo"]
    yes_no = Literal["Yes", "No"]
    flag = Literal[0, 1]
    return DynamicRefConfig.from_constraints(
        {
            "Input!C6": Literal["On", "Off"],
            "Input!C8": countries,
            "Trigger!A4": Literal["Country code"],
            "Trigger!B4": Literal["Country"],
            "Trigger!C4": Literal["Eurobond"],
            "Trigger!D4": Literal["PRGT market access"],
            "Trigger!E4": Literal["Market access (for stress tests and market fin. Module)"],
            "Trigger!A5": countries,
            "Trigger!A6": countries,
            "Trigger!A7": countries,
            "Trigger!A8": countries,
            "Trigger!B5": names,
            "Trigger!B6": names,
            "Trigger!B7": names,
            "Trigger!B8": names,
            "Trigger!C5": yes_no,
            "Trigger!C6": yes_no,
            "Trigger!C7": yes_no,
            "Trigger!C8": yes_no,
            "Trigger!D5": yes_no,
            "Trigger!D6": yes_no,
            "Trigger!D7": yes_no,
            "Trigger!D8": yes_no,
            "Trigger!E5": flag,
            "Trigger!E6": flag,
            "Trigger!E7": flag,
            "Trigger!E8": flag,
        },
        {},
    )


def test_index_named_range_match_uses_worksheet_code_column(tmp_path: Path) -> None:
    workbook = _mcve_workbook(tmp_path)
    document = _mcve_bindings()
    dynamic_refs = _mcve_dynamic_refs()
    targets = ["Input!C11", "Chart!I19", "Chart!D251", "Chart!D242"]
    graph = create_dependency_graph(
        workbook,
        targets,
        load_values=True,
        dynamic_refs=dynamic_refs,
    )
    expected = FormulaEvaluator(graph).evaluate(targets)
    assert expected == {
        "Input!C11": "Yes",
        "Chart!I19": 1.0,
        "Chart!D251": 67.17,
        "Chart!D242": 76.13,
    }

    modules = generate_inverted(workbook, document, dynamic_refs=dynamic_refs)
    internals = modules["internals.py"]
    assert "('Country code',), (4,), (50,), (652,), (668,)" in internals
    assert (
        "('Market access (for stress tests and market fin. Module)',), (0,), (0,), (1,), (0,)"
        in internals
    )
    pkg = load_package(modules, tmp_path, name="mkt_fin_index")
    kwargs = {
        "enabled": pkg.data.ENABLED_DEFAULT,
        "country_code": pkg.data.COUNTRY_CODE_DEFAULT,
    }
    assert pkg.compute_yes_no(**kwargs) == "Yes"
    assert pkg.compute_chart(**kwargs) == 67.17
    assert pkg.compute_b2(**kwargs) == 76.13
    assert_package_matches_evaluator(
        workbook, document, tmp_path, "mkt_fin_index_parity", dynamic_refs=dynamic_refs
    )
