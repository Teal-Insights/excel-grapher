"""INDEX(named_range,,1) keeps the worksheet country-code column (#891, #916).

LIC-DSF looks up market access with:

    INDEX(MktFin, MATCH(country_code, INDEX(MktFin,,1), 0), COLUMNS(MktFin))

`MktFin` is a headered table whose first worksheet column is country codes
(including the header cell). The inverted-tree reconstruction must INDEX that
column — not a bound-series measure column after a row-label is stripped —
and must keep Excel cell types so `INDEX(...) = 1` is numeric, not `'1' = 1`.
A computed string 0/1 series (`as_measure(1, 'str')`) must keep the numeric
Excel type as well; leaf `xl_lookup_cell` restore is not enough (#916).
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from excel_grapher.core.grid import Range
from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.inverted_tree import excel
from excel_grapher.grapher import create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    all_param_names,
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
_FLAG = _HEADERS[-1]


def test_xl_lookup_cell_restores_stringified_numbers_not_overrides() -> None:
    assert excel.xl_lookup_cell("1", 1) == 1
    assert excel.xl_lookup_cell("0", 0) == 0
    assert excel.xl_lookup_cell(1, 1) == 1
    assert excel.xl_lookup_cell("2", 1) == "2"
    assert excel.xl_lookup_cell("Ghana", "Ghana") == "Ghana"


def test_xl_typed_range_restores_stringified_cells_not_overrides() -> None:
    values = Range(
        "",
        1,
        1,
        1,
        2,
        lambda address: None,
        _coord_resolver=lambda _row, column: "1" if column == 1 else "0",
    )
    typed = excel.xl_typed_range(values, ((1, 0),))
    assert typed.cell(1, 1) == 1
    assert typed.cell(1, 2) == 0
    overridden = Range(
        "",
        1,
        1,
        1,
        1,
        lambda address: None,
        _coord_resolver=lambda _row, _column: "0",
    )
    assert excel.xl_typed_range(overridden, ((1,),)).cell(1, 1) == "0"


def _mcve_workbook(tmp_path: Path, *, computed_flag: bool = False) -> Path:
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
        excel_row = 5 + offset
        for col, value in enumerate(row, start=1):
            if computed_flag and col == 5:
                trigger[f"E{excel_row}"] = f'=IF(C{excel_row}="Yes",1,0)'
            else:
                trigger[f"{'ABCDE'[col - 1]}{excel_row}"] = value
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


def _headers_entry() -> dict[str, Any]:
    return {
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
                    "bind": {"kind": "column_header", "header_row": 3, "read": "int"},
                }
            ],
        },
        "key": ["POSITION"],
    }


def _country_code_entry() -> dict[str, Any]:
    return {
        "id": "mkt_fin_country_code",
        "sheet": "Trigger",
        "data_range": "Trigger!A5:A8",
        "layout": "series",
        "constant": {},
        "structure": {
            "measure": _measure("int"),
            "dimensions": [
                {
                    "id": "COUNTRY",
                    "concept": "COUNTRY",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "row_label", "label_column": "B", "read": "string"},
                }
            ],
        },
        "key": ["COUNTRY"],
    }


def _labels_entry() -> dict[str, Any]:
    return {
        "id": "mkt_fin_labels",
        "sheet": "Trigger",
        "data_range": "Trigger!C5:D8",
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
                    "bind": {"kind": "row_label", "label_column": "B", "read": "string"},
                },
                {
                    "id": "INDICATOR",
                    "concept": "INDICATOR",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "column_header", "header_row": 4, "read": "string"},
                },
            ],
        },
        "key": ["COUNTRY", "INDICATOR"],
    }


def _flag_entry() -> dict[str, Any]:
    return {
        "id": "mkt_fin_flag",
        "sheet": "Trigger",
        "data_range": "Trigger!E5:E8",
        "layout": "series",
        "internal": {},
        "structure": {
            "measure": _measure("string"),
            "dimensions": [
                {
                    "id": "COUNTRY",
                    "concept": "COUNTRY",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "row_label", "label_column": "B", "read": "string"},
                }
            ],
        },
        "key": ["COUNTRY"],
    }


def _mkt_fin_entry(
    *,
    data_range: str,
    label_column: str,
    label_read: str,
    direction: str = "constant",
) -> dict[str, Any]:
    entry: dict[str, Any] = {
        "id": "mkt_fin",
        "sheet": "Trigger",
        "data_range": data_range,
        "layout": "matrix",
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
                        "label_column": label_column,
                        "read": label_read,
                    },
                },
                {
                    "id": "INDICATOR",
                    "concept": "INDICATOR",
                    "role": "key",
                    "scope": "cell",
                    "bind": {"kind": "column_header", "header_row": 4, "read": "string"},
                },
            ],
        },
        "key": ["COUNTRY", "INDICATOR"],
    }
    entry[direction] = {}
    return entry


def _output_bindings(*tables: dict[str, Any]) -> dict[str, Any]:
    document = bindings_document(
        *tables,
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


def _mcve_bindings(*, mkt_fin_direction: str = "constant") -> dict[str, Any]:
    return _output_bindings(
        _headers_entry(),
        _mkt_fin_entry(
            data_range="Trigger!A5:E8",
            label_column="B",
            label_read="string",
            direction=mkt_fin_direction,
        ),
    )


def _stripped_label_bindings() -> dict[str, Any]:
    """Codes live in column A as row labels, outside the measure `data_range`."""
    return _output_bindings(
        _headers_entry(),
        _mkt_fin_entry(
            data_range="Trigger!B5:E8",
            label_column="A",
            label_read="int",
        ),
    )


def _load(tmp_path: Path, *, document: dict[str, Any] | None = None, name: str) -> Any:
    workbook = _mcve_workbook(tmp_path)
    bindings = _mcve_bindings() if document is None else document
    return load_package(generate_inverted(workbook, bindings), tmp_path, name=name)


def _kwargs(pkg: Any, **overrides: object) -> dict[str, object]:
    mkt_fin = getattr(pkg.data, "MKT_FIN_DEFAULT", None)
    if mkt_fin is None:
        mkt_fin = getattr(pkg.data, "MKT_FIN", None)
    values: dict[str, object] = {
        "enabled": pkg.data.ENABLED_DEFAULT,
        "country_code": pkg.data.COUNTRY_CODE_DEFAULT,
        "mkt_fin": mkt_fin,
        "mkt_fin_headers": pkg.data.MKT_FIN_HEADERS,
        "mkt_fin_country_code": getattr(pkg.data, "MKT_FIN_COUNTRY_CODE", None),
        "mkt_fin_labels": getattr(pkg.data, "MKT_FIN_LABELS", None),
    }
    values.update(overrides)
    accepted = all_param_names(pkg.compute_yes_no)
    return {key: value for key, value in values.items() if key in accepted}


def _computed_flag_bindings() -> dict[str, Any]:
    """LIC-DSF-style split: int codes, string labels, computed string 0/1 flag."""
    return _output_bindings(_headers_entry(), _country_code_entry(), _labels_entry(), _flag_entry())


def _ghana_market_access_results(tmp_path: Path, *, computed_flag: bool) -> None:
    workbook = _mcve_workbook(tmp_path, computed_flag=computed_flag)
    document = _computed_flag_bindings() if computed_flag else _mcve_bindings()
    targets = ["Input!C11", "Chart!I19", "Chart!D251", "Chart!D242"]
    graph = create_dependency_graph(
        workbook, targets, load_values=True, use_cached_dynamic_refs=True
    )
    assert FormulaEvaluator(graph).evaluate(targets) == {
        "Input!C11": "Yes",
        "Chart!I19": 1.0,
        "Chart!D251": 67.17,
        "Chart!D242": 76.13,
    }
    name = "mkt_fin_computed_flag" if computed_flag else "mkt_fin_index"
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name=name)
    kwargs = _kwargs(pkg)
    assert pkg.compute_yes_no(**kwargs) == "Yes"
    assert pkg.compute_chart(**kwargs) == 67.17
    assert pkg.compute_b2(**kwargs) == 76.13
    assert_package_matches_evaluator(workbook, document, tmp_path, f"{name}_parity")


def test_index_named_range_market_access_matches_evaluator(tmp_path: Path) -> None:
    _ghana_market_access_results(tmp_path, computed_flag=False)


def test_index_computed_string_flag_equals_one(tmp_path: Path) -> None:
    """`INDEX(...)=1` on a computed string 0/1 series matches the evaluator (#916)."""
    _ghana_market_access_results(tmp_path, computed_flag=True)


def test_index_named_range_off_and_other_country(tmp_path: Path) -> None:
    pkg = _load(tmp_path, name="mkt_fin_off")
    assert pkg.compute_yes_no(**_kwargs(pkg, enabled="Off")) == "No"
    assert pkg.compute_b2(**_kwargs(pkg, enabled="Off")) == 74.12
    assert pkg.compute_yes_no(**_kwargs(pkg, country_code=668)) == "No"
    assert pkg.compute_b2(**_kwargs(pkg, country_code=668)) == 74.12


def test_index_computed_string_flag_off_and_other_country(tmp_path: Path) -> None:
    workbook = _mcve_workbook(tmp_path, computed_flag=True)
    document = _computed_flag_bindings()
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="mkt_fin_flag_off")
    assert pkg.compute_yes_no(**_kwargs(pkg, enabled="Off")) == "No"
    assert pkg.compute_b2(**_kwargs(pkg, enabled="Off")) == 74.12
    assert pkg.compute_yes_no(**_kwargs(pkg, country_code=668)) == "No"
    assert pkg.compute_b2(**_kwargs(pkg, country_code=668)) == 74.12


def test_index_computed_string_flag_follows_label_overrides(tmp_path: Path) -> None:
    workbook = _mcve_workbook(tmp_path, computed_flag=True)
    document = _computed_flag_bindings()
    pkg = load_package(
        generate_inverted(workbook, document), tmp_path, name="mkt_fin_flag_override"
    )
    table = pkg.data.MKT_FIN_LABELS
    records = [
        (coord, "No" if coord == ("Ghana", "Eurobond") else table[coord]) for coord in table.domain
    ]
    with pkg.data.overrides(MKT_FIN_LABELS=table.with_records(records)):
        assert pkg.compute_yes_no(**_kwargs(pkg)) == "No"
    assert pkg.compute_yes_no(**_kwargs(pkg)) == "Yes"


def test_index_named_range_reads_bound_series_overrides(tmp_path: Path) -> None:
    pkg = _load(tmp_path, name="mkt_fin_override")
    table = pkg.data.MKT_FIN
    records = [
        (coord, "0" if coord == ("Ghana", _FLAG) else table[coord]) for coord in table.domain
    ]
    updated = table.with_records(records)
    with pkg.data.overrides(MKT_FIN=updated):
        assert pkg.compute_yes_no(**_kwargs(pkg)) == "No"
    assert pkg.compute_yes_no(**_kwargs(pkg)) == "Yes"


def test_index_named_range_includes_row_label_column(tmp_path: Path) -> None:
    workbook = _mcve_workbook(tmp_path)
    document = _stripped_label_bindings()
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="mkt_fin_stripped")
    kwargs = _kwargs(pkg)
    assert pkg.compute_yes_no(**kwargs) == "Yes"
    assert pkg.compute_chart(**kwargs) == 67.17
    assert pkg.compute_b2(**kwargs) == 76.13
    assert_package_matches_evaluator(workbook, document, tmp_path, "mkt_fin_stripped_parity")


def test_index_named_range_input_matrix_keeps_excel_types(tmp_path: Path) -> None:
    workbook = _mcve_workbook(tmp_path)
    document = _mcve_bindings(mkt_fin_direction="input")
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="mkt_fin_input")
    kwargs = _kwargs(pkg)
    assert pkg.compute_yes_no(**kwargs) == "Yes"
    assert pkg.compute_chart(**kwargs) == 67.17
    table = pkg.data.MKT_FIN_DEFAULT
    records = [
        (coord, "0" if coord == ("Ghana", _FLAG) else table[coord]) for coord in table.domain
    ]
    assert pkg.compute_yes_no(**_kwargs(pkg, mkt_fin=table.with_records(records))) == "No"
