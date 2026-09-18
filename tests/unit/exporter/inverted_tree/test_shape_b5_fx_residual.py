"""LIC-DSF B5 residual financing diverges only after the 30ppt FX-shock floor (#893).

`Chart Data` B5 later years pull `PV_ResFin_pub` residual stock. REER
overvaluation (`Input 1 - Basics!C31`) enters through the standard-test FX
shock `MAX(30, C37)`. At 0 and 20, `C37≤30` so the floor holds and residual
extra is zero. At 40 the shock exceeds 30, residual financing moves, and the
inverted-tree library must still match `FormulaEvaluator`.

The extra residual is gated by market access
`INDEX(MktFin, MATCH(code, INDEX(MktFin,,1), 0), COLUMNS(MktFin))`. When that
lookup misses the worksheet country-code column, C11 stays `No` and the extra
collapses — the same Off path C4/B2/B6 already take, visible on B5 only once
the floor no longer hides it.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Literal

import pytest
from fastpyxl.utils.cell import get_column_letter

from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.grapher import DynamicRefConfig, create_dependency_graph
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)

_INFL_MINUS_COMMODITY = 7.788070682246872


def _measure(dtype: str = "float") -> dict[str, Any]:
    read = "float" if dtype == "number" else dtype
    return {
        "concept": "OBS_VALUE",
        "dtype": dtype,
        "bind": {"kind": "data_cell", "read": read},
    }


def _scalar(
    series_id: str,
    data_range: str,
    *,
    direction: str,
    dtype: str = "float",
    domain: dict[str, Any] | None = None,
    compute_name: str | None = None,
) -> dict[str, Any]:
    entry = series_entry(
        series_id,
        data_range,
        direction=direction,
        dtype=dtype,
        domain=domain,
        compute_name=compute_name,
    )
    entry["structure"]["measure"] = _measure(dtype)
    return entry


def _fx_floor_workbook(tmp_path: Path, overval: float) -> Path:
    """Isolated MAX-floor cone: residual extra is `max(0, shock-30)`."""
    return write_workbook(
        tmp_path / f"b5_fx_floor_{overval:g}.xlsx",
        {
            "Input": {
                "C4": "New",
                "C31": overval,
                "C36": '=IF(C4="NEW",C31,0)',
                "C37": f"=C36+{_INFL_MINUS_COMMODITY}",
                "C38": '=IF(Input!$C$4="New",MAX(30,C37),30)',
                "D38": "=MAX(30,C37)",
            },
            "B5": {
                "AA62": '=IF(Input!$C$4="New",Input!D38,Input!C38)',
                "F52": "=-AA62",
                "F82": "=10+(AA62-30)*2",
                "F13": "=F82",
            },
        },
    )


def _fx_floor_bindings() -> dict[str, Any]:
    return bindings_document(
        _scalar("mode", "Input!C4", direction="input", dtype="string"),
        _scalar("overval", "Input!C31", direction="input"),
        _scalar("c36", "Input!C36", direction="internal"),
        _scalar("c37", "Input!C37", direction="internal"),
        _scalar("c38", "Input!C38", direction="internal"),
        _scalar("d38", "Input!D38", direction="internal"),
        _scalar("aa62", "B5!AA62", direction="internal"),
        _scalar("f52", "B5!F52", direction="internal"),
        _scalar("f82", "B5!F82", direction="internal"),
        _scalar("b5", "B5!F13", direction="output", compute_name="compute_b5"),
        schema_version="1.16.0",
    )


def _mktfin_workbook(tmp_path: Path, overval: float) -> Path:
    """B5 extra residual above the FX floor is gated by `INDEX(MktFin,,1)`."""
    trigger = {
        "A3": 1,
        "B3": 2,
        "C3": 3,
        "D3": 4,
        "E3": 5,
        "A4": "Country code",
        "B4": "Country",
        "C4": "Eurobond",
        "D4": "PRGT market access",
        "E4": "Market access (for stress tests and market fin. Module)",
    }
    rows = (
        (4, "Afghanistan", "No", "No", 0),
        (50, "Benin", "No", "No", 0),
        (652, "Ghana", "Yes", "Yes", 1),
        (668, "Togo", "No", "No", 0),
    )
    for offset, row in enumerate(rows):
        for col, value in enumerate(row, start=1):
            trigger[f"{get_column_letter(col)}{5 + offset}"] = value
    return write_workbook(
        tmp_path / f"b5_mktfin_{overval:g}.xlsx",
        {
            "Trigger": trigger,
            "Input": {
                "C4": "New",
                "C6": "On",
                "C8": 652,
                "C11": (
                    '=IF($C$6="On",IFERROR(IF(INDEX(MktFin,MATCH(C8,INDEX(MktFin,,1),0),'
                    'COLUMNS(MktFin))=1,"Yes","No"),"No"),"No")'
                ),
                "C31": overval,
                "C36": '=IF(C4="NEW",C31,0)',
                "C37": f'=IF(C4="NEW",C36+{_INFL_MINUS_COMMODITY},0)',
                "C38": '=IF(Input!$C$4="New",MAX(30,C37),30)',
            },
            "B5": {
                "AA62": "=Input!C38",
                "F82": '=IF(Input!C11="Yes",10+(AA62-30)*2,10)',
                "F13": "=F82",
            },
        },
        defined_names={"MktFin": "Trigger!$A$4:$E$8"},
    )


def _mktfin_bindings() -> dict[str, Any]:
    headers = {
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
    table = {
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
    document = bindings_document(
        headers,
        table,
        _scalar(
            "enabled",
            "Input!C6",
            direction="input",
            dtype="string",
            domain={"enum": ["On", "Off"]},
        ),
        _scalar("country_code", "Input!C8", direction="input", dtype="int"),
        _scalar(
            "mode",
            "Input!C4",
            direction="input",
            dtype="string",
            domain={"enum": ["New", "Old"]},
        ),
        _scalar("overval", "Input!C31", direction="input"),
        _scalar("yes_no", "Input!C11", direction="internal", dtype="string"),
        _scalar("c36", "Input!C36", direction="internal"),
        _scalar("c37", "Input!C37", direction="internal"),
        _scalar("c38", "Input!C38", direction="internal"),
        _scalar("aa62", "B5!AA62", direction="internal"),
        _scalar("f82", "B5!F82", direction="internal"),
        _scalar("b5", "B5!F13", direction="output", compute_name="compute_b5"),
        schema_version="1.16.0",
    )
    concepts = document["concept_scheme"]["concepts"]
    if not any(item["id"] == "INDICATOR" for item in concepts):
        concepts.append({"id": "INDICATOR", "dtype": "string"})
    return document


def _mktfin_constraints() -> dict[str, object]:
    countries = Literal["Afghanistan", "Benin", "Ghana", "Togo"]
    codes = Literal[4, 50, 652, 668]
    yes_no = Literal["Yes", "No"]
    return {
        "Input!C4": Literal["New", "Old"],
        "Input!C6": Literal["On", "Off"],
        "Input!C8": codes,
        "Input!C31": Literal[0, 20, 40],
        "Trigger!A4": Literal["Country code"],
        "Trigger!B4": Literal["Country"],
        "Trigger!C4": Literal["Eurobond"],
        "Trigger!D4": Literal["PRGT market access"],
        "Trigger!E4": Literal["Market access (for stress tests and market fin. Module)"],
        **{f"Trigger!A{row}": codes for row in range(5, 9)},
        **{f"Trigger!B{row}": countries for row in range(5, 9)},
        **{f"Trigger!C{row}": yes_no for row in range(5, 9)},
        **{f"Trigger!D{row}": yes_no for row in range(5, 9)},
        **{f"Trigger!E{row}": Literal[0, 1] for row in range(5, 9)},
    }


def _evaluate(workbook: Path, targets: list[str], constraints: dict[str, object] | None = None):
    graph = create_dependency_graph(
        workbook,
        targets,
        load_values=True,
        dynamic_refs=DynamicRefConfig.from_constraints(constraints, {}) if constraints else None,
    )
    return FormulaEvaluator(graph).evaluate(targets)


def _export_b5(
    workbook: Path,
    document: dict[str, Any],
    tmp_path: Path,
    name: str,
    *,
    overval: float,
    constraints: dict[str, object] | None = None,
    extra_kwargs: dict[str, object] | None = None,
) -> object:
    dynamic_refs = (
        DynamicRefConfig.from_constraints(constraints, {}) if constraints is not None else None
    )
    pkg = load_package(
        generate_inverted(workbook, document, dynamic_refs=dynamic_refs),
        tmp_path,
        name=name,
    )
    kwargs: dict[str, object] = {"overval": overval, "mode": "New"}
    if extra_kwargs:
        kwargs.update(extra_kwargs)
    accepted = set(pkg.compute_b5.__code__.co_varnames)
    return pkg.compute_b5(**{key: value for key, value in kwargs.items() if key in accepted})


@pytest.mark.parametrize("overval", [0.0, 20.0, 40.0])
def test_fx_floor_residual_matches_evaluator_above_and_below_floor(
    tmp_path: Path, overval: float
) -> None:
    """Control: MAX floor plus residual extra lower without the market-access lookup."""
    workbook = _fx_floor_workbook(tmp_path, overval)
    expected = _evaluate(workbook, ["B5!F13", "Input!C38", "B5!AA62"])
    got = _export_b5(
        workbook, _fx_floor_bindings(), tmp_path, f"b5_floor_{overval:g}", overval=overval
    )
    assert expected["Input!C38"] == pytest.approx(max(30.0, overval + _INFL_MINUS_COMMODITY))
    assert got == pytest.approx(expected["B5!F13"])
    if overval <= 20.0:
        assert got == pytest.approx(10.0)
    else:
        assert got == pytest.approx(10.0 + (overval + _INFL_MINUS_COMMODITY - 30.0) * 2)


@pytest.mark.parametrize("overval", [0.0, 20.0, 40.0])
def test_b5_residual_above_fx_floor_uses_mktfin_country_code_column(
    tmp_path: Path, overval: float
) -> None:
    """B5@40 extra residual requires INDEX(MktFin,,1) to be the country-code column."""
    workbook = _mktfin_workbook(tmp_path, overval)
    constraints = _mktfin_constraints()
    expected = _evaluate(
        workbook,
        ["Input!C11", "Input!C38", "B5!F13"],
        constraints,
    )
    got = _export_b5(
        workbook,
        _mktfin_bindings(),
        tmp_path,
        f"b5_mktfin_{overval:g}",
        overval=overval,
        constraints=constraints,
        extra_kwargs={"enabled": "On", "country_code": 652},
    )
    assert expected["Input!C11"] == "Yes"
    shock = max(30.0, overval + _INFL_MINUS_COMMODITY)
    assert expected["Input!C38"] == pytest.approx(shock)
    assert expected["B5!F13"] == pytest.approx(10.0 + (shock - 30.0) * 2)
    assert got == pytest.approx(expected["B5!F13"])
    if overval <= 20.0:
        assert got == pytest.approx(10.0)
    else:
        assert got == pytest.approx(10.0 + (shock - 30.0) * 2)
