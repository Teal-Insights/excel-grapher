"""Read Excel data validations as address-keyed domain suggestions (#983)."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import fastpyxl
import pytest
from fastpyxl.workbook.defined_name import DefinedName
from fastpyxl.worksheet.datavalidation import DataValidation

from excel_grapher.core.cell_types import CellKind
from excel_grapher.series_bindings.validation_domains import (
    ValidationDomainSuggestion,
    compare_series_validation_domains,
    suggest_validation_domains,
    validation_cell_type_env,
)


def _dv(sqref: str, **kwargs: Any) -> DataValidation:
    dv = DataValidation(**kwargs)
    for part in sqref.split():
        dv.add(part)
    return dv


def _book(path: Path, *rules: DataValidation, extra: Any = None) -> Path:
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Dash"
    lists = wb.create_sheet("My Lists")
    lists["A1"] = "North"
    lists["A2"] = "South"
    lists["A3"] = "North"
    lists["A4"] = None
    lists["B1"] = 1
    lists["B2"] = "=B1+1"
    for rule in rules:
        ws.add_data_validation(rule)
    if extra is not None:
        extra(wb)
    wb.save(path)
    wb.close()
    return path


def _only(suggestions: tuple[ValidationDomainSuggestion, ...]) -> ValidationDomainSuggestion:
    assert len(suggestions) == 1, suggestions
    return suggestions[0]


def test_inline_list_becomes_enum_with_quoting_and_blanks(tmp_path: Path) -> None:
    path = _book(
        tmp_path / "wb.xlsx",
        _dv("C12", type="list", formula1='"a,""b"",,a,2,2.5"', allow_blank=True),
    )
    s = _only(suggest_validation_domains(path))
    assert s.address == "Dash!C12"
    assert s.rule_id == "Dash#0"
    assert s.status == "domain"
    assert s.domain == {"enum": ["a", '"b"', 2, 2.5]}
    assert s.allow_blank is True
    assert s.reason is None
    assert s.validation.type == "list"
    assert s.validation.formula1 == '"a,""b"",,a,2,2.5"'


def test_shared_sqref_shares_rule_id_and_addresses_filter(tmp_path: Path) -> None:
    path = _book(
        tmp_path / "wb.xlsx",
        _dv("A1:A2 C3", type="whole", operator="between", formula1="1", formula2="5"),
        _dv("B1", type="decimal", formula1="0", formula2="0.5"),
    )
    all_ = suggest_validation_domains(path)
    assert [s.address for s in all_] == ["Dash!A1", "Dash!B1", "Dash!A2", "Dash!C3"]
    assert {s.rule_id for s in all_ if s.address != "Dash!B1"} == {"Dash#0"}
    b1 = next(s for s in all_ if s.address == "Dash!B1")
    assert b1.domain == {"real_between": {"min": 0, "max": 0.5}}
    a1 = next(s for s in all_ if s.address == "Dash!A1")
    assert a1.domain == {"between": {"min": 1, "max": 5}}

    picked = suggest_validation_domains(path, addresses=["'Dash'!$c$3", "Dash!Z9"])
    assert [s.address for s in picked] == ["Dash!C3"]


def test_range_and_named_range_lists_read_cached_constants(tmp_path: Path) -> None:
    def names(wb: Any) -> None:
        wb.defined_names["Regions"] = DefinedName("Regions", attr_text="'My Lists'!$A$1:$A$4")
        wb.defined_names["Calc"] = DefinedName("Calc", attr_text="'My Lists'!$B$1:$B$2")
        wb.defined_names["Two"] = DefinedName("Two", attr_text="'My Lists'!$A$1,'My Lists'!$A$2")

    path = _book(
        tmp_path / "wb.xlsx",
        _dv("A1", type="list", formula1="'My Lists'!$A$1:$A$4"),
        _dv("A2", type="list", formula1="=Regions"),
        _dv("A3", type="list", formula1="Calc"),
        _dv("A4", type="list", formula1="Two"),
        _dv("A5", type="list", formula1="INDIRECT($B$1)"),
        _dv("A6", type="list", formula1="'My Lists'!$A$4"),
        extra=names,
    )
    by = {s.address: s for s in suggest_validation_domains(path)}
    assert by["Dash!A1"].domain == {"enum": ["North", "South"]}
    assert by["Dash!A2"].domain == {"enum": ["North", "South"]}
    assert by["Dash!A3"].reason == "list_formula_cells"
    assert by["Dash!A4"].reason == "list_multi_area"
    assert by["Dash!A5"].reason == "list_formula"
    assert by["Dash!A6"].reason == "empty_enum"
    assert by["Dash!A6"].domain is None


@pytest.mark.parametrize(
    ("kwargs", "reason"),
    [
        ({"type": "list", "formula1": '"a,b"', "errorStyle": "warning"}, "non_stop"),
        ({"type": "date", "formula1": "1", "formula2": "2"}, "unsupported_type"),
        ({"type": "custom", "formula1": "A1>0"}, "unsupported_type"),
        ({"type": "textLength", "formula1": "1", "formula2": "2"}, "unsupported_type"),
        (
            {"type": "whole", "operator": "greaterThanOrEqual", "formula1": "1"},
            "unsupported_operator",
        ),
        (
            {"type": "decimal", "operator": "notBetween", "formula1": "1", "formula2": "2"},
            "unsupported_operator",
        ),
        ({"type": "whole", "formula1": "$B$1", "formula2": "5"}, "non_literal_bounds"),
        ({"type": "whole", "formula1": "1.5", "formula2": "5"}, "non_literal_bounds"),
        ({"type": "list", "formula1": '""'}, "empty_enum"),
    ],
)
def test_unsupported_reasons(tmp_path: Path, kwargs: dict[str, Any], reason: str) -> None:
    path = _book(tmp_path / "wb.xlsx", _dv("B2", **kwargs))
    s = _only(suggest_validation_domains(path))
    assert s.status == "unsupported"
    assert s.reason == reason
    assert s.domain is None


def test_suggestions_compile_through_existing_env_path(tmp_path: Path) -> None:
    path = _book(
        tmp_path / "wb.xlsx",
        _dv("A1", type="list", formula1='"x,y"'),
        _dv("A2", type="whole", formula1="1", formula2="3"),
        _dv("A3", type="date", formula1="1", formula2="2"),
    )
    env = validation_cell_type_env(suggest_validation_domains(path))
    assert set(env) == {"Dash!A1", "Dash!A2"}
    assert env["Dash!A1"].kind is CellKind.STRING
    assert env["Dash!A1"].enum is not None
    assert env["Dash!A1"].enum.values == frozenset({"x", "y"})
    assert env["Dash!A2"].interval is not None
    assert (env["Dash!A2"].interval.min, env["Dash!A2"].interval.max) == (1, 3)


# --- bindings already exist -------------------------------------------------


def _series(sid: str, data_range: str, **extra: Any) -> dict[str, Any]:
    series: dict[str, Any] = {
        "id": sid,
        "sheet": "Dash",
        "data_range": data_range,
        "layout": "scalar",
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "string",
                "bind": {"kind": "data_cell", "read": "string"},
            },
            "dimensions": [],
        },
        "key": [],
    }
    series.update(extra)
    return series


def _doc(*series: dict[str, Any]) -> dict[str, Any]:
    return {"schema_version": "1.13.0", "series": list(series)}


def _codes(findings: Any) -> dict[str, str]:
    return {f.series_id: f.code for f in findings}


def test_compare_uniform_partial_mixed_and_blank(tmp_path: Path) -> None:
    path = _book(
        tmp_path / "wb.xlsx",
        _dv("A1:A2", type="list", formula1='"x,y"'),
        _dv("B1", type="list", formula1='"x,y"'),
        _dv("B2", type="list", formula1='"x,z"'),
        _dv("C1", type="list", formula1='"x,y"', allow_blank=True),
        _dv("C2", type="list", formula1='"x,y"'),
        _dv("D1", type="list", formula1='"x,y"'),
    )
    doc = _doc(
        _series("uniform", "Dash!A1:A2", input={}),
        _series("mixed", "Dash!B1:B2", input={}),
        _series("blank", "Dash!C1:C2", input={}),
        _series("partial", "Dash!D1:D2", input={}),
        _series("untouched", "Dash!F1", input={}),
    )
    findings = compare_series_validation_domains(doc, path)
    assert _codes(findings) == {
        "uniform": "proposed_domain",
        "mixed": "mixed_rules",
        "blank": "mixed_blank",
        "partial": "partial_coverage",
    }
    proposed = next(f for f in findings if f.series_id == "uniform")
    assert proposed.workbook_domain == {"enum": ["x", "y"]}


def test_compare_authored_domain_wins_and_constants_are_skipped(tmp_path: Path) -> None:
    path = _book(
        tmp_path / "wb.xlsx",
        _dv("A1", type="list", formula1='"x,y"'),
        _dv("B1", type="list", formula1='"x,y"'),
        _dv("C1", type="list", formula1='"x,y"'),
        _dv("D1", type="list", formula1='"x,y"'),
    )
    doc = _doc(
        _series("agree", "Dash!A1", input={"domain": {"enum": ["y", "x"]}}),
        _series("conflict", "Dash!B1", domain={"enum": ["x"]}, input={}),
        _series("const", "Dash!C1", constant={}),
        _series("from_wb", "Dash!D1", domain={"from_workbook": True}, input={}),
    )
    findings = compare_series_validation_domains(doc, path)
    assert _codes(findings) == {"conflict": "domain_validation_conflict"}
    (f,) = findings
    assert f.declared_domain == {"enum": ["x"]}
    assert f.workbook_domain == {"enum": ["x", "y"]}


def test_compare_value_map_sides(tmp_path: Path) -> None:
    path = _book(
        tmp_path / "wb.xlsx",
        _dv("A1", type="list", formula1='"On,Off"'),
        _dv("B1", type="list", formula1='"on,off"'),
        _dv("C1", type="list", formula1='"maybe"'),
    )
    value_map = {"on": "On", "off": "Off"}
    doc = _doc(
        _series("needles", "Dash!A1", input={"value_map": value_map}),
        _series("keys", "Dash!B1", input={"value_map": value_map}),
        _series("neither", "Dash!C1", input={"value_map": value_map}),
    )
    assert _codes(compare_series_validation_domains(doc, path)) == {
        "needles": "matches_value_map",
        "keys": "value_map_side_mismatch",
        "neither": "value_map_validation_conflict",
    }
