"""Schema, validation, and compiler tests for series-level relations (schema 1.18.0)."""

from __future__ import annotations

from collections.abc import Sequence
from pathlib import Path
from typing import Annotated, Any

import pytest
import xlsxwriter

from excel_grapher.core.cell_types import (
    Between,
    GreaterThanCell,
    NotEqualCell,
    constraints_to_cell_type_env,
)
from excel_grapher.grapher import DependencyGraph, create_dependency_graph
from excel_grapher.series_bindings import (
    SeriesBindingsLoadError,
    SeriesBindingsSchemaError,
    load_series_bindings,
    merge_series_binding_documents,
    validate_bindings_document,
    validate_series_bindings,
)
from excel_grapher.series_bindings.domains import (
    SeriesRelationError,
    cell_type_env_from_bindings,
)
from tests.paths import SERIES_BINDINGS_FIXTURES as FIXTURES

_QUOTED_SHEET = "Input 4 - External Financing"


def _scenario_dimension() -> dict[str, Any]:
    return {
        "id": "SCENARIO",
        "concept": "SCENARIO",
        "role": "key",
        "scope": "series",
        "bind": {"kind": "constant", "value": "baseline"},
    }


def _keyed_int_series(
    *,
    series_id: str,
    data_range: str,
    domain: dict[str, Any] | None = None,
    relations: list[dict[str, str]] | None = None,
    dtype: str = "int",
    label_column: str = "A",
    sheet: str = "Inputs",
    key: Sequence[str] | None = None,
    extra_dimensions: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    input_block: dict[str, Any] = {}
    if domain is not None:
        input_block["domain"] = domain
    dimensions: list[dict[str, Any]] = [
        {
            "id": "INSTRUMENT",
            "concept": "INSTRUMENT",
            "role": "key",
            "scope": "cell",
            "bind": {
                "kind": "row_label",
                "label_column": label_column,
                "read": "string",
                "normalize": "strip",
            },
        }
    ]
    if extra_dimensions:
        dimensions.extend(extra_dimensions)
    series: dict[str, Any] = {
        "id": series_id,
        "sheet": sheet,
        "data_range": data_range,
        "layout": "series",
        "input": input_block,
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": dtype,
                "bind": {"kind": "data_cell", "read": dtype if dtype != "number" else "float"},
            },
            "dimensions": dimensions,
        },
        "key": list(key) if key is not None else ["INSTRUMENT"],
    }
    if relations is not None:
        series["relations"] = relations
    return series


def _relations_doc(*series: dict[str, Any]) -> dict[str, Any]:
    return {"schema_version": "1.18.0", "series": list(series)}


def _write_pair_workbook(
    path: Path,
    *,
    sheet: str = "Inputs",
    maturity_rows: tuple[tuple[str, int], tuple[str, int]] = (("loan_a", 10), ("loan_b", 20)),
    grace_rows: tuple[tuple[str, int], tuple[str, int]] = (("loan_a", 1), ("loan_b", 2)),
    maturity_start_row: int = 2,
    tenor_rows: tuple[tuple[str, int], tuple[str, int]] | None = None,
) -> Path:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet(sheet)
    for offset, (name, value) in enumerate(grace_rows):
        row = 2 + offset
        ws.write(f"A{row}", name)
        ws.write_number(f"G{row}", value)
    for offset, (name, value) in enumerate(maturity_rows):
        row = maturity_start_row + offset
        ws.write(f"A{row}", name)
        ws.write_number(f"H{row}", value)
    if tenor_rows is not None:
        for offset, (name, value) in enumerate(tenor_rows):
            row = 2 + offset
            ws.write(f"A{row}", name)
            ws.write_number(f"I{row}", value)
    wb.close()
    return path


def _pair_graph(tmp_path: Path) -> tuple[Path, DependencyGraph]:
    workbook = _write_pair_workbook(tmp_path / "pair.xlsx")
    graph = create_dependency_graph(
        workbook, ["Inputs!G2", "Inputs!G3", "Inputs!H2", "Inputs!H3"], load_values=True
    )
    return workbook, graph


def _issue_codes(report: dict[str, Any]) -> set[str]:
    return {issue["code"] for issue in report["issues"]}


def test_schema_accepts_greater_than_and_not_equal_relations() -> None:
    doc = _relations_doc(
        _keyed_int_series(
            series_id="input4_grace_period",
            data_range="Inputs!G2:G3",
            domain={"between": {"min": 0, "max": 50}},
        ),
        _keyed_int_series(
            series_id="input4_loan_maturity",
            data_range="Inputs!H2:H3",
            domain={"between": {"min": 0, "max": 80}},
            relations=[{"greater_than": "input4_grace_period"}],
        ),
        _keyed_int_series(
            series_id="pv_lc_tenor",
            data_range="Inputs!G2:G3",
            domain={"between": {"min": 1, "max": 40}},
        ),
        _keyed_int_series(
            series_id="pv_lc_tenor_mirror",
            data_range="Inputs!H2:H3",
            domain={"between": {"min": 1, "max": 40}},
            relations=[{"not_equal": "pv_lc_tenor"}],
        ),
    )
    bindings = validate_bindings_document(doc)
    assert bindings["series"][1]["relations"] == [{"greater_than": "input4_grace_period"}]
    assert bindings["series"][3]["relations"] == [{"not_equal": "pv_lc_tenor"}]


def test_schema_rejects_address_valued_and_malformed_relations() -> None:
    maturity = _keyed_int_series(
        series_id="input4_loan_maturity",
        data_range="Inputs!H2:H3",
        domain={"between": {"min": 0, "max": 80}},
        relations=[{"greater_than": "Inputs!G2"}],
    )
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(
            _relations_doc(
                _keyed_int_series(series_id="input4_grace_period", data_range="Inputs!G2:G3"),
                maturity,
            )
        )

    maturity["relations"] = [{"greater_than": "input4_grace_period", "not_equal": "other"}]
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(
            _relations_doc(
                _keyed_int_series(series_id="input4_grace_period", data_range="Inputs!G2:G3"),
                maturity,
            )
        )

    maturity["relations"] = [{}]
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(
            _relations_doc(
                _keyed_int_series(series_id="input4_grace_period", data_range="Inputs!G2:G3"),
                maturity,
            )
        )


def test_validate_reports_unknown_relation_partner(tmp_path: Path) -> None:
    workbook, graph = _pair_graph(tmp_path)
    bindings = validate_bindings_document(
        _relations_doc(
            _keyed_int_series(
                series_id="input4_loan_maturity",
                data_range="Inputs!H2:H3",
                domain={"between": {"min": 0, "max": 80}},
                relations=[{"greater_than": "missing_partner"}],
            )
        )
    )
    report = validate_series_bindings(graph, bindings, workbook=workbook)
    assert "unknown_relation_partner" in _issue_codes(report)


def test_validate_reports_incomparable_relation_dtype(tmp_path: Path) -> None:
    workbook, graph = _pair_graph(tmp_path)
    bindings = validate_bindings_document(
        _relations_doc(
            _keyed_int_series(
                series_id="input4_grace_period",
                data_range="Inputs!G2:G3",
                domain={"enum": ["loan_a", "loan_b"]},
                dtype="string",
            ),
            _keyed_int_series(
                series_id="input4_loan_maturity",
                data_range="Inputs!H2:H3",
                domain={"between": {"min": 0, "max": 80}},
                relations=[{"greater_than": "input4_grace_period"}],
            ),
        )
    )
    report = validate_series_bindings(graph, bindings, workbook=workbook)
    assert "incomparable_relation_dtype" in _issue_codes(report)


def test_validate_reports_cyclic_greater_than(tmp_path: Path) -> None:
    workbook, graph = _pair_graph(tmp_path)
    bindings = validate_bindings_document(
        _relations_doc(
            _keyed_int_series(
                series_id="input4_grace_period",
                data_range="Inputs!G2:G3",
                domain={"between": {"min": 0, "max": 50}},
                relations=[{"greater_than": "input4_loan_maturity"}],
            ),
            _keyed_int_series(
                series_id="input4_loan_maturity",
                data_range="Inputs!H2:H3",
                domain={"between": {"min": 0, "max": 80}},
                relations=[{"greater_than": "input4_grace_period"}],
            ),
        )
    )
    report = validate_series_bindings(graph, bindings, workbook=workbook)
    assert "cyclic_relation" in _issue_codes(report)


def test_validate_reports_reflexive_greater_than(tmp_path: Path) -> None:
    workbook, graph = _pair_graph(tmp_path)
    bindings = validate_bindings_document(
        _relations_doc(
            _keyed_int_series(
                series_id="input4_loan_maturity",
                data_range="Inputs!H2:H3",
                domain={"between": {"min": 0, "max": 80}},
                relations=[{"greater_than": "input4_loan_maturity"}],
            )
        )
    )
    report = validate_series_bindings(graph, bindings, workbook=workbook)
    assert "reflexive_relation" in _issue_codes(report)


def test_validate_reports_reflexive_not_equal(tmp_path: Path) -> None:
    workbook, graph = _pair_graph(tmp_path)
    bindings = validate_bindings_document(
        _relations_doc(
            _keyed_int_series(
                series_id="pv_lc_tenor",
                data_range="Inputs!G2:G3",
                domain={"between": {"min": 1, "max": 40}},
                relations=[{"not_equal": "pv_lc_tenor"}],
            )
        )
    )
    report = validate_series_bindings(graph, bindings, workbook=workbook)
    assert "reflexive_relation" in _issue_codes(report)
    with pytest.raises(SeriesRelationError, match="not_equal itself"):
        cell_type_env_from_bindings(bindings, workbook=workbook)


def test_validate_reports_incompatible_relation_key(tmp_path: Path) -> None:
    workbook, graph = _pair_graph(tmp_path)
    scalar = {
        "id": "input4_grace_period",
        "sheet": "Inputs",
        "data_range": "Inputs!G2",
        "layout": "scalar",
        "input": {"domain": {"between": {"min": 0, "max": 50}}},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": "int",
                "bind": {"kind": "data_cell", "read": "int"},
            },
            "dimensions": [],
        },
        "key": [],
    }
    bindings = validate_bindings_document(
        _relations_doc(
            scalar,
            _keyed_int_series(
                series_id="input4_loan_maturity",
                data_range="Inputs!H2:H3",
                domain={"between": {"min": 0, "max": 80}},
                relations=[{"greater_than": "input4_grace_period"}],
            ),
        )
    )
    report = validate_series_bindings(graph, bindings, workbook=workbook)
    assert "incompatible_relation_key" in _issue_codes(report)


def test_validate_accepts_reversed_key_field_order(tmp_path: Path) -> None:
    workbook, graph = _pair_graph(tmp_path)
    extra = [_scenario_dimension()]
    bindings = validate_bindings_document(
        _relations_doc(
            _keyed_int_series(
                series_id="input4_grace_period",
                data_range="Inputs!G2:G3",
                domain={"between": {"min": 0, "max": 50}},
                extra_dimensions=extra,
                key=["INSTRUMENT", "SCENARIO"],
            ),
            _keyed_int_series(
                series_id="input4_loan_maturity",
                data_range="Inputs!H2:H3",
                domain={"between": {"min": 0, "max": 80}},
                relations=[{"greater_than": "input4_grace_period"}],
                extra_dimensions=extra,
                key=["SCENARIO", "INSTRUMENT"],
            ),
        )
    )
    report = validate_series_bindings(graph, bindings, workbook=workbook)
    assert "incompatible_relation_key" not in _issue_codes(report)
    env = cell_type_env_from_bindings(bindings, workbook=workbook)
    assert env["Inputs!H2"].relations == (GreaterThanCell("Inputs!G2"),)


def test_validate_reports_missing_partner_cell_at_key(tmp_path: Path) -> None:
    workbook = _write_pair_workbook(tmp_path / "pair.xlsx")
    graph = create_dependency_graph(
        workbook, ["Inputs!G2", "Inputs!H2", "Inputs!H3"], load_values=True
    )
    bindings = validate_bindings_document(
        _relations_doc(
            _keyed_int_series(
                series_id="input4_grace_period",
                data_range="Inputs!G2",
                domain={"between": {"min": 0, "max": 50}},
            ),
            _keyed_int_series(
                series_id="input4_loan_maturity",
                data_range="Inputs!H2:H3",
                domain={"between": {"min": 0, "max": 80}},
                relations=[{"greater_than": "input4_grace_period"}],
            ),
        )
    )
    report = validate_series_bindings(graph, bindings, workbook=workbook)
    assert "missing_relation_partner_key" in _issue_codes(report)


def test_validate_reports_ambiguous_relation_partner_key(tmp_path: Path) -> None:
    workbook = _write_pair_workbook(
        tmp_path / "dup.xlsx",
        grace_rows=(("loan_a", 1), ("loan_a", 2)),
        maturity_rows=(("loan_a", 10), ("loan_b", 20)),
        maturity_start_row=5,
    )
    graph = create_dependency_graph(
        workbook, ["Inputs!G2", "Inputs!G3", "Inputs!H5", "Inputs!H6"], load_values=True
    )
    bindings = validate_bindings_document(
        _relations_doc(
            _keyed_int_series(
                series_id="input4_grace_period",
                data_range="Inputs!G2:G3",
                domain={"between": {"min": 0, "max": 50}},
            ),
            _keyed_int_series(
                series_id="input4_loan_maturity",
                data_range="Inputs!H5:H6",
                domain={"between": {"min": 0, "max": 80}},
                relations=[{"greater_than": "input4_grace_period"}],
            ),
        )
    )
    report = validate_series_bindings(graph, bindings, workbook=workbook)
    assert "ambiguous_relation_partner_key" in _issue_codes(report)


def test_validate_reports_unresolved_relation_key(tmp_path: Path) -> None:
    workbook = tmp_path / "blank.xlsx"
    wb = xlsxwriter.Workbook(workbook)
    ws = wb.add_worksheet("Inputs")
    ws.write("A2", "loan_a")
    ws.write("A3", "loan_b")
    ws.write_number("G2", 1)
    ws.write_number("G3", 2)
    ws.write_number("H2", 10)
    ws.write_number("H3", 20)
    wb.close()
    graph = create_dependency_graph(
        workbook, ["Inputs!G2", "Inputs!G3", "Inputs!H2", "Inputs!H3"], load_values=True
    )
    bindings = validate_bindings_document(
        _relations_doc(
            _keyed_int_series(
                series_id="input4_grace_period",
                data_range="Inputs!G2:G3",
                domain={"between": {"min": 0, "max": 50}},
                label_column="B",
            ),
            _keyed_int_series(
                series_id="input4_loan_maturity",
                data_range="Inputs!H2:H3",
                domain={"between": {"min": 0, "max": 80}},
                relations=[{"greater_than": "input4_grace_period"}],
            ),
        )
    )
    report = validate_series_bindings(graph, bindings, workbook=workbook)
    assert "unresolved_relation_key" in _issue_codes(report)


def test_fixture_compiles_greater_than_cell_for_cell(tmp_path: Path) -> None:
    workbook = _write_pair_workbook(tmp_path / "relations_pair.xlsx")
    bindings = load_series_bindings(FIXTURES / "relations_greater_than.yaml")
    env = cell_type_env_from_bindings(bindings, workbook=workbook)
    expected = constraints_to_cell_type_env(
        {
            "Inputs!G2": Annotated[int, Between(0, 50)],
            "Inputs!G3": Annotated[int, Between(0, 50)],
            "Inputs!H2": Annotated[int, Between(0, 80), GreaterThanCell("Inputs!G2")],
            "Inputs!H3": Annotated[int, Between(0, 80), GreaterThanCell("Inputs!G3")],
        },
        {},
    )
    assert env == expected


def test_compiler_aligns_partner_by_key_not_position(tmp_path: Path) -> None:
    workbook = _write_pair_workbook(
        tmp_path / "swapped.xlsx",
        maturity_rows=(("loan_b", 20), ("loan_a", 10)),
        maturity_start_row=5,
    )
    bindings = validate_bindings_document(
        _relations_doc(
            _keyed_int_series(
                series_id="input4_grace_period",
                data_range="Inputs!G2:G3",
                domain={"between": {"min": 0, "max": 50}},
            ),
            _keyed_int_series(
                series_id="input4_loan_maturity",
                data_range="Inputs!H5:H6",
                domain={"between": {"min": 0, "max": 80}},
                relations=[{"greater_than": "input4_grace_period"}],
            ),
        )
    )
    env = cell_type_env_from_bindings(bindings, workbook=workbook)
    assert env["Inputs!H5"].relations == (GreaterThanCell("Inputs!G3"),)
    assert env["Inputs!H6"].relations == (GreaterThanCell("Inputs!G2"),)


def test_compiler_emits_not_equal_relations(tmp_path: Path) -> None:
    workbook = _write_pair_workbook(tmp_path / "tenor.xlsx")
    bindings = validate_bindings_document(
        _relations_doc(
            _keyed_int_series(
                series_id="pv_lc_tenor",
                data_range="Inputs!G2:G3",
                domain={"between": {"min": 1, "max": 40}},
            ),
            _keyed_int_series(
                series_id="pv_lc_tenor_mirror",
                data_range="Inputs!H2:H3",
                domain={"between": {"min": 1, "max": 40}},
                relations=[{"not_equal": "pv_lc_tenor"}],
            ),
        )
    )
    env = cell_type_env_from_bindings(bindings, workbook=workbook)
    expected = constraints_to_cell_type_env(
        {
            "Inputs!G2": Annotated[int, Between(1, 40)],
            "Inputs!G3": Annotated[int, Between(1, 40)],
            "Inputs!H2": Annotated[int, Between(1, 40), NotEqualCell("Inputs!G2")],
            "Inputs!H3": Annotated[int, Between(1, 40), NotEqualCell("Inputs!G3")],
        },
        {},
    )
    assert env == expected


def test_compiler_emits_multiple_relations_on_one_series(tmp_path: Path) -> None:
    workbook = _write_pair_workbook(
        tmp_path / "multi.xlsx",
        tenor_rows=(("loan_a", 3), ("loan_b", 4)),
    )
    bindings = validate_bindings_document(
        _relations_doc(
            _keyed_int_series(
                series_id="input4_grace_period",
                data_range="Inputs!G2:G3",
                domain={"between": {"min": 0, "max": 50}},
            ),
            _keyed_int_series(
                series_id="pv_lc_tenor",
                data_range="Inputs!I2:I3",
                domain={"between": {"min": 1, "max": 40}},
            ),
            _keyed_int_series(
                series_id="input4_loan_maturity",
                data_range="Inputs!H2:H3",
                domain={"between": {"min": 0, "max": 80}},
                relations=[
                    {"greater_than": "input4_grace_period"},
                    {"not_equal": "pv_lc_tenor"},
                ],
            ),
        )
    )
    env = cell_type_env_from_bindings(bindings, workbook=workbook)
    assert env["Inputs!H2"].relations == (
        GreaterThanCell("Inputs!G2"),
        NotEqualCell("Inputs!I2"),
    )
    assert env["Inputs!H3"].relations == (
        GreaterThanCell("Inputs!G3"),
        NotEqualCell("Inputs!I3"),
    )


def test_compiler_normalizes_quoted_sheet_relation_addresses(tmp_path: Path) -> None:
    workbook = _write_pair_workbook(tmp_path / "quoted.xlsx", sheet=_QUOTED_SHEET)
    quoted = f"'{_QUOTED_SHEET}'"
    bindings = validate_bindings_document(
        _relations_doc(
            _keyed_int_series(
                series_id="input4_grace_period",
                sheet=_QUOTED_SHEET,
                data_range=f"{quoted}!G2:G3",
                domain={"between": {"min": 0, "max": 50}},
            ),
            _keyed_int_series(
                series_id="input4_loan_maturity",
                sheet=_QUOTED_SHEET,
                data_range=f"{quoted}!H2:H3",
                domain={"between": {"min": 0, "max": 80}},
                relations=[{"greater_than": "input4_grace_period"}],
            ),
        )
    )
    env = cell_type_env_from_bindings(bindings, workbook=workbook)
    assert env[f"{_QUOTED_SHEET}!H2"].relations == (GreaterThanCell(f"{_QUOTED_SHEET}!G2"),)
    assert env[f"{_QUOTED_SHEET}!H3"].relations == (GreaterThanCell(f"{_QUOTED_SHEET}!G3"),)


def test_compiler_fails_closed_when_partner_has_no_cell_at_key(tmp_path: Path) -> None:
    workbook = _write_pair_workbook(tmp_path / "pair.xlsx")
    bindings = validate_bindings_document(
        _relations_doc(
            _keyed_int_series(
                series_id="input4_grace_period",
                data_range="Inputs!G2",
                domain={"between": {"min": 0, "max": 50}},
            ),
            _keyed_int_series(
                series_id="input4_loan_maturity",
                data_range="Inputs!H2:H3",
                domain={"between": {"min": 0, "max": 80}},
                relations=[{"greater_than": "input4_grace_period"}],
            ),
        )
    )
    with pytest.raises(SeriesRelationError, match="loan_b"):
        cell_type_env_from_bindings(bindings, workbook=workbook)


def test_complementary_shards_must_repeat_relations() -> None:
    left = _relations_doc(
        _keyed_int_series(
            series_id="input4_loan_maturity",
            data_range="Inputs!H2:H3",
            domain={"between": {"min": 0, "max": 80}},
            relations=[{"greater_than": "input4_grace_period"}],
        )
    )
    right = _relations_doc(
        _keyed_int_series(
            series_id="input4_loan_maturity",
            sheet="Alt",
            data_range="Alt!H2:H3",
            domain={"between": {"min": 0, "max": 80}},
        )
    )
    with pytest.raises(SeriesBindingsLoadError, match="structural fields differ"):
        merge_series_binding_documents([left, right])
