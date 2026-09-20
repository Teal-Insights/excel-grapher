"""Schema, validation, and compiler tests for series-level relations (schema 1.18.0)."""

from __future__ import annotations

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
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import (
    SeriesBindingsSchemaError,
    load_series_bindings,
    validate_bindings_document,
    validate_series_bindings,
)
from excel_grapher.series_bindings.domains import (
    SeriesRelationError,
    cell_type_env_from_bindings,
)
from tests.paths import SERIES_BINDINGS_FIXTURES as FIXTURES


def _keyed_int_series(
    *,
    series_id: str,
    data_range: str,
    domain: dict[str, Any] | None = None,
    relations: list[dict[str, str]] | None = None,
    dtype: str = "int",
    label_column: str = "A",
) -> dict[str, Any]:
    input_block: dict[str, Any] = {}
    if domain is not None:
        input_block["domain"] = domain
    series: dict[str, Any] = {
        "id": series_id,
        "sheet": "Inputs",
        "data_range": data_range,
        "layout": "series",
        "input": input_block,
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": dtype,
                "bind": {"kind": "data_cell", "read": dtype if dtype != "number" else "float"},
            },
            "dimensions": [
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
            ],
        },
        "key": ["INSTRUMENT"],
    }
    if relations is not None:
        series["relations"] = relations
    return series


def _relations_doc(*series: dict[str, Any]) -> dict[str, Any]:
    return {"schema_version": "1.18.0", "series": list(series)}


def _write_pair_workbook(
    path: Path,
    *,
    maturity_rows: tuple[tuple[str, int], tuple[str, int]] = (("loan_a", 10), ("loan_b", 20)),
    grace_rows: tuple[tuple[str, int], tuple[str, int]] = (("loan_a", 1), ("loan_b", 2)),
    maturity_start_row: int = 2,
) -> Path:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Inputs")
    for offset, (name, value) in enumerate(grace_rows):
        row = 2 + offset
        ws.write(f"A{row}", name)
        ws.write_number(f"G{row}", value)
    for offset, (name, value) in enumerate(maturity_rows):
        row = maturity_start_row + offset
        ws.write(f"A{row}", name)
        ws.write_number(f"H{row}", value)
    wb.close()
    return path


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


def test_validate_reports_unknown_partner_incomparable_dtype_cycle_and_reflexive(
    tmp_path: Path,
) -> None:
    workbook = _write_pair_workbook(tmp_path / "pair.xlsx")
    graph = create_dependency_graph(
        workbook, ["Inputs!G2", "Inputs!G3", "Inputs!H2", "Inputs!H3"], load_values=True
    )

    unknown = validate_bindings_document(
        _relations_doc(
            _keyed_int_series(
                series_id="input4_loan_maturity",
                data_range="Inputs!H2:H3",
                domain={"between": {"min": 0, "max": 80}},
                relations=[{"greater_than": "missing_partner"}],
            )
        )
    )
    report = validate_series_bindings(graph, unknown, workbook=workbook)
    assert any(issue["code"] == "unknown_relation_partner" for issue in report["issues"])

    incomparable = validate_bindings_document(
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
    report = validate_series_bindings(graph, incomparable, workbook=workbook)
    assert any(issue["code"] == "incomparable_relation_dtype" for issue in report["issues"])

    cyclic = validate_bindings_document(
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
    report = validate_series_bindings(graph, cyclic, workbook=workbook)
    assert any(issue["code"] == "cyclic_relation" for issue in report["issues"])

    reflexive = validate_bindings_document(
        _relations_doc(
            _keyed_int_series(
                series_id="input4_loan_maturity",
                data_range="Inputs!H2:H3",
                domain={"between": {"min": 0, "max": 80}},
                relations=[{"greater_than": "input4_loan_maturity"}],
            )
        )
    )
    report = validate_series_bindings(graph, reflexive, workbook=workbook)
    assert any(issue["code"] == "reflexive_greater_than" for issue in report["issues"])


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
    assert any(issue["code"] == "missing_relation_partner_key" for issue in report["issues"])


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
