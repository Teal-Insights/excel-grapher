"""Schema 1.19.0 series-level domain, from_workbook, and from_bindings compiler."""

from __future__ import annotations

from collections.abc import Mapping
from pathlib import Path
from typing import Annotated, Any, Literal

import pytest
import xlsxwriter

from excel_grapher.core.cell_types import (
    Between,
    CellType,
    EnumDomain,
    GreaterThanCell,
    RealBetween,
    constraints_to_cell_type_env,
    normalize_cell_type_env_key,
)
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.grapher.dynamic_refs import DynamicRefConfig
from excel_grapher.grapher.graph_pickle import dump_graph, load_graph
from excel_grapher.series_bindings import (
    SeriesBindingsSchemaError,
    load_series_bindings,
    normalize_series_entry,
    validate_bindings_document,
    validate_series_bindings,
)
from excel_grapher.series_bindings.domains import (
    SeriesDomainIndex,
    cell_type_env_from_bindings,
    undomained_leaves,
)
from excel_grapher.series_bindings.input_coerce import measure_domain_from_series
from tests.paths import INVERTED_TREE_TINY_DSA, INVERTED_TREE_TINY_DSA_LABELLED
from tests.paths import SERIES_BINDINGS_FIXTURES as FIXTURES


def _scalar_series(
    *,
    series_id: str = "shock_year",
    data_range: str = "Inputs!B21",
    dtype: str = "int",
    direction: str = "input",
    domain: dict[str, Any] | None = None,
    input_domain: dict[str, Any] | None = None,
    schema_version: str = "1.19.0",
) -> dict[str, Any]:
    series: dict[str, Any] = {
        "id": series_id,
        "sheet": "Inputs",
        "data_range": data_range,
        "layout": "scalar",
        direction: {},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": dtype,
                "bind": {"kind": "data_cell", "read": dtype if dtype != "number" else "float"},
            },
            "dimensions": [],
        },
        "key": [],
    }
    if domain is not None:
        series["domain"] = domain
    if input_domain is not None:
        series["input"] = {"domain": input_domain}
    return {"schema_version": schema_version, "series": [series]}


def _write_inputs_workbook(path: Path, cells: Mapping[str, object]) -> Path:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Inputs")
    for coord, value in cells.items():
        if isinstance(value, str):
            ws.write(coord, value)
        else:
            ws.write_number(coord, value)
    wb.close()
    return path


def test_schema_accepts_series_level_domain_kinds() -> None:
    for domain in (
        {"enum": ["Borvelia", "Litellia"]},
        {"between": {"min": 1, "max": 5}},
        {"real_between": {"min": 0.0, "max": 200.0}},
        {"from_workbook": True},
    ):
        dtype = "string" if "enum" in domain else ("int" if "between" in domain else "float")
        doc = _scalar_series(domain=domain, dtype=dtype)
        bindings = validate_bindings_document(doc)
        assert bindings["series"][0]["domain"] == domain


def test_schema_rejects_from_workbook_false_and_mixed_kinds() -> None:
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(_scalar_series(domain={"from_workbook": False}))
    with pytest.raises(SeriesBindingsSchemaError):
        validate_bindings_document(
            _scalar_series(domain={"from_workbook": True, "between": {"min": 0, "max": 1}})
        )


def test_input_domain_normalizes_to_series_level() -> None:
    domain = {"between": {"min": 1, "max": 5}}
    normalized = normalize_series_entry(_scalar_series(input_domain=domain)["series"][0])
    assert normalized["domain"] == domain
    assert normalized["input"]["domain"] == domain


def test_conflicting_series_and_input_domain_is_an_error(tmp_path: Path) -> None:
    workbook = _write_inputs_workbook(tmp_path / "wb.xlsx", {"B21": 2})
    graph = create_dependency_graph(workbook, ["Inputs!B21"], load_values=True)
    doc = _scalar_series(
        domain={"between": {"min": 1, "max": 5}},
        input_domain={"between": {"min": 0, "max": 10}},
    )
    bindings = validate_bindings_document(doc)
    report = validate_series_bindings(graph, bindings, workbook=workbook)
    assert any(issue["code"] == "conflicting_domain" for issue in report["issues"])


def test_measure_domain_from_series_reads_series_level_key() -> None:
    series = normalize_series_entry(
        _scalar_series(domain={"between": {"min": 1, "max": 5}})["series"][0]
    )
    assert measure_domain_from_series(series) == {"between": {"min": 1, "max": 5}}


def test_from_workbook_is_not_a_compute_domain() -> None:
    series = _scalar_series(domain={"from_workbook": True})["series"][0]
    assert measure_domain_from_series(series) is None


def test_constant_series_compile_from_workbook_values(tmp_path: Path) -> None:
    workbook = _write_inputs_workbook(tmp_path / "wb.xlsx", {"A10": "Borvelia", "A11": "Litellia"})
    doc = {
        "schema_version": "1.19.0",
        "series": [
            {
                "id": "country_profile_names",
                "sheet": "Inputs",
                "data_range": "Inputs!A10:A11",
                "layout": "series",
                "constant": {},
                "structure": {
                    "measure": {
                        "concept": "OBS_VALUE",
                        "dtype": "string",
                        "bind": {"kind": "data_cell", "read": "string"},
                    },
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
                        }
                    ],
                },
                "key": ["COUNTRY"],
            }
        ],
    }
    bindings = validate_bindings_document(doc)
    env = cell_type_env_from_bindings(bindings, workbook=workbook)
    expected = constraints_to_cell_type_env(
        {"Inputs!A10": Literal["Borvelia"], "Inputs!A11": Literal["Litellia"]},
        {},
    )
    assert env == expected


def test_internal_without_domain_is_omitted(tmp_path: Path) -> None:
    workbook = _write_inputs_workbook(tmp_path / "wb.xlsx", {"C5": 1})
    doc = _scalar_series(
        series_id="engine_year_labels",
        data_range="Inputs!C5",
        direction="internal",
    )
    bindings = validate_bindings_document(doc)
    assert cell_type_env_from_bindings(bindings, workbook=workbook) == {}


def test_internal_from_workbook_pins_formula_cells(tmp_path: Path) -> None:
    workbook = tmp_path / "wb.xlsx"
    wb = xlsxwriter.Workbook(workbook)
    ws = wb.add_worksheet("Inputs")
    ws.write_number("B3", 2026)
    ws.write_formula("C5", "=Inputs!B3", None, 2026)
    wb.close()
    doc = _scalar_series(
        series_id="engine_year_labels",
        data_range="Inputs!C5",
        direction="internal",
        domain={"from_workbook": True},
    )
    bindings = validate_bindings_document(doc)
    env = cell_type_env_from_bindings(bindings, workbook=workbook)
    expected = constraints_to_cell_type_env({"Inputs!C5": Literal[2026]}, {})
    assert env == expected


def test_from_workbook_missing_value_is_an_error(tmp_path: Path) -> None:
    workbook = _write_inputs_workbook(tmp_path / "wb.xlsx", {})
    wb = xlsxwriter.Workbook(workbook)
    ws = wb.add_worksheet("Inputs")
    ws.write_blank("B21", None)
    wb.close()
    graph = create_dependency_graph(workbook, ["Inputs!B21"], load_values=True)
    doc = _scalar_series(domain={"from_workbook": True})
    bindings = validate_bindings_document(doc)
    report = validate_series_bindings(graph, bindings, workbook=workbook)
    assert any(issue["code"] == "from_workbook_missing_value" for issue in report["issues"])


def test_domain_cell_not_in_graph_is_a_warning(tmp_path: Path) -> None:
    workbook = _write_inputs_workbook(tmp_path / "wb.xlsx", {"B21": 2, "B22": 3})
    graph = create_dependency_graph(workbook, ["Inputs!B21"], load_values=True)
    doc = _scalar_series(
        data_range="Inputs!B22",
        domain={"between": {"min": 1, "max": 5}},
    )
    bindings = validate_bindings_document(doc)
    report = validate_series_bindings(graph, bindings, workbook=workbook)
    assert any(issue["code"] == "domain_cell_not_in_graph" for issue in report["issues"])


def test_series_level_domain_compiles_with_relations(tmp_path: Path) -> None:
    workbook = tmp_path / "pair.xlsx"
    wb = xlsxwriter.Workbook(workbook)
    ws = wb.add_worksheet("Inputs")
    ws.write("A2", "loan_a")
    ws.write_number("G2", 1)
    ws.write_number("H2", 10)
    wb.close()
    doc = {
        "schema_version": "1.19.0",
        "series": [
            {
                "id": "grace",
                "sheet": "Inputs",
                "data_range": "Inputs!G2",
                "layout": "scalar",
                "input": {},
                "domain": {"between": {"min": 0, "max": 50}},
                "structure": {
                    "measure": {
                        "concept": "OBS_VALUE",
                        "dtype": "int",
                        "bind": {"kind": "data_cell", "read": "int"},
                    },
                    "dimensions": [],
                },
                "key": [],
            },
            {
                "id": "maturity",
                "sheet": "Inputs",
                "data_range": "Inputs!H2",
                "layout": "scalar",
                "input": {},
                "domain": {"between": {"min": 0, "max": 80}},
                "relations": [{"greater_than": "grace"}],
                "structure": {
                    "measure": {
                        "concept": "OBS_VALUE",
                        "dtype": "int",
                        "bind": {"kind": "data_cell", "read": "int"},
                    },
                    "dimensions": [],
                },
                "key": [],
            },
        ],
    }
    bindings = validate_bindings_document(doc)
    env = cell_type_env_from_bindings(bindings, workbook=workbook)
    expected = constraints_to_cell_type_env(
        {
            "Inputs!G2": Annotated[int, Between(0, 50)],
            "Inputs!H2": Annotated[int, Between(0, 80), GreaterThanCell("Inputs!G2")],
        },
        {},
    )
    assert env == expected


def test_tiny_dsa_bindings_compile_input_and_constant_domains() -> None:
    workbook = INVERTED_TREE_TINY_DSA / "tiny-dsa.xlsx"
    bindings = load_series_bindings(INVERTED_TREE_TINY_DSA / "bindings")
    compiled = cell_type_env_from_bindings(bindings, workbook=workbook)
    expected = constraints_to_cell_type_env(
        {
            "Inputs!B5": Literal["Borvelia", "Litellia", "Aurelium"],
            "Inputs!B21": Annotated[int, Between(1, 5)],
            "Inputs!B22": Literal[1, 2, 3],
            "Inputs!B10": Annotated[float, RealBetween(0.0, 200.0)],
        },
        {},
    )
    for address, cell_type in expected.items():
        assert compiled[address] == cell_type
    names = compiled["Inputs!A10"].enum
    assert names == EnumDomain(values=frozenset({"Borvelia"}))


def test_tiny_dsa_labelled_bindings_omit_uncached_formula_pins() -> None:
    workbook = INVERTED_TREE_TINY_DSA_LABELLED / "tiny-dsa-labelled.xlsx"
    bindings = load_series_bindings(INVERTED_TREE_TINY_DSA_LABELLED / "bindings")
    compiled = cell_type_env_from_bindings(bindings, workbook=workbook)
    # Engine!C5:G5 are formula pins (`from_workbook`) without cached values.
    # shock_year has no series-level domain (public axis is calendar years).
    for address in (
        "Engine!C5",
        "Engine!D5",
        "Engine!E5",
        "Engine!F5",
        "Engine!G5",
        "Inputs!B21",
    ):
        assert address not in compiled
    expected = constraints_to_cell_type_env(
        {
            "Inputs!B5": Literal["Borvelia", "Litellia", "Aurelium"],
            "Inputs!B22": Literal[1, 2, 3],
        },
        {},
    )
    for address, cell_type in expected.items():
        assert compiled[address] == cell_type


def test_input_value_map_needles_compile_when_domain_omitted(tmp_path: Path) -> None:
    workbook = _write_inputs_workbook(tmp_path / "wb.xlsx", {"B5": "B"})
    series = _scalar_series(
        series_id="country_name",
        data_range="Inputs!B5",
        dtype="string",
    )["series"][0]
    series["input"] = {"value_map": {"Borvelia": "B", "Litellia": "L"}}
    bindings = validate_bindings_document({"schema_version": "1.19.0", "series": [series]})
    env = cell_type_env_from_bindings(bindings, workbook=workbook)
    expected = constraints_to_cell_type_env({"Inputs!B5": Literal["B", "L"]}, {})
    assert env == expected
    assert measure_domain_from_series(bindings["series"][0]) == {
        "enum": frozenset({"Borvelia", "Litellia"})
    }


def test_overlay_constraints_win_per_key(tmp_path: Path) -> None:
    workbook = _write_inputs_workbook(tmp_path / "wb.xlsx", {"B21": 2, "B22": 3})
    bindings = validate_bindings_document(_scalar_series(domain={"between": {"min": 1, "max": 5}}))
    base = DynamicRefConfig.from_bindings(bindings, workbook)
    overlay = DynamicRefConfig.from_constraints(
        {"Inputs!B21": Annotated[int, Between(0, 10)], "Inputs!B22": Literal[1, 2, 3]},
    )
    merged, overrides = base.overlay(overlay)
    assert overrides == ("Inputs!B21",)
    interval = merged.cell_type_env["Inputs!B21"].interval
    assert interval is not None
    assert (interval.min, interval.max) == (0, 10)
    assert "Inputs!B22" in merged.cell_type_env


def test_from_bindings_matches_compiler(tmp_path: Path) -> None:
    workbook = _write_inputs_workbook(tmp_path / "wb.xlsx", {"B21": 2})
    bindings = validate_bindings_document(_scalar_series(domain={"between": {"min": 1, "max": 5}}))
    config = DynamicRefConfig.from_bindings(bindings, workbook)
    compiled = cell_type_env_from_bindings(bindings, workbook=workbook)
    assert dict(config.cell_type_env) == compiled
    assert "Inputs!B21" in config.cell_type_env
    assert isinstance(config.cell_type_env["Inputs!B21"], CellType)


def test_series_domain_index_is_lazy_mapping(tmp_path: Path) -> None:
    workbook = _write_inputs_workbook(tmp_path / "wb.xlsx", {"B21": 2, "B22": 3, "A10": "Borvelia"})
    doc = {
        "schema_version": "1.19.0",
        "series": [
            _scalar_series(domain={"between": {"min": 1, "max": 5}})["series"][0],
            {
                "id": "names",
                "sheet": "Inputs",
                "data_range": "Inputs!A10",
                "layout": "scalar",
                "constant": {},
                "structure": {
                    "measure": {
                        "concept": "OBS_VALUE",
                        "dtype": "string",
                        "bind": {"kind": "data_cell", "read": "string"},
                    },
                    "dimensions": [],
                },
                "key": [],
            },
        ],
    }
    bindings = validate_bindings_document(doc)
    index = SeriesDomainIndex.from_bindings(bindings, workbook=workbook)
    assert index.domain_for("Inputs!B22") is None
    assert "Inputs!B22" not in index
    shock = index.domain_for("Inputs!B21")
    assert shock is not None
    assert shock.interval is not None
    assert index["Inputs!A10"].enum is not None
    assert set(index) == {"Inputs!B21", "Inputs!A10"}


def test_cell_type_env_from_bindings_stays_lazy(tmp_path: Path) -> None:
    workbook = tmp_path / "mcve.xlsx"
    wb = xlsxwriter.Workbook(workbook)
    ws = wb.add_worksheet("Engine")
    for col in range(20):
        for row in range(50):
            ws.write_number(row, col, 0.0)
    wb.close()
    doc = {
        "schema_version": "1.19.0",
        "series": [
            {
                "id": "engine_grid",
                "sheet": "Engine",
                "data_range": "Engine!A1:T50",
                "layout": "scalar",
                "constant": {},
                "structure": {
                    "measure": {
                        "concept": "OBS_VALUE",
                        "dtype": "float",
                        "bind": {"kind": "data_cell", "read": "float"},
                    },
                    "dimensions": [],
                },
                "key": [],
            }
        ],
    }
    bindings = validate_bindings_document(doc)
    index = SeriesDomainIndex.from_bindings(bindings, workbook=workbook)
    assert index._expanded is None
    key = normalize_cell_type_env_key("Engine!A1")
    _ = index[key]
    assert index._expanded is None
    assert len(index._memo) == 1

    env = cell_type_env_from_bindings(bindings, workbook=workbook)
    assert isinstance(env, SeriesDomainIndex)
    assert env._expanded is None
    _ = env[key]
    assert env._expanded is None, "single lookup should stay lazy"
    assert len(env._memo) == 1
    assert env[key].enum is not None


def test_attach_domains_and_undomained_leaves(tmp_path: Path) -> None:
    workbook = _write_inputs_workbook(tmp_path / "wb.xlsx", {"B21": 2, "B22": 3})
    graph = create_dependency_graph(workbook, ["Inputs!B21", "Inputs!B22"], load_values=True)
    bindings = validate_bindings_document(_scalar_series(domain={"between": {"min": 1, "max": 5}}))
    index = graph.attach_domains(bindings, workbook=workbook)
    assert graph.domains is index
    assert graph.cell_type_env is index
    assert graph.domain_for("Inputs!B21") is not None
    assert undomained_leaves(graph, bindings, workbook=workbook) == ["Inputs!B22"]


def test_pickle_stores_domain_handle(tmp_path: Path) -> None:
    workbook = INVERTED_TREE_TINY_DSA / "tiny-dsa.xlsx"
    bindings_path = INVERTED_TREE_TINY_DSA / "bindings"
    bindings = load_series_bindings(bindings_path)
    config = DynamicRefConfig.from_bindings(bindings, workbook, bindings_path=bindings_path)
    graph = create_dependency_graph(
        workbook,
        ["Inputs!B5", "Inputs!B21"],
        load_values=True,
        dynamic_refs=config,
    )
    assert graph.domains is not None
    blob = tmp_path / "graph.pkl"
    dump_graph(graph, blob)
    loaded = load_graph(blob)
    assert loaded.domains is not None
    assert loaded.domain_for("Inputs!B21") is not None
    assert loaded.domain_for("Inputs!B5") is not None


def test_load_normalizes_legacy_input_domain_on_relations_fixture() -> None:
    bindings = load_series_bindings(FIXTURES / "relations_greater_than.yaml")
    maturity = next(
        series for series in bindings["series"] if series["id"] == "input4_loan_maturity"
    )
    assert maturity["domain"] == {"between": {"min": 0, "max": 80}}
    assert maturity["input"]["domain"] == {"between": {"min": 0, "max": 80}}
