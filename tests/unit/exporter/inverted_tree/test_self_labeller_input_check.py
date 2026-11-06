"""Issue 1070 — a self-labelling input must bind `.required` before iterating.

The widest formula-header input labels its own axis, so that axis stays off
the checker signature. `.required` is still a `DomainTemplate`. The checker
binds it to the value under validation. An external labeller stays a keyword
parameter. A required template can be a proper subset of the series domain.
"""

from __future__ import annotations

from pathlib import Path
from typing import Annotated

import pytest

from excel_grapher.core.cell_types import RealBetween
from excel_grapher.exporter.inverted_tree.emit import generate_inverted_tree_modules
from excel_grapher.grapher import DynamicRefConfig, create_dependency_graph
from excel_grapher.series_bindings import validate_bindings_document
from excel_grapher.series_bindings.workflow import output_series_targets
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)
from tests.unit.exporter.inverted_tree.test_model_from_workbook import _plant_formula_cache

_WIDE = {"real_between": {"min": -1.0e15, "max": 1.0e15}}


def _plant_years(path: Path, years: dict[str, int]) -> None:
    for coordinate, year in years.items():
        _plant_formula_cache(path, coordinate, str(year))


def _function_source(module: str, name: str) -> str:
    marker = f"def {name}"
    start = module.index(marker)
    nxt = module.find("\n\ndef ", start + 1)
    return module[start:] if nxt < 0 else module[start:nxt]


def _account_workbook(path: Path) -> Path:
    """Formula year headers 2017–2019 over values 1, 2, and 3."""
    write_workbook(
        path,
        {
            "Data": {
                "A1": 2017,
                "B1": "=A1",
                "C1": "=B1+1",
                "D1": "=C1+1",
                "B2": 1.0,
                "C2": 2.0,
                "D2": 3.0,
                "E2": "=B2+C2+D2",
            }
        },
    )
    _plant_years(path, {"B1": 2017, "C1": 2018, "D1": 2019})
    return path


def _account_bindings(*, domain: dict | None = _WIDE) -> dict:
    return bindings_document(
        series_entry(
            "current_account",
            "Data!B2:D2",
            layout="series",
            direction="input",
            header_row=1,
            domain=domain,
        ),
        series_entry("total", "Data!E2", direction="output"),
    )


def test_self_labeller_binds_required_to_the_value_under_validation(tmp_path: Path) -> None:
    """`from_defaults` walks the bound template and returns the workbook total."""
    path = _account_workbook(tmp_path / "self_labeller.xlsx")
    modules = generate_inverted(path, _account_bindings())
    check = _function_source(modules["validation.py"], "_check_current_account")
    assert check.startswith(
        "def _check_current_account(current_account: data.CurrentAccount) -> data.CurrentAccount:"
    )
    assert "TIME_PERIOD" not in check.splitlines()[0]
    assert "data.CURRENT_ACCOUNT.schema.validate(current_account)" in check
    assert ".schema.bind(" not in check
    assert "data.CURRENT_ACCOUNT.required.bind(TIME_PERIOD=current_account)" in check
    assert "CURRENT_ACCOUNT_POSITIONS" not in modules["data.py"]
    assert "for coordinate in data.CURRENT_ACCOUNT.required:" not in check
    assert "current_account.domain" not in check

    pkg = load_package(modules, tmp_path, name="self_labeller_check")
    assert pkg.compute_total(pkg.TotalInputs.from_defaults()) == 6.0
    assert pkg.Model.from_defaults().total == 6.0
    bad = pkg.data.CURRENT_ACCOUNT.with_nested((1.0e16, 2.0, 3.0))
    with pytest.raises(ValueError, match=r"current_account.*out of domain"):
        pkg.TotalInputs(current_account=bad)


def test_self_labeller_binds_by_axis_name_when_it_is_not_first(tmp_path: Path) -> None:
    """`TIME_PERIOD` is axes[1]; the checker does not bind `INSTRUMENT`."""
    path = tmp_path / "grid.xlsx"
    write_workbook(
        path,
        {
            "Inputs": {
                "Z1": 2017,
                "A2": 1,
                "A3": 2,
                "B1": "=Z1",
                "C1": "=B1+1",
                "B2": 10.0,
                "C2": 20.0,
                "B3": 30.0,
                "C3": 40.0,
                "D2": "=B2+C2+B3+C3",
            }
        },
    )
    _plant_years(path, {"B1": 2017, "C1": 2018})
    document = bindings_document(
        {
            "id": "grid",
            "sheet": "Inputs",
            "data_range": "Inputs!B2:C3",
            "layout": "series",
            "input": {},
            "domain": dict(_WIDE),
            "key": ["INSTRUMENT", "TIME_PERIOD"],
            "structure": {
                "measure": {
                    "concept": "OBS_VALUE",
                    "dtype": "float",
                    "bind": {"kind": "data_cell", "read": "float"},
                },
                "dimensions": [
                    {
                        "id": "INSTRUMENT",
                        "concept": "INSTRUMENT",
                        "role": "key",
                        "scope": "cell",
                        "bind": {"kind": "row_label", "label_column": "A", "read": "int"},
                    },
                    {
                        "id": "TIME_PERIOD",
                        "concept": "TIME_PERIOD",
                        "role": "key",
                        "scope": "cell",
                        "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
                    },
                ],
            },
        },
        series_entry("total", "Inputs!D2", direction="output"),
    )
    document["concept_scheme"]["concepts"].append({"id": "INSTRUMENT", "dtype": "int"})
    modules = generate_inverted(path, document)
    check = _function_source(modules["validation.py"], "_check_grid")
    assert "data.GRID.schema.validate(grid)" in check
    assert "data.GRID.required.bind(TIME_PERIOD=grid)" in check
    assert "INSTRUMENT=" not in check
    pkg = load_package(modules, tmp_path, name="named_axis_labeller")
    assert pkg.compute_total(pkg.TotalInputs.from_defaults()) == 100.0


def test_external_labeller_stays_a_keyword_parameter(tmp_path: Path) -> None:
    """A narrower window binds `TIME_PERIOD=TIME_PERIOD`; the widest binds itself."""
    path = tmp_path / "windows.xlsx"
    write_workbook(
        path,
        {
            "Inputs": {
                "A1": 2024,
                "A2": "=A1",
                "B2": "=A2+1",
                "C2": "=B2+1",
                "D2": "=C2+1",
                "A3": 1.0,
                "B3": 2.0,
                "C3": 3.0,
                "D3": 4.0,
                "A4": 10.0,
                "B4": 20.0,
                "E3": "=A3+B3+C3+D3",
                "E4": "=A4+B4",
            }
        },
    )
    _plant_years(path, {"A2": 2024, "B2": 2025, "C2": 2026, "D2": 2027})
    document = bindings_document(
        series_entry(
            "span",
            "Inputs!A3:D3",
            layout="series",
            direction="input",
            header_row=2,
            domain=_WIDE,
        ),
        series_entry(
            "early",
            "Inputs!A4:B4",
            layout="series",
            direction="input",
            header_row=2,
            domain=_WIDE,
        ),
        series_entry("span_total", "Inputs!E3", direction="output"),
        series_entry("early_total", "Inputs!E4", direction="output"),
    )
    modules = generate_inverted(path, document)
    span = _function_source(modules["validation.py"], "_check_span")
    early = _function_source(modules["validation.py"], "_check_early")
    assert "data.SPAN.required.bind(TIME_PERIOD=span)" in span
    assert "TIME_PERIOD:" not in span.splitlines()[0]
    assert "data.EARLY.schema.bind(TIME_PERIOD=TIME_PERIOD).validate(early)" in early
    assert "data.EARLY.required.bind(TIME_PERIOD=TIME_PERIOD)" in early
    assert "TIME_PERIOD:" in early
    pkg = load_package(modules, tmp_path, name="external_labeller_check")
    assert pkg.compute_span_total(pkg.SpanTotalInputs.from_defaults()) == 10.0
    assert pkg.compute_early_total(pkg.EarlyTotalInputs.from_defaults()) == 30.0


def test_mixed_self_and_external_axes_bind_by_name(tmp_path: Path) -> None:
    """`INSTRUMENT` is an external labeller; `TIME_PERIOD` is this series."""
    path = tmp_path / "mixed.xlsx"
    write_workbook(
        path,
        {
            "Inputs": {
                "A1": 1,
                "A2": "=A1",
                "A3": "=A2+1",
                "A4": "=A3+1",
                "B2": 1.0,
                "B3": 2.0,
                "B4": 3.0,
                "C2": "=B2+B3+B4",
                "Z1": 2017,
                "D1": "=Z1",
                "E1": "=D1+1",
                "D2": 10.0,
                "E2": 20.0,
                "D3": 30.0,
                "E3": 40.0,
                "F2": "=D2+E2+D3+E3",
            }
        },
    )
    _plant_years(path, {"A2": 1, "A3": 2, "A4": 3, "D1": 2017, "E1": 2018})
    document = bindings_document(
        series_entry(
            "codes",
            "Inputs!B2:B4",
            layout="series",
            direction="input",
            label_column="A",
            key_concept="INSTRUMENT",
            domain=_WIDE,
        ),
        {
            "id": "grid",
            "sheet": "Inputs",
            "data_range": "Inputs!D2:E3",
            "layout": "series",
            "input": {},
            "domain": dict(_WIDE),
            "key": ["INSTRUMENT", "TIME_PERIOD"],
            "structure": {
                "measure": {
                    "concept": "OBS_VALUE",
                    "dtype": "float",
                    "bind": {"kind": "data_cell", "read": "float"},
                },
                "dimensions": [
                    {
                        "id": "INSTRUMENT",
                        "concept": "INSTRUMENT",
                        "role": "key",
                        "scope": "cell",
                        "bind": {"kind": "row_label", "label_column": "A", "read": "int"},
                    },
                    {
                        "id": "TIME_PERIOD",
                        "concept": "TIME_PERIOD",
                        "role": "key",
                        "scope": "cell",
                        "bind": {"kind": "column_header", "header_row": 1, "read": "int"},
                    },
                ],
            },
        },
        series_entry("codes_total", "Inputs!C2", direction="output"),
        series_entry("grid_total", "Inputs!F2", direction="output"),
    )
    document["concept_scheme"]["concepts"].append({"id": "INSTRUMENT", "dtype": "int"})
    modules = generate_inverted(path, document)
    grid = _function_source(modules["validation.py"], "_check_grid")
    codes = _function_source(modules["validation.py"], "_check_codes")
    assert "data.GRID.schema.bind(INSTRUMENT=INSTRUMENT, TIME_PERIOD=grid).validate(grid)" in grid
    assert "data.GRID.required.bind(INSTRUMENT=INSTRUMENT, TIME_PERIOD=grid)" in grid
    assert (
        "data.GRID.schema.bind(INSTRUMENT=codes, TIME_PERIOD=grid).validate(grid)"
        in modules["internals.py"]
    )
    assert "TIME_PERIOD:" not in grid.split("required.bind")[0]
    assert "data.CODES.required.bind(INSTRUMENT=codes)" in codes
    assert "TIME_PERIOD" not in codes
    pkg = load_package(modules, tmp_path, name="mixed_axis_labeller")
    assert pkg.compute_codes_total(pkg.CodesTotalInputs.from_defaults()) == 6.0
    assert pkg.compute_grid_total(pkg.GridTotalInputs.from_defaults()) == 100.0
    nested = ((10.0, 20.0), (30.0, 40.0))
    with pytest.raises(pkg.tensor.SchemaError, match="grid: expected Tensor"):
        pkg.GridTotalInputs(grid=nested, codes=pkg.data.CODES)
    with pytest.raises(pkg.tensor.SchemaError, match="grid: expected Tensor"):
        pkg.internals.grid_total(grid=nested, codes=pkg.data.CODES)
    with pytest.raises(pkg.tensor.SchemaError, match="axis 'TIME_PERIOD'"):
        pkg.GridTotalInputs(grid=pkg.data.CODES, codes=pkg.data.CODES)


def test_required_subset_skips_coordinates_left_out_of_extraction(tmp_path: Path) -> None:
    """Off-graph years stay out of the checker even when they fail the domain."""
    path = tmp_path / "subset.xlsx"
    write_workbook(
        path,
        {
            "Data": {
                "A1": 2017,
                "B1": "=A1",
                "C1": "=B1+1",
                "D1": "=C1+1",
                "E1": "=D1+1",
                "B2": 100.0,
                "C2": 100.0,
                "D2": 3.0,
                "E2": 4.0,
                "F2": "=D2+E2",
            }
        },
    )
    _plant_years(path, {"B1": 2017, "C1": 2018, "D1": 2019, "E1": 2020})
    document = bindings_document(
        series_entry(
            "current_account",
            "Data!B2:E2",
            layout="series",
            direction="input",
            header_row=1,
            domain={"real_between": {"min": 0, "max": 10}},
        ),
        series_entry("total", "Data!F2", direction="output"),
    )
    bindings = validate_bindings_document(document)
    graph = create_dependency_graph(
        path,
        output_series_targets(bindings, workbook=path),
        load_values=True,
        capture_dependency_provenance=True,
    )
    modules = generate_inverted_tree_modules(
        graph,
        series_bindings=bindings,
        bindings_workbook=path,
    )
    check = _function_source(modules["validation.py"], "_check_current_account")
    assert "data.CURRENT_ACCOUNT.required.bind(TIME_PERIOD=current_account)" in check
    assert "CURRENT_ACCOUNT_POSITIONS" not in check
    assert "current_account.domain" not in check
    pkg = load_package(modules, tmp_path, name="required_subset")
    account = pkg.data.CURRENT_ACCOUNT
    assert list(account.required.bind(TIME_PERIOD=account)) == [(2019,), (2020,)]
    assert pkg.compute_total(pkg.TotalInputs.from_defaults()) == 7.0
    rejected = account.with_nested((100.0, 100.0, 1.0e16, 4.0))
    with pytest.raises(ValueError, match=r"current_account.*out of domain"):
        pkg.TotalInputs(current_account=rejected)


def test_element_constraint_loop_binds_the_self_labeller(tmp_path: Path) -> None:
    """Annotated element checks use the same required bind as measure domains."""
    path = _account_workbook(tmp_path / "annotated.xlsx")
    bounds = RealBetween(-1.0e15, 1.0e15)
    modules = generate_inverted(
        path,
        _account_bindings(domain=None),
        dynamic_refs=DynamicRefConfig.from_constraints(
            {
                "Data!B2": Annotated[float, bounds],
                "Data!C2": Annotated[float, bounds],
                "Data!D2": Annotated[float, bounds],
            }
        ),
    )
    check = _function_source(modules["validation.py"], "_check_current_account")
    assert "require_annotated_domain" in check
    assert "require_input_domain" not in check
    assert "data.CURRENT_ACCOUNT.required.bind(TIME_PERIOD=current_account)" in check
    pkg = load_package(modules, tmp_path, name="annotated_self_labeller")
    assert pkg.compute_total(pkg.TotalInputs.from_defaults()) == 6.0
    bad = pkg.data.CURRENT_ACCOUNT.with_nested((1.0e16, 2.0, 3.0))
    with pytest.raises(ValueError, match=r"current_account.*out of domain"):
        pkg.TotalInputs(current_account=bad)
