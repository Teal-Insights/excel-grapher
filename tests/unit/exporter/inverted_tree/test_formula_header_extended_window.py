"""Formula-header timelines whose formula series extend beyond the covering input."""

from __future__ import annotations

import subprocess
import zipfile
from pathlib import Path
from typing import Any

import pytest
from fastpyxl import Workbook

from excel_grapher import create_dependency_graph
from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.export_runtime.provenance import column_letter
from excel_grapher.exporter.inverted_tree import generate_inverted_tree_modules
from excel_grapher.series_bindings import validate_bindings_document
from excel_grapher.series_bindings.workflow import all_series_targets
from tests.unit.exporter.inverted_tree.helpers import load_package

FIRST = 2024
PRIOR = FIRST - 1
CUTOFF = FIRST + 1
LAST = FIRST + 2
IMF = "IMF"
TOTAL = "Total"
_REPO_ROOT = Path(__file__).resolve().parents[4]


def _assert_generated_package_typechecks(pkg_dir: Path) -> None:
    public_targets = [str(pkg_dir / name) for name in ("api.py", "model.py", "validation.py")]
    for cmd in (
        ["uv", "run", "--no-sync", "ruff", "check", *public_targets],
        ["uv", "run", "--no-sync", "ruff", "format", "--check", *public_targets],
        [
            "uv",
            "run",
            "--no-sync",
            "ruff",
            "check",
            "--select",
            "F821",
            str(pkg_dir / "internals.py"),
            str(pkg_dir / "data.py"),
        ],
        [
            "uv",
            "run",
            "--no-sync",
            "ty",
            "check",
            "--project",
            str(_REPO_ROOT),
            "--extra-search-path",
            str(pkg_dir.parent),
            "--ignore",
            "unresolved-attribute",
            str(pkg_dir / "api.py"),
            str(pkg_dir / "model.py"),
            str(pkg_dir / "internals.py"),
            str(pkg_dir / "data.py"),
        ],
    ):
        proc = subprocess.run(
            cmd,
            cwd=str(_REPO_ROOT),
            capture_output=True,
            text=True,
            check=False,
        )
        assert proc.returncode == 0, proc.stdout + proc.stderr


def _plant_formula_caches(path: Path, caches: dict[str, dict[str, int | float]]) -> None:
    """Keep each stored formula and set its saved calculated value."""
    with zipfile.ZipFile(path) as archive:
        files = {name: archive.read(name) for name in archive.namelist()}
    for sheet_xml, sheet_caches in caches.items():
        xml = files[sheet_xml].decode()
        for coordinate, cached in sheet_caches.items():
            marker = f'<c r="{coordinate}"'
            start = xml.find(marker)
            if start < 0:
                raise RuntimeError(f"missing cell {coordinate} in {sheet_xml}")
            end = xml.find("</c>", start)
            cell = xml[start:end]
            rendered = str(cached)
            if "<v />" in cell:
                cell = cell.replace("<v />", f"<v>{rendered}</v>", 1)
            elif "<v/>" in cell:
                cell = cell.replace("<v/>", f"<v>{rendered}</v>", 1)
            elif "<v>" in cell:
                close = cell.find("</v>")
                open_end = cell.find(">", cell.find("<v"))
                cell = cell[: open_end + 1] + rendered + cell[close:]
            else:
                cell = cell + f"<v>{rendered}</v>"
            xml = xml[:start] + cell + xml[end:]
        files[sheet_xml] = xml.encode()
    with zipfile.ZipFile(path, "w") as archive:
        for name, payload in files.items():
            archive.writestr(name, payload)


def _dimension(axis_id: str, bind: dict[str, object]) -> dict[str, object]:
    return {
        "id": axis_id,
        "concept": axis_id,
        "role": "key",
        "scope": "cell",
        "bind": bind,
    }


def _measure() -> dict[str, object]:
    return {
        "concept": "OBS_VALUE",
        "dtype": "float",
        "bind": {"kind": "data_cell", "read": "float"},
    }


def _write_mcve_workbook(path: Path, first_year: int) -> None:
    book = Workbook()
    inputs = book.active
    assert inputs is not None
    inputs.title = "Input"
    ext = book.create_sheet("Ext")
    inputs["A2"] = IMF
    inputs["B2"] = 10.0
    inputs["C2"] = 20.0
    inputs["D2"] = 30.0
    inputs["B1"] = f"={first_year}"
    inputs["C1"] = "=B1+1"
    inputs["D1"] = "=C1+1"
    ext["A2"] = IMF
    ext["A3"] = TOTAL
    ext["B1"] = "=Input!B1-1"
    ext["C1"] = "=Input!B1"
    ext["D1"] = "=Input!C1"
    ext["E1"] = "=Input!D1"
    ext["B2"] = "=0"
    ext["B3"] = "=0"
    ext["C2"] = "=IF(ISNUMBER(Input!B2),Input!B2,0)"
    ext["D2"] = "=IF(ISNUMBER(Input!C2),Input!C2,0)"
    ext["E2"] = "=IF(ISNUMBER(Input!D2),Input!D2,0)"
    ext["C3"] = "=C2"
    ext["D3"] = "=D2"
    book.save(path)
    book.close()
    _plant_formula_caches(
        path,
        {
            "xl/worksheets/sheet1.xml": {
                "B1": first_year,
                "C1": first_year + 1,
                "D1": first_year + 2,
            },
            "xl/worksheets/sheet2.xml": {
                "B1": first_year - 1,
                "C1": first_year,
                "D1": first_year + 1,
                "E1": first_year + 2,
                "B2": 0,
                "B3": 0,
                "C2": 10,
                "D2": 20,
                "E2": 30,
                "C3": 10,
                "D3": 20,
            },
        },
    )


def _mcve_bindings() -> dict[str, Any]:
    header = {"kind": "column_header", "header_row": 1, "read": "int"}
    instrument = {
        "kind": "row_label",
        "label_column": "A",
        "read": "string",
        "normalize": "strip",
    }
    return {
        "schema_version": "1.22.0",
        "concept_scheme": {
            "id": "old_debt",
            "concepts": [
                {"id": "TIME_PERIOD", "dtype": "int"},
                {"id": "INSTRUMENT", "dtype": "string"},
                {"id": "OBS_VALUE", "dtype": "number"},
            ],
        },
        "series": [
            {
                "id": "debt",
                "sheet": "Input",
                "data_range": "Input!B2:D2",
                "layout": "series",
                "input": {},
                "key": ["INSTRUMENT", "TIME_PERIOD"],
                "structure": {
                    "measure": _measure(),
                    "dimensions": [
                        _dimension("INSTRUMENT", instrument),
                        _dimension("TIME_PERIOD", header),
                    ],
                },
            },
            {
                "id": "service",
                "sheet": "Ext",
                "data_range": "Ext!B2:E3",
                "layout": "matrix",
                "output": {"compute": {"name": "compute_service"}},
                "key": ["INSTRUMENT", "TIME_PERIOD"],
                "structure": {
                    "measure": _measure(),
                    "dimensions": [
                        _dimension("INSTRUMENT", instrument),
                        _dimension("TIME_PERIOD", header),
                    ],
                },
            },
        ],
    }


def test_old_debt_service_prior_year_header_shifts_with_input_axis(tmp_path: Path) -> None:
    """Issue 1086 MCVE: prior-year formula header binds to `debt` and shifts cutoff."""
    template = tmp_path / "template.xlsx"
    _write_mcve_workbook(template, FIRST)
    document = validate_bindings_document(_mcve_bindings())
    targets = [
        "Input!B2",
        "Input!C2",
        "Input!D2",
        "Ext!C2",
        "Ext!D2",
        "Ext!E2",
        "Ext!C3",
        "Ext!D3",
    ]
    graph = create_dependency_graph(template, targets, load_values=True)
    with pytest.warns(UserWarning, match="Skipped 3 cell"):
        modules = generate_inverted_tree_modules(
            graph, series_bindings=document, bindings_workbook=template
        )
    body = modules["internals.py"][modules["internals.py"].index("def service") :]
    assert "2025" not in body
    assert ".required.bind(TIME_PERIOD=debt)" in body
    assert "_p_time_period <= 2" in body

    pkg = load_package(modules, tmp_path, name="olddebt")
    _assert_generated_package_typechecks(tmp_path / "olddebt")
    template_model = pkg.Model.from_workbook(template)
    assert template_model.debt.domain.axis("TIME_PERIOD").keys == (FIRST, CUTOFF, LAST)
    assert list(template_model.service.items()) == [
        ((IMF, FIRST), 10.0),
        ((IMF, CUTOFF), 20.0),
        ((IMF, LAST), 30.0),
        ((TOTAL, FIRST), 10.0),
        ((TOTAL, CUTOFF), 20.0),
    ]
    assert template_model.cells("service")[(IMF, PRIOR)] == "Ext!B2"
    assert template_model.cells("service")[(IMF, FIRST)] == "Ext!C2"

    for delta in (1, -2, 10):
        case = tmp_path / f"shifted_{delta}.xlsx"
        first = FIRST + delta
        _write_mcve_workbook(case, first)
        case_graph = create_dependency_graph(case, targets, load_values=True)
        expected = FormulaEvaluator(case_graph).evaluate(
            ["Ext!C2", "Ext!D2", "Ext!E2", "Ext!C3", "Ext!D3"]
        )

        shifted_model = pkg.Model.from_workbook(case)
        assert shifted_model.debt.domain.axis("TIME_PERIOD").keys == (first, first + 1, first + 2)
        assert list(shifted_model.service.items()) == [
            ((IMF, first), expected["Ext!C2"]),
            ((IMF, first + 1), expected["Ext!D2"]),
            ((IMF, first + 2), expected["Ext!E2"]),
            ((TOTAL, first), expected["Ext!C3"]),
            ((TOTAL, first + 1), expected["Ext!D3"]),
        ]
        via_compute = pkg.compute_service(pkg.ServiceInputs.from_workbook(case))
        assert list(via_compute.items()) == list(shifted_model.service.items())
        via_init = pkg.Model(debt=shifted_model.debt).service
        assert list(via_init.items()) == list(shifted_model.service.items())
        assert shifted_model.cells("service")[(IMF, first - 1)] == "Ext!B2"
        assert shifted_model.cells("service")[(IMF, first)] == "Ext!C2"
        assert shifted_model.cells("service")[(TOTAL, first + 1)] == "Ext!D3"


def test_on_graph_prior_and_trailing_years_without_input_carry_zero(tmp_path: Path) -> None:
    """On-graph years outside `debt`'s axis evaluate `ISNUMBER` blank refs to 0 across shifts."""

    def write_case(path: Path, first_year: int) -> None:
        book = Workbook()
        inputs = book.active
        assert inputs is not None
        inputs.title = "Input"
        ext = book.create_sheet("Ext")
        inputs["A2"] = IMF
        for idx, offset in enumerate(range(-1, 4)):
            col = column_letter(idx + 2)
            inputs[f"{col}1"] = (
                f"={first_year - 1}" if idx == 0 else f"={column_letter(idx + 1)}1+1"
            )
            ext[f"{col}1"] = f"=Input!{col}1"
            ext[f"{col}2"] = f"=IF(ISNUMBER(Input!{col}2),Input!{col}2,0)"
            if offset in (0, 1, 2):
                inputs[f"{col}2"] = float((offset + 1) * 10)
            if offset <= 1:
                ext[f"{col}3"] = f"={col}2"
        ext["A2"] = IMF
        ext["A3"] = TOTAL
        book.save(path)
        book.close()
        header_caches = {f"{column_letter(idx + 2)}1": first_year - 1 + idx for idx in range(5)}
        _plant_formula_caches(
            path,
            {
                "xl/worksheets/sheet1.xml": header_caches,
                "xl/worksheets/sheet2.xml": {
                    **header_caches,
                    "B2": 0,
                    "C2": 10,
                    "D2": 20,
                    "E2": 30,
                    "F2": 0,
                    "B3": 0,
                    "C3": 10,
                    "D3": 20,
                },
            },
        )

    header = {"kind": "column_header", "header_row": 1, "read": "int"}
    instrument = {"kind": "row_label", "label_column": "A", "read": "string", "normalize": "strip"}
    doc = validate_bindings_document(
        {
            "schema_version": "1.22.0",
            "concept_scheme": {
                "id": "extended_both_ends",
                "concepts": [
                    {"id": "TIME_PERIOD", "dtype": "int"},
                    {"id": "INSTRUMENT", "dtype": "string"},
                    {"id": "OBS_VALUE", "dtype": "number"},
                ],
            },
            "series": [
                {
                    "id": "debt",
                    "sheet": "Input",
                    "data_range": "Input!B2:F2",
                    "layout": "series",
                    "input": {},
                    "key": ["INSTRUMENT", "TIME_PERIOD"],
                    "structure": {
                        "measure": _measure(),
                        "dimensions": [
                            _dimension("INSTRUMENT", instrument),
                            _dimension("TIME_PERIOD", header),
                        ],
                    },
                },
                {
                    "id": "service",
                    "sheet": "Ext",
                    "data_range": "Ext!B2:F3",
                    "layout": "matrix",
                    "output": {"compute": {"name": "compute_service"}},
                    "key": ["INSTRUMENT", "TIME_PERIOD"],
                    "structure": {
                        "measure": _measure(),
                        "dimensions": [
                            _dimension("INSTRUMENT", instrument),
                            _dimension("TIME_PERIOD", header),
                        ],
                    },
                },
            ],
        }
    )
    blanks = ["Input!B2", "Input!F2"]
    template = tmp_path / "template.xlsx"
    write_case(template, FIRST)
    targets = [
        "Input!C2",
        "Input!D2",
        "Input!E2",
        "Ext!B2",
        "Ext!C2",
        "Ext!D2",
        "Ext!E2",
        "Ext!F2",
        "Ext!B3",
        "Ext!C3",
        "Ext!D3",
    ]
    graph = create_dependency_graph(template, targets, load_values=True, blank_ranges=blanks)
    with pytest.warns(UserWarning, match="Skipped 2 cell"):
        modules = generate_inverted_tree_modules(
            graph,
            series_bindings=doc,
            bindings_workbook=template,
            blank_ranges=blanks,
        )
    pkg = load_package(modules, tmp_path, name="extended_ends")

    for delta in (0, 1, 5):
        case = tmp_path / f"case_{delta}.xlsx"
        first = FIRST + delta
        write_case(case, first)
        case_graph = create_dependency_graph(case, targets, load_values=True, blank_ranges=blanks)
        expected = FormulaEvaluator(case_graph, blank_ranges=blanks).evaluate(targets[3:])
        model = pkg.Model.from_workbook(case)
        assert model.debt.domain.axis("TIME_PERIOD").keys == (first, first + 1, first + 2)
        assert list(model.service.items()) == [
            ((IMF, first - 1), expected["Ext!B2"]),
            ((IMF, first), expected["Ext!C2"]),
            ((IMF, first + 1), expected["Ext!D2"]),
            ((IMF, first + 2), expected["Ext!E2"]),
            ((IMF, first + 3), expected["Ext!F2"]),
            ((TOTAL, first - 1), expected["Ext!B3"]),
            ((TOTAL, first), expected["Ext!C3"]),
            ((TOTAL, first + 1), expected["Ext!D3"]),
        ]
        assert model.service[IMF, first - 1] == 0.0
        assert model.service[IMF, first + 3] == 0.0


def test_multi_labeller_extended_window_coexists_with_static_history_and_neighbours(
    tmp_path: Path,
) -> None:
    """LIC-DSF topology: multiple partial labellers, static history, neighbour total, and constants."""

    def write_case(path: Path, first_year: int) -> None:
        book = Workbook()
        sheet = book.active
        assert sheet is not None
        sheet.title = "Model"
        # Static literal history on columns B:C (2015, 2016)
        sheet["B1"] = 2015
        sheet["C1"] = 2016
        sheet["B2"] = 5.0
        sheet["C2"] = 6.0
        sheet["B3"] = "=B2+1"
        sheet["C3"] = "=C2+1"
        # Macro formula-header input on columns E:H (first-4 .. first-1, i.e. 2020..2023 -> overlaps 2023)
        macro_caches: dict[str, int | float] = {}
        for idx in range(4):
            col = column_letter(5 + idx)
            yr = first_year - 4 + idx
            sheet[f"{col}5"] = f"={yr}" if idx == 0 else f"={column_letter(4 + idx)}5+1"
            sheet[f"{col}6"] = float(idx + 1)
            macro_caches[f"{col}5"] = yr
        # Old debt formula-header input on columns J:M (first .. first+3, i.e. 2024..2027, wider overlap)
        for idx in range(4):
            col = column_letter(10 + idx)
            yr = first_year + idx
            sheet[f"{col}8"] = f"={yr}" if idx == 0 else f"={column_letter(9 + idx)}8+1"
            sheet[f"{col}9"] = float((idx + 1) * 10)
            macro_caches[f"{col}8"] = yr
        # Extended constant (I11:M11) and extended service (I12:M13) on first-1 .. first+3 (2023..2027)
        sheet["A12"] = IMF
        sheet["A13"] = TOTAL
        for idx in range(5):
            col = column_letter(9 + idx)
            yr = first_year - 1 + idx
            sheet[f"{col}10"] = "=J8-1" if idx == 0 else f"={column_letter(8 + idx)}10+1"
            macro_caches[f"{col}10"] = yr
            sheet[f"{col}11"] = float(idx)
            if idx == 0:
                sheet[f"{col}12"] = "=0"
                sheet[f"{col}13"] = "=0"
            else:
                debt_col = column_letter(9 + idx)
                sheet[f"{col}12"] = f"=IF(ISNUMBER({debt_col}9),{debt_col}9+{col}11,0)"
                if idx <= 2:
                    sheet[f"{col}13"] = f"={col}12"
        # Neighbour total on J15:M15 (first .. first+3) summing service row and pinned/range reads
        for idx in range(4):
            col = column_letter(10 + idx)
            sheet[f"{col}15"] = f"={col}12+$I$12"
        sheet["B17"] = "=SUM(J12:M12)"
        book.save(path)
        book.close()
        _plant_formula_caches(path, {"xl/worksheets/sheet1.xml": macro_caches})

    instrument = {"kind": "row_label", "label_column": "A", "read": "string", "normalize": "strip"}
    doc = validate_bindings_document(
        {
            "schema_version": "1.22.0",
            "concept_scheme": {
                "id": "lic_dsf_shape",
                "concepts": [
                    {"id": "TIME_PERIOD", "dtype": "int"},
                    {"id": "INSTRUMENT", "dtype": "string"},
                    {"id": "OBS_VALUE", "dtype": "number"},
                ],
            },
            "series": [
                {
                    "id": "static_hist",
                    "sheet": "Model",
                    "data_range": "Model!B2:C2",
                    "layout": "series",
                    "input": {},
                    "key": ["TIME_PERIOD"],
                    "structure": {
                        "measure": _measure(),
                        "dimensions": [
                            _dimension(
                                "TIME_PERIOD",
                                {"kind": "column_header", "header_row": 1, "read": "int"},
                            )
                        ],
                    },
                },
                {
                    "id": "macro",
                    "sheet": "Model",
                    "data_range": "Model!E6:H6",
                    "layout": "series",
                    "input": {},
                    "key": ["TIME_PERIOD"],
                    "structure": {
                        "measure": _measure(),
                        "dimensions": [
                            _dimension(
                                "TIME_PERIOD",
                                {"kind": "column_header", "header_row": 5, "read": "int"},
                            )
                        ],
                    },
                },
                {
                    "id": "debt",
                    "sheet": "Model",
                    "data_range": "Model!J9:M9",
                    "layout": "series",
                    "input": {},
                    "key": ["TIME_PERIOD"],
                    "structure": {
                        "measure": _measure(),
                        "dimensions": [
                            _dimension(
                                "TIME_PERIOD",
                                {"kind": "column_header", "header_row": 8, "read": "int"},
                            )
                        ],
                    },
                },
                {
                    "id": "spread",
                    "sheet": "Model",
                    "data_range": "Model!I11:M11",
                    "layout": "series",
                    "constant": {},
                    "key": ["TIME_PERIOD"],
                    "structure": {
                        "measure": _measure(),
                        "dimensions": [
                            _dimension(
                                "TIME_PERIOD",
                                {"kind": "column_header", "header_row": 10, "read": "int"},
                            )
                        ],
                    },
                },
                {
                    "id": "neighbour_total",
                    "sheet": "Model",
                    "data_range": "Model!J15:M15",
                    "layout": "series",
                    "output": {"compute": {"name": "compute_neighbour_total"}},
                    "key": ["TIME_PERIOD"],
                    "structure": {
                        "measure": _measure(),
                        "dimensions": [
                            _dimension(
                                "TIME_PERIOD",
                                {"kind": "column_header", "header_row": 8, "read": "int"},
                            )
                        ],
                    },
                },
                {
                    "id": "service",
                    "sheet": "Model",
                    "data_range": "Model!I12:M13",
                    "layout": "matrix",
                    "output": {"compute": {"name": "compute_service"}},
                    "key": ["INSTRUMENT", "TIME_PERIOD"],
                    "structure": {
                        "measure": _measure(),
                        "dimensions": [
                            _dimension("INSTRUMENT", instrument),
                            _dimension(
                                "TIME_PERIOD",
                                {"kind": "column_header", "header_row": 10, "read": "int"},
                            ),
                        ],
                    },
                },
                {
                    "id": "service_sum",
                    "sheet": "Model",
                    "data_range": "Model!B17",
                    "layout": "scalar",
                    "output": {"compute": {"name": "compute_service_sum"}},
                    "key": [],
                    "structure": {"measure": _measure(), "dimensions": []},
                },
                {
                    "id": "static_out",
                    "sheet": "Model",
                    "data_range": "Model!B3:C3",
                    "layout": "series",
                    "output": {"compute": {"name": "compute_static_out"}},
                    "key": ["TIME_PERIOD"],
                    "structure": {
                        "measure": _measure(),
                        "dimensions": [
                            _dimension(
                                "TIME_PERIOD",
                                {"kind": "column_header", "header_row": 1, "read": "int"},
                            )
                        ],
                    },
                },
            ],
        }
    )

    template = tmp_path / "template.xlsx"
    write_case(template, FIRST)
    all_targets = [
        t
        for t in all_series_targets(doc, workbook=template)
        if t not in {"Model!I13", "Model!L13", "Model!M13"}
    ]
    graph = create_dependency_graph(template, all_targets, load_values=True)
    with pytest.warns(UserWarning, match="Skipped 3 cell"):
        modules = generate_inverted_tree_modules(
            graph, series_bindings=doc, bindings_workbook=template
        )
    pkg = load_package(modules, tmp_path, name="lic_dsf_topology")
    _assert_generated_package_typechecks(tmp_path / "lic_dsf_topology")
    # `service` chooses `debt` (overlap 4 > macro overlap 1), so `macro` is not in ServiceInputs.
    assert set(pkg.ServiceInputs.describe()) == {"debt"}

    for delta in (0, 1, -2):
        case = tmp_path / f"shifted_{delta}.xlsx"
        first = FIRST + delta
        write_case(case, first)
        case_graph = create_dependency_graph(case, all_targets, load_values=True)
        eval_cells = [
            "Model!B3",
            "Model!C3",
            "Model!I12",
            "Model!J12",
            "Model!K12",
            "Model!L12",
            "Model!M12",
            "Model!J13",
            "Model!K13",
            "Model!J15",
            "Model!K15",
            "Model!L15",
            "Model!M15",
            "Model!B17",
        ]
        expected = FormulaEvaluator(case_graph).evaluate(eval_cells)
        model = pkg.Model.from_workbook(case)
        assert list(model.static_out.items()) == [
            ((2015,), expected["Model!B3"]),
            ((2016,), expected["Model!C3"]),
        ]
        assert list(model.service.items()) == [
            ((IMF, first - 1), expected["Model!I12"]),
            ((IMF, first), expected["Model!J12"]),
            ((IMF, first + 1), expected["Model!K12"]),
            ((IMF, first + 2), expected["Model!L12"]),
            ((IMF, first + 3), expected["Model!M12"]),
            ((TOTAL, first), expected["Model!J13"]),
            ((TOTAL, first + 1), expected["Model!K13"]),
        ]
        assert list(model.neighbour_total.items()) == [
            ((first,), expected["Model!J15"]),
            ((first + 1,), expected["Model!K15"]),
            ((first + 2,), expected["Model!L15"]),
            ((first + 3,), expected["Model!M15"]),
        ]
        assert model.service_sum == pytest.approx(expected["Model!B17"])
