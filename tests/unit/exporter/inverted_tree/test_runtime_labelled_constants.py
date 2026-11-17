"""Constants on runtime-labelled axes align snapshot keys onto active labellers."""

from __future__ import annotations

import subprocess
from pathlib import Path
from typing import Any

import pytest
from fastpyxl import Workbook

from excel_grapher import create_dependency_graph
from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter.export_runtime.provenance import column_letter
from excel_grapher.exporter.inverted_tree import generate_inverted_tree_modules
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.series_bindings import validate_bindings_document
from excel_grapher.series_bindings.workflow import all_series_targets, output_series_targets
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    series_entry,
    write_workbook,
)
from tests.unit.exporter.inverted_tree.test_formula_header_string_remap import _plant_strings
from tests.unit.exporter.inverted_tree.test_model_from_workbook import _plant_formula_cache
from tests.unit.exporter.inverted_tree.test_sparse_domain_emit import _ragged_labelled_bindings

_REPO_ROOT = Path(__file__).resolve().parents[4]

FIRST_YEAR = 2011
LAST_YEAR = 2044
SNAPSHOT = (2018, 2019, 2020, 2021, 2022)
HEADER_ROW = 7
INPUT_ROW = 8
CONSTANT_ROW = 9
OUTPUT_ROW = 10


def _year_column(year: int) -> str:
    return column_letter(year - FIRST_YEAR + 1)


def _plant_int_caches(path: Path, caches: dict[str, int]) -> None:
    for coordinate, cached in caches.items():
        _plant_formula_cache(path, coordinate, str(cached))


def _write_year_shift_workbook(
    path: Path,
    first_year: int,
    *,
    input_offset: float = 0.0,
) -> None:
    book = Workbook()
    sheet = book.active
    assert sheet is not None
    sheet.title = "Input"
    caches: dict[str, int] = {}
    previous = ""
    for offset in range(LAST_YEAR - FIRST_YEAR + 1):
        column = column_letter(offset + 1)
        header = f"{column}{HEADER_ROW}"
        shifted = first_year + offset
        sheet[header] = f"={shifted}" if offset == 0 else f"={previous}+1"
        caches[header] = shifted
        sheet[f"{column}{INPUT_ROW}"] = input_offset
        previous = header
    for index, year in enumerate(SNAPSHOT, start=1):
        column = _year_column(year)
        sheet[f"{column}{CONSTANT_ROW}"] = float(index)
        sheet[f"{column}{OUTPUT_ROW}"] = f"={column}{CONSTANT_ROW}+{column}{INPUT_ROW}"
    first_col = _year_column(SNAPSHOT[0])
    last_col = _year_column(SNAPSHOT[-1])
    sheet["A12"] = f"=SUM({first_col}{CONSTANT_ROW}:{last_col}{CONSTANT_ROW})"
    sheet["B12"] = f"=${first_col}${CONSTANT_ROW}"
    sheet["C12"] = f"=INDEX({first_col}{CONSTANT_ROW}:{last_col}{CONSTANT_ROW},3)"
    book.save(path)
    book.close()
    _plant_int_caches(path, caches)


def _dimension(axis_id: str, *, header_row: int = HEADER_ROW, read: str = "int") -> dict[str, Any]:
    return {
        "id": axis_id,
        "concept": axis_id,
        "role": "key",
        "scope": "cell",
        "bind": {"kind": "column_header", "header_row": header_row, "read": read},
    }


def _measure() -> dict[str, Any]:
    return {
        "concept": "OBS_VALUE",
        "dtype": "float",
        "bind": {"kind": "data_cell", "read": "float"},
    }


def _year_shift_bindings(
    *,
    account_id: str = "current_account",
    deflator_id: str = "us_deflator",
    include_extras: bool = False,
) -> dict[str, Any]:
    span = f"{_year_column(FIRST_YEAR)}{INPUT_ROW}:{_year_column(LAST_YEAR)}{INPUT_ROW}"
    window = f"{_year_column(SNAPSHOT[0])}{CONSTANT_ROW}:{_year_column(SNAPSHOT[-1])}{CONSTANT_ROW}"
    output = f"{_year_column(SNAPSHOT[0])}{OUTPUT_ROW}:{_year_column(SNAPSHOT[-1])}{OUTPUT_ROW}"
    series: list[dict[str, Any]] = [
        {
            "id": account_id,
            "sheet": "Input",
            "data_range": f"Input!{span}",
            "layout": "series",
            "input": {},
            "key": ["TIME_PERIOD"],
            "structure": {"measure": _measure(), "dimensions": [_dimension("TIME_PERIOD")]},
        },
        {
            "id": deflator_id,
            "sheet": "Input",
            "data_range": f"Input!{window}",
            "layout": "series",
            "constant": {},
            "key": ["TIME_PERIOD"],
            "structure": {"measure": _measure(), "dimensions": [_dimension("TIME_PERIOD")]},
        },
        {
            "id": "deflator_path",
            "sheet": "Input",
            "data_range": f"Input!{output}",
            "layout": "series",
            "output": {"compute": {"name": "compute_deflator_path"}},
            "key": ["TIME_PERIOD"],
            "structure": {"measure": _measure(), "dimensions": [_dimension("TIME_PERIOD")]},
        },
    ]
    if include_extras:
        series.extend(
            [
                {
                    "id": "window_sum",
                    "sheet": "Input",
                    "data_range": "Input!A12",
                    "layout": "scalar",
                    "output": {"compute": {"name": "compute_window_sum"}},
                    "key": [],
                    "structure": {"measure": _measure(), "dimensions": []},
                },
                {
                    "id": "pinned_first",
                    "sheet": "Input",
                    "data_range": "Input!B12",
                    "layout": "scalar",
                    "output": {"compute": {"name": "compute_pinned_first"}},
                    "key": [],
                    "structure": {"measure": _measure(), "dimensions": []},
                },
                {
                    "id": "indexed_third",
                    "sheet": "Input",
                    "data_range": "Input!C12",
                    "layout": "scalar",
                    "output": {"compute": {"name": "compute_indexed_third"}},
                    "key": [],
                    "structure": {"measure": _measure(), "dimensions": []},
                },
            ]
        )
    return {
        "schema_version": "1.22.0",
        "concept_scheme": {
            "id": "year_shift",
            "concepts": [
                {"id": "TIME_PERIOD", "dtype": "int"},
                {"id": "OBS_VALUE", "dtype": "number"},
            ],
        },
        "series": series,
    }


def test_formula_header_constant_window_survives_axis_shifts(tmp_path: Path) -> None:
    """Issue 1083 MCVE plus backward and disjoint shifts match FormulaEvaluator."""
    template = tmp_path / "template.xlsx"
    _write_year_shift_workbook(template, FIRST_YEAR)
    document = validate_bindings_document(_year_shift_bindings(include_extras=True))
    graph = create_dependency_graph(
        template, all_series_targets(document, workbook=template), load_values=True
    )
    modules = generate_inverted_tree_modules(
        graph, series_bindings=document, bindings_workbook=template
    )
    pkg = load_package(modules, tmp_path, name="yearshift")

    axis = next(
        item for item in pkg.data.US_DEFLATOR.schema.domain.axes if item.name == "TIME_PERIOD"
    )
    assert axis.snapshot == SNAPSHOT
    assert axis.source == tuple(range(FIRST_YEAR, LAST_YEAR + 1))
    assert axis.labeller == "current_account"
    assert axis.size == len(SNAPSHOT)
    assert pkg.data.US_DEFLATOR.domain.axis("TIME_PERIOD").keys == SNAPSHOT

    template_model = pkg.Model.from_workbook(template)
    assert template_model.deflator_path.domain.axis("TIME_PERIOD").keys == SNAPSHOT
    assert tuple(v for _c, v in template_model.deflator_path.items()) == (1.0, 2.0, 3.0, 4.0, 5.0)
    assert template_model.window_sum == pytest.approx(15.0)
    assert template_model.pinned_first == pytest.approx(1.0)
    assert template_model.indexed_third == pytest.approx(3.0)

    for delta, input_offset in ((1, 0.0), (-2, 10.0), (19, 0.5)):
        case = tmp_path / f"shifted_{delta}.xlsx"
        _write_year_shift_workbook(case, FIRST_YEAR + delta, input_offset=input_offset)
        case_graph = create_dependency_graph(
            case, all_series_targets(document, workbook=case), load_values=True
        )
        deflator_cells = [f"Input!{_year_column(year)}{OUTPUT_ROW}" for year in SNAPSHOT]
        expected = FormulaEvaluator(case_graph).evaluate(
            [*deflator_cells, "Input!A12", "Input!B12", "Input!C12"]
        )

        model = pkg.Model.from_workbook(case)
        assert "_us_deflator" not in vars(model)
        expected_years = tuple(year + delta for year in SNAPSHOT)
        assert model.deflator_path.domain.axis("TIME_PERIOD").keys == expected_years
        assert "_us_deflator" in vars(model)
        for (_coord, value), cell in zip(model.deflator_path.items(), deflator_cells, strict=True):
            assert value == pytest.approx(expected[cell])
        assert model.window_sum == pytest.approx(expected["Input!A12"])
        assert model.pinned_first == pytest.approx(expected["Input!B12"])
        assert model.indexed_third == pytest.approx(expected["Input!C12"])

        bundle = pkg.DeflatorPathInputs.from_workbook(case)
        via_compute = pkg.compute_deflator_path(bundle)
        assert list(via_compute.items()) == list(model.deflator_path.items())

        via_init = pkg.Model(current_account=model.current_account).deflator_path
        assert list(via_init.items()) == list(model.deflator_path.items())

        cells = dict(model.cells("us_deflator"))
        assert tuple(cells) == tuple((year,) for year in expected_years)
        assert cells[(expected_years[0],)] == f"Input!{_year_column(SNAPSHOT[0])}{CONSTANT_ROW}"


def test_constant_overrides_align_on_shifted_axis_and_fail_closed(tmp_path: Path) -> None:
    """`data.overrides` accepts snapshot Series, shifted Series, and plain Tensors."""
    template = tmp_path / "template.xlsx"
    shifted = tmp_path / "shifted.xlsx"
    _write_year_shift_workbook(template, FIRST_YEAR)
    _write_year_shift_workbook(shifted, FIRST_YEAR + 1)
    pkg = load_package(
        generate_inverted(template, _year_shift_bindings()),
        tmp_path,
        name="yearshift_overrides",
    )

    doubled = pkg.data.US_DEFLATOR.with_values((2.0, 4.0, 6.0, 8.0, 10.0))
    with pkg.data.overrides(US_DEFLATOR=doubled):
        got = pkg.Model.from_workbook(shifted).deflator_path
        assert got.domain.axis("TIME_PERIOD").keys == (2019, 2020, 2021, 2022, 2023)
        assert tuple(v for _c, v in got.items()) == (2.0, 4.0, 6.0, 8.0, 10.0)

    already_shifted = pkg.Series(
        pkg.Domain.product(pkg.Axis("TIME_PERIOD", (2019, 2020, 2021, 2022, 2023), int)),
        (10.0, 20.0, 30.0, 40.0, 50.0),
        schema=pkg.data.US_DEFLATOR.schema,
        cells=pkg.data.US_DEFLATOR.cells,
    )
    with pkg.data.overrides(US_DEFLATOR=already_shifted):
        got = pkg.Model.from_workbook(shifted).deflator_path
        assert tuple(v for _c, v in got.items()) == (10.0, 20.0, 30.0, 40.0, 50.0)

    plain_snapshot = pkg.Tensor.from_nested(
        domain=pkg.Domain.product(pkg.Axis("TIME_PERIOD", SNAPSHOT, int)),
        values=(5.0, 4.0, 3.0, 2.0, 1.0),
    )
    with pkg.data.overrides(US_DEFLATOR=plain_snapshot):
        got = pkg.Model.from_workbook(shifted).deflator_path
        assert tuple(v for _c, v in got.items()) == (5.0, 4.0, 3.0, 2.0, 1.0)

    concrete_template_series = pkg.Model.from_workbook(template).deflator_path
    with pkg.data.overrides(US_DEFLATOR=concrete_template_series):
        got = pkg.Model.from_workbook(shifted).deflator_path
        assert tuple(v for _c, v in got.items()) == (1.0, 2.0, 3.0, 4.0, 5.0)

    cellless_same_schema = pkg.Series(
        pkg.Domain.product(pkg.Axis("TIME_PERIOD", SNAPSHOT, int)),
        (9.0, 8.0, 7.0, 6.0, 5.0),
        schema=pkg.data.US_DEFLATOR.schema,
        cells=None,
    )
    with pkg.data.overrides(US_DEFLATOR=cellless_same_schema):
        shifted_model = pkg.Model.from_workbook(shifted)
        assert tuple(v for _c, v in shifted_model.deflator_path.items()) == (
            9.0,
            8.0,
            7.0,
            6.0,
            5.0,
        )
        assert dict(shifted_model.cells("us_deflator")) == {
            (2019,): f"Input!{_year_column(2018)}{CONSTANT_ROW}",
            (2020,): f"Input!{_year_column(2019)}{CONSTANT_ROW}",
            (2021,): f"Input!{_year_column(2020)}{CONSTANT_ROW}",
            (2022,): f"Input!{_year_column(2021)}{CONSTANT_ROW}",
            (2023,): f"Input!{_year_column(2022)}{CONSTANT_ROW}",
        }

    foreign = pkg.Series(
        pkg.Domain.product(pkg.Axis("TIME_PERIOD", (2090, 2091, 2092, 2093, 2094), int)),
        (1.0, 2.0, 3.0, 4.0, 5.0),
        schema=pkg.data.US_DEFLATOR.schema,
        cells=None,
    )
    with (
        pkg.data.overrides(US_DEFLATOR=foreign),
        pytest.raises(pkg.tensor.SchemaError, match="2090"),
    ):
        _ = pkg.Model.from_workbook(shifted).deflator_path


def test_authored_axis_labels_constant_and_ragged_matrix_shift(tmp_path: Path) -> None:
    """Shifted `axis_labels` keep ragged stock positions and per-year outputs."""
    workbook = write_workbook(
        tmp_path / "ragged_labelled_pos.xlsx",
        {
            "Origin": {"A1": 2017},
            "Data": {
                "B1": "=Origin!A1",
                "C1": "=B1+1",
                "D1": "=C1+1",
                "A4": "outstanding",
                "B4": 10.0,
                "C4": 20.0,
                "D4": 30.0,
                "A5": "short",
                "B5": 1.0,
                "C5": 2.0,
                "E4": "=1*B4+2*C4+3*D4+4*B5+5*C5",
                "B6": "=B4+B5",
                "C6": "=C4+C5",
                "D6": "=D4",
            },
        },
    )
    document = _ragged_labelled_bindings()
    document["series"].append(
        series_entry(
            "annual_stock",
            "Data!B6:D6",
            layout="series",
            direction="output",
            header_row=1,
        )
    )
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="ragged_const_shift")
    catalog, _deps, graph = inverted_graph_parts(workbook, document)
    graph.set_node_value("Origin!A1", 2025)
    annual_cells = list(catalog.get("annual_stock").cells)
    expected = FormulaEvaluator(graph).evaluate(["Data!E4", *annual_cells])

    shifted_model = pkg.Model(first_year=2025)
    assert shifted_model.total == pytest.approx(expected["Data!E4"])
    assert shifted_model.total == pytest.approx(154.0)
    assert pkg.Model.from_defaults(first_year=2025).total == pytest.approx(154.0)
    assert pkg.compute_total(pkg.TotalInputs(first_year=2025)) == pytest.approx(154.0)

    stock = shifted_model._stock
    assert stock.domain.axis("TIME_PERIOD").keys == (2025, 2026, 2027)
    assert ("short", 2027) not in set(stock.domain)
    assert stock[("short", 2027)] is None
    assert list(stock.items()) == [
        (("outstanding", 2025), 10.0),
        (("outstanding", 2026), 20.0),
        (("outstanding", 2027), 30.0),
        (("short", 2025), 1.0),
        (("short", 2026), 2.0),
    ]

    assert annual_cells == ["Data!B6", "Data!C6", "Data!D6"]
    assert [expected[cell] for cell in annual_cells] == pytest.approx([11.0, 22.0, 30.0])
    assert shifted_model.annual_stock.domain.axis("TIME_PERIOD").keys == (2025, 2026, 2027)
    assert list(shifted_model.annual_stock.items()) == [
        (coord, pytest.approx(expected[cell]))
        for coord, cell in zip(shifted_model.annual_stock.domain, annual_cells, strict=True)
    ]


def test_permuted_full_length_input_reads_unswapped_snapshot_cells_from_workbook(
    tmp_path: Path,
) -> None:
    """`from_workbook` materializes snapshot cells without permuting `snapshot` into `source`."""

    def write_case(path: Path, *, rows: tuple[str, str]) -> Path:
        write_workbook(
            path,
            {
                "Inputs": {
                    "A2": f'="{rows[0]}"',
                    "A3": f'="{rows[1]}"',
                    "B2": 1.0,
                    "B3": 2.0,
                    # Literal row labels in reverse order ("Beta", "Alpha") on template
                    "D2": "Beta",
                    "D3": "Alpha",
                    "E2": 100.0,
                    "E3": 10.0,
                    "C2": "=B2+E3",
                    "C3": "=B3+E2",
                }
            },
        )
        _plant_strings(path, {"A2": rows[0], "A3": rows[1]})
        return path

    template = write_case(tmp_path / "perm_in_template.xlsx", rows=("Alpha", "Beta"))
    document = bindings_document(
        series_entry(
            "flags",
            "Inputs!B2:B3",
            layout="series",
            direction="input",
            label_column="A",
            key_concept="COUNTRY",
            key_read="string",
        ),
        series_entry(
            "rates",
            "Inputs!E2:E3",
            layout="series",
            direction="input",
            label_column="D",
            key_concept="COUNTRY",
            key_read="string",
        ),
        series_entry(
            "out",
            "Inputs!C2:C3",
            layout="series",
            direction="output",
            label_column="A",
            key_concept="COUNTRY",
            key_read="string",
        ),
    )
    pkg = load_package(generate_inverted(template, document), tmp_path, name="perm_input_wb")
    case = write_case(tmp_path / "perm_in_case.xlsx", rows=("Public", "Private"))
    case_catalog, _deps, case_graph = inverted_graph_parts(case, document)
    expected = FormulaEvaluator(case_graph).evaluate(list(case_catalog.get("out").cells))

    model = pkg.Model.from_workbook(case)
    assert model.rates.domain.axis("COUNTRY").keys == ("Private", "Public")
    assert dict(model.rates.items()) == {("Private",): 100.0, ("Public",): 10.0}
    assert dict(model.out.items()) == {
        ("Public",): pytest.approx(expected["Inputs!C2"]),
        ("Private",): pytest.approx(expected["Inputs!C3"]),
    }


def test_recurrence_group_and_matrix_constants_on_formula_axes(tmp_path: Path) -> None:
    """A 2-D constant on two formula-header axes aligns inside a recurrence scan."""

    def write_case(path: Path, *, rows: tuple[str, str], first_year: int) -> Path:
        write_workbook(
            path,
            {
                "Inputs": {
                    "B1": f"={first_year}",
                    "C1": "=B1+1",
                    "D1": "=C1+1",
                    "A2": f'="{rows[0]}"',
                    "A3": f'="{rows[1]}"',
                    "B2": 1.0,
                    "C2": 2.0,
                    "D2": 3.0,
                    "B3": 4.0,
                    "C3": 5.0,
                    "D3": 6.0,
                    # Constant window on the last two years (C..D), with permuted row labels
                    "F2": f'="{rows[1]}"',
                    "F3": f'="{rows[0]}"',
                    "G2": 100.0,
                    "H2": 200.0,
                    "G3": 10.0,
                    "H3": 20.0,
                    # Mutually recursive scan over C..D reading the 2-D constant
                    "C5": "=G3+C2",
                    "D5": "=C6+H3+D2",
                    "C6": "=C5+G2+C3",
                    "D6": "=D5+H2+D3",
                }
            },
        )
        _plant_int_caches(
            path,
            {"B1": first_year, "C1": first_year + 1, "D1": first_year + 2},
        )
        _plant_strings(
            path,
            {"A2": rows[0], "A3": rows[1], "F2": rows[1], "F3": rows[0]},
        )
        return path

    template = write_case(tmp_path / "scan_template.xlsx", rows=("Alpha", "Beta"), first_year=2020)
    document = {
        "schema_version": "1.22.0",
        "concept_scheme": {
            "id": "matrix_scan",
            "concepts": [
                {"id": "SECTOR", "dtype": "string"},
                {"id": "TIME_PERIOD", "dtype": "int"},
                {"id": "OBS_VALUE", "dtype": "number"},
            ],
        },
        "series": [
            {
                "id": "flows",
                "sheet": "Inputs",
                "data_range": "Inputs!B2:D3",
                "layout": "matrix",
                "input": {},
                "key": ["SECTOR", "TIME_PERIOD"],
                "structure": {
                    "measure": _measure(),
                    "dimensions": [
                        {
                            "id": "SECTOR",
                            "concept": "SECTOR",
                            "role": "key",
                            "scope": "cell",
                            "bind": {"kind": "row_label", "label_column": "A", "read": "string"},
                        },
                        _dimension("TIME_PERIOD", header_row=1),
                    ],
                },
            },
            {
                "id": "stocks",
                "sheet": "Inputs",
                "data_range": "Inputs!G2:H3",
                "layout": "matrix",
                "constant": {},
                "key": ["SECTOR", "TIME_PERIOD"],
                "structure": {
                    "measure": _measure(),
                    "dimensions": [
                        {
                            "id": "SECTOR",
                            "concept": "SECTOR",
                            "role": "key",
                            "scope": "cell",
                            "bind": {"kind": "row_label", "label_column": "F", "read": "string"},
                        },
                        {
                            "id": "TIME_PERIOD",
                            "concept": "TIME_PERIOD",
                            "role": "key",
                            "scope": "cell",
                            "bind": {
                                "kind": "value_map",
                                "values": {2021: "G", 2022: "H"},
                                "read": "int",
                            },
                        },
                    ],
                },
            },
            {
                "id": "path_alpha",
                "sheet": "Inputs",
                "data_range": "Inputs!C5:D5",
                "layout": "series",
                "output": {"compute": {"name": "compute_path_alpha"}},
                "key": ["TIME_PERIOD"],
                "structure": {
                    "measure": _measure(),
                    "dimensions": [_dimension("TIME_PERIOD", header_row=1)],
                },
            },
            {
                "id": "path_beta",
                "sheet": "Inputs",
                "data_range": "Inputs!C6:D6",
                "layout": "series",
                "output": {"compute": {"name": "compute_path_beta"}},
                "key": ["TIME_PERIOD"],
                "structure": {
                    "measure": _measure(),
                    "dimensions": [_dimension("TIME_PERIOD", header_row=1)],
                },
            },
        ],
    }
    pkg = load_package(generate_inverted(template, document), tmp_path, name="matrix_scan")
    shifted = write_case(
        tmp_path / "scan_shifted.xlsx", rows=("Public", "Private"), first_year=2025
    )
    shifted_catalog, _deps, shifted_graph = inverted_graph_parts(shifted, document)
    expected = FormulaEvaluator(shifted_graph).evaluate(
        [*shifted_catalog.get("path_alpha").cells, *shifted_catalog.get("path_beta").cells]
    )

    model = pkg.Model.from_workbook(shifted)
    assert model.path_alpha.domain.axis("TIME_PERIOD").keys == (2026, 2027)
    assert model.path_beta.domain.axis("TIME_PERIOD").keys == (2026, 2027)
    for (_c, value), cell in zip(
        model.path_alpha.items(), shifted_catalog.get("path_alpha").cells, strict=True
    ):
        assert value == pytest.approx(expected[cell])
    for (_c, value), cell in zip(
        model.path_beta.items(), shifted_catalog.get("path_beta").cells, strict=True
    ):
        assert value == pytest.approx(expected[cell])


def test_output_only_extraction_retains_off_graph_labeller_for_constant(tmp_path: Path) -> None:
    """`output_series_targets` works when the output only touches the constant."""
    template = tmp_path / "template.xlsx"
    shifted = tmp_path / "shifted.xlsx"
    _write_year_shift_workbook(template, FIRST_YEAR)
    _write_year_shift_workbook(shifted, FIRST_YEAR + 1)
    raw = _year_shift_bindings(include_extras=True)
    raw["series"] = [
        entry
        for entry in raw["series"]
        if entry["id"] in {"current_account", "us_deflator", "window_sum"}
    ]
    document = validate_bindings_document(raw)
    graph = create_dependency_graph(
        template, output_series_targets(document, workbook=template), load_values=True
    )
    modules = generate_inverted_tree_modules(
        graph, series_bindings=document, bindings_workbook=template
    )
    pkg = load_package(modules, tmp_path, name="off_graph_labeller")
    assert pkg.Model.from_workbook(shifted).window_sum == pytest.approx(15.0)
    assert pkg.compute_window_sum(pkg.WindowSumInputs.from_workbook(shifted)) == pytest.approx(15.0)


def test_labeller_depending_on_constant_on_its_own_axis_fails_closed(tmp_path: Path) -> None:
    """An `axis_labels` series cannot depend on a constant keyed on that axis."""
    workbook = write_workbook(
        tmp_path / "cycle_const_labeller.xlsx",
        {
            "Inputs": {"A2": 2024, "B2": 2025, "C2": 2026, "A3": 2024, "B3": 2025, "C3": 2026},
            "Engine": {
                "A1": "=Inputs!A3",
                "B1": "=Inputs!B3",
                "C1": "=Inputs!C3",
                "A2": "=A1",
                "B2": "=B1",
                "C2": "=C1",
            },
        },
    )
    document = bindings_document(
        series_entry(
            "seed_years",
            "Inputs!A3:C3",
            layout="series",
            direction="constant",
            dtype="int",
            header_row=2,
        ),
        series_entry(
            "year_labels",
            "Engine!A1:C1",
            layout="series",
            direction="internal",
            dtype="int",
            header_row=1,
            axis_labels="TIME_PERIOD",
        ),
        series_entry(
            "path",
            "Engine!A2:C2",
            layout="series",
            direction="output",
            header_row=1,
        ),
        schema_version="1.17.0",
    )
    with pytest.raises(InvertedTreeExportError, match="keyed on labelled axis"):
        generate_inverted(workbook, document)


def test_generated_long_name_constant_alignment_passes_ruff_and_ty(tmp_path: Path) -> None:
    """Long LIC-DSF-style series ids wrap cleanly for ruff format, ruff check, and ty."""
    template = tmp_path / "template.xlsx"
    _write_year_shift_workbook(template, FIRST_YEAR)
    document = _year_shift_bindings(
        account_id="input3_input_3_macro_current_account",
        deflator_id="input3_private_external_debt_stock_historical_override",
    )
    modules = generate_inverted(template, document)
    package = tmp_path / "long_const_lint"
    package.mkdir()
    for filename, content in modules.items():
        (package / filename).write_text(content, encoding="utf-8")

    targets = [str(package / name) for name in ("api.py", "model.py", "validation.py")]
    check = subprocess.run(
        ["uv", "run", "--no-sync", "ruff", "check", *targets],
        cwd=str(_REPO_ROOT),
        capture_output=True,
        text=True,
        check=False,
    )
    assert check.returncode == 0, check.stdout + check.stderr

    fmt = subprocess.run(
        ["uv", "run", "--no-sync", "ruff", "format", "--check", *targets],
        cwd=str(_REPO_ROOT),
        capture_output=True,
        text=True,
        check=False,
    )
    assert fmt.returncode == 0, fmt.stdout + fmt.stderr

    undefined = subprocess.run(
        [
            "uv",
            "run",
            "--no-sync",
            "ruff",
            "check",
            "--select",
            "F821",
            str(package / "internals.py"),
            str(package / "data.py"),
        ],
        cwd=str(_REPO_ROOT),
        capture_output=True,
        text=True,
        check=False,
    )
    assert undefined.returncode == 0, undefined.stdout + undefined.stderr

    ty = subprocess.run(
        [
            "uv",
            "run",
            "--no-sync",
            "ty",
            "check",
            "--extra-search-path",
            str(tmp_path),
            "--project",
            str(_REPO_ROOT),
            "--ignore",
            "unresolved-attribute",
            str(package / "api.py"),
            str(package / "model.py"),
            str(package / "internals.py"),
            str(package / "data.py"),
        ],
        cwd=str(_REPO_ROOT),
        capture_output=True,
        text=True,
        check=False,
    )
    assert ty.returncode == 0, ty.stdout + ty.stderr
