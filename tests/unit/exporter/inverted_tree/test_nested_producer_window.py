"""Nested recurrences index external producers in their own full domain."""

from pathlib import Path

import pytest

from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    write_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a20_matrix_join import _matrix_entry


@pytest.mark.parametrize("pinned", [False, True])
def test_nested_stock_does_not_compact_amortization_to_consumed_years(
    tmp_path: Path, pinned: bool
) -> None:
    cells = {"A2": "a", "A3": "b", "A5": "a", "A6": "b"}
    for year, column in enumerate("BCDE", start=2020):
        cells[f"{column}1"] = year
        for row in (5, 6):
            cells[f"{column}{row}"] = f"={(row - 4) * (year - 2019)}"
        for row in (2, 3):
            previous = chr(ord(column) - 1)
            cells[f"{column}{row}"] = (
                100.0
                if column == "B"
                else f"={previous}{row}-{'$B' if pinned else column}{row + 3}"
            )
    workbook = write_workbook(tmp_path / "nested.xlsx", {"M": cells})
    document = bindings_document(
        _matrix_entry("amortization", "M!B5:E6", header_row=1, direction="internal"),
        _matrix_entry("opening", "M!B2:B3", header_row=1),
        _matrix_entry("stock", "M!C2:E3", header_row=1, direction="output"),
    )
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="nested_window")
    result = pkg.compute_stock()
    assert dict(result.items()) == {
        (area, year): 100.0 - ((year - 2020) if pinned else sum(range(2, year - 2018))) * payment
        for area, payment in (("a", 1), ("b", 2))
        for year in range(2021, 2024)
    }


@pytest.mark.parametrize("force_rung", [None, 3])
def test_fused_resolver_keeps_external_producer_coordinate_origin(
    tmp_path: Path, force_rung
) -> None:
    from tests.unit.exporter.inverted_tree.helpers import (
        inverted_graph_parts,
        named_input_kwargs,
        series_entry,
    )

    workbook = write_workbook(
        tmp_path / "origin.xlsx",
        {
            "M": {
                "A1": 2020,
                "B1": 2021,
                "C1": 2022,
                "D1": 2023,
                "A2": 100.0,
                "B2": 10.0,
                "C2": 20.0,
                "D2": 30.0,
                "B3": "=A4+B2",
                "C3": "=B4+C2",
                "D3": "=C4+D2",
                "A4": 1.0,
                "B4": "=B3*2",
                "C4": "=C3*2",
                "D4": "=D3*2",
            }
        },
    )
    document = bindings_document(
        series_entry("source", "M!A2:D2", layout="series", direction="input", header_row=1),
        series_entry("seed", "M!A4", direction="constant"),
        series_entry("a", "M!B3:D3", layout="series", direction="output", header_row=1),
        series_entry("b", "M!B4:D4", layout="series", direction="internal", header_row=1),
    )
    catalog, _, graph = inverted_graph_parts(workbook, document)
    pkg = load_package(
        generate_inverted(workbook, document, force_rung=force_rung), tmp_path, name="fused_origin"
    )
    result = pkg.compute_a(**named_input_kwargs(pkg, catalog, graph))
    assert dict(result.items()) == {(2021,): 11.0, (2022,): 42.0, (2023,): 114.0}


@pytest.mark.parametrize("force_rung", [None, 3])
def test_nested_scan_reads_area_only_parameter_without_advancing_year(
    tmp_path: Path, force_rung
) -> None:
    from tests.unit.exporter.inverted_tree.helpers import series_entry

    cells = {"A2": "a", "A3": "b", "A5": "a", "A6": "b", "B5": "=1", "B6": "=2"}
    for year, column in enumerate("BCDEFG", start=2020):
        cells[f"{column}1"] = year
        for row in (2, 3):
            operator = "-" if row == 2 else "+"
            cells[f"{column}{row}"] = (
                10.0 if column == "B" else f"={chr(ord(column) - 1)}{row}{operator}B{row + 3}"
            )
    workbook = write_workbook(tmp_path / "area_only.xlsx", {"M": cells})
    document = bindings_document(
        series_entry(
            "payment",
            "M!B5:B6",
            layout="series",
            direction="internal",
            label_column="A",
            key_concept="REF_AREA",
            key_read="string",
        ),
        _matrix_entry("opening", "M!B2:B3", header_row=1),
        _matrix_entry("stock", "M!C2:G3", header_row=1, direction="output"),
    )
    pkg = load_package(
        generate_inverted(workbook, document, force_rung=force_rung), tmp_path, name="area_only"
    )
    assert dict(pkg.compute_stock().items()) == {
        (area, year): 10.0 - (year - 2020) * payment
        for area, payment in (("a", 1), ("b", -2))
        for year in range(2021, 2026)
    }


@pytest.mark.parametrize("force_rung", [None, 3])
def test_lookup_table_demands_in_scc_producer_before_materialization(
    tmp_path: Path, force_rung
) -> None:
    from tests.unit.exporter.inverted_tree.helpers import series_entry

    workbook = write_workbook(
        tmp_path / "lookup_scc.xlsx",
        {
            "M": {
                "A1": 2020,
                "B1": 2021,
                "C1": 2022,
                "A2": "=B3+1",
                "B2": "=C3+1",
                "C2": "=1",
                "A3": "=HLOOKUP(A1,$A$1:$C$2,2,FALSE)",
                "B3": "=HLOOKUP(B1,$A$1:$C$2,2,FALSE)",
                "C3": "=0",
            }
        },
    )
    document = bindings_document(
        series_entry("years", "M!A1:C1", layout="series", direction="constant", header_row=1),
        series_entry("values", "M!A2:C2", layout="series", direction="internal", header_row=1),
        series_entry("result", "M!A3:C3", layout="series", direction="output", header_row=1),
    )
    pkg = load_package(
        generate_inverted(workbook, document, force_rung=force_rung), tmp_path, name="lookup_scc"
    )
    assert dict(pkg.compute_result().items()) == {(2020,): 2.0, (2021,): 1.0, (2022,): 0.0}
    named = pkg.internals.scan_values_result(years=pkg.data.YEARS)
    assert dict(named.result.items()) == {(2020,): 2.0, (2021,): 1.0, (2022,): 0.0}


@pytest.mark.parametrize("force_rung", [None, 2, 3])
def test_nested_keyed_reads_use_host_catalog_origin(tmp_path: Path, force_rung) -> None:
    from tests.unit.exporter.inverted_tree.helpers import series_entry

    cells = {"A2": "a", "A3": "b"}
    for year, column in enumerate("BCDEFG", start=2020):
        cells[f"{column}1"] = year
        cells[f"{column}5"] = year - 2019
        for row in (2, 3):
            cells[f"{column}{row}"] = (
                10 if column == "B" else f"={chr(ord(column) - 1)}{row}+{column}$5+$B$5"
            )
    workbook = write_workbook(tmp_path / "keyed_origin.xlsx", {"M": cells})
    document = bindings_document(
        series_entry("source", "M!B5:G5", layout="series", direction="constant", header_row=1),
        _matrix_entry("opening", "M!B2:B3", header_row=1),
        _matrix_entry("result", "M!C2:G3", header_row=1, direction="output"),
    )
    pkg = load_package(
        generate_inverted(workbook, document, force_rung=force_rung), tmp_path, name="keyed_origin"
    )
    assert dict(pkg.compute_result().items()) == {
        (area, year): 10 + sum(range(3, year - 2017))
        for area in ("a", "b")
        for year in range(2021, 2026)
    }


@pytest.mark.parametrize("force_rung", [None, 3])
def test_choose_does_not_request_unselected_recurrence(tmp_path: Path, force_rung) -> None:
    from tests.unit.exporter.inverted_tree.helpers import series_entry

    workbook = write_workbook(
        tmp_path / "choose_lazy.xlsx",
        {
            "M": {
                "A1": 2020,
                "B1": 2021,
                "C1": 2022,
                "A2": "=CHOOSE(1,2,B3)",
                "B2": "=CHOOSE(1,3,C3)",
                "C2": "=4",
                "A3": "=A2",
                "B3": "=A3+B2",
                "C3": "=B3+C2",
            }
        },
    )
    document = bindings_document(
        series_entry("chosen", "M!A2:C2", layout="series", direction="internal", header_row=1),
        series_entry("result", "M!A3:C3", layout="series", direction="output", header_row=1),
    )
    pkg = load_package(
        generate_inverted(workbook, document, force_rung=force_rung), tmp_path, name="choose_lazy"
    )
    assert dict(pkg.compute_result().items()) == {(2020,): 2.0, (2021,): 5.0, (2022,): 9.0}


@pytest.mark.parametrize("force_rung", [None, 2, 3])
def test_each_partition_uses_its_own_formula_regions(tmp_path: Path, force_rung) -> None:
    cells = {
        "A2": "a",
        "A3": "b",
        "A5": "a",
        "A6": "b",
        "B5": 10,
        "C5": 11,
        "D5": 12,
        "E5": 13,
        "B6": 20,
    }
    for year, column in enumerate("BCDE", start=2020):
        cells[f"{column}1"] = year
        cells[f"{column}2"] = f"={column}5-1"
        cells[f"{column}3"] = (
            "=B6-1" if column == "B" else "=7" if column == "C" else f"={chr(ord(column) - 1)}3"
        )
    workbook = write_workbook(tmp_path / "regions.xlsx", {"M": cells})
    source = _matrix_entry("source", "M!B5:E6", header_row=1)
    source["data_range"] = ["M!B5:E5", "M!B6"]
    document = bindings_document(
        source, _matrix_entry("result", "M!B2:E3", header_row=1, direction="output")
    )
    pkg = load_package(
        generate_inverted(workbook, document, force_rung=force_rung), tmp_path, name="regions"
    )
    assert dict(pkg.compute_result().items()) == {
        ("a", 2020): 9.0,
        ("a", 2021): 10.0,
        ("a", 2022): 11.0,
        ("a", 2023): 12.0,
        ("b", 2020): 19.0,
        ("b", 2021): 7.0,
        ("b", 2022): 7.0,
        ("b", 2023): 7.0,
    }
