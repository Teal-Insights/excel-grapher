"""Projected calls must retain the coordinate origin of formula regions."""

from pathlib import Path

from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    named_input_kwargs,
    series_entry,
    write_workbook,
)


def test_projected_mixed_region_helper_keeps_full_argument_window(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "window.xlsx",
        {
            "M": {
                "A1": 2020,
                "B1": 2021,
                "C1": 2022,
                "D1": 2023,
                "E1": 2024,
                "F1": 2025,
                "B4": 10.0,
                "C4": 20.0,
                "D4": 30.0,
                "E4": 40.0,
                "F4": 50.0,
                "A5": 1.0,
                "B5": 2.0,
                "C5": 3.0,
                "D5": 4.0,
                "E5": 5.0,
                "B2": "=B4+A5",
                "C2": "=C4+B5",
                "D2": "=D4+C5*2",
                "E2": "=E4+D5*2",
                "F2": "=F4+E5*2",
                "D3": "=D2",
                "E3": "=E2",
                "F3": "=F2",
            }
        },
    )
    document = bindings_document(
        series_entry("source", "M!B4:F4", layout="series", direction="input", header_row=1),
        series_entry("shifted", "M!A5:E5", layout="series", direction="input", header_row=1),
        series_entry("middle", "M!B2:F2", layout="series", direction="internal", header_row=1),
        series_entry("result", "M!D3:F3", layout="series", direction="output", header_row=1),
    )
    catalog, _, graph = inverted_graph_parts(workbook, document)
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="partial_window")
    result = pkg.compute_result(**named_input_kwargs(pkg, catalog, graph))
    assert dict(result.items()) == {(2023,): 36.0, (2024,): 48.0, (2025,): 60.0}
