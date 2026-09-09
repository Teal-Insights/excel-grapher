"""A scan seed refers to one coordinate of its producer."""

from pathlib import Path

import pytest

from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    named_input_kwargs,
    series_entry,
    write_workbook,
)


@pytest.mark.parametrize("recurrence", [False, True])
@pytest.mark.parametrize("first", [10.0, "#N/A"])
def test_sequence_seed_selects_value_not_producer(
    tmp_path: Path, recurrence: bool, first: float | str
) -> None:
    workbook = write_workbook(
        tmp_path / "seed.xlsx",
        {
            "M": {
                "A1": 2020,
                "B1": 2021,
                "C1": 2022,
                "D1": 2023,
                "A2": first,
                "B2": 20.0,
                "C2": 30.0,
                "B3": "=A2+1",
                "C3": "=B3+1" if recurrence else "=B2+1",
                "D3": "=C3+1" if recurrence else "=C2+1",
            }
        },
    )
    document = bindings_document(
        series_entry(
            "seeds",
            "M!A2:C2",
            layout="series",
            direction="input",
            header_row=1,
            key_concept="INSTRUMENT",
        ),
        series_entry("result", "M!B3:D3", layout="series", direction="output", header_row=1),
    )
    catalog, _, graph = inverted_graph_parts(workbook, document)
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="sequence_seed")
    result = pkg.compute_result(**named_input_kwargs(pkg, catalog, graph))
    assert dict(result.items()) == {
        (2021,): "#N/A" if first == "#N/A" else 11.0,
        (2022,): ("#N/A" if first == "#N/A" else 12.0) if recurrence else 21.0,
        (2023,): ("#N/A" if first == "#N/A" else 13.0) if recurrence else 31.0,
    }
