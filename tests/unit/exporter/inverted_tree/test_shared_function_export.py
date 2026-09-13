"""Shared worksheet functions used by real inverted-tree exports."""

from pathlib import Path
from typing import Literal

import pytest

from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)


@pytest.mark.parametrize("force_rung", [None, 3])
@pytest.mark.parametrize(
    "formula, expected",
    [
        ("=NPV(0.1,A1:A3)", 4.815927873779113),
        ("=RANK(A2,A1:A3)", 2.0),
        ("=LARGE(A1:A3,2)", 2.0),
        ("=STDEV(A1:A3)", 1.0),
        ('=COUNTIF(A1:A3,">1")', 2.0),
        ("=ROUND(1.25,1)", 1.3),
        ("=ROUNDDOWN(-1.29,1)", -1.2),
        ('=NUMBERVALUE("1,234.5",".",",")', 1234.5),
        ('=LEFT("abc",2)', "ab"),
        ("=HLOOKUP(1,A1:A3,2,FALSE())", 2.0),
        ("=LOOKUP(2,A1:A3,C1:C3)", 20.0),
        ("=XLOOKUP(2,A1:A3,C1:C3)", 20.0),
        ("=IFERROR(1/0,7)", 7.0),
        ("=IFERROR(2,1/0)", 2.0),
        ("=IFNA(NA(),7)", 7.0),
        ("=IFNA(1/0,7)", "#DIV/0!"),
        ("=ISERROR(1/0)", 1.0),
        ("=ISNA(NA())", 1.0),
        ("=ISNA(1/0)", 0.0),
        ("=ISBLANK(A1)", 0.0),
        ("=ISNUMBER(1/0)", 0.0),
        ("=NA()", "#N/A"),
        ("=ROW()", 5.0),
        ("=ROW(A1)", 1.0),
        ("=COLUMNS(A1:C3)", 3.0),
    ],
)
def test_shared_function_export(
    tmp_path: Path, formula: str, expected: float | str, force_rung: Literal[3] | None
) -> None:
    cells = {"A1": 1, "A2": 2, "A3": 3, "C1": 10, "C2": 20, "C3": 30, "B5": formula}
    workbook = write_workbook(tmp_path / "functions.xlsx", {"Engine": cells})
    document = bindings_document(
        *(
            series_entry(f"value_{i}", f"Engine!{cell}", direction="constant")
            for i, cell in enumerate(cells)
            if cell != "B5"
        ),
        series_entry("result", "Engine!B5", direction="output"),
    )
    package = load_package(generate_inverted(workbook, document, force_rung=force_rung), tmp_path)
    result = package.api.compute_result()
    assert result == (pytest.approx(expected) if isinstance(expected, float) else expected)


def test_shared_helpers_preserve_function_specific_error_handling() -> None:
    from excel_grapher.exporter.inverted_tree.runtime import xl_countif, xl_xlookup

    assert xl_countif((1, "#DIV/0!", 3), ">0") == 2
    assert xl_xlookup(2, (1, 2), ("#DIV/0!", 20)) == 20


@pytest.mark.parametrize("force_rung", [None, 3])
@pytest.mark.parametrize(
    "expression, expected", [("ROW()", (2.0, 3.0, 4.0)), ("ROW($C$2)", (2.0, 2.0, 2.0))]
)
def test_row_geometry_tracks_series_members(
    tmp_path: Path, expression: str, expected: tuple[float, ...], force_rung: Literal[3] | None
) -> None:
    workbook = write_workbook(
        tmp_path / "rows.xlsx",
        {
            "Engine": {
                "C2": 2020,
                "C3": 2021,
                "C4": 2022,
                "D2": f"={expression}",
                "D3": f"={expression}",
                "D4": f"={expression}",
            }
        },
    )
    document = bindings_document(
        series_entry("base_year", "Engine!C2", direction="constant"),
        series_entry(
            "result",
            "Engine!D2:D4",
            layout="series",
            direction="output",
            label_column="C",
        ),
    )
    modules = generate_inverted(workbook, document, force_rung=force_rung)
    assert "_kernels.result(" not in modules["internals.py"]
    package = load_package(modules, tmp_path)
    result = package.api.compute_result()
    assert tuple(result.domain) == ((2020,), (2021,), (2022,))
    internal = package.internals.result(**({"base_year": 2020} if "C$2" in expression else {}))
    assert tuple(internal[year] for year in (2020, 2021, 2022)) == expected
    assert tuple(result[year] for year in (2020, 2021, 2022)) == expected
