import numpy as np

from excel_grapher.evaluator.types import ExcelRange


def test_excel_range_cell_addresses() -> None:
    r = ExcelRange(sheet="Sheet1", start_row=1, start_col=1, end_row=2, end_col=2)
    assert list(r.cell_addresses()) == [
        "Sheet1!A1",
        "Sheet1!B1",
        "Sheet1!A2",
        "Sheet1!B2",
    ]


def test_excel_range_cell_addresses_quote_sheet_with_space() -> None:
    """Keys must match DependencyGraph / Node.key (format_cell_key rules)."""
    r = ExcelRange(sheet="Imported data", start_row=126, start_col=1, end_row=126, end_col=1)
    assert list(r.cell_addresses()) == ["'Imported data'!A126"]


def test_excel_range_resolve_shapes_array() -> None:
    r = ExcelRange(sheet="S", start_row=1, start_col=1, end_row=2, end_col=3)
    mapping = {
        "S!A1": 1,
        "S!B1": 2,
        "S!C1": 3,
        "S!A2": 4,
        "S!B2": 5,
        "S!C2": 6,
    }

    arr = r.resolve(lambda addr: mapping.get(addr))
    assert isinstance(arr, np.ndarray)
    assert arr.shape == (2, 3)
    assert arr.tolist() == [[1, 2, 3], [4, 5, 6]]
