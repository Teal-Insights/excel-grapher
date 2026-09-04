"""Cached text in a float-dtyped constant is a measure, not a crash."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)


def _mcve_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a15_thousands.xlsx",
        {
            "Store": {
                "A2": "1 000",
                "B2": "=A2",
            },
        },
    )


def _mcve_bindings() -> dict:
    return bindings_document(
        series_entry("store", "Store!A2", layout="scalar", direction="constant"),
        series_entry("out", "Store!B2", layout="scalar", direction="output"),
    )


def _sentinel_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "a15_sentinels.xlsx",
        {
            "Store": {
                "A1": 1,
                "B1": 2,
                "C1": 3,
                "D1": 4,
                "E1": 5,
                "F1": 6,
                "A2": "1 000",
                "B2": "n/a",
                "C2": "..",
                "D2": "--",
                "E2": "",
                "F2": 3.5,
                "A3": "=A2",
            },
        },
    )


def _sentinel_bindings() -> dict:
    return bindings_document(
        series_entry(
            "store",
            "Store!A2:F2",
            layout="series",
            direction="constant",
            header_row=1,
        ),
        series_entry("out", "Store!A3", layout="scalar", direction="output"),
    )


def _scalar(value: object) -> object:
    if isinstance(value, tuple):
        assert len(value) == 1
        return value[0]
    return value


def test_thousands_space_shared_string_emits_as_float(tmp_path: Path) -> None:
    modules = generate_inverted(_mcve_workbook(tmp_path), _mcve_bindings())
    assert "1000.0" in modules["data.py"]
    assert "1 000" not in modules["data.py"]
    pkg = load_package(modules, tmp_path, name="a15_thousands")
    assert _scalar(pkg.compute_out()) == pytest.approx(1000.0)


def test_imf_sentinels_stay_strings_in_float_constant(tmp_path: Path) -> None:
    modules = generate_inverted(_sentinel_workbook(tmp_path), _sentinel_bindings())
    data = modules["data.py"]
    assert "1000.0" in data
    assert "'n/a'" in data
    assert "'..'" in data
    assert "'--'" in data
    assert "3.5" in data
    assert "STORE: tuple[float | str, ...] =" in data
    pkg = load_package(modules, tmp_path, name="a15_sentinels")
    store = pkg.data.STORE
    assert store[0] == pytest.approx(1000.0)
    assert store[1:4] == ("n/a", "..", "--")
    assert store[4] in {0, ""}
    assert store[5] == pytest.approx(3.5)
    assert _scalar(pkg.compute_out()) == pytest.approx(1000.0)


def test_empty_and_nbsp_cached_text() -> None:
    from excel_grapher.exporter.inverted_tree.emit import _cell_value

    class _Graph:
        def __init__(self, value: object) -> None:
            self._value = value

        def get_node(self, address: str) -> object:
            del address
            return type("Node", (), {"value": self._value})()

    assert _cell_value(_Graph(""), "Store!A2", "float") == ""  # type: ignore[arg-type]
    assert _cell_value(_Graph("1\u00a0000"), "Store!A2", "float") == 1000.0  # type: ignore[arg-type]


def test_unreadable_int_constant_names_the_cell(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "a15_bad_int.xlsx",
        {"Store": {"A2": "n/a", "B2": "=A2"}},
    )
    document = bindings_document(
        series_entry("store", "Store!A2", layout="scalar", direction="constant", dtype="int"),
        series_entry("out", "Store!B2", layout="scalar", direction="output", dtype="int"),
    )
    with pytest.raises(InvertedTreeExportError, match="Store!A2"):
        generate_inverted(workbook, document)
