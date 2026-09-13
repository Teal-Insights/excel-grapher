"""Sandbox binding reconciliation: numeric cells declared as text read as numbers."""

from __future__ import annotations

import importlib.util
from pathlib import Path
from types import SimpleNamespace

_SCRIPT = Path(__file__).resolve().parents[3] / "scripts" / "measure_named_export.py"
_spec = importlib.util.spec_from_file_location("measure_named_export", _SCRIPT)
assert _spec is not None and _spec.loader is not None
measure = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(measure)


class _Graph:
    def __init__(self, values: dict[str, object]) -> None:
        self._values = values

    def get_node(self, address: str) -> SimpleNamespace | None:
        if address not in self._values:
            return None
        return SimpleNamespace(value=self._values[address])


def _entry(dtype: str) -> dict:
    return {
        "id": "flow",
        "input": {"setter": {"name": "set_flow"}},
        "structure": {
            "measure": {
                "concept": "OBS_VALUE",
                "dtype": dtype,
                "bind": {"kind": "data_cell", "read": dtype},
            }
        },
    }


def test_numeric_cells_declared_as_text_become_float_inputs() -> None:
    graph = _Graph({"S!B2": 1.0, "S!C2": 0})
    entry = _entry("string")
    assert measure.numeric_text_measure(graph, entry, ("S!B2", "S!C2")) is True
    measure.read_as_numbers(entry)
    assert entry["structure"]["measure"]["dtype"] == "float"
    assert entry["structure"]["measure"]["bind"]["read"] == "float"


def test_text_cells_and_typed_measures_are_left_alone() -> None:
    assert (
        measure.numeric_text_measure(_Graph({"S!B2": "n/a"}), _entry("string"), ("S!B2",)) is False
    )
    assert measure.numeric_text_measure(_Graph({"S!B2": 1.0}), _entry("float"), ("S!B2",)) is False
    assert measure.numeric_text_measure(_Graph({}), _entry("string"), ("S!B2",)) is False


def test_integer_measures_over_fractional_cells_become_float() -> None:
    assert (
        measure.fractional_int_measure(
            _Graph({"S!B2": 11.94, "S!C2": 3}), _entry("int"), ("S!B2", "S!C2")
        )
        is True
    )
    assert (
        measure.fractional_int_measure(
            _Graph({"S!B2": 11.0, "S!C2": 3}), _entry("int"), ("S!B2", "S!C2")
        )
        is False
    )
    assert (
        measure.fractional_int_measure(_Graph({"S!B2": 11.94}), _entry("float"), ("S!B2",)) is False
    )
