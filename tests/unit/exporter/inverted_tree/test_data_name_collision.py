"""Codegen fails closed when two `data.py` bindings share a Python name."""

from __future__ import annotations

from pathlib import Path

import pytest

from excel_grapher.exporter.inverted_tree import InvertedTreeExportError
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    unload_package,
    write_workbook,
)


def _series(
    series_id: str,
    data_range: str,
    *,
    direction: str,
    layout: str = "series",
    mode: str | None = None,
) -> dict[str, object]:
    entry = series_entry(
        series_id,
        data_range,
        layout=layout,
        direction=direction,
        label_column="A" if layout == "series" else None,
        key_concept="COUNTRY",
        key_read="string",
    )
    if mode is not None:
        entry["input"]["mode"] = mode
    return entry


def _export(tmp_path: Path, *series: dict[str, object], cells: dict[str, object]) -> dict[str, str]:
    workbook = write_workbook(tmp_path / "book.xlsx", {"Inputs": cells})
    return generate_inverted(workbook, bindings_document(*series))


def _labels() -> dict[str, object]:
    return {"A2": "grace", "A3": "maturity"}


def test_rate_default_without_rate_input_keeps_one_binding(tmp_path: Path) -> None:
    """A series id `rate_default` is legal when no input sentinel uses that name."""
    modules = _export(
        tmp_path,
        _series("rate_default", "Inputs!C2:C3", direction="internal"),
        _series("rate_out", "Inputs!I2:I3", direction="output"),
        cells={**_labels(), "C2": "=5", "C3": "=10", "I2": "=C2", "I3": "=C3"},
    )
    assignments = [
        line
        for line in modules["data.py"].splitlines()
        if line.startswith("RATE_DEFAULT =") or line.startswith("RATE_DEFAULT:")
    ]
    assert len(assignments) == 1
    assert assignments[0].startswith("RATE_DEFAULT:")
    pkg = load_package(modules, tmp_path, name="rate_default_pkg")
    try:
        assert pkg.data.RATE_DEFAULT.schema.series_id == "rate_default"
    finally:
        unload_package("rate_default_pkg", tmp_path)


@pytest.mark.parametrize("leading", ["rate", "rate_default"])
def test_override_sentinel_collides_with_series_id(tmp_path: Path, leading: str) -> None:
    """`rate` override sentinel and series `rate_default` both bind `RATE_DEFAULT`."""
    series = {
        "rate": _series("rate", "Inputs!H2:H3", direction="input", mode="override"),
        "rate_default": _series("rate_default", "Inputs!C2:C3", direction="internal"),
        "rate_out": _series("rate_out", "Inputs!I2:I3", direction="output"),
    }
    order = ["rate", "rate_default", "rate_out"]
    if leading == "rate_default":
        order = ["rate_default", "rate", "rate_out"]
    with pytest.raises(InvertedTreeExportError, match="generated name RATE_DEFAULT") as raised:
        _export(
            tmp_path,
            *(series[series_id] for series_id in order),
            cells={
                **_labels(),
                "C2": "=5",
                "C3": "=10",
                "H2": "=C2",
                "H3": "=C3",
                "I2": "=H2",
                "I3": "=H3",
            },
        )
    message = str(raised.value)
    assert "input 'rate' sentinel" in message
    assert "series 'rate_default'" in message


def test_leaf_input_sentinel_collides_with_series_id(tmp_path: Path) -> None:
    """A non-override input still owns `{NAME}_DEFAULT`."""
    with pytest.raises(InvertedTreeExportError, match="generated name RATE_DEFAULT") as raised:
        _export(
            tmp_path,
            _series("rate", "Inputs!H2:H3", direction="input"),
            _series("rate_default", "Inputs!C2:C3", direction="internal"),
            _series("rate_out", "Inputs!I2:I3", direction="output"),
            cells={
                **_labels(),
                "C2": "=5",
                "C3": "=10",
                "H2": 5,
                "H3": 10,
                "I2": "=H2",
                "I3": "=H3",
            },
        )
    message = str(raised.value)
    assert "input 'rate' sentinel" in message
    assert "series 'rate_default'" in message


def test_series_id_collides_with_dtype_values(tmp_path: Path) -> None:
    """`float_values` is the shared dtype tuple, not only an input sentinel."""
    with pytest.raises(InvertedTreeExportError, match="generated name FLOAT_VALUES") as raised:
        _export(
            tmp_path,
            _series("float_values", "Inputs!C2:C3", direction="internal"),
            _series("rate_out", "Inputs!I2:I3", direction="output"),
            cells={**_labels(), "C2": "=5", "C3": "=10", "I2": "=C2", "I3": "=C3"},
        )
    message = str(raised.value)
    assert "dtype 'float' values" in message
    assert "series 'float_values'" in message


def test_series_id_collides_with_cells_suffix(tmp_path: Path) -> None:
    """Scalar `{NAME}_CELLS` and a series id `rate_cells` share one binding."""
    with pytest.raises(InvertedTreeExportError, match="generated name RATE_CELLS") as raised:
        _export(
            tmp_path,
            _series("rate", "Inputs!H1", direction="input", layout="scalar"),
            _series("rate_cells", "Inputs!C2:C3", direction="internal"),
            _series("rate_out", "Inputs!I2:I3", direction="output"),
            cells={**_labels(), "H1": 5, "C2": "=5", "C3": "=10", "I2": "=C2", "I3": "=C3"},
        )
    message = str(raised.value)
    assert "series 'rate' cells" in message
    assert "series 'rate_cells'" in message


def test_series_id_collides_with_axis_constant(tmp_path: Path) -> None:
    """Axis constants live in the same module namespace as series bindings."""
    with pytest.raises(InvertedTreeExportError, match="generated name COUNTRY_AXIS") as raised:
        _export(
            tmp_path,
            _series("country_axis", "Inputs!H2:H3", direction="input"),
            _series("rate_out", "Inputs!I2:I3", direction="output"),
            cells={**_labels(), "H2": 5, "H3": 10, "I2": "=H2", "I3": "=H3"},
        )
    message = str(raised.value)
    assert "axis 'COUNTRY'" in message
    assert "series 'country_axis'" in message


def test_public_alias_collides_with_import(tmp_path: Path) -> None:
    """The facade `Series` would overwrite the imported tensor type."""
    with pytest.raises(InvertedTreeExportError, match="generated name Series is used by") as raised:
        _export(
            tmp_path,
            _series("series", "Inputs!H2:H3", direction="input"),
            _series("rate_out", "Inputs!I2:I3", direction="output"),
            cells={**_labels(), "H2": 5, "H3": 10, "I2": "=H2", "I3": "=H3"},
        )
    message = str(raised.value)
    assert "import Series" in message
    assert "series 'series' alias" in message
