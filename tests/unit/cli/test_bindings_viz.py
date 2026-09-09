"""Tests for ``excel-grapher bindings viz``."""

from __future__ import annotations

from pathlib import Path

import pytest
import yaml

from excel_grapher.cli import main
from tests.unit.exporter.inverted_tree.test_shape_a11_zipper import (
    _zipper_bindings,
    _zipper_workbook,
)


def test_main_bindings_viz_writes_html(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    workbook = _zipper_workbook(tmp_path)
    bindings_path = tmp_path / "zipper.bindings.yaml"
    bindings_path.write_text(yaml.safe_dump(_zipper_bindings(), sort_keys=False), encoding="utf-8")
    out = tmp_path / "zipper.html"
    exit_code = main(
        [
            "bindings",
            "viz",
            str(workbook),
            "--bindings",
            str(bindings_path),
            "--output",
            str(out),
            "--json",
        ]
    )
    captured = capsys.readouterr()
    assert exit_code == 0, captured.err
    assert out.is_file()
    html = out.read_text(encoding="utf-8")
    assert "debt" in html
    assert "adjustment" in html
    json_path = out.with_suffix(".viz.json")
    assert json_path.is_file()
    assert "wrote" in captured.out
    assert "statements=" in captured.out


def test_main_bindings_viz_forwards_blank_ranges(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    from tests.unit.exporter.inverted_tree.test_blank_ranges import (
        _mcve_bindings,
        _mcve_workbook,
    )

    workbook = _mcve_workbook(tmp_path)
    bindings_path = tmp_path / "blank.bindings.yaml"
    bindings_path.write_text(yaml.safe_dump(_mcve_bindings(), sort_keys=False), encoding="utf-8")
    blanks = tmp_path / "blank_ranges.py"
    blanks.write_text('BLANK_RANGES = ("Lookup!A1:C3",)\n', encoding="utf-8")
    out = tmp_path / "blank.html"
    exit_code = main(
        [
            "bindings",
            "viz",
            str(workbook),
            "--bindings",
            str(bindings_path),
            "--blank-ranges",
            str(blanks),
            "--output",
            str(out),
        ]
    )
    captured = capsys.readouterr()
    assert exit_code == 0, captured.err
    assert out.is_file()
    assert "statements=" in captured.out


def test_main_bindings_viz_missing_workbook(tmp_path: Path) -> None:
    exit_code = main(
        [
            "bindings",
            "viz",
            str(tmp_path / "missing.xlsx"),
            "--output",
            str(tmp_path / "out.html"),
        ]
    )
    assert exit_code == 1
