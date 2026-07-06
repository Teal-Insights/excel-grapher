"""Tests for series binding smoke helpers."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
import xlsxwriter
import yaml

from excel_grapher.series_bindings.smoke import BindingsSmokeError, smoke_test_bindings_module
from excel_grapher.series_bindings.workflow import (
    generate_bindings_modules,
    run_binding_checks,
    validate_bindings_workbook,
)
from tests.integration.user_flows.utils import write_ffv2_workbook
from tests.paths import SERIES_BINDINGS_FIXTURES as FIXTURES

FFV2_BINDINGS = FIXTURES / "ffv2.yaml"
BORVELIA_BINDINGS = FIXTURES / "borvelia_primary_balance.yaml"


def _write_borvelia_workbook(path: Path) -> None:
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Inputs")
    ws.write("A2", "Borvelia")
    ws.write("A5", "Primary balance (% of GDP)")
    for col, year in enumerate([1, 2, 3, 4, 5], start=5):
        ws.write(0, col, year)
        ws.write_number(4, col, float(year - 3))
    wb.close()


def _load_ffv2_bindings_document() -> dict[str, Any]:
    return yaml.safe_load(FFV2_BINDINGS.read_text(encoding="utf-8"))


def _write_bindings_variant(path: Path, mutate: Callable[[dict[str, Any]], None]) -> None:
    document = _load_ffv2_bindings_document()
    mutate(document)
    path.write_text(yaml.safe_dump(document, sort_keys=False), encoding="utf-8")


def test_smoke_test_bindings_module_ffv2_fixture(tmp_path: Path) -> None:
    workbook = tmp_path / "ffv2.xlsx"
    write_ffv2_workbook(workbook)
    bindings_path = FFV2_BINDINGS
    result = validate_bindings_workbook(workbook, bindings_path)
    files = generate_bindings_modules(
        result["graph"],
        targets=result["targets"],
        bindings=result["bindings"],
        workbook=workbook,
    )
    module_dir = tmp_path / "bindings_module"

    smoke_test_bindings_module(
        files,
        bindings=result["bindings"],
        graph=result["graph"],
        workbook=workbook,
        module_dir=module_dir,
        package_name="bindings_module",
    )


@pytest.fixture
def ffv2_workbook(tmp_path: Path) -> Path:
    path = tmp_path / "ffv2.xlsx"
    write_ffv2_workbook(path)
    return path


def test_smoke_test_setters_positional_and_dataframe_inputs(tmp_path: Path) -> None:
    """Single-key setters accept positional values and tidy DataFrames in smoke."""
    pytest.importorskip("pandas")
    workbook = tmp_path / "lic_inputs.xlsx"
    _write_borvelia_workbook(workbook)
    result = validate_bindings_workbook(workbook, BORVELIA_BINDINGS)
    files = generate_bindings_modules(
        result["graph"],
        targets=result["targets"],
        bindings=result["bindings"],
        workbook=workbook,
    )
    smoke_test_bindings_module(
        files,
        bindings=result["bindings"],
        graph=result["graph"],
        workbook=workbook,
        module_dir=tmp_path / "bindings_module",
        package_name="bindings_module",
    )


def test_smoke_fails_when_setter_names_collide(
    ffv2_workbook: Path,
    tmp_path: Path,
) -> None:
    """Duplicate setter names validate but smoke fails when the second setter is a no-op."""
    bindings_path = tmp_path / "dup_setter.yaml"

    def _duplicate_first_setter(document: dict[str, Any]) -> None:
        duplicate_name = document["series"][0]["input"]["setter"]["name"]
        document["series"][1]["input"]["setter"]["name"] = duplicate_name

    _write_bindings_variant(bindings_path, _duplicate_first_setter)

    result = validate_bindings_workbook(ffv2_workbook, bindings_path)
    assert result["report"]["ok"] is True

    files = generate_bindings_modules(
        result["graph"],
        targets=result["targets"],
        bindings=result["bindings"],
        workbook=ffv2_workbook,
    )

    with pytest.raises(BindingsSmokeError, match=r"Setter 'set_puka_receptions' did not update"):
        smoke_test_bindings_module(
            files,
            bindings=result["bindings"],
            graph=result["graph"],
            workbook=ffv2_workbook,
            module_dir=tmp_path / "bindings_module",
            package_name="bindings_module",
        )


def test_run_binding_checks_raises_when_validation_fails(
    ffv2_workbook: Path,
    tmp_path: Path,
) -> None:
    """Invalid key concepts fail validation before smoke is attempted."""
    bindings_path = tmp_path / "missing_key.yaml"

    def _reference_unknown_key(document: dict[str, Any]) -> None:
        document["series"][0]["key"] = ["NONEXISTENT"]

    _write_bindings_variant(bindings_path, _reference_unknown_key)

    result = validate_bindings_workbook(ffv2_workbook, bindings_path)
    assert result["report"]["ok"] is False
    assert any(issue["code"] == "key_not_in_dimensions" for issue in result["report"]["issues"])

    with pytest.raises(ValueError, match="Binding validation failed"):
        run_binding_checks(
            ffv2_workbook,
            bindings_path,
            module_dir=tmp_path / "bindings_module",
            package_name="bindings_module",
        )


def test_smoke_fails_when_compute_returns_wrong_record_count(
    ffv2_workbook: Path,
    tmp_path: Path,
) -> None:
    """Tampered generated compute code fails smoke record-count checks."""
    bindings_path = FFV2_BINDINGS
    result = validate_bindings_workbook(ffv2_workbook, bindings_path)
    files = generate_bindings_modules(
        result["graph"],
        targets=result["targets"],
        bindings=result["bindings"],
        workbook=ffv2_workbook,
    )
    api_py = files["api.py"]
    files["api.py"] = api_py.replace("    return records", "    return records[:1]")

    with pytest.raises(BindingsSmokeError, match=r"returned 1 records, expected 16"):
        smoke_test_bindings_module(
            files,
            bindings=result["bindings"],
            graph=result["graph"],
            workbook=ffv2_workbook,
            module_dir=tmp_path / "bindings_module",
            package_name="bindings_module",
        )
