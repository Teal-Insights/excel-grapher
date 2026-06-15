"""Tests for series binding smoke helpers."""

from __future__ import annotations

from pathlib import Path

from excel_grapher.series_bindings.smoke import smoke_test_bindings_module
from excel_grapher.series_bindings.workflow import (
    generate_bindings_modules,
    validate_bindings_workbook,
)
from tests.integration.user_flows.utils import write_ffv2_workbook

FIXTURES = Path(__file__).resolve().parents[2] / "fixtures" / "series_bindings"


def test_smoke_test_bindings_module_ffv2_fixture(tmp_path: Path) -> None:
    workbook = tmp_path / "ffv2.xlsx"
    write_ffv2_workbook(workbook)
    bindings_path = FIXTURES / "ffv2.yaml"
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
