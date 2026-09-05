"""Tests for series binding smoke helpers."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
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


def _load_ffv2_bindings_document() -> dict[str, Any]:
    return yaml.safe_load(FFV2_BINDINGS.read_text(encoding="utf-8"))


def _write_bindings_variant(path: Path, mutate: Callable[[dict[str, Any]], None]) -> None:
    document = _load_ffv2_bindings_document()
    mutate(document)
    path.write_text(yaml.safe_dump(document, sort_keys=False), encoding="utf-8")


@pytest.fixture
def ffv2_workbook(tmp_path: Path) -> Path:
    path = tmp_path / "ffv2.xlsx"
    write_ffv2_workbook(path)
    return path


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


def test_run_binding_checks_smokes_inverted_tree_computes(tmp_path: Path) -> None:
    import yaml

    from tests.unit.exporter.inverted_tree.helpers import (
        bindings_document,
        series_entry,
        write_workbook,
    )

    workbook = write_workbook(
        tmp_path / "inv.xlsx",
        {
            "Inputs": {"A1": 2.0},
            "Engine": {"A1": "=Inputs!A1*3"},
            "Outputs": {"A1": "=Engine!A1"},
        },
    )
    document = bindings_document(
        series_entry("x", "Inputs!A1", layout="scalar", direction="input"),
        series_entry("y", "Engine!A1", layout="scalar", direction="internal"),
        series_entry(
            "z",
            "Outputs!A1",
            layout="scalar",
            direction="output",
            compute_name="compute_z",
        ),
    )
    document["workbook"] = "inv.xlsx"
    bindings_path = tmp_path / "inv.bindings.yaml"
    bindings_path.write_text(yaml.safe_dump(document, sort_keys=False), encoding="utf-8")

    result = run_binding_checks(
        workbook,
        bindings_path,
        module_dir=tmp_path / "inv_pkg",
        package_name="inv_pkg",
        smoke_test=True,
    )
    assert "api.py" in result["generated_files"]
    assert "def compute_z" in result["generated_files"]["api.py"]


def test_run_binding_checks_inverted_tree_smoke_with_in_domain_default(
    tmp_path: Path,
) -> None:
    import yaml

    from tests.unit.exporter.inverted_tree.helpers import (
        bindings_document,
        series_entry,
        write_workbook,
    )

    workbook = write_workbook(
        tmp_path / "inv.xlsx",
        {
            "Inputs": {"A1": 0},
            "Engine": {"A1": "=Inputs!A1"},
            "Outputs": {"A1": "=Engine!A1"},
        },
    )
    document = bindings_document(
        series_entry(
            "flag",
            "Inputs!A1",
            layout="scalar",
            direction="input",
            dtype="int",
            domain={"enum": [0, 1]},
        ),
        series_entry("y", "Engine!A1", layout="scalar", direction="internal", dtype="int"),
        series_entry(
            "z",
            "Outputs!A1",
            layout="scalar",
            direction="output",
            dtype="int",
            compute_name="compute_z",
        ),
    )
    document["workbook"] = "inv.xlsx"
    bindings_path = tmp_path / "inv.bindings.yaml"
    bindings_path.write_text(yaml.safe_dump(document, sort_keys=False), encoding="utf-8")

    result = run_binding_checks(
        workbook,
        bindings_path,
        module_dir=tmp_path / "inv_domain_pkg",
        package_name="inv_domain_pkg",
        smoke_test=True,
    )
    assert "require_input_domain" in result["generated_files"]["api.py"]


def test_run_binding_checks_inverted_tree_smoke_with_value_map(tmp_path: Path) -> None:
    import yaml

    from tests.unit.exporter.inverted_tree.helpers import (
        bindings_document,
        series_entry,
        write_workbook,
    )

    workbook = write_workbook(
        tmp_path / "inv.xlsx",
        {
            "Inputs": {"A1": "High "},
            "Engine": {"A1": '=IF(Inputs!A1="High ",10,0)'},
            "Outputs": {"A1": "=Engine!A1"},
        },
    )
    document = bindings_document(
        series_entry(
            "selector",
            "Inputs!A1",
            layout="scalar",
            direction="input",
            dtype="string",
            value_map={"High": "High ", "Medium": "Medium", "Low": "Low "},
        ),
        series_entry("y", "Engine!A1", layout="scalar", direction="internal", dtype="float"),
        series_entry(
            "z",
            "Outputs!A1",
            layout="scalar",
            direction="output",
            dtype="float",
            compute_name="compute_z",
        ),
        schema_version="1.15.0",
    )
    document["workbook"] = "inv.xlsx"
    bindings_path = tmp_path / "inv.bindings.yaml"
    bindings_path.write_text(yaml.safe_dump(document, sort_keys=False), encoding="utf-8")

    result = run_binding_checks(
        workbook,
        bindings_path,
        module_dir=tmp_path / "inv_value_map_pkg",
        package_name="inv_value_map_pkg",
        smoke_test=True,
    )
    assert "apply_input_value_map" in result["generated_files"]["api.py"]


def test_inverted_tree_smoke_fails_when_compute_returns_wrong_length(tmp_path: Path) -> None:
    from tests.unit.exporter.inverted_tree.helpers import (
        bindings_document,
        series_entry,
        write_workbook,
    )

    workbook = write_workbook(
        tmp_path / "inv.xlsx",
        {
            "Inputs": {"A1": 2.0},
            "Engine": {"A1": "=Inputs!A1*3"},
            "Outputs": {"A1": "=Engine!A1"},
        },
    )
    document = bindings_document(
        series_entry("x", "Inputs!A1", layout="scalar", direction="input"),
        series_entry("y", "Engine!A1", layout="scalar", direction="internal"),
        series_entry(
            "z",
            "Outputs!A1",
            layout="scalar",
            direction="output",
            compute_name="compute_z",
        ),
    )
    result = validate_bindings_workbook(workbook, _write_inv_sidecar(tmp_path, document))
    files = generate_bindings_modules(
        result["graph"],
        bindings=result["bindings"],
        workbook=workbook,
    )
    files["api.py"] = files["api.py"].replace("    return (z,)", "    return (z, z)")
    with pytest.raises(BindingsSmokeError, match=r"returned 2 values, expected 1"):
        smoke_test_bindings_module(
            files,
            bindings=result["bindings"],
            graph=result["graph"],
            workbook=workbook,
            module_dir=tmp_path / "inv_pkg",
            package_name="inv_pkg",
        )


def _write_inv_sidecar(tmp_path: Path, document: dict[str, Any]) -> Path:
    import yaml

    document = dict(document)
    document["workbook"] = "inv.xlsx"
    path = tmp_path / "inv.bindings.yaml"
    path.write_text(yaml.safe_dump(document, sort_keys=False), encoding="utf-8")
    return path
