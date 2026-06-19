"""Formula and export gaps discovered via ``financial_model.xlsx`` parity."""

from __future__ import annotations

import importlib
import sys
import tempfile
from pathlib import Path

import pytest
import yaml

from excel_grapher import FormulaEvaluator, create_dependency_graph
from excel_grapher.evaluator.types import XlError
from excel_grapher.series_bindings.workflow import (
    generate_bindings_modules,
    validate_bindings_workbook,
)
from tests.unit.gaps.assertions import assert_evaluator_and_codegen_disagree
from tests.unit.gaps.workbook_helpers import (
    write_sumproduct_category_filter,
    write_sumproduct_std_dev,
    write_sumproduct_threshold_count,
    write_text_index_match,
    write_vlookup_false,
)


def _evaluate(path: Path, address: str) -> object:
    graph = create_dependency_graph(
        path,
        [address],
        load_values=True,
        use_cached_dynamic_refs=True,
    )
    with FormulaEvaluator(graph) as evaluator:
        return evaluator.evaluate(address)


def test_sumproduct_variance_std_dev_returns_na(tmp_path: Path) -> None:
    """``SUMPRODUCT((range-AVG(range))^2)/COUNT`` std-dev returns ``#N/A`` in evaluator."""
    workbook = write_sumproduct_std_dev(tmp_path / "std_dev.xlsx")
    assert _evaluate(workbook, "Statistical Analysis!F6") == XlError.NA


def test_sumproduct_string_filter_returns_value_error(tmp_path: Path) -> None:
    """Boolean string filter inside ``SUMPRODUCT`` returns ``#VALUE!``."""
    workbook = write_sumproduct_category_filter(tmp_path / "category_filter.xlsx")
    assert _evaluate(workbook, "Product Lookup!I14") == XlError.VALUE


def test_text_index_match_returns_array_not_formatted_string(tmp_path: Path) -> None:
    r"""``TEXT(INDEX(...))`` evaluates to a stringified array rather than ``"2029"``."""
    workbook = write_text_index_match(tmp_path / "text_index.xlsx")
    result = _evaluate(workbook, "Revenue Model!B22")
    assert result != "2029"
    assert result == "[[2029]]"


def test_sumproduct_threshold_count_eval_codegen_mismatch(tmp_path: Path) -> None:
    """``SUMPRODUCT((range>threshold)*1)`` disagrees between evaluator and codegen."""
    workbook = write_sumproduct_threshold_count(tmp_path / "threshold_count.xlsx")
    graph = create_dependency_graph(
        workbook,
        ["Product Lookup!I18"],
        load_values=True,
        use_cached_dynamic_refs=True,
    )
    assert_evaluator_and_codegen_disagree(graph, "Product Lookup!I18")


def _minimal_vlookup_bindings_yaml() -> dict:
    return {
        "schema_version": "1.3.0",
        "workbook": "vlookup.xlsx",
        "series": [
            {
                "id": "lookup_name",
                "sheet": "Product Lookup",
                "data_range": "Product Lookup!J5",
                "layout": "scalar",
                "editable": False,
                "structure": {
                    "measure": {
                        "concept": "OBS_VALUE",
                        "dtype": "string",
                        "bind": {"kind": "data_cell", "read": "string"},
                    },
                    "dimensions": [],
                },
                "key": [],
                "output": {"compute": {"name": "compute_lookup_name"}},
            }
        ],
    }


def test_modular_vlookup_false_export_missing_xl_false_runtime(tmp_path: Path) -> None:
    """Modular export imports ``xl_false`` for ``VLOOKUP(...,FALSE())`` but runtime lacks it."""
    workbook = write_vlookup_false(tmp_path / "vlookup.xlsx")
    bindings_path = tmp_path / "lookup.bindings.yaml"
    bindings_path.write_text(yaml.safe_dump(_minimal_vlookup_bindings_yaml()), encoding="utf-8")
    result = validate_bindings_workbook(workbook, bindings_path)
    files = generate_bindings_modules(
        result["graph"],
        targets=result["targets"],
        bindings=result["bindings"],
        workbook=workbook,
    )
    assert any("xl_false" in content for content in files.values())
    with tempfile.TemporaryDirectory() as temp_dir:
        module_dir = Path(temp_dir) / "bindings_module"
        module_dir.mkdir()
        for filename, content in files.items():
            (module_dir / filename).write_text(content, encoding="utf-8")
        sys.path.insert(0, temp_dir)
        try:
            with pytest.raises(ImportError, match="xl_false"):
                importlib.import_module("bindings_module")
        finally:
            sys.path.pop(0)
            sys.modules.pop("bindings_module", None)
