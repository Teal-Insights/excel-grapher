"""Unit tests for prefix-variant binding shard load and validation."""

from __future__ import annotations

import importlib
import sys
from pathlib import Path

import pytest
import yaml

from excel_grapher.series_bindings import load_series_bindings
from excel_grapher.series_bindings.smoke import smoke_test_bindings_module
from excel_grapher.series_bindings.workflow import (
    generate_bindings_modules,
    validate_bindings_workbook,
)
from tests.unit.fixtures.prefix_variant_workbook import (
    PrefixVariant,
    prefix_variant_binding_document,
    workbook_filename,
    write_prefix_variant_binding_shard,
    write_prefix_variant_workbook,
)


@pytest.mark.parametrize("variant", ["xlfn", "xludf"])
def test_prefix_variant_binding_shard_loads_with_matching_workbook_name(
    variant: PrefixVariant,
) -> None:
    document = prefix_variant_binding_document(variant=variant)
    assert document["workbook"] == workbook_filename(variant)


@pytest.mark.parametrize("variant", ["xlfn", "xludf"])
def test_prefix_variant_bindings_validate(tmp_path: Path, variant: PrefixVariant) -> None:
    workbook = write_prefix_variant_workbook(tmp_path / workbook_filename(variant), variant=variant)
    bindings_path = tmp_path / f"{variant}.bindings.yaml"
    bindings_path.write_text(
        yaml.safe_dump(prefix_variant_binding_document(variant=variant), sort_keys=False),
        encoding="utf-8",
    )

    result = validate_bindings_workbook(workbook, bindings_path)
    assert result["report"]["ok"], result["report"]["issues"]
    assert result["bindings"]["workbook"] == workbook_filename(variant)


@pytest.mark.parametrize("variant", ["xlfn", "xludf"])
def test_prefix_variant_binding_directory_merge(tmp_path: Path, variant: PrefixVariant) -> None:
    shard_dir = tmp_path / f"{variant}.bindings"
    shard_dir.mkdir()
    write_prefix_variant_binding_shard(
        shard_dir / "lookups.bindings.yaml",
        variant=variant,
        series_id="lookup_result",
    )
    write_prefix_variant_binding_shard(
        shard_dir / "mirror.bindings.yaml",
        variant=variant,
        series_id="lookup_mirror",
    )

    merged = load_series_bindings(shard_dir)
    assert merged["workbook"] == workbook_filename(variant)
    assert {series["id"] for series in merged["series"]} == {"lookup_result", "lookup_mirror"}


@pytest.mark.parametrize("variant", ["xlfn", "xludf"])
def test_prefix_variant_bindings_generate_modules_and_smoke(
    tmp_path: Path,
    variant: PrefixVariant,
) -> None:
    workbook = write_prefix_variant_workbook(tmp_path / workbook_filename(variant), variant=variant)
    bindings_path = tmp_path / f"{variant}.bindings.yaml"
    bindings_path.write_text(
        yaml.safe_dump(prefix_variant_binding_document(variant=variant), sort_keys=False),
        encoding="utf-8",
    )
    result = validate_bindings_workbook(workbook, bindings_path)
    files = generate_bindings_modules(
        result["graph"],
        targets=result["targets"],
        bindings=result["bindings"],
        workbook=workbook,
    )

    module_dir = tmp_path / f"{variant}_module"
    smoke_test_bindings_module(
        files,
        bindings=result["bindings"],
        graph=result["graph"],
        workbook=workbook,
        module_dir=module_dir,
        package_name=f"{variant}_module",
    )

    sys.path.insert(0, str(tmp_path))
    try:
        package = importlib.import_module(f"{variant}_module")
        records = package.compute_lookup_result(ctx=package.make_context())
        assert records[0]["OBS_VALUE"] == "hit"
    finally:
        sys.path.pop(0)
        sys.modules.pop(f"{variant}_module", None)
