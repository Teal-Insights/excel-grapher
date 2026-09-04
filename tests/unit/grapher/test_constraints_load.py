"""Tests for loading `CONSTRAINTS` tables from a constraints.py module."""

from __future__ import annotations

from pathlib import Path
from typing import Literal

import pytest

from excel_grapher.grapher.constraints import (
    ConstraintsLoadError,
    dynamic_refs_from_path,
    load_constraints_module,
)
from excel_grapher.grapher.dynamic_refs import DynamicRefConfig
from tests.paths import INVERTED_TREE_TINY_DSA
from tests.unit.exporter.inverted_tree.local_corpus import load_constraints_module as corpus_load


def test_load_constraints_module_reads_constraints_table(tmp_path: Path) -> None:
    path = tmp_path / "constraints.py"
    path.write_text(
        "from typing import Literal\n\nCONSTRAINTS = {'Sheet1!A1': Literal[1, 2]}\n",
        encoding="utf-8",
    )

    module = load_constraints_module(path)

    assert module.CONSTRAINTS["Sheet1!A1"] == Literal[1, 2]


def test_dynamic_refs_from_path_builds_config() -> None:
    config = dynamic_refs_from_path(INVERTED_TREE_TINY_DSA / "constraints.py")

    assert isinstance(config, DynamicRefConfig)
    assert config.cell_type_env


def test_load_constraints_module_requires_constraints_attr(tmp_path: Path) -> None:
    path = tmp_path / "constraints.py"
    path.write_text("NOTES = 'no table'\n", encoding="utf-8")

    with pytest.raises(ConstraintsLoadError, match="CONSTRAINTS"):
        load_constraints_module(path)


def test_load_constraints_module_requires_mapping(tmp_path: Path) -> None:
    path = tmp_path / "constraints.py"
    path.write_text("CONSTRAINTS = ['Sheet1!A1']\n", encoding="utf-8")

    with pytest.raises(ConstraintsLoadError, match="mapping"):
        load_constraints_module(path)


def test_load_constraints_module_missing_file(tmp_path: Path) -> None:
    missing = tmp_path / "missing.py"

    with pytest.raises(ConstraintsLoadError, match="not found"):
        load_constraints_module(missing)


def test_corpus_loader_shares_production_contract() -> None:
    path = INVERTED_TREE_TINY_DSA / "constraints.py"
    corpus = corpus_load(path)
    production = load_constraints_module(path)

    assert corpus is not None
    assert corpus.CONSTRAINTS == production.CONSTRAINTS
    assert "Inputs!B5" in production.CONSTRAINTS
