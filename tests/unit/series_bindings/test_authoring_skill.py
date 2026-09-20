"""Tests for the packaged author-bindings agent skill."""

from __future__ import annotations

from pathlib import Path

import yaml

from excel_grapher.series_bindings.versions import CURRENT_SCHEMA_VERSION
from excel_grapher.skills import author_bindings_skill_dir

_REPO_ROOT = Path(__file__).resolve().parents[3]
_CANONICAL = _REPO_ROOT / "excel_grapher" / "skills" / "author-bindings"
_COPIES = (
    _REPO_ROOT / "skills" / "author-bindings",
    _REPO_ROOT / ".cursor" / "skills" / "author-bindings",
)


def _skill_files(root: Path) -> dict[str, str]:
    files: dict[str, str] = {}
    for path in sorted(root.rglob("*")):
        if path.is_file():
            files[str(path.relative_to(root))] = path.read_text(encoding="utf-8")
    return files


def test_author_bindings_skill_dir_matches_package_copy() -> None:
    packaged = author_bindings_skill_dir()
    assert packaged.is_dir()
    assert (packaged / "SKILL.md").is_file()
    assert _skill_files(packaged) == _skill_files(_CANONICAL)


def test_skill_copies_stay_in_sync() -> None:
    canonical = _skill_files(_CANONICAL)
    assert canonical
    for copy in _COPIES:
        assert _skill_files(copy) == canonical


def test_skill_documents_authoring_loop_without_tiny_dsa() -> None:
    texts = [
        (_CANONICAL / "SKILL.md").read_text(encoding="utf-8"),
        (_CANONICAL / "references" / "conventions.md").read_text(encoding="utf-8"),
        (_CANONICAL / "references" / "pitfalls.md").read_text(encoding="utf-8"),
        (_CANONICAL / "references" / "validation-loop.md").read_text(encoding="utf-8"),
    ]
    joined = "\n".join(texts)
    lowered = joined.lower()
    assert "CURRENT_SCHEMA_VERSION" in joined
    assert "fill: true" in joined
    assert "constant: {}" in joined
    assert "bind.kind: constant" in joined
    assert "unreachable" in lowered
    assert "uniquify" in lowered or "unique" in lowered
    assert "matrix" in lowered
    assert "semantic family" in lowered
    assert "coverage worklist" in lowered
    assert "upsert" in lowered
    assert "tiny dsa" not in lowered
    assert "tiny-dsa" not in lowered
    assert "workbook_config" not in lowered
    assert "load_pipeline_config" not in lowered
    assert "bindings emit" not in lowered
    assert "author_bindings" not in lowered


def test_measure_shard_example_keeps_fill_and_shared_compute_name() -> None:
    example = yaml.safe_load(
        (_CANONICAL / "assets" / "measure-shards.example.yaml").read_text(encoding="utf-8")
    )
    assert example["schema_version"] == CURRENT_SCHEMA_VERSION
    outputs = example["outputs"]["series"]
    assert {series["output"]["compute"]["name"] for series in outputs} == {"compute_gap_milestones"}
    for series in outputs:
        time_binds = [
            dim["bind"] for dim in series["structure"]["dimensions"] if dim["id"] == "TIME_PERIOD"
        ]
        assert time_binds[0]["fill"] is True
    internals = example["internals"]["series"]
    assert set(internals[0]["key"]) == {"SCENARIO", "TIME_PERIOD", "MEASURE"}


def test_catalog_example_is_pedagogical_not_an_emitter_input() -> None:
    example = yaml.safe_load(
        (_CANONICAL / "assets" / "catalog.example.yaml").read_text(encoding="utf-8")
    )
    assert example["schema_version"] == CURRENT_SCHEMA_VERSION
    for direction in ("inputs", "outputs", "internals", "constants"):
        assert isinstance(example[direction]["series"], list)
    skill = (_CANONICAL / "SKILL.md").read_text(encoding="utf-8")
    assert "not an input to a generator" in skill.lower() or "pedagogical" in skill.lower()
