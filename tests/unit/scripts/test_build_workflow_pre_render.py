from __future__ import annotations

from pathlib import Path

from scripts.build_workflow import (
    WORKFLOW_HTML_EXCLUDE,
    WORKFLOW_RESOURCE,
    copy_workflow_fallback,
    register_workflow_resources,
)


def test_register_workflow_resources_adds_resource_and_render_exclude(tmp_path: Path) -> None:
    quarto_yml = tmp_path / "_quarto.yml"
    quarto_yml.write_text(
        "project:\n  resources: []\n  render:\n    - '**'\n",
        encoding="utf-8",
    )

    changed = register_workflow_resources(quarto_yml)

    assert changed is True
    config = quarto_yml.read_text(encoding="utf-8")
    assert WORKFLOW_RESOURCE in config
    assert WORKFLOW_HTML_EXCLUDE in config


def test_register_workflow_resources_is_idempotent(tmp_path: Path) -> None:
    quarto_yml = tmp_path / "_quarto.yml"
    quarto_yml.write_text(
        "project:\n"
        f"  resources:\n    - '{WORKFLOW_RESOURCE}'\n"
        "  render:\n"
        "    - '**'\n"
        f"    - '{WORKFLOW_HTML_EXCLUDE}'\n",
        encoding="utf-8",
    )

    changed = register_workflow_resources(quarto_yml)

    assert changed is False


def test_copy_workflow_fallback_copies_snapshot(tmp_path: Path) -> None:
    source = tmp_path / "custom" / "workflow"
    source.mkdir(parents=True)
    (source / "index.html").write_text("<html></html>", encoding="utf-8")
    (source / "workflow.json").write_text("{}", encoding="utf-8")
    output = tmp_path / "great-docs" / "workflow"

    copied = copy_workflow_fallback(source_dir=source, output_dir=output)

    assert copied is True
    assert (output / "index.html").read_text(encoding="utf-8") == "<html></html>"
    assert (output / "workflow.json").read_text(encoding="utf-8") == "{}"


def test_copy_workflow_fallback_returns_false_when_incomplete(tmp_path: Path) -> None:
    source = tmp_path / "custom" / "workflow"
    source.mkdir(parents=True)
    (source / "index.html").write_text("<html></html>", encoding="utf-8")
    output = tmp_path / "workflow"

    copied = copy_workflow_fallback(source_dir=source, output_dir=output)

    assert copied is False
    assert not output.exists()
