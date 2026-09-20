"""Packaged agent skills shipped with excel-grapher."""

from __future__ import annotations

from importlib.resources import files
from pathlib import Path


def author_bindings_skill_dir() -> Path:
    """Return the packaged `author-bindings` skill directory.

    Downstream repos can copy this tree into `.cursor/skills/author-bindings/`
    or Claude's skill path so agents pick up the same loop as this repository.
    """
    return Path(str(files("excel_grapher.skills").joinpath("author-bindings")))
