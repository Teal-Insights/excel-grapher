from __future__ import annotations

from pathlib import Path

from scripts.format_qmd import format_qmd_source


def test_formats_executable_chunk_bodies() -> None:
    source = "```{python}\nx = foo(\n    1,\n    2\n)\n```\n"

    assert format_qmd_source(source, "demo.qmd") == "```{python}\nx = foo(1, 2)\n```\n"


def test_preserves_chunk_option_lines_and_their_blank_separator() -> None:
    source = "```{python}\n#| echo: false\n\nx = {'a':1}\n```\n"

    formatted = format_qmd_source(source, "demo.qmd")

    assert formatted == '```{python}\n#| echo: false\n\nx = {"a": 1}\n```\n'


def test_formats_non_executable_python_fences() -> None:
    source = "```python\nlist_setters()   # comment\n```\n"

    assert format_qmd_source(source, "demo.qmd") == "```python\nlist_setters()  # comment\n```\n"


def test_leaves_non_python_fences_and_prose_untouched() -> None:
    source = "Some prose.\n\n```yaml\nkey:   value\n```\n\nMore prose.\n"

    assert format_qmd_source(source, "demo.qmd") == source


def test_backticks_inside_a_chunk_do_not_end_the_fence() -> None:
    source = '```{python}\nprint(f"```text\\n{x}\\n```")\ny = foo(\n    1\n)\n```\n'

    formatted = format_qmd_source(source, "demo.qmd")

    assert formatted == '```{python}\nprint(f"```text\\n{x}\\n```")\ny = foo(1)\n```\n'


def test_formatting_is_idempotent() -> None:
    source = "```{python}\nx = foo(\n    1,\n    2\n)\n```\n"

    once = format_qmd_source(source, "demo.qmd")

    assert format_qmd_source(once, "demo.qmd") == once


def test_repository_qmd_sources_are_formatted() -> None:
    qmd_files = sorted(Path("examples").rglob("*.qmd"))
    assert qmd_files, "expected at least one .qmd source under examples/"

    unformatted = [
        path
        for path in qmd_files
        if format_qmd_source(path.read_text(encoding="utf-8"), path.name)
        != path.read_text(encoding="utf-8")
    ]

    assert unformatted == []
