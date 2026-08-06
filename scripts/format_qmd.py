"""Run `ruff format` over the Python code blocks inside Quarto `.qmd` sources.

Ruff 0.16 formats Python embedded in Markdown, which means the rendered
`examples/**/*.md` files are now checked by `ruff format --check`. Those files
are generated from `.qmd` sources, so formatting only the Markdown would be
undone by the next `quarto render`. This script formats the `.qmd` code blocks
themselves, keeping source and rendered output in agreement.

Both executable chunks (```` ```{python} ````) and plain illustrative fences
(```` ```python ````) are formatted; Quarto chunk options (`#|` lines) are
preserved verbatim at the top of the chunk.

Usage:
    uv run python scripts/format_qmd.py [--check] [PATH ...]

Paths default to every `.qmd` file under `examples/`. With `--check`, nothing is
written and a non-zero exit status reports files that need formatting.
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

CLOSING_FENCE = "```"

# Fence headers whose bodies are Python, mapped to whether the block is an
# executable Quarto chunk (and therefore may carry `#|` option lines).
PYTHON_FENCES = {
    "```{python}": True,
    "```python": False,
}


class RuffFormatError(RuntimeError):
    """Raised when the `ruff format` subprocess fails on a code block."""


def _ruff_format(source: str, filename: str) -> str:
    """Format `source` with `ruff format`, reading from stdin.

    Args:
        source: Python source text for a single code block.
        filename: Name reported to ruff so configuration and error messages
            resolve against the right file.

    Returns:
        The formatted source text.

    Raises:
        RuffFormatError: If ruff exits non-zero.
    """
    result = subprocess.run(
        [sys.executable, "-m", "ruff", "format", "-", "--stdin-filename", filename],
        input=source,
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        raise RuffFormatError(f"ruff format failed for {filename}:\n{result.stderr}")
    return result.stdout


def format_qmd_source(text: str, filename: str = "document.qmd") -> str:
    """Format every Python code block in a `.qmd` document.

    Args:
        text: Full contents of the `.qmd` document.
        filename: Name reported to ruff for configuration resolution.

    Returns:
        The document with its Python code blocks formatted. Prose, non-Python
        fences, and Quarto chunk options are returned unchanged.

    Raises:
        RuffFormatError: If a code block cannot be formatted.
        ValueError: If a Python code fence is never closed.
    """
    lines = text.split("\n")
    out: list[str] = []
    index = 0

    while index < len(lines):
        line = lines[index]
        is_chunk = PYTHON_FENCES.get(line.strip())
        if is_chunk is None:
            out.append(line)
            index += 1
            continue

        end = index + 1
        while end < len(lines) and lines[end].rstrip() != CLOSING_FENCE:
            end += 1
        if end == len(lines):
            raise ValueError(f"unterminated code fence at {filename}:{index + 1}")

        body = lines[index + 1 : end]
        options: list[str] = []
        if is_chunk:
            while body and body[0].startswith("#|"):
                options.append(body.pop(0))
            # Keep the blank line authors put between options and code.
            if options and body and not body[0].strip():
                options.append(body.pop(0))

        formatted = _ruff_format("\n".join(body) + "\n", filename)

        out.append(line)
        out.extend(options)
        out.extend(formatted.rstrip("\n").split("\n"))
        out.append(lines[end])
        index = end + 1

    return "\n".join(out)


def main(argv: list[str] | None = None) -> int:
    """Format (or check) the Python code blocks in the given `.qmd` files.

    Args:
        argv: Command-line arguments, defaulting to `sys.argv[1:]`.

    Returns:
        `0` when every file is formatted, `1` when `--check` finds a file that
        would change.
    """
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--check",
        action="store_true",
        help="report files that need formatting instead of rewriting them",
    )
    parser.add_argument(
        "paths",
        nargs="*",
        type=Path,
        help="`.qmd` files to format (default: every .qmd under examples/)",
    )
    args = parser.parse_args(argv)

    paths = args.paths or sorted(Path("examples").rglob("*.qmd"))
    changed: list[Path] = []
    for path in paths:
        original = path.read_text(encoding="utf-8")
        formatted = format_qmd_source(original, path.name)
        if formatted == original:
            continue
        changed.append(path)
        if not args.check:
            path.write_text(formatted, encoding="utf-8")

    for path in changed:
        print(f"{'would reformat' if args.check else 'reformatted'}: {path}")

    return 1 if (args.check and changed) else 0


if __name__ == "__main__":
    raise SystemExit(main())
