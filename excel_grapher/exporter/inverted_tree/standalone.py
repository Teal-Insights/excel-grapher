"""Build the self-contained runtime modules of a generated package.

The generated package must run without `excel_grapher` installed. The
inverted-tree runtime imports its Excel semantics from the shared core; this
module embeds exactly those symbols into an `excel.py` module and rewrites the
runtime's imports to point at it.
"""

from __future__ import annotations

import ast
from pathlib import Path

from excel_grapher.exporter.embed import _core_modules, emit_runtime

_PACKAGE_ROOT = Path(__file__).resolve().parents[2]
_EXPORT_RUNTIME = _PACKAGE_ROOT / "exporter" / "export_runtime"
# Later modules override earlier symbols: the core value semantics win over
# the export aliases they are aligned with, and the export wrappers the
# runtime imports by name register last.
_MODULES = (
    ("export_runtime.ranges", _EXPORT_RUNTIME / "ranges.py"),
    ("export_runtime.values", _EXPORT_RUNTIME / "values.py"),
    ("export_runtime.errors", _EXPORT_RUNTIME / "errors.py"),
    *_core_modules(include_operators_fastpath=False),
    ("core.numpy_support", _PACKAGE_ROOT / "core" / "numpy_support.py"),
    ("export_runtime.error_funcs", _EXPORT_RUNTIME / "error_funcs.py"),
    ("export_runtime.lookup", _EXPORT_RUNTIME / "lookup.py"),
    ("export_runtime.math", _EXPORT_RUNTIME / "math.py"),
    ("export_runtime.text", _EXPORT_RUNTIME / "text.py"),
    ("series_bindings.input_coerce", _PACKAGE_ROOT / "series_bindings" / "input_coerce.py"),
)
# Names the shared core reaches only through annotations or guarded imports.
_ALWAYS_REQUIRED = frozenset({"T", "np"})
_TENSOR_MODULE = "excel_grapher.exporter.export_runtime.tensor"
_COLUMN_LETTER = '''
def _column_letter(index: int) -> str:
    """Return worksheet column letters for a 1-based column index."""
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters
'''


def _shared_imports(tree: ast.Module) -> list[ast.ImportFrom]:
    return [
        node
        for node in ast.walk(tree)
        if isinstance(node, ast.ImportFrom)
        and node.module is not None
        and node.module.startswith("excel_grapher.")
    ]


def _rewrite_imports(source: str, tree: ast.Module, imports: list[ast.ImportFrom]) -> str:
    lines = source.splitlines(keepends=True)
    replacements: dict[int, tuple[int, str]] = {}
    for node in imports:
        target = ".tensor" if node.module == _TENSOR_MODULE else ".excel"
        names = ", ".join(
            alias.name if alias.asname is None else f"{alias.name} as {alias.asname}"
            for alias in node.names
        )
        assert node.end_lineno is not None
        indent = " " * node.col_offset
        replacements[node.lineno] = (node.end_lineno, f"{indent}from {target} import {names}\n")
    rewritten: list[str] = []
    skip_until = 0
    for number, line in enumerate(lines, start=1):
        if number <= skip_until:
            continue
        if number in replacements:
            end, replacement = replacements[number]
            rewritten.append(replacement)
            skip_until = end
            continue
        rewritten.append(line)
    return "".join(rewritten)


def build_runtime_modules(runtime_source: str) -> dict[str, str]:
    """Return `runtime.py` and `excel.py` sources with no `excel_grapher` imports."""
    tree = ast.parse(runtime_source)
    imports = _shared_imports(tree)
    required = {
        alias.name for node in imports if node.module != _TENSOR_MODULE for alias in node.names
    } | set(_ALWAYS_REQUIRED)
    excel = emit_runtime(
        required,
        include_offset_table=False,
        include_dep_tracking=False,
        include_operators_fastpath=False,
        modules=_MODULES,
    )
    excel = excel.replace(
        '"""Standalone runtime for generated Excel formula code."""',
        '"""Excel value semantics shared by the generated model functions."""',
    )
    if "fastpyxl.utils.cell" in excel:
        excel = excel.replace("import fastpyxl.utils.cell\n", "")
        excel = excel.replace("fastpyxl.utils.cell.get_column_letter(", "_column_letter(")
        marker = "\n\n\n"
        head, _, tail = excel.partition(marker)
        excel = head + "\n" + _COLUMN_LETTER + "\n" + tail
    runtime = _rewrite_imports(runtime_source, tree, imports)
    return {"runtime.py": runtime.rstrip() + "\n", "excel.py": excel.rstrip() + "\n"}
