"""Build the self-contained runtime modules of a generated package.

The generated package must run without `excel_grapher` installed. That takes
two complementary modules:

- `excel.py` embeds shared Excel value semantics and the inverted-tree
  adapters (`xl_add`, `XlError`, `as_measure`, …) that generated internals
  call.
- `runtime.py` is named-axis primitives (`view`, `span`, `take`, `publish`,
  readers) with imports rewritten to `.excel` and `.tensor`.
"""

from __future__ import annotations

import ast
from pathlib import Path

from excel_grapher.exporter.embed import _core_modules, emit_runtime

_PACKAGE_ROOT = Path(__file__).resolve().parents[2]
_EXPORT_RUNTIME = _PACKAGE_ROOT / "exporter" / "export_runtime"
# Later modules override earlier symbols: the core value semantics win over
# the export aliases they are aligned with, and the export wrappers the
# adapters import by name register last.
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
_INVERTED_EXCEL_MODULE = "excel_grapher.exporter.inverted_tree.excel"
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


def _rewrite_imports(source: str, imports: list[ast.ImportFrom]) -> str:
    lines = source.splitlines(keepends=True)
    replacements: dict[int, tuple[int, str]] = {}
    for node in imports:
        if node.module == _INVERTED_EXCEL_MODULE:
            target = ".excel"
        elif node.module == _TENSOR_MODULE:
            target = ".tensor"
        else:
            target = ".excel"
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


def _top_level_def_names(tree: ast.Module) -> set[str]:
    names: set[str] = set()
    for node in tree.body:
        if isinstance(node, ast.FunctionDef | ast.AsyncFunctionDef | ast.ClassDef):
            names.add(node.name)
        elif isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name):
                    names.add(target.id)
        elif isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name):
            names.add(node.target.id)
    return names


def _collision_renames(adapter_tree: ast.Module) -> dict[str, str]:
    """Map embedded public names to adapter aliases when the adapter redefines them."""
    defs = _top_level_def_names(adapter_tree)
    mapping: dict[str, str] = {}
    for node in _shared_imports(adapter_tree):
        if node.module == _TENSOR_MODULE:
            continue
        for alias in node.names:
            if alias.asname and alias.asname != alias.name and alias.name in defs:
                mapping[alias.name] = alias.asname
    return mapping


def _passthrough_aliases(adapter_tree: ast.Module, renamed: dict[str, str]) -> list[str]:
    """Assign import aliases that the adapter uses without redefining the original."""
    defs = _top_level_def_names(adapter_tree)
    lines: list[str] = []
    seen: set[str] = set()
    for node in _shared_imports(adapter_tree):
        if node.module == _TENSOR_MODULE:
            continue
        for alias in node.names:
            if not alias.asname or alias.asname == alias.name:
                continue
            if alias.name in renamed or alias.name in defs:
                continue
            if alias.asname in seen:
                continue
            seen.add(alias.asname)
            lines.append(f"{alias.asname} = {alias.name}")
    return lines


class _RenameIdentifiers(ast.NodeTransformer):
    """Rename colliding identifiers in the embedded core so adapters can wrap them."""

    def __init__(self, mapping: dict[str, str]) -> None:
        self.mapping = mapping

    def visit_Name(self, node: ast.Name) -> ast.Name:
        if node.id in self.mapping:
            return ast.copy_location(ast.Name(id=self.mapping[node.id], ctx=node.ctx), node)
        return node

    def visit_FunctionDef(self, node: ast.FunctionDef) -> ast.AST:
        self.generic_visit(node)
        if node.name in self.mapping:
            node.name = self.mapping[node.name]
        return node

    def visit_AsyncFunctionDef(self, node: ast.AsyncFunctionDef) -> ast.AST:
        self.generic_visit(node)
        if node.name in self.mapping:
            node.name = self.mapping[node.name]
        return node

    def visit_ClassDef(self, node: ast.ClassDef) -> ast.AST:
        self.generic_visit(node)
        if node.name in self.mapping:
            node.name = self.mapping[node.name]
        return node


def _rename_identifiers(source: str, mapping: dict[str, str]) -> str:
    if not mapping:
        return source
    tree = ast.parse(source)
    renamed = _RenameIdentifiers(mapping).visit(tree)
    ast.fix_missing_locations(renamed)
    return ast.unparse(renamed) + "\n"


def _adapter_body(adapter_source: str) -> str:
    """Return adapter source with library imports and duplicate `T` removed."""
    tree = ast.parse(adapter_source)
    lines = adapter_source.splitlines(keepends=True)
    skip: set[int] = set()
    for node in tree.body:
        drop = (
            isinstance(node, ast.ImportFrom)
            and node.module is not None
            and (node.module == "__future__" or node.module.startswith("excel_grapher."))
        ) or (
            isinstance(node, ast.Assign)
            and len(node.targets) == 1
            and isinstance(node.targets[0], ast.Name)
            and node.targets[0].id == "T"
        )
        if drop:
            assert node.end_lineno is not None
            skip.update(range(node.lineno, node.end_lineno + 1))
    kept = [line for number, line in enumerate(lines, start=1) if number not in skip]
    return "".join(kept).strip() + "\n"


def _patch_column_letter(excel: str) -> str:
    if "fastpyxl.utils.cell" not in excel:
        return excel
    excel = excel.replace("import fastpyxl.utils.cell\n", "")
    excel = excel.replace("fastpyxl.utils.cell.get_column_letter(", "_column_letter(")
    marker = "\n\n\n"
    head, _, tail = excel.partition(marker)
    return head + "\n" + _COLUMN_LETTER + "\n" + tail


def build_runtime_modules(runtime_source: str, excel_source: str) -> dict[str, str]:
    """Return `runtime.py` and `excel.py` sources with no `excel_grapher` imports."""
    adapter_tree = ast.parse(excel_source)
    adapter_imports = [
        node for node in _shared_imports(adapter_tree) if node.module != _TENSOR_MODULE
    ]
    required = {alias.name for node in adapter_imports for alias in node.names} | set(
        _ALWAYS_REQUIRED
    )
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
    excel = _patch_column_letter(excel)
    renamed = _collision_renames(adapter_tree)
    excel = _rename_identifiers(excel, renamed)
    aliases = _passthrough_aliases(adapter_tree, renamed)
    if aliases:
        excel = excel.rstrip() + "\n\n" + "\n".join(aliases) + "\n"
    excel = excel.rstrip() + "\n\n" + _adapter_body(excel_source)
    runtime_tree = ast.parse(runtime_source)
    runtime = _rewrite_imports(runtime_source, _shared_imports(runtime_tree))
    return {"runtime.py": runtime.rstrip() + "\n", "excel.py": excel.rstrip() + "\n"}
