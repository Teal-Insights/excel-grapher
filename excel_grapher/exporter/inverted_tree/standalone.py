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
_PROVENANCE_MODULE = "excel_grapher.exporter.export_runtime.provenance"
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
        elif node.module == _PROVENANCE_MODULE:
            target = ".provenance"
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


def _adapter_stdlib_imports(adapter_tree: ast.Module) -> list[ast.ImportFrom]:
    """Return adapter `from … import` nodes that must live in the excel header."""
    imports: list[ast.ImportFrom] = []
    for node in adapter_tree.body:
        if not isinstance(node, ast.ImportFrom) or node.module is None:
            continue
        if node.module == "__future__" or node.module.startswith("excel_grapher."):
            continue
        imports.append(node)
    return imports


def _merge_stdlib_imports(excel: str, extra: list[ast.ImportFrom]) -> str:
    """Add adapter stdlib names to the embed header without mid-file imports."""
    tree = ast.parse(excel)
    bound: set[str] = set()
    last_import_end = 0
    existing_from: dict[str, ast.ImportFrom] = {}
    for node in tree.body:
        if isinstance(node, ast.Expr) and isinstance(node.value, ast.Constant):
            continue
        if isinstance(node, ast.ImportFrom) and node.module == "__future__":
            last_import_end = node.end_lineno or node.lineno
            continue
        if isinstance(node, ast.ImportFrom):
            last_import_end = node.end_lineno or node.lineno
            if node.module is not None:
                existing_from[node.module] = node
            for alias in node.names:
                bound.add(alias.asname or alias.name)
            continue
        if isinstance(node, ast.Import):
            last_import_end = node.end_lineno or node.lineno
            for alias in node.names:
                bound.add((alias.asname or alias.name).split(".", 1)[0])
            continue
        break
    missing_by_module: dict[str, list[str]] = {}
    for node in extra:
        assert node.module is not None
        for alias in node.names:
            name = alias.asname or alias.name
            if name in bound:
                continue
            bound.add(name)
            part = f"{alias.name} as {alias.asname}" if alias.asname else alias.name
            missing_by_module.setdefault(node.module, []).append(part)
    if not missing_by_module:
        return excel
    lines = excel.splitlines(keepends=True)
    insertions: list[str] = []
    for module, parts in missing_by_module.items():
        existing = existing_from.get(module)
        if (
            existing is not None
            and existing.lineno is not None
            and existing.lineno == existing.end_lineno
        ):
            line = lines[existing.lineno - 1]
            stripped = line.rstrip("\n")
            newline = line[len(stripped) :]
            lines[existing.lineno - 1] = f"{stripped}, {', '.join(parts)}{newline}"
            continue
        insertions.append(f"from {module} import {', '.join(parts)}\n")
    for insert in insertions:
        lines.insert(last_import_end, insert)
        last_import_end += 1
    return "".join(lines)


def _adapter_body(adapter_source: str) -> str:
    """Return adapter source with imports, module docstring, and duplicate `T` removed."""
    tree = ast.parse(adapter_source)
    lines = adapter_source.splitlines(keepends=True)
    skip: set[int] = set()
    for index, node in enumerate(tree.body):
        drop = isinstance(node, ast.Import | ast.ImportFrom) or (
            isinstance(node, ast.Assign)
            and len(node.targets) == 1
            and isinstance(node.targets[0], ast.Name)
            and node.targets[0].id == "T"
        )
        if index == 0 and isinstance(node, ast.Expr) and isinstance(node.value, ast.Constant):
            drop = isinstance(node.value.value, str)
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
        modules=_MODULES,
    )
    excel = excel.replace(
        '"""Standalone runtime for generated Excel formula code."""',
        '"""Excel value semantics shared by the generated model functions."""',
    )
    excel = _patch_column_letter(excel)
    renamed = _collision_renames(adapter_tree)
    excel = _rename_identifiers(excel, renamed)
    excel = _merge_stdlib_imports(excel, _adapter_stdlib_imports(adapter_tree))
    aliases = _passthrough_aliases(adapter_tree, renamed)
    if aliases:
        excel = excel.rstrip() + "\n\n" + "\n".join(aliases) + "\n"
    excel = excel.rstrip() + "\n\n" + _adapter_body(excel_source)
    runtime_tree = ast.parse(runtime_source)
    runtime = _rewrite_imports(runtime_source, _shared_imports(runtime_tree))
    return {"runtime.py": runtime.rstrip() + "\n", "excel.py": excel.rstrip() + "\n"}
