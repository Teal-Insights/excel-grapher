from __future__ import annotations

import ast
from collections import deque
from pathlib import Path

_RUNTIME_DIR = Path(__file__).resolve().parent
_CORE_DIR = _RUNTIME_DIR.parent.parent / "core"

# Core package modules define types, coercions, scalar operators, and addressing (canonical source).
_CORE_MODULES: list[tuple[str, Path]] = [
    ("core.types", _CORE_DIR / "types.py"),
    ("core.coercions", _CORE_DIR / "coercions.py"),
    ("core.operators", _CORE_DIR / "operators.py"),
    ("core.addressing", _CORE_DIR / "addressing.py"),
]

# Export runtime modules (representation-specific or re-exports); order preserved for iteration.
_RUNTIME_MODULES: list[tuple[str, Path]] = [
    ("core", _RUNTIME_DIR / "core.py"),
    ("operators", _RUNTIME_DIR / "operators.py"),
    ("math", _RUNTIME_DIR / "math.py"),
    ("text", _RUNTIME_DIR / "text.py"),
    ("info", _RUNTIME_DIR / "info.py"),
    ("logic", _RUNTIME_DIR / "logic.py"),
    ("lookup", _RUNTIME_DIR / "lookup.py"),
    ("reference", _RUNTIME_DIR / "reference.py"),
    ("offset_runtime", _RUNTIME_DIR / "offset_runtime.py"),
    ("cache", _RUNTIME_DIR / "cache.py"),
]

# All modules: core first so their definitions win when symbols are defined in both.
_ALL_MODULES: list[tuple[str, Path]] = _CORE_MODULES + _RUNTIME_MODULES
_ALL_MODULE_NAMES: list[str] = [name for name, _ in _ALL_MODULES]

# Top-level names that are stdlib so emitted "import X" order satisfies ruff isort (I001).
_ISORT_STDLIB: frozenset[str] = frozenset(
    {"collections", "dataclasses", "enum", "typing", "warnings"}
)


class _RuntimeNameCollector(ast.NodeVisitor):
    """Collect runtime-relevant Name identifiers, ignoring type annotations."""

    def __init__(self) -> None:
        self.names: set[str] = set()

    def visit_Name(self, node: ast.Name) -> None:  # noqa: N802
        self.names.add(node.id)

    def visit_arg(self, node: ast.arg) -> None:  # noqa: N802
        return

    def visit_AnnAssign(self, node: ast.AnnAssign) -> None:  # noqa: N802
        if node.value is not None:
            self.visit(node.value)

    def visit_FunctionDef(self, node: ast.FunctionDef) -> None:  # noqa: N802
        for deco in node.decorator_list:
            self.visit(deco)
        for d in node.args.defaults:
            self.visit(d)
        for d in node.args.kw_defaults:
            if d is not None:
                self.visit(d)
        for stmt in node.body:
            self.visit(stmt)

    def visit_AsyncFunctionDef(self, node: ast.AsyncFunctionDef) -> None:  # noqa: N802
        return self.visit_FunctionDef(node)  # type: ignore[arg-type]

    def visit_ClassDef(self, node: ast.ClassDef) -> None:  # noqa: N802
        for base in node.bases:
            self.visit(base)
        for kw in node.keywords:
            self.visit(kw)
        for deco in node.decorator_list:
            self.visit(deco)
        for stmt in node.body:
            self.visit(stmt)


def _referenced_names(node: ast.AST) -> set[str]:
    collector = _RuntimeNameCollector()
    collector.visit(node)
    return collector.names


def _top_level_defs(module: ast.Module) -> dict[str, ast.AST]:
    out: dict[str, ast.AST] = {}
    for node in module.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            out[node.name] = node
        elif isinstance(node, ast.Assign):
            for t in node.targets:
                if isinstance(t, ast.Name):
                    out[t.id] = node
        elif isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name):
            out[node.target.id] = node
    return out


def _extract_source_segment(src: str, node: ast.AST) -> str:
    # For decorated defs/classes, ast.get_source_segment() may start at the "def"/"class"
    # line and omit leading decorators. Since the generated runtime must preserve
    # decorators (e.g. @dataclass), slice by line numbers instead.
    if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
        start = node.lineno
        decorators = getattr(node, "decorator_list", None) or []
        if decorators:
            start = min(start, *(d.lineno for d in decorators if hasattr(d, "lineno")))
        end = getattr(node, "end_lineno", None)
        if end is not None:
            lines = src.splitlines()
            return "\n".join(lines[start - 1 : end]).rstrip()

    seg = ast.get_source_segment(src, node)
    if seg is None:
        raise ValueError("Could not extract source segment for node")
    return seg.rstrip()


def _collect_external_import_lines(module: ast.Module, src: str) -> list[str]:
    lines: list[str] = []
    for node in module.body:
        if isinstance(node, ast.Import):
            seg = ast.get_source_segment(src, node)
            if seg:
                lines.append(seg.rstrip())
        elif isinstance(node, ast.ImportFrom):
            # Skip relative imports (from .foo import bar)
            if node.level and node.level > 0:
                continue
            # The generated output always includes the __future__ annotations import.
            if node.module == "__future__":
                continue
            seg = ast.get_source_segment(src, node)
            if seg:
                lines.append(seg.rstrip())
    return lines


def _consolidate_import_lines(import_lines: list[str]) -> list[str]:
    """Consolidate compatible import statements to avoid duplicates.

    The runtime is emitted as a single module, so repeated imports such as:

    - from collections.abc import Callable
    - from collections.abc import Iterable, Iterator

    can be merged into one statement. This keeps the generated output cleaner
    and prevents redefinition lint errors (e.g. ruff F811).
    """

    # Keep any unparsable/unsupported lines in their original order.
    passthrough: list[str] = []

    # Consolidate "import ..." statements by alias tuple.
    # Example key: ("numpy", "np") for "import numpy as np"
    import_aliases: dict[tuple[str, str | None], None] = {}

    # Consolidate "from X import ..." statements by (module, level).
    # Example key: ("collections.abc", 0)
    from_aliases: dict[tuple[str | None, int], dict[str, str | None]] = {}

    for line in import_lines:
        try:
            mod = ast.parse(line)
        except SyntaxError:
            passthrough.append(line)
            continue

        if len(mod.body) != 1:
            passthrough.append(line)
            continue

        stmt = mod.body[0]
        if isinstance(stmt, ast.Import):
            for alias in stmt.names:
                import_aliases[(alias.name, alias.asname)] = None
            continue

        if isinstance(stmt, ast.ImportFrom):
            key = (stmt.module, stmt.level or 0)
            bucket = from_aliases.setdefault(key, {})
            for alias in stmt.names:
                # Keep the first asname we see for a given imported symbol.
                bucket.setdefault(alias.name, alias.asname)
            continue

        passthrough.append(line)

    out: list[str] = []

    # Preserve passthrough lines first (rare for this repo, but safe).
    out.extend(passthrough)

    # Emit in ruff isort (I001) order: stdlib "import", then stdlib "from", then third-party "import".
    def _import_sort_key(item: tuple[str, str | None]) -> tuple[int, str, str]:
        name, asname = item
        top = name.split(".", 1)[0]
        return (0 if top in _ISORT_STDLIB else 1, name, asname or "")

    import_sorted = sorted(import_aliases.keys(), key=_import_sort_key)
    stdlib_imports = [(n, a) for (n, a) in import_sorted if n.split(".", 1)[0] in _ISORT_STDLIB]
    third_party_imports = [(n, a) for (n, a) in import_sorted if n.split(".", 1)[0] not in _ISORT_STDLIB]

    for (name, asname) in stdlib_imports:
        out.append(f"import {name} as {asname}" if asname else f"import {name}")
    for (module, level), names in sorted(from_aliases.items(), key=lambda kv: (kv[0][1], kv[0][0] or "")):
        if level != 0 or module is None:
            continue
        parts = [f"{n} as {a}" if a else n for n, a in sorted(names.items(), key=lambda x: x[0])]
        out.append(f"from {module} import {', '.join(parts)}")
    if third_party_imports:
        out.append("")
    for (name, asname) in third_party_imports:
        out.append(f"import {name} as {asname}" if asname else f"import {name}")

    return out


class _AllNameCollector(ast.NodeVisitor):
    """Collect all Name identifiers, including those in annotations.

    The generated runtime uses `from __future__ import annotations`, so many
    imported typing symbols are only referenced for static analysis. We still
    treat those as "used" so that type-checkers can resolve them, while pruning
    truly-unused imports to keep `ruff check` clean.
    """

    def __init__(self) -> None:
        self.names: set[str] = set()

    def visit_Name(self, node: ast.Name) -> None:  # noqa: N802
        self.names.add(node.id)


def _binding_name_for_import(alias: ast.alias) -> str:
    if alias.asname:
        return alias.asname
    # `import fastpyxl.utils.cell` binds `fastpyxl`
    return alias.name.split(".", 1)[0]


def _prune_import_lines(import_lines: list[str], *, used_names: set[str]) -> list[str]:
    """Drop imported symbols that aren't referenced by the emitted runtime."""
    out: list[str] = []
    for line in import_lines:
        try:
            mod = ast.parse(line)
        except SyntaxError:
            out.append(line)
            continue

        if len(mod.body) != 1:
            out.append(line)
            continue

        stmt = mod.body[0]
        if isinstance(stmt, ast.Import):
            kept = [
                alias
                for alias in stmt.names
                if _binding_name_for_import(alias) in used_names
            ]
            for alias in kept:
                out.append(
                    f"import {alias.name} as {alias.asname}"
                    if alias.asname
                    else f"import {alias.name}"
                )
            continue

        if isinstance(stmt, ast.ImportFrom):
            # Preserve any relative imports verbatim (shouldn't appear here).
            if stmt.level and stmt.level > 0:
                out.append(line)
                continue
            if stmt.module is None:
                out.append(line)
                continue

            kept: list[str] = []
            for alias in stmt.names:
                binding = alias.asname or alias.name
                if binding in used_names:
                    kept.append(
                        f"{alias.name} as {alias.asname}" if alias.asname else alias.name
                    )
            if kept:
                out.append(f"from {stmt.module} import {', '.join(sorted(kept))}")
            continue

        out.append(line)

    # De-dupe while preserving order (pruning can reintroduce duplicates).
    deduped: list[str] = []
    seen: set[str] = set()
    for line in out:
        if line in seen:
            continue
        seen.add(line)
        deduped.append(line)
    return deduped


def emit_runtime(required_symbols: set[str], *, include_offset_table: bool) -> str:
    """Emit standalone runtime code for generated output.

    This uses AST-based extraction from curated runtime modules and core,
    including only the requested symbols (and their transitive runtime dependencies).
    """
    # Parse all runtime and core modules.
    module_src: dict[str, str] = {}
    module_ast: dict[str, ast.Module] = {}
    defs_by_module: dict[str, dict[str, ast.AST]] = {}
    imports_by_module: dict[str, list[str]] = {}

    for mod_name, mod_path in _ALL_MODULES:
        src = mod_path.read_text(encoding="utf-8")
        module_src[mod_name] = src
        mod_ast = ast.parse(src, filename=str(mod_path))
        module_ast[mod_name] = mod_ast
        defs_by_module[mod_name] = _top_level_defs(mod_ast)
        imports_by_module[mod_name] = _collect_external_import_lines(mod_ast, src)

    symbol_to_node: dict[str, ast.AST] = {}
    symbol_to_module: dict[str, str] = {}
    for mod, defs in defs_by_module.items():
        for name, node in defs.items():
            symbol_to_node[name] = node
            symbol_to_module[name] = mod

    # Dependency graph between runtime symbols.
    symbol_deps: dict[str, set[str]] = {}
    for name, node in symbol_to_node.items():
        refs = _referenced_names(node)
        symbol_deps[name] = {r for r in refs if r in symbol_to_node and r != name}

    seed = set(required_symbols) | {"XlError", "ExcelRange", "CellValue"}

    # Close over symbol dependencies.
    needed: set[str] = set()
    q: deque[str] = deque(seed)
    while q:
        s = q.popleft()
        if s in needed:
            continue
        needed.add(s)
        for dep in symbol_deps.get(s, set()):
            if dep not in needed:
                q.append(dep)

    # Imports: union external imports from modules that contribute symbols.
    used_modules = {symbol_to_module[s] for s in needed if s in symbol_to_module}
    import_lines: list[str] = []
    for mod in _ALL_MODULE_NAMES:
        if mod not in used_modules:
            continue
        for line in imports_by_module[mod]:
            if line not in import_lines:
                import_lines.append(line)
    import_lines = _consolidate_import_lines(import_lines)

    # Prune unused imported names for the subset runtime we emit (keeps ruff happy).
    used_names: set[str] = set()
    collector = _AllNameCollector()
    for s in needed:
        node = symbol_to_node.get(s)
        if node is None:
            continue
        collector.visit(node)
    used_names = collector.names
    import_lines = _prune_import_lines(import_lines, used_names=used_names)

    # Topo order symbols.
    ordered: list[str] = []
    remaining = {s for s in needed if s in symbol_to_node}
    while remaining:
        progressed = False
        for s in sorted(remaining):
            deps = symbol_deps.get(s, set())
            if deps.issubset(set(ordered)):
                ordered.append(s)
                remaining.remove(s)
                progressed = True
                break
        if not progressed:
            raise ValueError(f"Runtime symbol dependency cycle: {sorted(remaining)}")

    out: list[str] = [
        '"""Standalone runtime for generated Excel formula code."""',
        "",
        "from __future__ import annotations",
        "",
    ]
    out.extend(import_lines)
    if import_lines:
        out.append("")

    for s in ordered:
        mod = symbol_to_module[s]
        node = symbol_to_node[s]
        out.append(_extract_source_segment(module_src[mod], node))
        out.append("")

    return "\n".join(out).rstrip()

