"""
Export dependency graphs to standalone Python packages.

Canonical implementation: :class:`~excel_grapher.exporter.codegen.CodeGenerator`.
The shared runtime embedded in generated code lives at :mod:`excel_grapher.runtime`.
"""

from .codegen import CodeGenerator

__all__ = ["CodeGenerator"]
