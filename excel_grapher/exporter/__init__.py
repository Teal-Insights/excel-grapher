"""
Export dependency graphs to standalone Python packages.

Canonical implementation: :class:`~excel_grapher.exporter.codegen.CodeGenerator`.
The shared runtime embedded in generated code lives at :mod:`excel_grapher.runtime`.
"""

from .codegen import CodeGenerator
from .lightweight_viz import (
    WebVizPayload,
    to_web_viz_payload,
)
from .web_viz_layout import (
    LAYOUT_FORCEATLAS2,
    LAYOUT_GRAPHVIZ_DOT,
    LAYOUT_GRAPHVIZ_SFDP,
    LAYOUT_MULTIPARTITE,
    LAYOUT_SPRING,
    LAYOUT_STRATIFIED_MULTIPARTITE,
    list_web_viz_layouts,
    register_web_viz_layout,
)

__all__ = [
    "CodeGenerator",
    "WebVizPayload",
    "to_web_viz_payload",
    "LAYOUT_STRATIFIED_MULTIPARTITE",
    "LAYOUT_SPRING",
    "LAYOUT_FORCEATLAS2",
    "LAYOUT_MULTIPARTITE",
    "LAYOUT_GRAPHVIZ_DOT",
    "LAYOUT_GRAPHVIZ_SFDP",
    "list_web_viz_layouts",
    "register_web_viz_layout",
]
