"""Export dependency graphs to standalone Python packages.

Canonical implementation: `excel_grapher.exporter.codegen.CodeGenerator`.
Package export is `generate_modules()` and requires series bindings.
"""

from .codegen import CodeGenerator
from .lightweight_viz import (
    WebVizPayload,
    to_web_viz_payload,
)
from .projection import (
    BaseProjectionManifest,
    CollapsedGroup,
    CompositeProjectionManifest,
    FormulaRewrite,
    IdentityTransitCompression,
    OptimalCompression,
    ProjectedNodeSnapshot,
    ProjectionManifest,
    ProjectionResult,
    ProjectionStep,
    apply_projection,
    build_forwarding_projection_manifest,
    build_optimal_projection_manifest,
    project_optimal,
    register_projection_manifest,
    resolve_projection_manifest,
    unregister_projection_manifest,
)
from .semantic_catalog import (
    SemanticCatalogError,
    SemanticCatalogView,
    load_semantic_catalog,
)
from .semantic_graph import StatementGraph, build_statement_graph
from .semantic_viz import (
    SEMANTIC_VIZ_PAYLOAD_VERSION,
    SemanticVizPayload,
    to_semantic_viz_payload,
    write_semantic_viz_html,
)
from .series_graph import to_series_graph
from .web_viz_layout import (
    LAYOUT_FORCEATLAS2,
    LAYOUT_GRAPHVIZ_DOT,
    LAYOUT_GRAPHVIZ_SFDP,
    LAYOUT_MULTIPARTITE,
    LAYOUT_SPRING,
    LAYOUT_STRATIFIED_MULTIPARTITE,
    WebVizLayoutPlugin,
    WebVizLayoutSpec,
    list_web_viz_layouts,
    register_web_viz_layout,
    resolve_web_viz_layout,
    unregister_web_viz_layout,
)

__all__ = [
    "CodeGenerator",
    "WebVizPayload",
    "to_web_viz_payload",
    "SemanticCatalogError",
    "SemanticCatalogView",
    "SemanticVizPayload",
    "SEMANTIC_VIZ_PAYLOAD_VERSION",
    "StatementGraph",
    "build_statement_graph",
    "load_semantic_catalog",
    "to_semantic_viz_payload",
    "to_series_graph",
    "write_semantic_viz_html",
    "BaseProjectionManifest",
    "CollapsedGroup",
    "CompositeProjectionManifest",
    "FormulaRewrite",
    "IdentityTransitCompression",
    "OptimalCompression",
    "ProjectionManifest",
    "ProjectionResult",
    "ProjectionStep",
    "ProjectedNodeSnapshot",
    "apply_projection",
    "build_forwarding_projection_manifest",
    "build_optimal_projection_manifest",
    "project_optimal",
    "register_projection_manifest",
    "resolve_projection_manifest",
    "unregister_projection_manifest",
    "LAYOUT_STRATIFIED_MULTIPARTITE",
    "LAYOUT_SPRING",
    "LAYOUT_FORCEATLAS2",
    "LAYOUT_MULTIPARTITE",
    "LAYOUT_GRAPHVIZ_DOT",
    "LAYOUT_GRAPHVIZ_SFDP",
    "WebVizLayoutPlugin",
    "WebVizLayoutSpec",
    "list_web_viz_layouts",
    "register_web_viz_layout",
    "resolve_web_viz_layout",
    "unregister_web_viz_layout",
]
