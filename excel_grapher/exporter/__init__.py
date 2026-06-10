"""Export dependency graphs to standalone Python packages.

Canonical implementation: `excel_grapher.exporter.codegen.CodeGenerator`.
The shared runtime embedded in generated code lives at `excel_grapher.runtime`.
"""

from excel_grapher.series_bindings.docstring_renderers import (
    GoogleSeriesDocstringRenderer,
    NumpySeriesDocstringRenderer,
    PlainSeriesDocstringRenderer,
    RstSeriesDocstringRenderer,
    SeriesDocstringRenderCallable,
    SeriesDocstringRenderer,
    SeriesDocstringRendererName,
    SeriesDocstringRendererSpec,
    resolve_series_docstring_renderer,
)
from excel_grapher.series_bindings.docstrings import (
    FieldDoc,
    SeriesBindingDocstringCallback,
    SeriesBindingDocstringCallbackSpec,
    SeriesBindingDocstringContext,
    SeriesBindingDocstringContract,
    SeriesFunctionDoc,
    list_series_docstring_callbacks,
    register_series_docstring_callback,
    resolve_series_docstring_callback,
)

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
    "FieldDoc",
    "GoogleSeriesDocstringRenderer",
    "NumpySeriesDocstringRenderer",
    "PlainSeriesDocstringRenderer",
    "RstSeriesDocstringRenderer",
    "SeriesDocstringRenderCallable",
    "SeriesDocstringRenderer",
    "SeriesDocstringRendererName",
    "SeriesDocstringRendererSpec",
    "SeriesBindingDocstringCallback",
    "SeriesBindingDocstringCallbackSpec",
    "SeriesBindingDocstringContext",
    "SeriesBindingDocstringContract",
    "SeriesFunctionDoc",
    "LAYOUT_STRATIFIED_MULTIPARTITE",
    "LAYOUT_SPRING",
    "LAYOUT_FORCEATLAS2",
    "LAYOUT_MULTIPARTITE",
    "LAYOUT_GRAPHVIZ_DOT",
    "LAYOUT_GRAPHVIZ_SFDP",
    "list_web_viz_layouts",
    "register_web_viz_layout",
    "list_series_docstring_callbacks",
    "register_series_docstring_callback",
    "resolve_series_docstring_callback",
    "resolve_series_docstring_renderer",
]
