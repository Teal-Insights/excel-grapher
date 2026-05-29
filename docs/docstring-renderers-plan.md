# Implementation plan: series docstring renderers

GitHub issue: [#212 Add docstring renderers for different documentation/markup systems](https://github.com/Teal-Insights/excel-grapher/issues/212)

## Goal

Generated series-binding functions should support multiple docstring formatting conventions while keeping docstring content structured and renderer-independent.

The existing callback mechanism should continue to own the content-generation step:

- callbacks receive `SeriesBindingDocstringContext`
- callbacks return `SeriesFunctionDoc | None`
- codegen emits the rendered text as the generated function docstring

The new renderer layer should own only the formatting step.

## Current architecture

Relevant modules:

- `excel_grapher/series_bindings/docstrings.py`
  - defines `SeriesFunctionDoc`, `SeriesBindingDocstringContract`, and callback registration
  - derives structured docstring contracts from series bindings
  - currently contains the single renderer function
- `excel_grapher/series_bindings/setter_codegen.py`
  - emits generated `set_*` functions
- `excel_grapher/series_bindings/compute_codegen.py`
  - emits generated output `compute_*` functions
- `excel_grapher/series_bindings/bindings_codegen.py`
  - combines setter and compute emission
- `excel_grapher/exporter/codegen.py`
  - exposes `CodeGenerator.generate(...)` and `CodeGenerator.generate_modules(...)`

The existing design already separates structured content from emitted source code. The missing layer is a pluggable renderer between `SeriesFunctionDoc` and the emitted docstring literal.

## Public API shape

Add a renderer protocol in `excel_grapher/series_bindings/docstrings.py`:

```python
class SeriesDocstringRenderer(Protocol):
    def render(
        self,
        contract: SeriesBindingDocstringContract,
        doc: SeriesFunctionDoc,
    ) -> str: ...
```

Add a renderer-name type:

```python
SeriesDocstringRendererName = Literal["plain", "rst", "google", "numpy"]
```

Thread a new keyword through codegen APIs:

```python
docstring_renderer: SeriesDocstringRendererName | SeriesDocstringRenderer = "plain"
```

This should be added to:

- `CodeGenerator.generate(...)`
- `CodeGenerator.generate_modules(...)`
- `CodeGenerator._emit_series_binding_setters(...)`
- `emit_series_bindings_block(...)`
- `emit_setters_block(...)`
- `emit_computes_block(...)`
- `emit_setter_function(...)`
- `emit_compute_function(...)`
- `resolve_series_function_docstring(...)`

The existing `series_docstring_callback` parameter should remain the content provider. The new `docstring_renderer` parameter should format the resulting `SeriesFunctionDoc`.

## Renderer resolution

Add a helper to resolve either a built-in renderer name or a custom renderer object:

```python
def resolve_series_docstring_renderer(
    renderer: SeriesDocstringRendererName | SeriesDocstringRenderer,
) -> SeriesDocstringRenderer:
    ...
```

Behavior:

- `"plain"`, `"rst"`, `"google"`, and `"numpy"` resolve to built-in renderer instances
- custom renderer objects are accepted when they satisfy the protocol
- unknown string names raise `ValueError` with the known renderer names

Because the project is greenfield, remove the current hard-coded rendering function instead of keeping a compatibility wrapper.

## Built-in renderers

Implement four renderers:

1. `PlainSeriesDocstringRenderer`
   - preserves the current plain-text structure
   - remains the default
2. `RstSeriesDocstringRenderer`
   - uses reStructuredText-friendly section labels and field markup
   - suitable for Sphinx-style documentation
3. `GoogleSeriesDocstringRenderer`
   - uses Google-style sections such as `Args:`, `Returns:`, and `Examples:`
4. `NumpySeriesDocstringRenderer`
   - uses NumPy-style underlined sections such as `Parameters`, `Returns`, and `Examples`

All renderers should include the same semantic information where available:

- summary
- purpose
- record matching guidance
- required record fields
- optional record fields
- source binding details
- example usage

## Default behavior

If `series_docstring_callback` is `None`, keep current behavior:

- use series notes when present
- otherwise emit the existing fallback summary, such as `Apply records for ...` or `Compute records for ...`

Renderers format structured `SeriesFunctionDoc` output. They should not affect the simple fallback docstring path unless a callback provides structured content.

## Test plan

Follow test-driven development for the implementation.

Unit tests in `tests/unit/series_bindings/test_docstrings.py`:

- built-in renderer names resolve successfully
- unknown renderer names raise a helpful `ValueError`
- custom renderer objects are accepted
- each built-in renderer includes summary, field sections, source binding details, and example usage
- `resolve_series_function_docstring(...)` applies the selected renderer to callback output
- fallback behavior is unchanged when no callback is supplied

Codegen tests:

- setter codegen applies a non-plain renderer
- output compute codegen applies a non-plain renderer
- `CodeGenerator.generate(...)` threads `docstring_renderer` through to generated series functions
- `CodeGenerator.generate_modules(...)` threads `docstring_renderer` through to generated series functions

Existing callback tests should be updated to call the new renderer path directly where needed rather than relying on the removed hard-coded render function.

## Public exports

Export the renderer API from:

- `excel_grapher.series_bindings`
- `excel_grapher.exporter`

Likely exports:

- `SeriesDocstringRenderer`
- `SeriesDocstringRendererName`
- `resolve_series_docstring_renderer`
- built-in renderer classes if they are intended as customization examples

## Non-goals

- Do not add a dependency for docstring rendering.
- Do not change callback semantics.
- Do not make renderer selection part of the series binding manifest yet.
- Do not apply renderer formatting to unstructured fallback notes unless structured doc content is available.
