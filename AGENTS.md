Use the `gh` CLI tool to manage issues and pull requests.

Always run Python code with `uv run`. Use `uv add` to add dependencies or `uvx` to use uninstalled Python command line tools.

We use `fastpyxl` as a drop-in replacement for `openpyxl`.

Always practice test-driven development. Write a stub (if necessary), write a test, watch it fail for the right reason (RED), write the code to make it pass (GREEN), and then refactor to clean up the code.

## Parity

The project aims for **behavioral parity** across **Excel** (reference), **`FormulaEvaluator`**, and **exported standalone code**. Semantics are centralized in `excel_grapher/exporter/export_runtime/`; the evaluator and codegen must both use that runtime so **evaluator ↔ export** stays aligned.

**Excel-facing tests:** Prefer validating against **live Excel** (xlwings on Windows/macOS, or Excel via COM from WSL) when comparing to the real engine. **Run-if-available:** if automation is missing, **`pytest.skip`** with a clear reason—do not fail CI. Cache-based comparisons (`excel_workbook_parity`) remain useful for environments without Excel. See `.cursor/rules/parity.mdc` for the full contract.