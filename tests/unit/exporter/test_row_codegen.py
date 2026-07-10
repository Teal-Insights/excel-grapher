"""Codegen emission for Option B row helpers + member wrappers (issue #377 sprint 4)."""

from __future__ import annotations

import re

from excel_grapher.evaluator.name_utils import address_to_python_name, row_key_to_helper_name
from excel_grapher.exporter.codegen import CodeGenerator
from tests.fixtures.row_nodes.option_b_stripe import (
    OPTION_B_ROW_KEY,
    build_option_b_product_graph,
    build_option_b_stripe_fixture,
)


def test_row_key_to_helper_name() -> None:
    assert row_key_to_helper_name(OPTION_B_ROW_KEY) == "_row_sheet1_d63_e63"


def test_codegen_emits_one_row_helper() -> None:
    g = build_option_b_product_graph()
    code = CodeGenerator(g).generate(["Sheet1!D63", "Sheet1!E63"])
    helper = row_key_to_helper_name(OPTION_B_ROW_KEY)
    defs = re.findall(rf"^def ({re.escape(helper)})\(", code, flags=re.M)
    assert defs == [helper]
    assert f"def {helper}(ctx, *, column: str):" in code
    assert 'f"Sheet1!{column}35"' in code or "f'Sheet1!{column}35'" in code


def test_codegen_wrappers_call_helper_with_column() -> None:
    g = build_option_b_product_graph()
    code = CodeGenerator(g).generate(["Sheet1!D63", "Sheet1!E63"])
    helper = row_key_to_helper_name(OPTION_B_ROW_KEY)
    d_name = address_to_python_name("Sheet1!D63")
    e_name = address_to_python_name("Sheet1!E63")
    assert f"def {d_name}(ctx):" in code
    assert f"def {e_name}(ctx):" in code
    assert f"{helper}(ctx, column='D')" in code or f'{helper}(ctx, column="D")' in code
    assert f"{helper}(ctx, column='E')" in code or f'{helper}(ctx, column="E")' in code


def test_codegen_does_not_emit_cell_fn_for_row_key() -> None:
    g = build_option_b_product_graph()
    code = CodeGenerator(g).generate(["Sheet1!D63", "Sheet1!E63"])
    row_cell_name = address_to_python_name(OPTION_B_ROW_KEY)
    assert f"def {row_cell_name}(ctx):" not in code


def test_codegen_exported_member_matches_evaluator() -> None:
    fixture = build_option_b_stripe_fixture()
    code = CodeGenerator(fixture.option_b).generate(list(fixture.member_keys))
    namespace: dict[str, object] = {}
    exec(code, namespace)
    make_context = namespace["make_context"]
    resolve = namespace["_resolve_formula"]
    ctx = make_context()  # type: ignore[operator]
    for member, expected in (("Sheet1!D63", 6), ("Sheet1!E63", 10)):
        fn = resolve(member)  # type: ignore[operator]
        assert fn is not None
        assert fn(ctx) == expected
