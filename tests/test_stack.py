import math

import pytest

from kns_plan.presets import default_stack, example_stack, save_default_stack
from kns_plan.schema import CostLine, CostStack, LineType
from kns_plan.stack import coefficient, evaluate, implied_fob, linear_terms, retail_price_idr, round_up
from kns_plan.validate import check_stack


def test_example_stack_price_chain_by_hand():
    # 100 × 1.10 freight × 1.20 duties = 132 landed; ÷ (1 − (32% + 4% + 1% + 8%)) = 240 wholesale;
    # ÷ (1 − 35%) = 369.23 before VAT; × 1.11 = 409.85 shelf
    r = evaluate(example_stack(), 100, 18200).values
    assert r["landed"] == pytest.approx(132)
    assert r["wholesale"] == pytest.approx(240)
    assert r["distributor"] == pytest.approx(0.32 * 240)
    assert r["before_vat"] == pytest.approx(240 / 0.65)
    assert r["shelf"] == pytest.approx(240 / 0.65 * 1.11)


def test_coefficient_retail_price_and_brand_price_to_ask_for():
    stack = example_stack()
    assert coefficient(stack, 18200) == pytest.approx(4.098462, abs=1e-6)
    # 8 × 4.098462 × 18,200 = 596,736 → rounded up to 597,000
    assert retail_price_idr(stack, 8, 18200, 1000) == 597_000
    # 650,000 ÷ 18,200 ÷ 4.098462 = 8.714071
    assert implied_fob(stack, 650_000, 18200) == pytest.approx(8.714071, abs=1e-6)


def test_company_default_stack_comes_from_the_data_folder(tmp_path):
    assert default_stack(data_dir=tmp_path).lines[0].rate == 0.10  # no file: example rates
    custom = example_stack()
    custom.lines[0].rate = 0.17
    save_default_stack(custom, tmp_path)
    loaded = default_stack("s2", "Second", data_dir=tmp_path)
    assert (loaded.id, loaded.name, loaded.lines[0].rate) == ("s2", "Second", 0.17)


def test_fixed_amounts_make_the_inverse_affine_and_still_exact():
    stack = CostStack(id="s", name="s", lines=[
        CostLine(id="fee", label="fee", type=LineType.FIXED, amount=3640, amount_currency="idr"),
        CostLine(id="m", label="m", type=LineType.PCT_OF_RESULT, rate=0.3),
        CostLine(id="shelf", label="shelf", type=LineType.SUBTOTAL),
    ])
    a, b = linear_terms(stack, 18200)
    assert a == pytest.approx(1 / 0.7)
    assert b == pytest.approx(0.2 / 0.7)
    fob = implied_fob(stack, 500_000, 18200)
    assert evaluate(stack, fob, 18200).shelf * 18200 == pytest.approx(500_000)


def test_round_up_ignores_float_noise():
    assert round_up(1_186_000.0000000002, 1000) == 1_186_000
    assert round_up(1_185_001, 1000) == 1_186_000
    assert round_up(123.4, 0) == 123.4


def test_validation_catches_bad_stacks():
    no_shelf = CostStack(id="s", name="s", lines=[
        CostLine(id="f", label="f", type=LineType.PCT_OF_BASE, rate=0.1, base="fob")])
    over_100 = CostStack(id="s", name="s", lines=[
        CostLine(id="m1", label="m1", type=LineType.PCT_OF_RESULT, rate=0.6),
        CostLine(id="m2", label="m2", type=LineType.PCT_OF_RESULT, rate=0.4),
        CostLine(id="shelf", label="shelf", type=LineType.SUBTOTAL)])
    later_base = CostStack(id="s", name="s", lines=[
        CostLine(id="f", label="f", type=LineType.PCT_OF_BASE, rate=0.1, base="shelf"),
        CostLine(id="shelf", label="shelf", type=LineType.SUBTOTAL)])
    for stack, text in ((no_shelf, "must end"), (over_100, "under 100%"), (later_base, "earlier line")):
        issues = check_stack(stack)
        assert any(text in i.message for i in issues), (text, issues)
    assert check_stack(example_stack()) == []
    assert not math.isnan(evaluate(example_stack(), 1, 1).values["anp"])
