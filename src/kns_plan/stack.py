"""Cost stack arithmetic.

A stack is evaluated top to bottom starting from the brand price (FOB/EXW):

* ``pct_of_base``   adds rate × base, where base is FOB, the running total, or an earlier subtotal
* ``fixed``         adds a fixed amount (IDR amounts are converted at the plan's FX rate)
* ``pct_of_result`` is a share of the *next* subtotal: subtotal = running / (1 − Σ those rates)
* ``subtotal``      closes a step; the last subtotal is the shelf price incl. VAT

Every line type is linear in FOB, so shelf = a × FOB + b exactly. That makes the brand price
for a target shelf price a closed-form inverse: FOB = (target − b) / a.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field

from .schema import BASE_FOB, BASE_RUNNING, CostStack, LineType


@dataclass
class StackResult:
    fob: float
    values: dict[str, float] = field(default_factory=dict)  # line id → amount in brand currency

    @property
    def shelf(self) -> float:
        return self.values[list(self.values)[-1]]


def evaluate(stack: CostStack, fob: float, fx_rate: float) -> StackResult:
    result = StackResult(fob=fob)
    running = fob
    pending: list[str] = []  # pct_of_result lines waiting for the next subtotal
    for line in stack.lines:
        if line.type is LineType.PCT_OF_BASE:
            if line.base == BASE_FOB:
                base = fob
            elif line.base in (None, BASE_RUNNING):
                base = running
            else:
                base = result.values[line.base]
            value = line.rate * base
            running += value
        elif line.type is LineType.FIXED:
            value = line.amount / fx_rate if line.amount_currency == "idr" else line.amount
            running += value
        elif line.type is LineType.PCT_OF_RESULT:
            pending.append(line.id)
            value = math.nan  # filled in at the next subtotal
        else:  # subtotal
            share = sum(_line(stack, i).rate for i in pending)
            value = running / (1 - share)
            for i in pending:
                result.values[i] = _line(stack, i).rate * value
            pending.clear()
            running = value
        result.values[line.id] = value
    return result


def linear_terms(stack: CostStack, fx_rate: float) -> tuple[float, float]:
    """(a, b) such that shelf price in brand currency = a × FOB + b."""
    b = evaluate(stack, 0.0, fx_rate).shelf
    a = evaluate(stack, 1.0, fx_rate).shelf - b
    return a, b


def coefficient(stack: CostStack, fx_rate: float) -> float:
    """Retail coefficient as shown on the price structure sheet: shelf at FOB 100, divided by 100."""
    return evaluate(stack, 100.0, fx_rate).shelf / 100.0


def round_up(value: float, step: int) -> float:
    if step == 0:
        return value
    # guard against float noise such as 1185999.9999999998 rounding up a whole step
    return math.ceil(round(value / step, 9)) * step


def retail_price_idr(stack: CostStack, fob: float, fx_rate: float, rounding_step: int) -> float:
    return round_up(evaluate(stack, fob, fx_rate).shelf * fx_rate, rounding_step)


def implied_fob(stack: CostStack, target_retail_idr: float, fx_rate: float) -> float:
    """Brand price (brand currency) that produces the target shelf price, before rounding."""
    a, b = linear_terms(stack, fx_rate)
    return (target_retail_idr / fx_rate - b) / a


def _line(stack: CostStack, line_id: str):
    return next(line for line in stack.lines if line.id == line_id)
