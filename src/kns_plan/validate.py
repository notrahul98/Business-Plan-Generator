"""Checks run before a workbook is generated. Errors block generation; warnings are shown only."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

from . import model
from .schema import BASE_FOB, BASE_RUNNING, CostStack, LineType, Plan


@dataclass(frozen=True)
class Issue:
    level: Literal["error", "warning"]
    where: str
    message: str

    def __str__(self) -> str:
        return f"[{self.level}] {self.where}: {self.message}"


def check_stack(stack: CostStack) -> list[Issue]:
    issues: list[Issue] = []
    where = f"cost stack '{stack.name}'"

    def err(msg: str) -> None:
        issues.append(Issue("error", where, msg))

    if not stack.lines or stack.lines[-1].type is not LineType.SUBTOTAL:
        err("must end with a subtotal line (the shelf price incl. VAT)")

    seen_values: set[str] = set()  # lines whose value is known by the time later lines are evaluated
    pending_share = 0.0
    pending: list[str] = []
    for line in stack.lines:
        label = f"line '{line.label}'"
        if line.type in (LineType.PCT_OF_BASE, LineType.PCT_OF_RESULT):
            if line.rate is None:
                err(f"{label} needs a rate")
                continue
            if not 0 <= line.rate < 10:
                err(f"{label}: rate {line.rate:.2%} is out of range")
        if line.type is LineType.PCT_OF_BASE:
            if line.base not in (None, BASE_FOB, BASE_RUNNING) and line.base not in seen_values:
                err(f"{label}: base '{line.base}' must be fob, running, or an earlier line")
        elif line.type is LineType.FIXED:
            if line.amount is None:
                err(f"{label} needs an amount")
        elif line.type is LineType.PCT_OF_RESULT:
            pending_share += line.rate
            pending.append(line.id)
            continue  # its value is only known after the next subtotal
        elif line.type is LineType.SUBTOTAL:
            if pending_share >= 1:
                err(f"{label}: '% of resulting price' lines add up to {pending_share:.0%}; must be under 100%")
            seen_values.update(pending)
            pending, pending_share = [], 0.0
        seen_values.add(line.id)
    return issues


def check_plan(plan: Plan) -> list[Issue]:
    issues: list[Issue] = []

    def warn(where: str, msg: str) -> None:
        issues.append(Issue("warning", where, msg))

    for stack in plan.stacks:
        issues += check_stack(stack)
        vat_lines = [line for line in stack.lines if line.id == "vat" and line.rate is not None]
        if vat_lines and abs(vat_lines[0].rate - plan.settings.vat_rate) > 1e-9:
            warn(f"cost stack '{stack.name}'", f"VAT line is {vat_lines[0].rate:.0%} but the plan's VAT setting is "
                                               f"{plan.settings.vat_rate:.0%}")
    if not plan.skus:
        warn("products", "the plan has no products yet")
    listed = {s for ch in plan.channels for s in model.channel_weights(plan, ch)}
    for sku in plan.skus:
        where = f"product '{sku.name}'"
        if sku.fob is None and sku.target_retail is None and sku.pricing_rule is None:
            warn(where, "has no brand price, target retail or pricing rule, so it gets no prices")
        if sku.pricing_rule:
            comp = plan.competitor(sku.pricing_rule.competitor)
            if not (comp.size and sku.size and model.same_unit(comp.unit, sku.unit)):
                warn(where, f"competitor '{comp.product}' cannot be size-matched (missing size or different unit); "
                            "its full price is used")
        if plan.channels and sku.id not in listed:
            warn(where, "is not listed in any channel, so it has no sales")
    for ch in plan.channels:
        where = f"channel '{ch.name}'"
        if ch.margin is None and not ch.margin_components:
            warn(where, "has no margin; 0% is used")
        if model.effective_margin(plan, ch) >= 1:
            issues.append(Issue("error", where, "margin and fees add up to 100% or more"))
        if not model.channel_weights(plan, ch):
            warn(where, "lists no products")
        if not any(d.channel == ch.id for d in plan.drivers):
            warn(where, "has no volume driver, so it has no sales")
        if ch.package:
            for comp in plan.package(ch.package).components:
                if plan.sku(comp.sku).size is None:
                    issues.append(Issue("error", where, f"package product '{plan.sku(comp.sku).name}' needs a size"))
    n = plan.settings.months
    for o in plan.overrides:
        if o.month > n:
            issues.append(Issue("error", "overrides", f"month {o.month} is beyond the {n}-month plan"))
    for d in plan.drivers:
        for f in d.factors:
            if f.kind == "monthly" and len(f.values) > n:
                warn(f"driver '{d.name}'", f"factor '{f.name}' has {len(f.values)} months; only {n} are used")
            if f.kind == "constant" and f.value is None:
                warn(f"driver '{d.name}'", f"factor '{f.name}' has no value; 0 is used")
    return issues


def errors(issues: list[Issue]) -> list[Issue]:
    return [i for i in issues if i.level == "error"]
