"""Plan calculations beyond the cost stack: proposals, volumes, year analysis and stock.

These mirror the workbook formulas one for one. Every generated workbook is recalculated and
compared against these numbers, so a change here must be matched in the Excel writer.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field

from .schema import SKU, Channel, Competitor, Plan
from .stack import evaluate, implied_fob, linear_terms, retail_price_idr, round_up


# ------------------------------------------------------------------ pricing per product

def same_unit(a: str | None, b: str | None) -> bool:
    return (a or "").strip().lower() == (b or "").strip().lower()


def size_matched_price(plan: Plan, comp: Competitor) -> float:
    """Competitor price scaled to our product's size, when both sizes are known in the same unit."""
    sku = plan.sku(comp.sku)
    if comp.size and sku.size and same_unit(comp.unit, sku.unit):
        return round_up(comp.price_idr / comp.size * sku.size, plan.settings.rounding_step)
    return comp.price_idr


def rule_price(plan: Plan, sku: SKU) -> float | None:
    if not sku.pricing_rule:
        return None
    matched = size_matched_price(plan, plan.competitor(sku.pricing_rule.competitor))
    return round_up(matched * (1 - sku.pricing_rule.discount), plan.settings.rounding_step)


def proposed_retail(plan: Plan, sku: SKU) -> float | None:
    """Typed target wins; otherwise the competitor rule; otherwise none."""
    return sku.target_retail if sku.target_retail is not None else rule_price(plan, sku)


def retail_from_fob(plan: Plan, sku: SKU) -> float | None:
    if sku.fob is None:
        return None
    s = plan.settings
    return retail_price_idr(plan.stack(sku.stack), sku.fob, s.fx_rate, s.rounding_step)


def forecast_price(plan: Plan, sku: SKU) -> float | None:
    """Shelf price used for sales: the proposed retail if there is one, else the price from the brand price."""
    p = proposed_retail(plan, sku)
    return p if p is not None else retail_from_fob(plan, sku)


def implied_brand_price(plan: Plan, sku: SKU) -> float | None:
    p = proposed_retail(plan, sku)
    return None if p is None else implied_fob(plan.stack(sku.stack), p, plan.settings.fx_rate)


def cost_basis_fob(plan: Plan, sku: SKU) -> float | None:
    """Brand price used for cost of goods and purchases: the current brand price, else the implied one."""
    return sku.fob if sku.fob is not None else implied_brand_price(plan, sku)


def landed_cost_idr(plan: Plan, sku: SKU) -> float:
    """KNS cost of goods per unit (the stack's cost-of-goods subtotal), IDR. 0 when unknown."""
    stack = plan.stack(sku.stack)
    fob = cost_basis_fob(plan, sku)
    if fob is None or stack.cogs_line is None:
        return 0.0
    return evaluate(stack, fob, plan.settings.fx_rate).values[stack.cogs_line] * plan.settings.fx_rate


def landed_terms(plan: Plan, stack_id: str) -> tuple[float, float]:
    """(a, b): cost-of-goods subtotal in brand currency = a × FOB + b."""
    stack = plan.stack(stack_id)
    fx = plan.settings.fx_rate
    b = evaluate(stack, 0.0, fx).values[stack.cogs_line]
    return evaluate(stack, 1.0, fx).values[stack.cogs_line] - b, b


# ------------------------------------------------------------------ channels and units

def effective_margin(plan: Plan, ch: Channel) -> float:
    if ch.margin_components:
        vat = plan.settings.vat_rate
        return sum(c.rate * (1 + vat if c.plus_vat else 1) for c in ch.margin_components)
    return ch.margin or 0.0


def channel_weights(plan: Plan, ch: Channel) -> dict[str, float]:
    """SKU → units of that product per driver unit. Treatment channels use their package's usage."""
    if ch.package:
        out: dict[str, float] = {}
        for comp in plan.package(ch.package).components:
            size = plan.sku(comp.sku).size or 1.0
            out[comp.sku] = out.get(comp.sku, 0.0) + comp.usage / size
        return out
    return dict(ch.listings)


def driver_units(plan: Plan, driver_id: str) -> list[float]:
    d = next(d for d in plan.drivers if d.id == driver_id)
    return [math.prod(f.at(m) for f in d.factors) for m in range(plan.settings.months)]


def units(plan: Plan) -> dict[tuple[str, str], list[float]]:
    """(channel, sku) → units per month for the whole horizon."""
    n = plan.settings.months
    per_channel: dict[str, list[float]] = {ch.id: [0.0] * n for ch in plan.channels}
    for d in plan.drivers:
        series = driver_units(plan, d.id)
        per_channel[d.channel] = [a + b for a, b in zip(per_channel[d.channel], series)]
    overrides = {(o.channel, o.sku, o.month): o.units for o in plan.overrides}
    out = {}
    for ch in plan.channels:
        for sku_id, weight in channel_weights(plan, ch).items():
            out[(ch.id, sku_id)] = [overrides.get((ch.id, sku_id, m + 1), weight * per_channel[ch.id][m])
                                    for m in range(n)]
    return out


def year_slice(series: list[float], year: int) -> list[float]:
    return series[(year - 1) * 12: year * 12]


# ------------------------------------------------------------------ year analysis

@dataclass
class ChannelYear:
    sales: float = 0.0         # retail sales incl. VAT
    margin_rate: float = 0.0
    margin: float = 0.0
    wholesale: float = 0.0     # incl. VAT
    vat: float = 0.0
    net_wholesale: float = 0.0
    cogs: float = 0.0
    gross_profit: float = 0.0
    units: float = 0.0


@dataclass
class Analysis:
    by_year: dict[int, dict[str, ChannelYear]] = field(default_factory=dict)  # year → channel → numbers

    def total(self, year: int) -> ChannelYear:
        t = ChannelYear()
        for cy in self.by_year[year].values():
            for k in ("sales", "margin", "wholesale", "vat", "net_wholesale", "cogs", "gross_profit", "units"):
                setattr(t, k, getattr(t, k) + getattr(cy, k))
        return t


def analyse(plan: Plan) -> Analysis:
    u = units(plan)
    prices = {s.id: forecast_price(plan, s) or 0.0 for s in plan.skus}
    costs = {s.id: landed_cost_idr(plan, s) for s in plan.skus}
    vat = plan.settings.vat_rate
    result = Analysis()
    for y in range(1, plan.settings.years + 1):
        result.by_year[y] = {}
        for ch in plan.channels:
            cy = ChannelYear(margin_rate=effective_margin(plan, ch))
            for (ch_id, sku_id), series in u.items():
                if ch_id != ch.id:
                    continue
                qty = sum(year_slice(series, y))
                cy.units += qty
                cy.sales += qty * prices[sku_id]
                cy.cogs += qty * costs[sku_id]
            cy.margin = cy.sales * cy.margin_rate
            cy.wholesale = cy.sales - cy.margin
            cy.vat = cy.wholesale * vat / (1 + vat)
            cy.net_wholesale = cy.wholesale - cy.vat
            cy.gross_profit = cy.net_wholesale - cy.cogs
            result.by_year[y][ch.id] = cy
    return result


# ------------------------------------------------------------------ add-ons

def treatment_cost(plan: Plan, package_id: str) -> float:
    """Product cost of one treatment at forecast prices, IDR."""
    total = 0.0
    for comp in plan.package(package_id).components:
        sku = plan.sku(comp.sku)
        total += comp.usage / (sku.size or 1.0) * (forecast_price(plan, sku) or 0.0)
    return total


def average_cost_fob(plan: Plan) -> float | None:
    values = [v for v in (cost_basis_fob(plan, s) for s in plan.skus) if v is not None]
    return sum(values) / len(values) if values else None


def freight_share(plan: Plan, cost: float, units_per_container: float) -> float | None:
    """Container cost as a share of the goods' brand-price value: a starting point for the freight %."""
    avg = average_cost_fob(plan)
    return None if not avg else cost / (units_per_container * avg)


# ------------------------------------------------------------------ stock and purchases

@dataclass
class StockYear:
    sales: float
    opening: float
    target_closing: float
    purchase: float
    closing: float
    cover_months: float | None


def stock(plan: Plan) -> dict[str, list[StockYear]]:
    """Per product, per year, in units. Purchases bring closing stock up to the cover target."""
    u = units(plan)
    years = plan.settings.years
    cover = plan.settings.stock_cover_months
    out = {}
    for sku in plan.skus:
        sales = [sum(sum(year_slice(series, y)) for (_, s), series in u.items() if s == sku.id)
                 for y in range(1, years + 1)]
        rows, opening = [], sku.opening_units
        for i in range(years):
            next_sales = sales[i + 1] if i + 1 < years else sales[i]
            target = cover / 12 * next_sales
            need = max(0.0, sales[i] + target - opening)
            purchase = math.ceil(round(need / sku.moq, 9)) * sku.moq if sku.moq else need
            closing = opening + purchase - sales[i]
            cover_m = closing / (next_sales / 12) if next_sales else None
            rows.append(StockYear(sales[i], opening, target, purchase, closing, cover_m))
            opening = closing
        out[sku.id] = rows
    return out
