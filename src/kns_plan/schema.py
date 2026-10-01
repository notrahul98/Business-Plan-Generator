"""Plan data model. A plan saved as JSON is the source of truth; the workbook is generated from it."""

from __future__ import annotations

import re
from enum import Enum
from typing import Literal

from pydantic import BaseModel, Field, model_validator

SCHEMA_VERSION = 1
ID_PATTERN = r"^[a-z0-9][a-z0-9_]{0,39}$"

# Bases a "% of base" line can point at, besides the id of an earlier line.
BASE_FOB = "fob"
BASE_RUNNING = "running"


# ------------------------------------------------------------------ pricing

class LineType(str, Enum):
    PCT_OF_BASE = "pct_of_base"      # rate × a base (FOB, running total, or an earlier line)
    PCT_OF_RESULT = "pct_of_result"  # rate × the next subtotal (margin on selling price, i.e. a gross-up)
    FIXED = "fixed"                  # a fixed amount per unit
    SUBTOTAL = "subtotal"            # Landed, Wholesale, Before VAT, Shelf…


class CostLine(BaseModel):
    id: str = Field(pattern=ID_PATTERN)
    label: str
    type: LineType
    rate: float | None = None
    base: str | None = None
    amount: float | None = None
    amount_currency: Literal["brand", "idr"] = "brand"
    note: str = ""


class CostStack(BaseModel):
    id: str = Field(pattern=ID_PATTERN)
    name: str
    lines: list[CostLine]
    cogs_line: str | None = "landed"  # subtotal used as KNS cost of goods; None = no gross profit lines

    @model_validator(mode="after")
    def _unique_line_ids(self) -> CostStack:
        ids = [line.id for line in self.lines]
        dupes = sorted({i for i in ids if ids.count(i) > 1})
        if dupes or BASE_FOB in ids or BASE_RUNNING in ids:
            raise ValueError(f"stack '{self.id}': line ids must be unique and not 'fob'/'running' ({dupes})")
        if self.cogs_line is not None and self.cogs_line not in ids:
            self.cogs_line = None
        return self

    @property
    def shelf(self) -> CostLine:
        return self.lines[-1]


class Settings(BaseModel):
    brand: str
    origin: str = ""
    incoterm: str = "FOB"
    brand_currency: str = "USD"
    fx_rate: float = Field(gt=0, description="IDR per 1 unit of brand currency")
    hs_code: str = ""
    rounding_step: Literal[0, 100, 500, 1000] = 1000  # retail prices round UP to this step; 0 = no rounding
    vat_rate: float = Field(default=0.11, ge=0, lt=1, description="VAT on sales, used in the year analysis")
    start_year: int = Field(default=2027, ge=2000, le=2100)
    start_month: int = Field(default=1, ge=1, le=12)
    years: int = Field(default=3, ge=1, le=5)
    stock_cover_months: float = Field(default=3, ge=0, le=24,
                                      description="Closing stock target, in months of the next year's sales")

    @property
    def months(self) -> int:
        return self.years * 12


# ------------------------------------------------------------------ products and competitors

class PricingRule(BaseModel):
    """Proposed retail = competitor price, size-matched to our product, less a discount."""
    competitor: str
    discount: float = Field(default=0.0, gt=-1, lt=1, description="0.15 = 15% below; -0.10 = 10% above")


class SKU(BaseModel):
    id: str = Field(pattern=ID_PATTERN)
    name: str = Field(min_length=1)
    size: float | None = Field(default=None, gt=0)
    size_label: str | None = None  # as written by the brand, e.g. "31 x 4"
    unit: str | None = None
    fob: float | None = Field(default=None, ge=0, description="Brand price in brand currency")
    target_retail: float | None = Field(default=None, gt=0, description="Proposed shelf price, IDR incl. VAT")
    pricing_rule: PricingRule | None = None  # used when target_retail is empty
    stack: str = "default"
    opening_units: float = Field(default=0, ge=0, description="Stock on hand at the start of the plan")
    moq: float | None = Field(default=None, gt=0, description="Purchases round up to a multiple of this")
    lead_time_months: int = Field(default=0, ge=0, le=12)
    attributes: dict[str, str] = Field(default_factory=dict)


class Competitor(BaseModel):
    id: str = Field(pattern=ID_PATTERN)
    sku: str  # our product it benchmarks
    brand: str = ""
    product: str
    origin: str = ""
    size: float | None = Field(default=None, gt=0)
    unit: str | None = None
    price_idr: float = Field(gt=0, description="Shelf price incl. VAT")
    source: str = ""
    registration: str = ""  # e.g. BPOM notification number
    notes: str = ""


# ------------------------------------------------------------------ channels and volume

class MarginComponent(BaseModel):
    label: str
    rate: float = Field(ge=0, lt=1)
    plus_vat: bool = False  # fee is charged with VAT on top (e.g. marketplace fees)


class Channel(BaseModel):
    id: str = Field(pattern=ID_PATTERN)
    name: str
    kind: Literal["retail", "member", "treatment", "marketplace", "other"] = "retail"
    margin: float | None = Field(default=None, ge=0, lt=1, description="Share of retail sales incl. VAT")
    margin_components: list[MarginComponent] = Field(default_factory=list)
    listings: dict[str, float] = Field(default_factory=dict,
                                       description="SKU id → weight; listed SKUs only. Units = driver × weight")
    package: str | None = None  # treatment channels: weights come from this package's usage


class Factor(BaseModel):
    name: str
    kind: Literal["constant", "monthly"] = "constant"
    value: float | None = None
    values: list[float] = Field(default_factory=list)  # monthly; shorter lists repeat their last value
    unit: str = ""

    def at(self, month: int) -> float:
        """Value for month index 0..n-1."""
        if self.kind == "constant":
            return self.value or 0.0
        if not self.values:
            return 0.0
        return self.values[min(month, len(self.values) - 1)]


class Driver(BaseModel):
    id: str = Field(pattern=ID_PATTERN)
    channel: str
    name: str
    preset: str = "custom"
    factors: list[Factor] = Field(min_length=1)


class Override(BaseModel):
    channel: str
    sku: str
    month: int = Field(ge=1, le=60, description="1 = first month of the plan")
    units: float = Field(ge=0)


# ------------------------------------------------------------------ add-on modules

class PackageComponent(BaseModel):
    sku: str
    usage: float = Field(gt=0, description="Amount used per treatment, in the product's size unit")


class Package(BaseModel):
    """A treatment (or kit) that consumes several products."""
    id: str = Field(pattern=ID_PATTERN)
    name: str
    components: list[PackageComponent] = Field(min_length=1)
    price_multiplier: float = Field(default=4.0, gt=0, description="Outlet price per treatment ÷ product cost")
    duration_minutes: int | None = None


class ContainerOption(BaseModel):
    name: str
    cost: float = Field(gt=0, description="Freight and charges per container, brand currency")
    units: float = Field(gt=0, description="Sellable units per container")


class MarketRow(BaseModel):
    segment: str
    metric: str = ""
    value: str = ""
    source: str = ""


# ------------------------------------------------------------------ plan

class Plan(BaseModel):
    schema_version: int = SCHEMA_VERSION
    settings: Settings
    stacks: list[CostStack]
    skus: list[SKU] = Field(default_factory=list)
    competitors: list[Competitor] = Field(default_factory=list)
    channels: list[Channel] = Field(default_factory=list)
    drivers: list[Driver] = Field(default_factory=list)
    overrides: list[Override] = Field(default_factory=list)
    packages: list[Package] = Field(default_factory=list)
    containers: list[ContainerOption] = Field(default_factory=list)
    market: list[MarketRow] = Field(default_factory=list)

    @model_validator(mode="after")
    def _references(self) -> Plan:
        def unique(items, what):
            ids = [i.id for i in items]
            dupes = sorted({i for i in ids if ids.count(i) > 1})
            if dupes:
                raise ValueError(f"{what} ids must be unique: {', '.join(dupes)}")
            return set(ids)

        stack_ids = unique(self.stacks, "cost stack")
        sku_ids = unique(self.skus, "product")
        comp_ids = unique(self.competitors, "competitor")
        chan_ids = unique(self.channels, "channel")
        unique(self.drivers, "driver")
        pkg_ids = unique(self.packages, "package")
        for s in self.skus:
            if s.stack not in stack_ids:
                raise ValueError(f"product '{s.id}' uses unknown cost stack '{s.stack}'")
            if s.pricing_rule and s.pricing_rule.competitor not in comp_ids:
                raise ValueError(f"product '{s.id}' prices against unknown competitor '{s.pricing_rule.competitor}'")
        for c in self.competitors:
            if c.sku not in sku_ids:
                raise ValueError(f"competitor '{c.id}' benchmarks unknown product '{c.sku}'")
        for ch in self.channels:
            bad = set(ch.listings) - sku_ids
            if bad:
                raise ValueError(f"channel '{ch.id}' lists unknown products: {', '.join(sorted(bad))}")
            if ch.package and ch.package not in pkg_ids:
                raise ValueError(f"channel '{ch.id}' uses unknown package '{ch.package}'")
        for d in self.drivers:
            if d.channel not in chan_ids:
                raise ValueError(f"driver '{d.id}' belongs to unknown channel '{d.channel}'")
        for o in self.overrides:
            if o.channel not in chan_ids or o.sku not in sku_ids:
                raise ValueError(f"override for unknown channel/product: {o.channel}/{o.sku}")
        for p in self.packages:
            bad = {c.sku for c in p.components} - sku_ids
            if bad:
                raise ValueError(f"package '{p.id}' uses unknown products: {', '.join(sorted(bad))}")
        return self

    def stack(self, stack_id: str) -> CostStack:
        return next(s for s in self.stacks if s.id == stack_id)

    def sku(self, sku_id: str) -> SKU:
        return next(s for s in self.skus if s.id == sku_id)

    def competitor(self, comp_id: str) -> Competitor:
        return next(c for c in self.competitors if c.id == comp_id)

    def package(self, pkg_id: str) -> Package:
        return next(p for p in self.packages if p.id == pkg_id)


def slugify(text: str, existing: set[str] | None = None) -> str:
    """Make a valid id from free text, unique within `existing`."""
    base = re.sub(r"[^a-z0-9]+", "_", text.lower()).strip("_")[:36] or "item"
    if not base[0].isalnum():
        base = "x" + base
    slug, n = base, 2
    while existing and slug in existing:
        slug, n = f"{base}_{n}", n + 1
    return slug
