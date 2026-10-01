"""Shared test plans. All brands, products and prices are made up."""

import pytest

from kns_plan.presets import example_stack
from kns_plan.schema import (SKU, Channel, Competitor, ContainerOption, CostLine, CostStack, Driver, Factor,
                             LineType, MarginComponent, MarketRow, Override, Package, PackageComponent, Plan,
                             PricingRule, Settings)


@pytest.fixture
def reference_plan() -> Plan:
    """Two products priced with the example stack; expected figures are worked out by hand in the tests."""
    return Plan(
        settings=Settings(brand="Brand A", origin="Korea", incoterm="FOB Korea", brand_currency="USD",
                          fx_rate=18200, hs_code="3304.99.00", rounding_step=1000),
        stacks=[example_stack()],
        skus=[
            SKU(id="eye_stick", name="Eye serum stick", size=15, unit="ml", fob=8, target_retail=650000),
            SKU(id="jelly_mask", name="Jelly mask", size=124, size_label="31 x 4", unit="grams", fob=5,
                target_retail=380000),
        ],
    )


@pytest.fixture
def mixed_plan() -> Plan:
    """A plan exercising everything the example stack does not: fixed amounts in both
    currencies, a base pointing at an earlier subtotal, two stacks, a product with no brand price,
    extra product attributes and no rounding."""
    food = CostStack(id="food", name="Food, reefer container", lines=[
        CostLine(id="freight", label="Reefer freight", type=LineType.PCT_OF_BASE, rate=0.12, base="fob"),
        CostLine(id="port", label="Port handling", type=LineType.FIXED, amount=1500, amount_currency="idr"),
        CostLine(id="duties", label="Duties", type=LineType.PCT_OF_BASE, rate=0.25, base="running"),
        CostLine(id="landed", label="Landed", type=LineType.SUBTOTAL),
        CostLine(id="dist", label="Distributor", type=LineType.PCT_OF_RESULT, rate=0.35),
        CostLine(id="coop", label="Brand co-op", type=LineType.FIXED, amount=0.05),
        CostLine(id="anp", label="A&P", type=LineType.PCT_OF_RESULT, rate=0.10),
        CostLine(id="wholesale", label="Wholesale", type=LineType.SUBTOTAL),
        CostLine(id="insurance", label="Insurance on landed", type=LineType.PCT_OF_BASE, rate=0.01, base="landed"),
        CostLine(id="retailer", label="Retailer", type=LineType.PCT_OF_RESULT, rate=0.20),
        CostLine(id="before_vat", label="Before VAT", type=LineType.SUBTOTAL),
        CostLine(id="vat", label="VAT", type=LineType.PCT_OF_BASE, rate=0.11, base="before_vat"),
        CostLine(id="shelf", label="Shelf", type=LineType.SUBTOTAL),
    ])
    return Plan(
        settings=Settings(brand="Test Brand", brand_currency="EUR", fx_rate=20000, rounding_step=0),
        stacks=[example_stack(), food],
        skus=[
            SKU(id="bar", name="Protein bar", size=55, unit="g", fob=1.16, stack="food",
                attributes={"Flavour": "Salty peanut", "EAN": "0000000000001"}),
            SKU(id="serum", name="Serum", size=30, unit="ml", fob=10, target_retail=650000),
            SKU(id="new_item", name="Item without brand price", target_retail=99000, stack="food"),
        ],
    )


@pytest.fixture
def full_plan(mixed_plan) -> Plan:
    """Every module at once: competitor rules, three kinds of channel, constant and monthly drivers,
    overrides, a treatment package, MOQ, containers and market notes, over two years from July."""
    p = mixed_plan
    p.settings.start_month, p.settings.start_year, p.settings.years = 7, 2027, 2
    p.settings.rounding_step = 500
    p.skus[1].target_retail = None
    p.skus[1].pricing_rule = PricingRule(competitor="brand_x", discount=0.15)
    p.skus[1].moq = 240
    p.skus[1].opening_units = 100
    p.skus.append(SKU(id="massage_oil", name="Massage oil 1L", size=1000, unit="ml", fob=24))
    p.competitors = [
        Competitor(id="brand_x", sku="serum", brand="Brand X", product="Eye cream", size=14, unit="ml",
                   price_idr=800000, source="retailer website"),
        Competitor(id="patch", sku="serum", brand="Brand Y", product="Eye patch", size=60, unit="pcs",
                   price_idr=340000),
    ]
    p.channels = [
        Channel(id="watsons", name="Watsons", margin=0.38,
                listings={"serum": 1, "bar": 0.5, "new_item": 2}),
        Channel(id="shopee", name="Shopee", kind="marketplace",
                margin_components=[MarginComponent(label="Revenue share", rate=0.12, plus_vat=True),
                                   MarginComponent(label="Payment fee", rate=0.02)],
                listings={"serum": 1, "bar": 1}),
        Channel(id="spas", name="Day spas", kind="treatment", margin=0.30, package="ritual"),
    ]
    p.drivers = [
        Driver(id="watsons_stores", channel="watsons", name="Stores", preset="retail_sellthrough", factors=[
            Factor(name="Stores", kind="monthly", values=[2, 4, 6, 10, 15], unit="stores"),
            Factor(name="Units per product per store", kind="monthly", values=[3, 5, 10, 8], unit="units"),
            Factor(name="Probability", kind="constant", value=0.7, unit="share")]),
        Driver(id="shopee_orders", channel="shopee", name="Orders", preset="marketplace", factors=[
            Factor(name="Orders", kind="monthly", values=[50, 80, 120, 200, 150], unit="orders"),
            Factor(name="Units per order", kind="constant", value=1.5)]),
        Driver(id="spa_treatments", channel="spas", name="Treatments", preset="treatments", factors=[
            Factor(name="Spas", kind="monthly", values=[1, 2, 3, 5, 4], unit="spas"),
            Factor(name="Treatments per spa", kind="constant", value=40)]),
        Driver(id="spa_extra", channel="spas", name="Hotel spas", factors=[
            Factor(name="Treatments", kind="monthly", values=[0, 0, 0, 30, 60, 90, 70])]),
    ]
    p.overrides = [Override(channel="watsons", sku="serum", month=1, units=12),
                   Override(channel="shopee", sku="bar", month=14, units=0)]
    p.packages = [Package(id="ritual", name="Signature ritual", duration_minutes=90, price_multiplier=4,
                          components=[PackageComponent(sku="massage_oil", usage=30),
                                      PackageComponent(sku="serum", usage=1.5)])]
    p.containers = [ContainerOption(name="20 ft reefer", cost=4200, units=24000),
                    ContainerOption(name="40 ft reefer", cost=6900, units=52000)]
    p.market = [MarketRow(segment="Hotel and resort spas", metric="Share of spa market", value="35-40%",
                          source="Industry estimate")]
    return Plan.model_validate(p.model_dump())
