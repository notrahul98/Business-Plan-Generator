"""Helpers for editing a plan from the app: keep references consistent when rows are added or removed."""

from __future__ import annotations

from .presets import default_stack, driver_factors
from .schema import Driver, Plan, Settings, slugify


def new_plan(brand: str, fx_rate: float, currency: str = "USD", data_dir=None) -> Plan:
    """A blank plan starting from the company's default cost stack (see presets.default_stack)."""
    return Plan(settings=Settings(brand=brand, brand_currency=currency, fx_rate=fx_rate),
                stacks=[default_stack(data_dir=data_dir)])


def prune(data: dict) -> dict:
    """Drop references to products, competitors, channels or packages that no longer exist.

    Works on a plain dict (as edited by forms) so it can run before validation.
    """
    skus = {s["id"] for s in data.get("skus", [])}
    stacks = {s["id"] for s in data.get("stacks", [])}
    first_stack = next(iter(s["id"] for s in data.get("stacks", [])), "default")

    data["competitors"] = [c for c in data.get("competitors", []) if c.get("sku") in skus]
    comps = {c["id"] for c in data["competitors"]}
    for p in data.get("packages", []):
        p["components"] = [c for c in p.get("components", []) if c["sku"] in skus]
    data["packages"] = [p for p in data.get("packages", []) if p["components"]]
    pkgs = {p["id"] for p in data["packages"]}
    for s in data.get("skus", []):
        if s.get("stack") not in stacks:
            s["stack"] = first_stack
        rule = s.get("pricing_rule")
        if rule and rule.get("competitor") not in comps:
            s["pricing_rule"] = None
    for ch in data.get("channels", []):
        ch["listings"] = {k: v for k, v in (ch.get("listings") or {}).items() if k in skus}
        if ch.get("package") and ch["package"] not in pkgs:
            ch["package"] = None
    chans = {c["id"] for c in data.get("channels", [])}
    data["drivers"] = [d for d in data.get("drivers", []) if d.get("channel") in chans]
    data["overrides"] = [o for o in data.get("overrides", []) if o.get("channel") in chans and o.get("sku") in skus]
    return data


def ensure_id(row: dict, existing: set[str], source: str) -> str:
    """Keep a row's id, or make one from its name."""
    rid = (row.get("id") or "").strip()
    if not rid:
        rid = slugify(source or "item", existing)
    existing.add(rid)
    return rid


def add_driver(plan: Plan, channel_id: str, preset: str, name: str | None = None) -> Driver:
    from .presets import DRIVER_PRESETS
    existing = {d.id for d in plan.drivers}
    driver = Driver(id=slugify(f"{channel_id} {preset}", existing), channel=channel_id,
                    name=name or DRIVER_PRESETS[preset]["label"], preset=preset, factors=driver_factors(preset))
    plan.drivers.append(driver)
    return driver
