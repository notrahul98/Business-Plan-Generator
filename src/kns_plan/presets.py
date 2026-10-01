"""Default cost stack and volume driver presets.

The rates in `example_stack` are placeholders. The company's real default stack is kept out of the
code, in `<data folder>/default_stack.json` on the server; new plans start from that file when it
exists. Create it with `kns-plan set-default-stack`.
"""

from __future__ import annotations

import os
from pathlib import Path

from .schema import CostLine, CostStack, Factor, LineType

P_BASE, P_RESULT, FIXED, SUB = (LineType.PCT_OF_BASE, LineType.PCT_OF_RESULT,
                                LineType.FIXED, LineType.SUBTOTAL)
DEFAULT_STACK_FILE = "default_stack.json"


def example_stack(stack_id: str = "default", name: str = "Default") -> CostStack:
    """The usual import-and-distribute chain with placeholder rates."""
    return CostStack(id=stack_id, name=name, lines=[
        CostLine(id="freight", label="Freight, insurance and handling", type=P_BASE, rate=0.10, base="fob"),
        CostLine(id="duties", label="Import duties, local tax and clearing", type=P_BASE, rate=0.20,
                 base="running", note="on brand price + freight"),
        CostLine(id="landed", label="Landed cost", type=SUB),
        CostLine(id="distributor", label="Distributor margin", type=P_RESULT, rate=0.32),
        CostLine(id="ga", label="General administrative and operating expenses", type=P_RESULT, rate=0.04),
        CostLine(id="incentive", label="Sales team incentive", type=P_RESULT, rate=0.01),
        CostLine(id="anp", label="Advertising and promotion allowance", type=P_RESULT, rate=0.08),
        CostLine(id="wholesale", label="Wholesale", type=SUB),
        CostLine(id="retailer", label="Retailer margin", type=P_RESULT, rate=0.35),
        CostLine(id="before_vat", label="Retail price before VAT", type=SUB),
        CostLine(id="vat", label="VAT", type=P_BASE, rate=0.11, base="before_vat"),
        CostLine(id="shelf", label="Retail price incl. VAT", type=SUB),
    ])


def default_stack_path(data_dir: Path | str | None = None) -> Path | None:
    folder = data_dir or os.environ.get("KNS_DATA_DIR")
    return Path(folder) / DEFAULT_STACK_FILE if folder else None


def default_stack(stack_id: str = "default", name: str | None = None,
                  data_dir: Path | str | None = None) -> CostStack:
    """The company default from the data folder if one is set up, else the example stack."""
    path = default_stack_path(data_dir)
    if path and path.exists():
        stack = CostStack.model_validate_json(path.read_text(encoding="utf-8"))
        return stack.model_copy(update={"id": stack_id, "name": name or stack.name}, deep=True)
    return example_stack(stack_id, name or "Default")


def save_default_stack(stack: CostStack, data_dir: Path | str) -> Path:
    path = default_stack_path(data_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(stack.model_dump_json(indent=2, exclude_none=True), encoding="utf-8")
    return path


# Volume driver presets: the factors a new driver starts with. Units = product of the factors.
DRIVER_PRESETS: dict[str, dict] = {
    "retail_sellthrough": {
        "label": "Retail sell-through",
        "factors": [Factor(name="Stores", kind="monthly", values=[0], unit="stores"),
                    Factor(name="Units per product per store", kind="monthly", values=[3], unit="units"),
                    Factor(name="Probability of achievement", kind="constant", value=1.0, unit="share")],
    },
    "member_based": {
        "label": "Member-based (gyms, clubs)",
        "factors": [Factor(name="Locations", kind="monthly", values=[0], unit="locations"),
                    Factor(name="Active members per location", kind="constant", value=1000, unit="members"),
                    Factor(name="Buying rate per month", kind="constant", value=0.0015, unit="share"),
                    Factor(name="Units per purchase", kind="constant", value=1, unit="units")],
    },
    "treatments": {
        "label": "Treatments (spas, clinics)",
        "factors": [Factor(name="Outlets", kind="monthly", values=[0], unit="outlets"),
                    Factor(name="Treatments per outlet per month", kind="monthly", values=[10], unit="treatments")],
    },
    "marketplace": {
        "label": "Marketplace orders",
        "factors": [Factor(name="Orders per month", kind="monthly", values=[0], unit="orders"),
                    Factor(name="Units per order", kind="constant", value=1, unit="units")],
    },
    "custom": {
        "label": "Custom",
        "factors": [Factor(name="Volume", kind="monthly", values=[0], unit="units")],
    },
}


def driver_factors(preset: str) -> list[Factor]:
    return [f.model_copy(deep=True) for f in DRIVER_PRESETS[preset]["factors"]]
