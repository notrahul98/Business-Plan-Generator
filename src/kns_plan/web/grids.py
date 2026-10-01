"""Editable tables in the app. A grid is described once; the same description renders the HTML table
and parses the submitted form back into rows. Inputs are named `<grid>-<row>-<column>`."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from ..intake import parse_number


@dataclass
class Column:
    name: str
    label: str
    kind: str = "text"  # text | number | idr | percent | select | check | readonly | hidden
    options: list[tuple[str, str]] = field(default_factory=list)  # (value, label) for selects
    width: int = 120
    hint: str = ""
    align: str = "right"  # readonly columns: "left" for text


@dataclass
class Grid:
    name: str
    columns: list[Column]
    rows: list[dict]
    add_label: str = "Add row"
    allow_add: bool = True
    allow_delete: bool = True


def display(col: Column, value) -> str:
    """Value as shown in an input."""
    if value is None or value == "":
        return ""
    if col.kind == "percent":
        return _trim(round(float(value) * 100, 6))
    if col.kind in ("number", "idr"):
        return _trim(value)
    if col.kind == "check":
        return "1" if value else ""
    return str(value)


def _trim(v) -> str:
    f = float(v)
    return str(int(f)) if f.is_integer() else f"{f:.6f}".rstrip("0").rstrip(".")


def parse(form, grid: Grid) -> list[dict]:
    """Rows in submitted order; rows whose visible cells are all blank are dropped."""
    pattern = re.compile(rf"^{re.escape(grid.name)}-(\d+)-(.+)$")
    raw: dict[int, dict[str, str]] = {}
    for key in form.keys():
        m = pattern.match(key)
        if m:
            raw.setdefault(int(m[1]), {})[m[2]] = form.get(key)
    order = sorted(raw)
    if grid.name + "-order" in form:
        wanted = [int(x) for x in form.get(grid.name + "-order").split(",") if x.strip().isdigit()]
        order = [i for i in wanted if i in raw] + [i for i in order if i not in wanted]
    cols = {c.name: c for c in grid.columns}
    out = []
    for i in order:
        cells = raw[i]
        visible = [c for c in grid.columns if c.kind not in ("hidden", "readonly", "check")]
        if not any((cells.get(c.name) or "").strip() for c in visible):
            continue
        row = {}
        for name, col in cols.items():
            v = (cells.get(name) or "").strip()
            if col.kind == "readonly":
                continue
            if col.kind == "check":
                row[name] = v in ("1", "on", "true", "yes")
            elif col.kind in ("number", "idr", "percent"):
                n = parse_number(v, idr=col.kind == "idr")
                row[name] = None if n is None else (n / 100 if col.kind == "percent" else n)
            else:
                row[name] = v or None
        out.append(row)
    return out
