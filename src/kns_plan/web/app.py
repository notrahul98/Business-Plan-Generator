"""KNS Business Plan Generator — web app.

Run with `kns-plan serve`. Configuration (environment variables):
  KNS_DATA_DIR       folder for the database, generated files, uploads and backups (default ./data)
  KNS_PASSWORD_HASH  shared office password hash, made with `kns-plan hash-password`
                     (or put the hash in <data folder>/password.hash)
  KNS_HTTPS          set to 1 behind HTTPS so the session cookie is marked secure
"""

from __future__ import annotations

import json
import os
import secrets
import threading
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import ValidationError
from starlette.middleware.sessions import SessionMiddleware

from .. import intake, model
from ..edit import add_driver, ensure_id, new_plan, prune
from ..excel import PlanInvalid, generate, read_inputs
from ..excel.inputs import month_labels
from ..presets import DRIVER_PRESETS, default_stack
from ..schema import LineType, Plan, slugify
from ..stack import evaluate, linear_terms
from ..validate import check_plan
from .auth import Gate
from .grids import Column, Grid, display, parse
from .store import Store

HERE = Path(__file__).parent
STEPS = [("setup", "Brand & market"), ("pricing", "Cost stack"), ("products", "Products"),
         ("competitors", "Competitors"), ("channels", "Channels"), ("volume", "Volume"),
         ("extras", "Extras"), ("review", "Review & generate")]
GENERATE_LOCK = threading.Lock()  # LibreOffice runs one file at a time


def create_app(data_dir: Path | str | None = None, password_hash: str | None = None,
               https: bool | None = None) -> FastAPI:
    data = Path(data_dir or os.environ.get("KNS_DATA_DIR", "data")).resolve()
    data.mkdir(parents=True, exist_ok=True)
    store = Store(data / "plans.sqlite")
    hash_file = data / "password.hash"
    pw_hash = (password_hash or os.environ.get("KNS_PASSWORD_HASH", "")
               or (hash_file.read_text().strip() if hash_file.exists() else ""))
    gate = Gate(store, pw_hash)
    secret_file = data / ".session_key"
    if not secret_file.exists():
        secret_file.write_text(secrets.token_urlsafe(48))
    https = https if https is not None else os.environ.get("KNS_HTTPS") == "1"

    app = FastAPI(title="KNS Business Plan Generator", docs_url=None, redoc_url=None, openapi_url=None)
    app.mount("/static", StaticFiles(directory=HERE / "static"), name="static")
    templates = Jinja2Templates(directory=HERE / "templates")
    templates.env.globals.update(STEPS=STEPS, display=display)
    # changes whenever a static file changes, so browsers fetch the new CSS/JS after an update
    templates.env.globals["asset_v"] = str(max(int(f.stat().st_mtime) for f in (HERE / "static").iterdir()))
    templates.env.filters["idr"] = lambda v: "" if v is None else f"{v:,.0f}"
    templates.env.filters["money"] = lambda v: "" if v is None else f"{v:,.2f}"
    templates.env.filters["pct"] = lambda v: "" if v is None else f"{v:.1%}"
    templates.env.filters["num"] = lambda v: display(Column("v", "v", "number"), v)
    templates.env.filters["units"] = lambda v: "" if v is None else f"{v:,.0f}"
    app.state.store, app.state.data = store, data

    # ------------------------------------------------------------ auth
    @app.middleware("http")
    async def require_login(request: Request, call_next):
        open_paths = ("/login", "/static/", "/health")
        if not request.url.path.startswith(open_paths) and not request.session.get("ok"):
            return RedirectResponse("/login", status_code=303)
        response = await call_next(request)
        response.headers["X-Frame-Options"] = "DENY"
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["Referrer-Policy"] = "same-origin"
        return response

    # added after the login check so it wraps it: the session exists before the check runs
    app.add_middleware(SessionMiddleware, secret_key=secret_file.read_text().strip(), https_only=https,
                       same_site="strict", max_age=12 * 3600)

    def ip_of(request: Request) -> str:
        fwd = request.headers.get("x-forwarded-for")
        return fwd.split(",")[0].strip() if fwd else (request.client.host if request.client else "?")

    def user_of(request: Request) -> str:
        return request.session.get("name", "")

    @app.get("/health")
    def health():
        return {"ok": True}

    @app.get("/login", response_class=HTMLResponse)
    def login_page(request: Request):
        return templates.TemplateResponse(request, "login.html", {"error": "", "configured": bool(pw_hash)})

    @app.post("/login", response_class=HTMLResponse)
    def login(request: Request, password: str = Form(""), name: str = Form("")):
        if not pw_hash:
            return templates.TemplateResponse(request, "login.html", {"error": "", "configured": False})
        ok, msg = gate.attempt(ip_of(request), password)
        if not ok:
            return templates.TemplateResponse(request, "login.html", {"error": msg, "configured": True},
                                              status_code=401)
        request.session.clear()
        request.session.update(ok=True, name=name.strip()[:40], csrf=secrets.token_urlsafe(24))
        return RedirectResponse("/", status_code=303)

    @app.post("/logout")
    def logout(request: Request):
        request.session.clear()
        return RedirectResponse("/login", status_code=303)

    async def form_of(request: Request):
        form = await request.form()
        if form.get("csrf") != request.session.get("csrf"):
            raise HTTPException(400, "This form has expired. Reload the page and try again.")
        return form

    def page(request: Request, name: str, ctx: dict, status: int = 200):
        ctx.setdefault("flash", request.session.pop("flash", None))
        ctx.setdefault("flash_kind", request.session.pop("flash_kind", "ok"))
        ctx["csrf"] = request.session.get("csrf", "")
        ctx["user"] = user_of(request)
        return templates.TemplateResponse(request, name, ctx, status_code=status)

    def flash(request: Request, msg: str, kind: str = "ok") -> None:
        request.session["flash"], request.session["flash_kind"] = msg, kind

    def load(plan_id: int, version: int | None = None):
        try:
            return store.get(plan_id, version)
        except KeyError:
            raise HTTPException(404, "Plan not found") from None

    # ------------------------------------------------------------ plan list
    @app.get("/", response_class=HTMLResponse)
    def home(request: Request, archived: int = 0):
        return page(request, "plans.html", {"plans": store.list_plans(bool(archived)), "archived": archived})

    @app.post("/plans")
    async def create_plan(request: Request):
        form = await form_of(request)
        brand = (form.get("brand") or "").strip()
        fx = intake.parse_number(form.get("fx_rate"), idr=True)
        if not brand or not fx or fx <= 0:
            flash(request, "Enter a brand name and an exchange rate above 0.", "error")
            return RedirectResponse("/", status_code=303)
        plan = new_plan(brand, fx, (form.get("currency") or "USD").strip().upper(), data)
        name = (form.get("name") or "").strip() or f"{brand} business plan"
        pid = store.create(name, plan, user_of(request))
        return RedirectResponse(f"/plans/{pid}/setup", status_code=303)

    @app.post("/plans/{plan_id}/copy")
    async def copy_plan(request: Request, plan_id: int):
        await form_of(request)
        row, plan = load(plan_id)
        pid = store.create(f"{row.name} (copy)", plan, user_of(request))
        return RedirectResponse(f"/plans/{pid}/setup", status_code=303)

    @app.post("/plans/{plan_id}/rename")
    async def rename_plan(request: Request, plan_id: int):
        form = await form_of(request)
        if (form.get("name") or "").strip():
            store.rename(plan_id, form.get("name").strip())
        return RedirectResponse(f"/plans/{plan_id}/setup", status_code=303)

    @app.post("/plans/{plan_id}/archive")
    async def archive_plan(request: Request, plan_id: int):
        form = await form_of(request)
        store.archive(plan_id, form.get("archived", "1") == "1")
        return RedirectResponse("/", status_code=303)

    @app.get("/plans/{plan_id}")
    def plan_home(plan_id: int):
        return RedirectResponse(f"/plans/{plan_id}/setup", status_code=303)

    # ------------------------------------------------------------ steps
    def step_page(request: Request, plan_id: int, step: str):
        if step not in dict(STEPS) and step != "versions":
            raise HTTPException(404)
        row, plan = load(plan_id)
        return page(request, f"step_{step}.html", build_context(step, row, plan))

    def build_context(step: str, row, plan: Plan, submitted: dict[str, list[dict]] | None = None) -> dict:
        ctx = {"row": row, "plan": plan, "step": step, "issues": check_plan(plan),
               "months": month_labels(plan.settings)}
        builders = {"pricing": pricing_ctx, "products": products_ctx, "competitors": competitors_ctx,
                    "channels": channels_ctx, "volume": volume_ctx, "extras": extras_ctx, "review": review_ctx,
                    "versions": lambda p: {"versions": store.versions(row.id)}}
        if step in builders:
            ctx.update(builders[step](plan))
        if submitted:
            for g in ctx.get("grids", {}).values():
                if g.name in submitted:
                    g.rows = submitted[g.name]
        return ctx

    async def save_step(request: Request, plan_id: int, step: str):
        if step not in appliers:
            raise HTTPException(404)
        form = await form_of(request)
        row, plan = load(plan_id)
        data = plan.model_dump()
        ctx = build_context(step, row, plan)
        grids = ctx.get("grids", {})
        submitted = {g.name: parse(form, g) for g in grids.values()}
        try:
            appliers[step](data, form, submitted, plan)
            new = Plan.model_validate(prune(data))
        except (ValidationError, ValueError) as e:
            msg = _errors(e)
            ctx = build_context(step, row, plan, submitted)
            ctx.update(flash=f"Not saved: {msg}", flash_kind="error")
            return page(request, f"step_{step}.html", ctx, status=422)
        version = store.save(plan_id, new, user_of(request), note=step)
        errs = [i for i in check_plan(new) if i.level == "error"]
        flash(request, f"Saved (version {version})." + (f" {len(errs)} problem(s) to fix before generating."
                                                         if errs else ""), "warn" if errs else "ok")
        nxt = form.get("next")
        target = nxt if nxt in dict(STEPS) else step
        return RedirectResponse(f"/plans/{plan_id}/{target}", status_code=303)

    # ------------------------------------------------------------ step: setup
    def apply_setup(data, form, submitted, plan):
        s = data["settings"]
        for key in ("brand", "origin", "incoterm", "brand_currency", "hs_code"):
            s[key] = (form.get(key) or "").strip()
        s["brand_currency"] = s["brand_currency"].upper() or "USD"
        s["fx_rate"] = intake.parse_number(form.get("fx_rate"), idr=True)
        s["rounding_step"] = int(form.get("rounding_step") or 0)
        s["vat_rate"] = (intake.parse_number(form.get("vat_rate")) or 0) / 100
        for key in ("start_year", "start_month", "years"):
            s[key] = int(intake.parse_number(form.get(key)) or 0)
        s["stock_cover_months"] = intake.parse_number(form.get("stock_cover_months")) or 0

    # ------------------------------------------------------------ step: cost stack
    def pricing_ctx(plan: Plan) -> dict:
        grids, previews = {}, {}
        types = [(t.value, t.value.replace("_", " ")) for t in LineType]
        for k, st in enumerate(plan.stacks):
            grids[f"lines_{k}"] = Grid(f"lines_{k}", [
                Column("id", "ID", "hidden"),
                Column("label", "Line", width=300),
                Column("type", "Type", "select", [("pct_of_base", "% of a base"),
                                                  ("pct_of_result", "% of resulting price"),
                                                  ("fixed", "fixed amount"), ("subtotal", "subtotal")], width=170),
                Column("rate", "Rate %", "percent", width=80),
                Column("base", "Base", "text", width=120, hint="fob, running, or a line ID"),
                Column("amount", "Amount", "number", width=90),
                Column("amount_currency", "Currency", "select", [("brand", plan.settings.brand_currency),
                                                                 ("idr", "IDR")], width=90),
                Column("note", "Note", width=220),
                Column("value", f"At brand price 100", "readonly", width=120),
            ], [{**line.model_dump(), "type": line.type.value} for line in st.lines], add_label="Add line")
            try:
                res = evaluate(st, 100, plan.settings.fx_rate)
                for r in grids[f"lines_{k}"].rows:
                    r["value"] = f"{res.values.get(r['id'], 0):,.2f}"
                a, b = linear_terms(st, plan.settings.fx_rate)
                previews[k] = {"coefficient": a, "fixed": b,
                               "retail_100": res.shelf * plan.settings.fx_rate}
            except Exception:  # noqa: BLE001 - an invalid stack shows no preview; issues list explains why
                previews[k] = None
        return {"grids": grids, "previews": previews, "types": types}

    def apply_pricing(data, form, submitted, plan):
        stacks = []
        for k, st in enumerate(plan.stacks):
            if form.get(f"stack-{k}-delete") == "1" and len(plan.stacks) > 1:
                continue
            ids: set[str] = set()
            lines = []
            for r in submitted[f"lines_{k}"]:
                r["id"] = ensure_id(r, ids, r.get("label") or "line")
                r["type"] = r.get("type") or "pct_of_base"
                r["amount_currency"] = r.get("amount_currency") or "brand"
                r["note"] = r.get("note") or ""
                lines.append(r)
            cogs = form.get(f"stack-{k}-cogs") or None
            stacks.append({"id": st.id, "name": (form.get(f"stack-{k}-name") or st.name).strip(),
                           "lines": lines, "cogs_line": cogs})
        if form.get("add_stack") == "1":
            existing = {s["id"] for s in stacks}
            new = default_stack(slugify("stack", existing), f"Cost stack {len(stacks) + 1}", data).model_dump()
            stacks.append(new)
        data["stacks"] = stacks

    # ------------------------------------------------------------ step: products
    def comp_options(plan: Plan):
        return [("", "—")] + [(c.id, f"{c.brand} {c.product}".strip() + f" (for {plan.sku(c.sku).name})")
                              for c in plan.competitors]

    def products_ctx(plan: Plan) -> dict:
        stacks = [(s.id, s.name) for s in plan.stacks]
        rows = []
        for s in plan.skus:
            r = s.model_dump()
            r["rule_competitor"] = s.pricing_rule.competitor if s.pricing_rule else None
            r["rule_discount"] = s.pricing_rule.discount if s.pricing_rule else None
            r["retail"] = _fmt_idr(model.retail_from_fob(plan, s))
            r["proposed"] = _fmt_idr(model.proposed_retail(plan, s))
            ask = model.implied_brand_price(plan, s)
            r["ask"] = "" if ask is None else f"{ask:,.2f}"
            r["attrs"] = len(s.attributes)
            rows.append(r)
        cur = plan.settings.brand_currency
        grid = Grid("skus", [
            Column("id", "ID", "hidden"),
            Column("name", "Product", width=260),
            Column("size", "Size", "number", width=70),
            Column("size_label", "Size label", width=80),
            Column("unit", "Unit", width=60),
            Column("fob", f"Brand price ({cur})", "number", width=95),
            Column("target_retail", "Target retail (IDR)", "idr", width=110),
            Column("rule_competitor", "Price vs competitor", "select", comp_options(plan), width=200),
            Column("rule_discount", "% below", "percent", width=70),
            Column("stack", "Cost stack", "select", stacks, width=120),
            Column("opening_units", "Opening stock", "number", width=80),
            Column("moq", "MOQ", "number", width=70),
            Column("lead_time_months", "Lead time (m)", "number", width=70),
            Column("retail", "Retail at brand price", "readonly", width=110),
            Column("proposed", "Proposed retail", "readonly", width=110),
            Column("ask", f"Brand price to ask ({cur})", "readonly", width=110),
        ], rows, add_label="Add product")
        return {"grids": {"skus": grid}}

    def apply_products(data, form, submitted, plan):
        old = {s.id: s for s in plan.skus}
        ids: set[str] = set()
        skus = []
        for r in submitted["skus"]:
            r["id"] = ensure_id(r, ids, r.get("name"))
            comp, disc = r.pop("rule_competitor", None), r.pop("rule_discount", None)
            r["pricing_rule"] = {"competitor": comp, "discount": disc or 0.0} if comp else None
            r["opening_units"] = r.get("opening_units") or 0
            r["lead_time_months"] = int(r.get("lead_time_months") or 0)
            r["stack"] = r.get("stack") or plan.stacks[0].id
            r["attributes"] = old[r["id"]].attributes if r["id"] in old else {}
            skus.append(r)
        data["skus"] = skus

    # ------------------------------------------------------------ step: competitors
    def competitors_ctx(plan: Plan) -> dict:
        rows = []
        for c in plan.competitors:
            r = c.model_dump()
            r["matched"] = _fmt_idr(model.size_matched_price(plan, c))
            rows.append(r)
        grid = Grid("competitors", [
            Column("id", "ID", "hidden"),
            Column("sku", "For our product", "select", [(s.id, s.name) for s in plan.skus], width=220),
            Column("brand", "Brand", width=110),
            Column("product", "Competitor product", width=220),
            Column("origin", "Origin", width=90),
            Column("size", "Size", "number", width=60),
            Column("unit", "Unit", width=55),
            Column("price_idr", "Price (IDR)", "idr", width=100),
            Column("source", "Source", width=160),
            Column("registration", "BPOM / reg. no.", width=120),
            Column("notes", "Notes", width=140),
            Column("matched", "Price at our size", "readonly", width=110),
        ], rows, add_label="Add competitor")
        return {"grids": {"competitors": grid}}

    def apply_competitors(data, form, submitted, plan):
        ids: set[str] = set()
        out = []
        for r in submitted["competitors"]:
            r["id"] = ensure_id(r, ids, f"{r.get('brand') or ''} {r.get('product') or ''}")
            for k in ("brand", "origin", "source", "registration", "notes"):
                r[k] = r.get(k) or ""
            out.append(r)
        data["competitors"] = out

    # ------------------------------------------------------------ step: channels
    def channels_ctx(plan: Plan) -> dict:
        pk = [("", "—")] + [(p.id, p.name) for p in plan.packages]
        ch_rows = []
        for ch in plan.channels:
            r = ch.model_dump()
            r["effective"] = f"{model.effective_margin(plan, ch):.1%}"
            ch_rows.append(r)
        channels = Grid("channels", [
            Column("id", "ID", "hidden"),
            Column("name", "Channel", width=200),
            Column("kind", "Kind", "select", [("retail", "Retail"), ("member", "Member-based"),
                                              ("treatment", "Treatments"), ("marketplace", "Marketplace"),
                                              ("other", "Other")], width=130),
            Column("margin", "Margin %", "percent", width=80, hint="share of retail sales incl. VAT"),
            Column("package", "Treatment package", "select", pk, width=160),
            Column("effective", "Margin used", "readonly", width=90),
        ], ch_rows, add_label="Add channel")
        fees = Grid("fees", [
            Column("channel", "Channel", "select", [(c.id, c.name) for c in plan.channels], width=200),
            Column("label", "Fee", width=220),
            Column("rate", "Rate %", "percent", width=80),
            Column("plus_vat", "Plus VAT", "check", width=70),
        ], [{"channel": ch.id, **f.model_dump()} for ch in plan.channels for f in ch.margin_components],
            add_label="Add fee")
        listing_channels = [c for c in plan.channels if not c.package]
        listings = Grid("listings", [Column("id", "ID", "hidden"), Column("name", "Product", "readonly", width=260, align="left")]
                        + [Column(c.id, c.name, "number", width=100) for c in listing_channels],
                        [{"id": s.id, "name": s.name, **{c.id: c.listings.get(s.id) for c in listing_channels}}
                         for s in plan.skus], allow_add=False, allow_delete=False)
        return {"grids": {"channels": channels, "fees": fees, "listings": listings}}

    def apply_channels(data, form, submitted, plan):
        ids: set[str] = set()
        old = {c.id: c for c in plan.channels}
        fees: dict[str, list] = {}
        for f in submitted["fees"]:
            if f.get("channel") and f.get("label"):
                fees.setdefault(f["channel"], []).append(
                    {"label": f["label"], "rate": f.get("rate") or 0.0, "plus_vat": bool(f.get("plus_vat"))})
        listing = {r["id"]: r for r in submitted["listings"]}
        out = []
        for r in submitted["channels"]:
            r["id"] = ensure_id(r, ids, r.get("name"))
            r["kind"] = r.pop("kind", None) or "retail"
            r["margin_components"] = fees.get(r["id"], [])
            prev = old.get(r["id"])
            if prev and not prev.package:
                r["listings"] = {sid: row[r["id"]] for sid, row in listing.items()
                                 if row.get(r["id"]) is not None}
            else:
                r["listings"] = prev.listings if prev else {}
            if form.get("list_all") == r["id"]:
                r["listings"] = {s.id: r["listings"].get(s.id, 1.0) for s in plan.skus}
            out.append(r)
        data["channels"] = out

    # ------------------------------------------------------------ step: volume
    def volume_ctx(plan: Plan) -> dict:
        labels = month_labels(plan.settings)
        n = plan.settings.months
        grids = {}
        for d in plan.drivers:
            rows = []
            for f in d.factors:
                r = {"name": f.name, "unit": f.unit, "kind": f.kind, "value": f.value}
                for m in range(n):
                    r[f"m{m + 1}"] = f.at(m) if f.kind == "monthly" else None
                rows.append(r)
            grids[f"factors_{d.id}"] = Grid(f"factors_{d.id}", [
                Column("name", "Factor", width=210), Column("unit", "Unit", width=80),
                Column("kind", "Kind", "select", [("constant", "constant"), ("monthly", "by month")], width=95),
                Column("value", "Constant", "number", width=80),
            ] + [Column(f"m{m + 1}", labels[m], "number", width=72) for m in range(n)], rows,
                add_label="Add factor")
        names = {c.id: c.name for c in plan.channels}
        grids["overrides"] = Grid("overrides", [
            Column("channel", "Channel", "select", list(names.items()), width=180),
            Column("sku", "Product", "select", [(s.id, s.name) for s in plan.skus], width=240),
            Column("month", "Month", "select", [(str(m + 1), labels[m]) for m in range(n)], width=100),
            Column("units", "Units", "number", width=90),
        ], [{**o.model_dump(), "month": str(o.month)} for o in plan.overrides], add_label="Add override")
        preview = None
        if plan.channels and plan.drivers:
            try:
                a = model.analyse(plan)
                preview = [(y, [(names[c], a.by_year[y][c].units, a.by_year[y][c].sales) for c in a.by_year[y]],
                            a.total(y)) for y in a.by_year]
            except Exception:  # noqa: BLE001
                preview = None
        return {"grids": grids, "presets": [(k, v["label"]) for k, v in DRIVER_PRESETS.items()],
                "channel_names": names, "preview": preview}

    def apply_volume(data, form, submitted, plan):
        drivers = []
        for d in plan.drivers:
            if form.get(f"driver-{d.id}-delete") == "1":
                continue
            factors = []
            for r in submitted[f"factors_{d.id}"]:
                kind = r.get("kind") or "constant"
                values = [r.get(f"m{m + 1}") or 0.0 for m in range(plan.settings.months)]
                while len(values) > 1 and values[-1] == values[-2]:
                    values.pop()
                factors.append({"name": r.get("name") or "Factor", "unit": r.get("unit") or "", "kind": kind,
                                "value": r.get("value") if kind == "constant" else None,
                                "values": values if kind == "monthly" else []})
            if factors:
                drivers.append({"id": d.id, "channel": d.channel, "preset": d.preset,
                                "name": (form.get(f"driver-{d.id}-name") or d.name).strip(), "factors": factors})
        data["drivers"] = drivers
        data["overrides"] = [{"channel": o["channel"], "sku": o["sku"], "month": int(o["month"]),
                              "units": o.get("units") or 0} for o in submitted["overrides"]
                             if o.get("channel") and o.get("sku") and o.get("month")]

    @app.post("/plans/{plan_id}/volume/add-driver")
    async def add_driver_route(request: Request, plan_id: int):
        form = await form_of(request)
        row, plan = load(plan_id)
        channel, preset = form.get("channel"), form.get("preset") or "custom"
        if channel not in {c.id for c in plan.channels} or preset not in DRIVER_PRESETS:
            flash(request, "Pick a channel and a driver type.", "error")
        else:
            add_driver(plan, channel, preset, (form.get("name") or "").strip() or None)
            store.save(plan_id, plan, user_of(request), "add driver")
            flash(request, "Driver added. Fill in its factors below.")
        return RedirectResponse(f"/plans/{plan_id}/volume", status_code=303)

    # ------------------------------------------------------------ step: extras
    def extras_ctx(plan: Plan) -> dict:
        prods = [(s.id, s.name) for s in plan.skus]
        packages = Grid("packages", [
            Column("package", "Package / treatment", width=220),
            Column("multiplier", "Price multiplier", "number", width=95),
            Column("duration", "Minutes", "number", width=70),
            Column("sku", "Product used", "select", prods, width=240),
            Column("usage", "Usage per treatment", "number", width=110, hint="in the product's size unit"),
        ], [{"package": p.name, "multiplier": p.price_multiplier, "duration": p.duration_minutes,
             "sku": c.sku, "usage": c.usage} for p in plan.packages for c in p.components],
            add_label="Add product to a package")
        rows = []
        for c in plan.containers:
            share = model.freight_share(plan, c.cost, c.units)
            rows.append({**c.model_dump(), "share": "" if share is None else f"{share:.1%}"})
        containers = Grid("containers", [
            Column("name", "Container", width=200),
            Column("cost", f"Cost ({plan.settings.brand_currency})", "number", width=100),
            Column("units", "Units per container", "number", width=120),
            Column("share", "Freight share", "readonly", width=100),
        ], rows, add_label="Add container")
        market = Grid("market", [Column("segment", "Segment", width=220), Column("metric", "Metric", width=200),
                                 Column("value", "Value", width=120), Column("source", "Source", width=260)],
                      [m.model_dump() for m in plan.market], add_label="Add row")
        costs = {p.name: model.treatment_cost(plan, p.id) for p in plan.packages}
        return {"grids": {"packages": packages, "containers": containers, "market": market},
                "treatment_costs": costs}

    def apply_extras(data, form, submitted, plan):
        old = {p.name: p for p in plan.packages}
        pkgs: dict[str, dict] = {}
        ids = set()
        for r in submitted["packages"]:
            name = r.get("package")
            if not name or not r.get("sku") or not r.get("usage"):
                continue
            if name not in pkgs:
                pid = old[name].id if name in old else slugify(name, ids)
                ids.add(pid)
                pkgs[name] = {"id": pid, "name": name, "price_multiplier": r.get("multiplier") or 4.0,
                              "duration_minutes": int(r["duration"]) if r.get("duration") else None,
                              "components": []}
            pkgs[name]["components"].append({"sku": r["sku"], "usage": r["usage"]})
        data["packages"] = list(pkgs.values())
        data["containers"] = [r for r in submitted["containers"] if r.get("name") and r.get("cost") and r.get("units")]
        data["market"] = [{k: r.get(k) or "" for k in ("segment", "metric", "value", "source")}
                          for r in submitted["market"] if r.get("segment")]

    # ------------------------------------------------------------ step: review
    def outputs_dir(plan_id: int) -> Path:
        return data / "outputs" / str(plan_id)

    def review_ctx(plan: Plan) -> dict:
        summary = None
        if plan.channels and plan.drivers:
            try:
                a = model.analyse(plan)
                summary = [(y, a.total(y)) for y in a.by_year]
            except Exception:  # noqa: BLE001
                summary = None
        return {"summary": summary}

    def apply_review(data, form, submitted, plan):
        pass

    @app.post("/plans/{plan_id}/generate")
    async def generate_route(request: Request, plan_id: int):
        await form_of(request)
        row, plan = load(plan_id)
        folder = outputs_dir(plan_id)
        folder.mkdir(parents=True, exist_ok=True)
        fname = f"{slugify(row.name)}-v{row.version}.xlsx"
        try:
            with GENERATE_LOCK:
                result = generate(plan, folder / fname)
        except PlanInvalid as e:
            flash(request, "Fix these before generating: " + "; ".join(i.message for i in e.issues
                                                                         if i.level == "error"), "error")
            return RedirectResponse(f"/plans/{plan_id}/review", status_code=303)
        report = {"file": fname, "version": row.version, "ok": result.ok, "backend": result.audit_backend,
                  "checked": len(result.expected), "issues": [str(i) for i in result.issues],
                  "by": user_of(request)}
        (folder / (fname + ".json")).write_text(json.dumps(report, indent=2))
        if result.ok:
            flash(request, f"Workbook ready: {len(result.expected)} values checked against the engine.")
        else:
            flash(request, "The workbook failed its checks, so it is not offered for download. See the report.",
                  "error")
        return RedirectResponse(f"/plans/{plan_id}/review", status_code=303)

    def reports(plan_id: int) -> list[dict]:
        folder = outputs_dir(plan_id)
        if not folder.exists():
            return []
        out = []
        for f in sorted(folder.glob("*.xlsx.json"), key=lambda p: p.stat().st_mtime, reverse=True)[:10]:
            out.append(json.loads(f.read_text()))
        return out

    templates.env.globals["reports"] = reports

    @app.get("/plans/{plan_id}/download/{fname}")
    def download(plan_id: int, fname: str):
        folder = outputs_dir(plan_id)
        path = (folder / fname).resolve()
        if path.parent != folder.resolve() or not path.exists() or path.suffix != ".xlsx":
            raise HTTPException(404)
        report = json.loads((folder / (fname + ".json")).read_text())
        if not report.get("ok"):
            raise HTTPException(409, "This workbook failed its checks.")
        return FileResponse(path, filename=fname,
                            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    @app.post("/plans/{plan_id}/reimport")
    async def reimport(request: Request, plan_id: int, file: UploadFile = File(...)):
        form = await form_of(request)  # noqa: F841 - validates the token
        load(plan_id)
        tmp = data / "uploads" / f"{secrets.token_hex(8)}.xlsx"
        tmp.parent.mkdir(parents=True, exist_ok=True)
        tmp.write_bytes(await file.read())
        try:
            plan = read_inputs(tmp)
        except (ValueError, ValidationError, KeyError) as e:
            flash(request, f"Could not read the _Inputs sheets: {_errors(e)}", "error")
            return RedirectResponse(f"/plans/{plan_id}/review", status_code=303)
        finally:
            tmp.unlink(missing_ok=True)
        v = store.save(plan_id, plan, user_of(request), f"re-imported {file.filename}")
        flash(request, f"Inputs re-imported from {file.filename} (version {v}).")
        return RedirectResponse(f"/plans/{plan_id}/review", status_code=303)

    @app.get("/plans/{plan_id}/export.json")
    def export_json(plan_id: int):
        row, _ = load(plan_id)
        return Response(store.export_json(plan_id), media_type="application/json",
                        headers={"Content-Disposition": f'attachment; filename="{slugify(row.name)}.json"'})

    @app.post("/plans/{plan_id}/restore/{version}")
    async def restore(request: Request, plan_id: int, version: int):
        await form_of(request)
        _, plan = load(plan_id, version)
        v = store.save(plan_id, plan, user_of(request), f"restored version {version}")
        flash(request, f"Version {version} restored as version {v}.")
        return RedirectResponse(f"/plans/{plan_id}/setup", status_code=303)

    # ------------------------------------------------------------ uploads with column mapping
    KINDS = {"products": ("Products", "products"), "competitors": ("Competitors", "competitors"),
             "stores": ("Store list", "volume")}

    @app.get("/template/products.xlsx")
    def template_download():
        return Response(intake.write_template(), headers={
            "Content-Disposition": 'attachment; filename="KNS product template.xlsx"'},
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    @app.post("/plans/{plan_id}/upload/{kind}", response_class=HTMLResponse)
    async def upload(request: Request, plan_id: int, kind: str, file: UploadFile = File(...)):
        await form_of(request)
        if kind not in KINDS:
            raise HTTPException(404)
        row, plan = load(plan_id)
        content = await file.read()
        if len(content) > 15 * 1024 * 1024:
            flash(request, "That file is over 15 MB.", "error")
            return RedirectResponse(f"/plans/{plan_id}/{KINDS[kind][1]}", status_code=303)
        token = secrets.token_hex(8)
        up = data / "uploads"
        up.mkdir(parents=True, exist_ok=True)
        (up / f"{token}.bin").write_bytes(content)
        (up / f"{token}.name").write_text(file.filename or "upload.xlsx")
        return mapping_page(request, row, plan, kind, token)

    def mapping_page(request, row, plan, kind, token, sheet=None, error=""):
        up = data / "uploads"
        content, filename = (up / f"{token}.bin").read_bytes(), (up / f"{token}.name").read_text()
        try:
            table = intake.read_table(content, filename, sheet)
        except Exception as e:  # noqa: BLE001
            flash(request, f"Could not read {filename}: {e}", "error")
            return RedirectResponse(f"/plans/{row.id}/{KINDS[kind][1]}", status_code=303)
        mapping = intake.guess_mapping(kind, table.headers)
        labels = {"id": "Brand product code", "name": "Product name", "size": "Size", "unit": "Unit",
                  "fob": f"Brand price ({plan.settings.brand_currency})", "target_retail": "Target retail (IDR)",
                  "moq": "MOQ", "sku": "Our product (name or ID)", "brand": "Competitor brand",
                  "product": "Competitor product", "origin": "Origin", "price_idr": "Price (IDR)",
                  "source": "Source", "registration": "Registration", "channel": "Channel", "store": "Store",
                  "opening": "Opening month"}
        return page(request, "mapping.html", {
            "row": row, "plan": plan, "kind": kind, "kind_label": KINDS[kind][0], "token": token,
            "filename": filename, "table": table, "mapping": mapping, "labels": labels,
            "required": intake.REQUIRED[kind], "sheets": intake.sheet_names(content, filename),
            "sheet": table.sheet, "preview": table.rows[:8], "error": error, "step": KINDS[kind][1],
            "issues": [], "months": month_labels(plan.settings)})

    @app.post("/plans/{plan_id}/import/{kind}/{token}", response_class=HTMLResponse)
    async def do_import(request: Request, plan_id: int, kind: str, token: str):
        form = await form_of(request)
        if kind not in KINDS or not token.isalnum():
            raise HTTPException(404)
        row, plan = load(plan_id)
        up = data / "uploads"
        if not (up / f"{token}.bin").exists():
            flash(request, "That upload has expired. Upload the file again.", "error")
            return RedirectResponse(f"/plans/{plan_id}/{KINDS[kind][1]}", status_code=303)
        sheet = form.get("sheet") or None
        if form.get("change_sheet") == "1":
            return mapping_page(request, row, plan, kind, token, sheet)
        content, filename = (up / f"{token}.bin").read_bytes(), (up / f"{token}.name").read_text()
        table = intake.read_table(content, filename, sheet)
        mapping = {f: (form.get(f"map-{f}") or None) for f in intake.SYNONYMS[kind]}
        missing = [f for f in intake.REQUIRED[kind] if not mapping.get(f)]
        if missing:
            return mapping_page(request, row, plan, kind, token, sheet,
                                error="Choose a column for: " + ", ".join(missing))
        replace = form.get("mode") == "replace"
        if kind == "products":
            existing = set() if replace else {s.id for s in plan.skus}
            res = intake.to_skus(table, mapping, existing)
            if replace:
                plan.skus = res.items
            else:
                by_name = {s.name.strip().lower(): s for s in plan.skus}
                for item in res.items:
                    match = by_name.get(item.name.strip().lower())
                    if match:  # update the existing product, keep its id and plan settings
                        for k in ("size", "size_label", "unit", "fob", "target_retail", "moq"):
                            if getattr(item, k) is not None:
                                setattr(match, k, getattr(item, k))
                        match.attributes.update(item.attributes)
                    else:
                        item.stack = plan.stacks[0].id
                        plan.skus.append(item)
            note = f"{len(res.items)} product(s) from {filename}"
        elif kind == "competitors":
            res = intake.to_competitors(table, mapping, plan)
            plan.competitors = res.items if replace else plan.competitors + res.items
            note = f"{len(res.items)} competitor(s) from {filename}"
        else:
            counts = intake.stores_by_month(table, mapping, plan.settings)
            res = intake.Converted()
            by_name = {c.name.strip().lower(): c for c in plan.channels}
            applied = []
            for ch_name, series in counts.items():
                ch = by_name.get(ch_name.strip().lower())
                if not ch:
                    res.skipped.append(f"channel '{ch_name}' is not in this plan; add it on the Channels step")
                    continue
                drv = next((d for d in plan.drivers if d.channel == ch.id), None) or \
                    add_driver(plan, ch.id, "retail_sellthrough")
                fac = drv.factors[0]
                fac.kind, fac.values, fac.value = "monthly", series, None
                applied.append(f"{ch.name}: {fac.name} = {int(series[-1])} by the end")
            note = "store counts from " + filename
            res.items = applied
        try:
            plan = Plan.model_validate(prune(plan.model_dump()))
        except ValidationError as e:
            flash(request, f"Import not saved: {_errors(e)}", "error")
            return RedirectResponse(f"/plans/{plan_id}/{KINDS[kind][1]}", status_code=303)
        v = store.save(plan_id, plan, user_of(request), note)
        (up / f"{token}.bin").unlink(missing_ok=True)
        (up / f"{token}.name").unlink(missing_ok=True)
        msg = f"Imported {note} (version {v})."
        if res.skipped:
            msg += f" Skipped {len(res.skipped)}: " + "; ".join(res.skipped[:5]) + (
                "…" if len(res.skipped) > 5 else "")
        flash(request, msg, "warn" if res.skipped else "ok")
        return RedirectResponse(f"/plans/{plan_id}/{KINDS[kind][1]}", status_code=303)

    # generic step routes last, so the specific routes above match first
    app.add_api_route("/plans/{plan_id}/{step}", step_page, methods=["GET"], response_class=HTMLResponse)
    app.add_api_route("/plans/{plan_id}/{step}", save_step, methods=["POST"], response_class=HTMLResponse)

    appliers = {"setup": apply_setup, "pricing": apply_pricing, "products": apply_products,
                "competitors": apply_competitors, "channels": apply_channels, "volume": apply_volume,
                "extras": apply_extras, "review": apply_review}
    return app


def _errors(e: Exception) -> str:
    if isinstance(e, ValidationError):
        parts = []
        for err in e.errors()[:4]:
            loc = " › ".join(str(x) for x in err["loc"] if not isinstance(x, int))
            parts.append(f"{loc}: {err['msg']}" if loc else err["msg"])
        return "; ".join(parts)
    return str(e)


def _fmt_idr(v) -> str:
    return "" if v is None else f"{v:,.0f}"

