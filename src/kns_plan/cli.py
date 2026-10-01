"""Command line: generate a workbook from a plan file, or read a plan back from a workbook.

    kns-plan generate plan.json -o out/plan.xlsx
    kns-plan import out/plan.xlsx -o plan.json
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .excel import PlanInvalid, generate, read_inputs
from .presets import default_stack, save_default_stack
from .schema import Plan, Settings


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="kns-plan")
    sub = parser.add_subparsers(dest="command", required=True)
    g = sub.add_parser("generate", help="plan JSON → workbook")
    g.add_argument("plan", type=Path)
    g.add_argument("-o", "--output", type=Path, required=True)
    g.add_argument("--no-audit", action="store_true", help="skip recalculation and checks")
    i = sub.add_parser("import", help="workbook _Inputs sheets → plan JSON")
    i.add_argument("workbook", type=Path)
    i.add_argument("-o", "--output", type=Path, required=True)
    n = sub.add_parser("new", help="start a blank plan with the KNS default cost stack")
    n.add_argument("--brand", required=True)
    n.add_argument("--currency", default="USD")
    n.add_argument("--fx", type=float, required=True, help="IDR per 1 unit of brand currency")
    n.add_argument("-o", "--output", type=Path, required=True)
    n.add_argument("--data", type=Path, default=None, help="data folder holding default_stack.json")
    ds = sub.add_parser("set-default-stack", help="make a plan's cost stack the default for new plans")
    ds.add_argument("plan", type=Path, help="plan JSON (e.g. downloaded from the app)")
    ds.add_argument("--stack", default=None, help="stack ID in that plan (default: the first)")
    ds.add_argument("--data", type=Path, default=None, help="data folder (default: KNS_DATA_DIR or ./data)")
    sv = sub.add_parser("serve", help="run the web app")
    sv.add_argument("--host", default="127.0.0.1", help="use 0.0.0.0 only behind the reverse proxy")
    sv.add_argument("--port", type=int, default=8400)
    sv.add_argument("--data", type=Path, default=None, help="data folder (default: KNS_DATA_DIR or ./data)")
    sub.add_parser("hash-password", help="make KNS_PASSWORD_HASH for the shared office password")
    bk = sub.add_parser("backup", help="copy the plan database (run nightly from Task Scheduler)")
    bk.add_argument("--data", type=Path, default=None)
    bk.add_argument("--to", type=Path, required=True, help="backup folder, ideally on another disk")
    bk.add_argument("--keep", type=int, default=30)
    args = parser.parse_args(argv)

    if args.command == "serve":
        import uvicorn

        from .web.app import create_app
        uvicorn.run(create_app(args.data), host=args.host, port=args.port, proxy_headers=True,
                    forwarded_allow_ips="127.0.0.1")
        return 0

    if args.command == "hash-password":
        import getpass

        from .web.auth import hash_password
        pw = getpass.getpass("Office password: ")
        if len(pw) < 10 or pw != getpass.getpass("Again: "):
            print("Passwords must match and be at least 10 characters.", file=sys.stderr)
            return 2
        print("Set this as the KNS_PASSWORD_HASH environment variable on the server:")
        print(hash_password(pw))
        return 0

    if args.command == "backup":
        import os

        from .web.store import Store
        data = args.data or Path(os.environ.get("KNS_DATA_DIR", "data"))
        target = Store(data / "plans.sqlite").backup(args.to, args.keep)
        print(f"backed up to {target}")
        return 0

    if args.command == "set-default-stack":
        import json
        import os

        raw = json.loads(args.plan.read_text(encoding="utf-8"))
        plan = Plan.model_validate(raw.get("plan", raw))  # accepts the app's export too
        stack = plan.stack(args.stack) if args.stack else plan.stacks[0]
        data = args.data or Path(os.environ.get("KNS_DATA_DIR", "data"))
        print(f"default stack for new plans: '{stack.name}' → {save_default_stack(stack, data)}")
        return 0

    if args.command == "new":
        plan = Plan(settings=Settings(brand=args.brand, brand_currency=args.currency, fx_rate=args.fx),
                    stacks=[default_stack(data_dir=args.data)])
        args.output.write_text(plan.model_dump_json(indent=2, exclude_none=True), encoding="utf-8")
        print(f"wrote {args.output}: add products under \"skus\", then run `kns-plan generate`")
        return 0

    if args.command == "generate":
        plan = Plan.model_validate_json(args.plan.read_text(encoding="utf-8"))
        try:
            result = generate(plan, args.output, audit=not args.no_audit)
        except PlanInvalid as e:
            print(f"Plan has errors, nothing generated:\n{e}", file=sys.stderr)
            return 2
        for issue in result.issues:
            print(issue)
        status = "OK" if result.ok else "FAILED CHECKS"
        checked = f", checked with {result.audit_backend}" if result.audit_backend else ""
        print(f"{status}: wrote {result.path} ({len(result.expected)} values cross-checked{checked})")
        return 0 if result.ok else 1

    plan = read_inputs(args.workbook)
    args.output.write_text(plan.model_dump_json(indent=2, exclude_none=True), encoding="utf-8")
    print(f"wrote {args.output}: {len(plan.skus)} products, {len(plan.stacks)} cost stack(s)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
