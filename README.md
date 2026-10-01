# KNS Business Plan Generator

Internal web app that turns a brand's product list into a market-entry business plan workbook
(.xlsx) with live Excel formulas: price structure, competitor benchmark, channel plan, monthly
forecast for 1–5 years, year analysis, summary in IDR and the brand's currency, and stock and
purchases. Works for any brand and any mix of business types (retail, gyms, spas, marketplaces).

## How a plan is built (the 8 steps in the app)

1. **Brand & market**: currency, exchange rate, VAT, price rounding, plan start and length, stock cover.
2. **Cost stack**: the price build-up from brand price to shelf price. Four line types:
   - *% of a base*: share of the brand price, the running total or an earlier line (freight, duties, VAT)
   - *% of resulting price*: a margin on the next subtotal (distributor margin, G&A, A&P, retailer margin)
   - *fixed amount*: per unit, in the brand currency or IDR
   - *subtotal*: Landed, Wholesale, Before VAT, Shelf

   New plans start from the company default stack, kept on the server in
   `<data folder>/default_stack.json` (not in this repository). Without that file, new plans use
   the example stack in `presets.py`, whose rates are placeholders. To set the default, build a
   stack in any plan, download the plan data (JSON) from the Review step and run
   `kns-plan set-default-stack plan.json`. A plan can have several stacks (e.g. food and cosmetics).
3. **Products**: upload the brand's file (any layout; columns are matched for you), or use the
   KNS product template. Unmatched columns are kept as attributes.
4. **Competitors** (optional): benchmarks per product, size-matched to our size. A product can be
   priced "X% below" a competitor; a typed target price wins.
5. **Channels**: margin per channel, or fees for marketplaces; which products each channel carries.
6. **Volume**: each channel has drivers; a driver's monthly volume is its factors multiplied
   (stores × units per store × probability, members × buying rate, spas × treatments…). Presets
   for retail, member-based, treatments and marketplaces; a store list can be uploaded. Any
   month can be overridden.
7. **Extras** (optional): treatment packages (costing per treatment), container freight, market notes.
8. **Review & generate**: checks, headline numbers, and the workbook.

Every save is a new version (nothing is overwritten). The workbook's `_Inputs` sheets can be
edited in Excel and re-imported; edits to calculation sheets are not read back.

## Every workbook is checked before download

The app computes every figure in Python and writes the workbook as formulas. Before a download is
offered, the workbook is recalculated (LibreOffice on the server) and each key figure is compared
with the Python result. Any `#REF!`/`#DIV/0!` or mismatch blocks the download.

## Development

```
python -m venv .venv
.venv\Scripts\python -m pip install -e ".[web,dev]"
.venv\Scripts\python -m pytest
```

Command line:

```
kns-plan new --brand "Brand" --fx 18200 -o plan.json     # blank plan with the KNS default stack
kns-plan generate plan.json -o out/plan.xlsx               # workbook, recalculated and checked
kns-plan import out/plan.xlsx -o plan.json                 # read the _Inputs sheets back
kns-plan serve --data data-dev                             # web app on http://127.0.0.1:8400
```

For local testing, `dev/dev_password.txt` holds the password for the `data-dev` folder, and
`tests/private/` can hold checks against real plans; all three are ignored by git. Without LibreOffice, checks run with the `formulas` package instead and the
downloaded file has no cached values until opened in Excel.

Deployment on the office server: see [docs/DEPLOY.md](docs/DEPLOY.md).

## Code map

| Path | What |
|---|---|
| `src/kns_plan/schema.py` | Plan data model (the JSON saved for each version) |
| `src/kns_plan/stack.py` | Cost stack arithmetic, retail coefficient, brand price to ask for |
| `src/kns_plan/model.py` | Proposals, volumes, year analysis, stock, treatment costing |
| `src/kns_plan/validate.py` | Checks shown in the app; errors block generating |
| `src/kns_plan/excel/` | Workbook writer: `_Inputs` sheets plus one module per output sheet |
| `src/kns_plan/recalc.py` | LibreOffice / Python recalculation for the pre-download check |
| `src/kns_plan/intake.py` | File upload, column matching, KNS product template, store lists |
| `src/kns_plan/web/` | FastAPI app, storage (SQLite, versioned), login with lockout |
