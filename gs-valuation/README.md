# Goldman Sachs (NYSE: GS) — Corporate Valuation Suite

A full sell-side-style valuation package for The Goldman Sachs Group, Inc., built as a portfolio project
demonstrating financial modeling, valuation, and quantitative-validation skills for equity research,
investment banking, corporate finance, and asset management roles.

> **Not investment research or advice.** See `memo/GS_Valuation_Memo.md` for the full disclaimer. All
> figures are sourced from public filings and secondary market-data sites as of the dates noted in each
> workbook's "Data Sources" tab.

## What's in here

| Deliverable | Location | Description |
|---|---|---|
| Historical financial analysis | `models/01_Historical_Financial_Analysis.xlsx` | FY2021–FY2025 income statement, ROE/ROTE/efficiency ratio trends, per-share & capital-return history, segment revenue |
| Integrated three-statement model | `models/02_Three_Statement_Model.xlsx` | Bank-appropriate Income Statement → Equity Roll-Forward → Shares Outstanding model, FY2026E–FY2030E, with a live Bear/Base/Bull scenario toggle |
| DCF + Reverse DCF + Sensitivity | `models/03_DCF_and_Reverse_DCF.xlsx` | FCFE-based DCF, closed-form reverse DCF, 2-way and 1-way sensitivity tables, independent Bull/Base/Bear valuation |
| Comparable company analysis | `models/04_Comparable_Company_Analysis.xlsx` | Trading comps vs. JPM/MS/BAC/C, implied valuation via peer multiples |
| Valuation memo | `memo/GS_Valuation_Memo.md` | Full write-up synthesizing every analysis into one investment view |
| Python validation scripts | `scripts/` | Independent Python re-implementations that cross-check every Excel output, plus a Monte Carlo simulation |
| Central assumptions file | `data/assumptions.py` | Single source of truth for every hardcoded input, with inline source citations |

## Headline finding

Every valuation method in this repo — DCF Base Case, comps, and a 20,000-trial Monte Carlo simulation —
lands **30–40% below** the current market price, and even this model's own Bull Case DCF falls short of
today's price. A reverse DCF shows the market is pricing in either a ~19% through-cycle ROE (above the
FY2025 actual of 15.0%, but inside GS's historical range) or sustained ~14%/year revenue growth for five
years (above this model's own Bull Case). See Section 6 of the memo for the full discussion — this is a
valuation-gap finding, not a trading recommendation.

## How the models are built

GS is a bank holding company, not an industrial firm, so this suite deliberately does **not** force a
traditional working-capital-driven three-statement model onto it. Assets and liabilities are a bank's
financial inventory, not operating capital; the real constraint is regulatory capital. Instead:

- The "three-statement model" is the industry-standard bank equivalent: **Income Statement → Common
  Equity Roll-Forward → Diluted Shares Walk**, fully linked.
- The DCF uses **Free Cash Flow to Equity (FCFE)**, per Aswath Damodaran's financial-services DCF
  framework: FCFE = Net income to common − required equity retention to fund balance-sheet/RWA growth at
  a constant capital ratio.
- Every workbook has a live scenario selector (`Assumptions!B3`, type `Bear`/`Base`/`Bull`) that flows
  through every downstream formula — this is not three static snapshots, it's one live model with a
  switch.

## Running the validation scripts

```bash
cd scripts
pip install numpy --break-system-packages   # if not already installed

python validate_three_statement.py   # rebuilds the 3-statement model in pure Python, diffs vs Excel
python validate_dcf.py               # cross-checks the DCF + solves the multi-stage reverse DCF numerically
python monte_carlo_sensitivity.py 20000   # Monte Carlo simulation over key DCF drivers
```

All three scripts read directly from the `.xlsx` files in `models/` (via `openpyxl`, `data_only=True`),
so they validate the *actual calculated Excel output*, not a re-statement of the formulas — a real
build-it-twice sanity check.

## Recalculating the Excel models after editing assumptions

If you change any blue (hardcoded) input cell — most usefully the scenario selector on the `Assumptions`
tab of workbook 2 or 3 — reopen the file in Excel or LibreOffice, or recalculate headlessly:

```bash
python /mnt/skills/public/xlsx/scripts/recalc.py models/02_Three_Statement_Model.xlsx
python /mnt/skills/public/xlsx/scripts/recalc.py models/03_DCF_and_Reverse_DCF.xlsx
```

(Outside this sandboxed environment, use any headless LibreOffice recalculation approach, or simply open
and save the file in Excel — formulas are stored but not evaluated by `openpyxl` until recalculated.)

## Excel model conventions

Standard IB color coding is used throughout:
- **Blue** = hardcoded input (source-cited, either inline or in a Data Sources tab)
- **Black** = formula / calculation
- **Green** = link pulled from another sheet in the same workbook
- Yellow highlight = key assumption cell (e.g., the scenario selector)

## Known limitations (see memo Section 8 for full discussion)

- Market-data snapshot as of ~July 2, 2026; secondary sources disagree by low single digits — refresh
  from a terminal (Bloomberg/CapIQ/FactSet) before using for a live decision or interview prep.
- Comps peer book-value data is incomplete (sourced for GS and Citigroup only) — flagged explicitly in
  the workbook and memo rather than silently estimated.
- Bear/Base/Bull driver assumptions are this analyst's own construction calibrated to GS's 2021–2025
  historical range, not sourced from a specific sell-side research note.

## Suggested next extensions

- Pull full peer book-value data from a terminal to complete the comps section
- Extend the FCFE required-retention mechanic with GS's actual CET1/SLR/stress-capital-buffer targets
  (Form 10-K Note 20, Regulation and Capital Adequacy)
- Add a segment-level (Global Banking & Markets / Asset & Wealth Management / Platform Solutions)
  sum-of-the-parts valuation using the FY2025 10-K Note 25 recast table
