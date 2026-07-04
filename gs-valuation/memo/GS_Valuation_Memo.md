# Goldman Sachs Group, Inc. (NYSE: GS) — Valuation Memo

**Prepared for:** Corporate Valuation Suite — Portfolio Project
**Sector:** Financials — Global Investment Bank / Diversified Financial Services
**Data as-of date:** July 2, 2026
**Current price:** $1,022.11 | **Market cap:** ~$301.5bn | **52-week range:** $691.30 – $1,125.00

> **Disclaimer:** This is a portfolio / educational analysis, not investment research, and is not a
> recommendation to buy, sell, or hold any security. All figures are sourced from public filings and
> secondary market-data sites as of the dates noted throughout the accompanying workbooks and may
> contain staleness or transcription error — verify against primary sources (SEC EDGAR) before relying
> on any of this for a real decision. The author is not a licensed financial advisor.

---

## 1. Executive Summary

Goldman Sachs enters mid-2026 coming off its strongest two-year stretch since the 2021 capital-markets
peak: net revenues grew from $46.25bn (FY2023) to $58.28bn (FY2025), net earnings more than doubled from
$8.52bn to $17.18bn, and ROE recovered from a 2023 trough of 7.5% to 15.0% in FY2025 — back inside the
firm's own 14–16% through-cycle target band. The stock has re-rated accordingly, up roughly 78% over the
trailing 52 weeks to an all-time high near $1,106 in late June 2026.

That re-rating is the crux of this memo. Every valuation lens built in this repository — a bottoms-up
FCFE DCF, a reverse DCF, a full Bear/Base/Bull scenario range, a comparable-company analysis, and a
20,000-trial Monte Carlo simulation — lands **below** the current $1,022 market price, including under
this model's own Bull Case ($898/share). The consistent conclusion across methods is not "the stock is
mispriced" (that is not a claim this model is positioned to make with confidence) but rather: **the
market is currently pricing in a more optimistic multi-year path than even a constructive scenario in
this model assumes**, and understanding *why* an investor might reasonably hold that more optimistic view
is the real analytical question, addressed in Section 6.

| Method | Implied value / share | vs. current price |
|---|---:|---:|
| DCF — Bear Case | $412.66 | (59.6)% |
| DCF — **Base Case** | **$631.24** | **(38.2)%** |
| DCF — Bull Case | $897.95 | (12.1)% |
| Comps — blended (P/B & P/E, peer mean/median) | $670.96 | (34.3)% |
| Monte Carlo — median (P50) | $609.62 | (40.4)% |
| Monte Carlo — P95 (upper tail) | $742.14 | (27.4)% |

---

## 2. Historical Financial Analysis (see `01_Historical_Financial_Analysis.xlsx`)

- **Net revenues** grew from $59.34bn (2021, a cyclical peak) to a 2023 trough of $46.25bn, then
  recovered to $58.28bn by 2025 — essentially round-tripping the 2021 peak in nominal terms after a
  sharp down-cycle.
- **ROE** tells the more dramatic story: 23.0% (2021) → 10.2% (2022) → 7.5% (2023) → 12.7% (2024) →
  15.0% (2025). The 2022–2023 trough coincided with the wind-down of GS's consumer strategy (GreenSky,
  the GM and Apple Card partnerships, Marcus) and a slow capital-markets environment; both headwinds have
  since faded.
- **Book value per share** grew every single year in this window regardless of the earnings cycle —
  $284.39 (2021) to $357.60 (2025), a 5.9% CAGR — because GS retains enough earnings to grow capital even
  in a soft year. Combined with a shrinking share count (355.0mm diluted shares in 2021 to 300.6mm in
  2025, a cumulative 15.3% reduction), BVPS growth has been a more reliable value-creation engine over
  this window than the earnings cycle itself.
- **Efficiency ratio** improved from 74.6% (2023) to 63.1% (2024), a meaningful operating-leverage
  story layered on top of the revenue recovery.

**Read-through for the model:** because 2021 was a genuine cyclical outlier and 2023 a genuine trough,
this model's Base Case deliberately does NOT extrapolate the 2024–2025 recovery rate forward — see
Section 4.

---

## 3. Three-Statement (Equity Roll-Forward) Model (see `02_Three_Statement_Model.xlsx`)

GS is a bank holding company, not an industrial firm, so this model uses the industry-standard bank
equivalent of a three-statement model: an integrated **Income Statement → Common Equity Roll-Forward →
Diluted Shares Walk**, fully linked, with a scenario toggle (`Assumptions!B3`) that flows through every
downstream tab. Assets and liabilities are a bank's financial inventory, not operating capital — modeling
them at the same granularity as a manufacturer's working capital would add complexity without adding
insight; capital adequacy (CET1, leverage ratio) is the real constraint, and this model proxies it via a
balance-sheet/RWA growth assumption that drives required equity retention.

**Base Case mechanics (FY2026E–FY2030E):**
- Revenue growth decelerates from 6.0% (2026E) to 4.5% (2030E) — a normalization off the 2025 cyclical
  strength, not a continuation of it.
- Efficiency ratio improves modestly from 63% to 61% on continued operating leverage.
- Payout ratio (dividends + buybacks) of 65% of net income to common, split ~28% dividends / ~72%
  buybacks, consistent with GS's actual FY2024–FY2025 capital-return mix.
- Result: diluted EPS grows from $57.05 (2026E) to $81.53 (2030E); BVPS grows from $387.30 to $538.46;
  ROE holds in a 15.3%–15.9% band, in line with the FY2025 actual.

This chain was independently re-implemented in Python (`scripts/validate_three_statement.py`) and matches
the Excel model to within 0.1% on every output field across all five years — see Section 7.

---

## 4. DCF and Reverse DCF (see `03_DCF_and_Reverse_DCF.xlsx`)

### 4.1 DCF methodology

Because GS is a bank, this model uses **Free Cash Flow to Equity (FCFE)**, not unlevered FCFF — the
standard adaptation for financial-services DCF (see Damodaran, *"Valuing Financial Service Firms"*).
FCFE = Net income to common − required equity retention to fund balance-sheet/RWA growth at a constant
capital ratio. Five years of explicit FCFE (2026E–2030E) are discounted at the CAPM cost of equity
(Base Case: Rf 4.48% + Beta 1.30 × ERP 5.0% ≈ 10.8%), plus a Gordon-growth terminal value at a 3.0%
terminal growth rate, discounted back 5 years, summed, and divided by current diluted shares outstanding
(295.0mm).

**Base Case result: $631.24/share, a 38.2% discount to the current $1,022.11 price.**

Terminal value is 70.9% of total equity value in the Base Case — a normal proportion for a 5-year
explicit window but a reminder that this valuation is highly sensitive to the terminal-growth and
cost-of-equity assumptions (see Sensitivity Tables, Section 5).

### 4.2 Reverse DCF — what does the current price already assume?

The more actionable question for a stock near all-time highs is not "what is it worth given my
assumptions" but "what does the price already assume, and is that reasonable?" Two independent solves:

- **Single-stage closed-form** (Excel, `Reverse DCF` tab): solving Price = FCFE₁/(Ke − g) for g given the
  current price implies a **perpetual FCFE growth rate of 6.6%** at the Base Case 10.8% cost of equity.
  Translated through the sustainable-growth identity (g = ROE × retention ratio, at a 35% Base Case
  retention ratio), that implies a **required through-cycle ROE of ~18.9%** — above the FY2025 actual
  (15.0%) but below the FY2021 cyclical peak (23.0%). This is an aggressive-but-not-unprecedented
  assumption; GS has hit that level of ROE before, just not persistently.

- **Multi-stage numeric solve** (`scripts/validate_dcf.py`, bisection): a more intuitive framing —
  what constant *revenue growth rate*, sustained for the full 5-year explicit window (before reverting to
  a 3.0% terminal rate), justifies the current price? Answer: **14.3% per year for 5 straight years** —
  roughly 2.4x the Base Case revenue growth rate and above even this model's Bull Case (11.0% in year
  one, tapering to 6.5%). This is the more demanding of the two framings, because it is not accompanied
  by any offsetting improvement in efficiency ratio or reduction in required capital retention.

**Takeaway:** the current price is not obviously irrational (an 18.9% through-cycle ROE is inside GS's
demonstrated historical range), but it is pricing in sustained outperformance relative to both the
FY2025 actual and this model's own Bull Case — an underwriting question, not a valuation-mechanics one.

---

## 5. Sensitivity Analysis (see `03_DCF_and_Reverse_DCF.xlsx`, `Sensitivity Tables` tab)

The 2-way grid (cost of equity 8.5%–12.5% × terminal growth 1.5%–4.0%) shows implied value/share ranging
from **$465.51** (Ke 12.5%, g 1.5%) to **$1,065.31** (Ke 8.5%, g 4.0%) — a wide range that underscores how
much of bank valuation lives in the discount-rate and terminal-growth assumptions rather than the
explicit-period forecast. Only the most aggressive corner of the grid (Ke ≤ 9.0%, g ≥ 3.5%) approaches
the current market price, which would require the market to be assigning GS a cost of equity meaningfully
below this model's CAPM-derived 10.8% — plausible if investors view GS's earnings as structurally less
volatile post-2022 (narrower consumer exposure, more durable financing revenue mix) than its pre-2022
history would suggest, but not something this model independently verifies.

One-way sensitivities (holding other drivers at Base Case) confirm cost of equity is the more powerful
lever: a 100bp move in Ke shifts implied value by roughly $45–60/share around the Base Case, versus
roughly $25–30/share for an equivalent 50bp move in terminal growth.

---

## 6. Comparable Company Analysis (see `04_Comparable_Company_Analysis.xlsx`)

Peer set: JPMorgan Chase, Morgan Stanley, Bank of America, Citigroup — the same primary peer group GS
references in its own investor materials for total-shareholder-return comparisons.

Applying peer P/B and P/E multiples to GS's own book value and TTM EPS gives a **blended implied price of
$670.96**, a 34.3% discount to the current price — directionally consistent with (though somewhat higher
than) the DCF Base Case. Two important caveats on this section, flagged explicitly in the workbook:

1. **Data completeness:** clean, sourced book-value-per-share figures were only available for GS and
   Citigroup in this pass, so the peer P/B statistic is effectively a single data point (Citigroup,
   ~1.25x). JPM, MS, and BAC book values need to be pulled from a terminal or their latest 10-Qs before
   this section should be treated as decision-grade. Similarly, peer P/E draws on only JPM and MS.
2. **Multiple basis mismatch:** peer multiples are trailing-twelve-month; GS's own historical multiples
   in this workbook are shown on both a FY2025A and TTM basis, which will differ modestly from
   quarter-mix effects — treat all comps output here as directional.

Directionally, though, the comps read is the same as the DCF read: **GS's current price sits above what
a relative-value framework built on its own historical peer group would suggest**, consistent with the
market either (a) awarding GS a premium multiple for its superior ROE/ROTE profile relative to peers, or
(b) simply re-rating alongside a broader bank-sector rally (JPM, MS, and BAC are all also near multi-year
highs as of this snapshot) that a peer-relative framework cannot, by construction, capture.

---

## 7. Validation & Reproducibility (see `scripts/`)

Three Python scripts independently validate the Excel models rather than simply reformatting their
output:

- **`validate_three_statement.py`** — rebuilds the entire Income Statement + Capital & Shares roll-forward
  chain from scratch in Python and diffs it against the calculated values in the Excel workbook.
  **Result: exact match (< 0.1% on every field, every year).**
- **`validate_dcf.py`** — (1) independently recomputes the Base Case DCF and confirms it matches Excel to
  the cent ($631.24 both), and (2) solves the multi-stage reverse DCF numerically via bisection (14.3%
  implied 5-year revenue growth), extending beyond what the closed-form Excel tab can solve on its own.
- **`monte_carlo_sensitivity.py`** — runs a 20,000-trial Monte Carlo simulation drawing revenue growth,
  efficiency ratio, cost of equity, and terminal growth from triangular distributions anchored to the
  Bear/Base/Bull range already built into the model. Result: **median implied price $609.62, and 0% of
  20,000 trials produced an implied price above the current $1,022.11 market price** — i.e., within the
  full Bear-to-Bull range this model considers plausible, none of it reaches today's price. That is the
  single most important quantitative finding in this project: it is not that the DCF Base Case happens to
  be conservative, it is that **the entire modeled scenario space falls short of the current market
  price**, which sharpens Section 4.2's question from "is 18.9% ROE plausible" to "is the market
  effectively underwriting a scenario more bullish than this analyst's own Bull Case, and if so, on what
  basis" — precisely the kind of question that belongs in a buy-side or sell-side thesis, not a mechanical
  model output.

---

## 8. Key Risks and Limitations of This Analysis

- **Market-data staleness:** all current price/shares/peer figures are a snapshot as of ~July 2, 2026,
  compiled from public secondary sources that disagree with each other by low single digits depending on
  exact pull time (documented in each workbook's Data Sources tab). Refresh from a terminal before using
  this for a live decision.
- **Simplified capital framework:** the FCFE required-retention mechanic (retention = BS/RWA growth ×
  beginning equity) is a standard simplification, not a full CET1/SLR/stress-capital-buffer build. A
  bank-specific analyst role would extend this with GS's actual Note 20 (Regulation and Capital Adequacy)
  disclosures.
- **Assumption provenance:** the Bear/Base/Bull driver values are this analyst's own construction,
  calibrated to GS's 2021–2025 historical range — they are not sourced from a specific sell-side note and
  should be clearly labeled as such in any interview discussion of this project.
- **Comps data gaps:** as flagged in Section 6, peer book-value data is incomplete; treat the comps
  section as directional pending a full terminal pull.
- **This is not a recommendation.** The consistent "current price is above every method's output" finding
  describes a valuation gap, not a trading signal — closing that gap could as easily happen through
  continued fundamental outperformance (GS re-rating structurally higher, matching the market's implied
  view) as through price correction.

---

## 9. Repository Map

```
gs-valuation/
├── README.md
├── memo/
│   └── GS_Valuation_Memo.md          <- this file
├── models/
│   ├── 01_Historical_Financial_Analysis.xlsx
│   ├── 02_Three_Statement_Model.xlsx
│   ├── 03_DCF_and_Reverse_DCF.xlsx
│   └── 04_Comparable_Company_Analysis.xlsx
├── scripts/
│   ├── validate_three_statement.py
│   ├── validate_dcf.py
│   └── monte_carlo_sensitivity.py
└── data/
    └── assumptions.py                 <- single source of truth for all hardcoded inputs
```
