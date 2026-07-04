"""
Goldman Sachs Group, Inc. (NYSE: GS) — Corporate Valuation Suite
Central data assumptions module. All hardcoded inputs sourced below;
Excel-building scripts import from here to keep every workbook consistent.

SOURCES (see README.md "Data Sources" for full citation list):
- GS FY2021-2025 Full Year & Q4 Earnings Results press releases / 8-Ks (SEC EDGAR)
- GS FY2023-2025 Form 10-K (SEC EDGAR)
- GS 4Q25 Earnings Results Presentation (Jan 15, 2026)
- Market data snapshot: week of June 29 - July 2, 2026 (Google Finance, StockAnalysis.com,
  Robinhood, MacroTrends, WallStreetZen, GuruFocus, Yahoo Finance) — NOTE: secondary market-data
  aggregators disagree by low single-digit percentages depending on exact pull time; snapshot
  values below were chosen for internal consistency (price x shares ~= market cap) and should be
  refreshed from a terminal (Bloomberg/CapIQ/FactSet) before using this model for live decisions.
- 10-Year US Treasury yield: 4.48% (Jul 2, 2026, US Treasury / Trading Economics)
"""

AS_OF_DATE = "2026-07-02"

# ---------------------------------------------------------------------------
# GS Historical Financials, FY2021-FY2025 ($ in millions unless noted)
# Source: GS Full Year & Q4 Earnings Results press releases (8-K Ex-99.1/99.2), each respective year
# ---------------------------------------------------------------------------
GS_HISTORY = {
    2021: dict(net_revenues=59339, net_earnings=21635, diluted_eps=59.45, roe=0.230, rote=0.243,
               bvps=284.39, efficiency_ratio=None, opex=None,
               source="GS 4Q21 Earnings Release, Jan 18 2022, SEC 8-K Ex-99.1/99.2"),
    2022: dict(net_revenues=47365, net_earnings=11261, diluted_eps=30.06, roe=0.102, rote=0.110,
               bvps=303.55, efficiency_ratio=None, opex=None,
               source="GS 4Q22 Earnings Release, Jan 17 2023, SEC 8-K Ex-99.2"),
    2023: dict(net_revenues=46254, net_earnings=8516, diluted_eps=22.87, roe=0.075, rote=0.081,
               bvps=313.56, efficiency_ratio=0.746, opex=34500,
               source="GS 4Q23 Earnings Release, Jan 16 2024, SEC 8-K Ex-99.1"),
    2024: dict(net_revenues=53510, net_earnings=14281, diluted_eps=40.54, roe=0.127, rote=0.137,
               bvps=336.72, efficiency_ratio=0.631, opex=33770,
               source="GS 4Q24 Earnings Release, Jan 15 2025, SEC 8-K Ex-99.1; BVPS derived from "
                      "disclosed +6.2% YoY growth to $357.60 in FY25 release"),
    2025: dict(net_revenues=58280, net_earnings=17180, diluted_eps=51.32, roe=0.150, rote=0.160,
               bvps=357.60, efficiency_ratio=None, opex=None,
               source="GS 4Q25 Earnings Release, Jan 15 2026, SEC 8-K / FY2025 10-K"),
}

# Preferred dividends (approx, $mm) — used to bridge net earnings to net earnings to common
GS_PREFERRED_DIV = {2021: 794, 2022: 858, 2023: 898, 2024: 850, 2025: 900}

# Diluted shares outstanding, period-end / weighted-avg approx (millions)
GS_SHARES_HISTORY = {2021: 355.0, 2022: 347.6, 2023: 341.9, 2024: 327.3, 2025: 300.6}
GS_SHARES_CURRENT_MM = 295.0  # source: Google Finance snapshot, ~Jul 2026

# ---------------------------------------------------------------------------
# Current market data (snapshot as of AS_OF_DATE)
# ---------------------------------------------------------------------------
GS_PRICE = 1022.11        # Robinhood close, 2026-07-02
GS_DIV_QUARTERLY = 4.51   # Google Finance, declared quarterly dividend
GS_BETA = 1.30            # Google Finance
GS_52WK_HIGH = 1125.00
GS_52WK_LOW = 691.30

RISK_FREE_RATE = 0.0448   # 10Y UST, Jul 2 2026 (Trading Economics / MacroMicro)
EQUITY_RISK_PREMIUM = 0.050  # standard long-run US ERP assumption (Damodaran-style)

TOTAL_ASSETS_APPROX = 1809000  # $mm, ~$1.809T, companiesmarketcap.com, 2026 snapshot

# ---------------------------------------------------------------------------
# Peer comparable companies (snapshot as of AS_OF_DATE, various sources — see README)
# ---------------------------------------------------------------------------
PEERS = {
    "JPM": dict(name="JPMorgan Chase & Co.", price=333.70, shares_mm=2680.0,
                ttm_eps=20.88, roe=0.1646, bvps=None, div_yield=None,
                source="StockAnalysis.com statistics, Jun 2026; MacroTrends P/E Jun 24 2026"),
    "MS":  dict(name="Morgan Stanley", price=212.68, shares_mm=1320.0,
                ttm_eps=12.69, roe=None, bvps=None, div_yield=None,
                source="MacroTrends market cap Jun 19 2026 ($280.92B); MacroTrends P/E Apr 3 2026 (16.76x)"),
    "BAC": dict(name="Bank of America Corp.", price=55.16, shares_mm=7096.6,
                ttm_eps=None, roe=None, bvps=None, div_yield=None,
                source="WallStreetZen, Jun 11 2026"),
    "C":   dict(name="Citigroup Inc.", price=140.00, shares_mm=1705.6,
                ttm_eps=None, roe=0.0765, bvps=112.22, div_yield=None,
                source="CompaniesMarketCap Jul 2026 (mkt cap $238.71B / shares); "
                       "BVPS from Citi 1Q26 8-K Ex-99.1; ROE(ttm) StockAnalysis.com"),
}

# ---------------------------------------------------------------------------
# Projection assumptions — Base Case (2026E-2030E)
# ---------------------------------------------------------------------------
PROJECTION_YEARS = [2026, 2027, 2028, 2029, 2030]

SCENARIOS = {
    "Bear": dict(rev_growth=[-0.03, 0.01, 0.02, 0.02, 0.02],
                 efficiency_ratio=[0.68, 0.67, 0.66, 0.66, 0.66],
                 tax_rate=0.23, payout_ratio=0.55, bs_growth=0.02,
                 terminal_growth=0.020, cost_of_equity=0.115),
    "Base": dict(rev_growth=[0.06, 0.05, 0.05, 0.045, 0.045],
                 efficiency_ratio=[0.63, 0.62, 0.615, 0.61, 0.61],
                 tax_rate=0.22, payout_ratio=0.65, bs_growth=0.04,
                 terminal_growth=0.030, cost_of_equity=0.108),
    "Bull": dict(rev_growth=[0.11, 0.09, 0.08, 0.07, 0.065],
                 efficiency_ratio=[0.60, 0.585, 0.575, 0.57, 0.565],
                 tax_rate=0.21, payout_ratio=0.75, bs_growth=0.06,
                 terminal_growth=0.035, cost_of_equity=0.101),
}

PREFERRED_DIV_PROJECTED = 900  # $mm/yr, held flat at 2025 level across projection
