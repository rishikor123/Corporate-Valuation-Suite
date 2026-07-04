"""
monte_carlo_sensitivity.py
===========================
Monte Carlo simulation of the DCF: draws revenue growth, efficiency ratio,
cost of equity, and terminal growth from probability distributions centered
on the Base Case (with spreads informed by the Bear/Bull cases already built
into the model) and produces a distribution of implied GS share prices.

This complements the deterministic Bear/Base/Bull scenarios (which are three
discrete points) with a continuous probability picture -- e.g. "what's the
probability the DCF implies a price above today's market price?"

Requires: numpy (pip install numpy --break-system-packages if missing)

Usage:
    python monte_carlo_sensitivity.py [n_trials]
"""
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'data'))
from assumptions import *
import numpy as np


def simulate(n_trials=20000, seed=42):
    rng = np.random.default_rng(seed)
    base = SCENARIOS['Base']

    # Triangular distributions anchored on Bear/Base/Bull for each driver,
    # applied as a single scalar shock per trial (not year-by-year) for
    # tractability -- i.e., "how much better/worse than Base Case, on average,
    # does this 5-year path turn out to be."
    growth_shock = rng.triangular(
        left=SCENARIOS['Bear']['rev_growth'][1] - base['rev_growth'][1],
        mode=0.0,
        right=SCENARIOS['Bull']['rev_growth'][1] - base['rev_growth'][1],
        size=n_trials)
    eff_shock = rng.triangular(
        left=SCENARIOS['Bull']['efficiency_ratio'][1] - base['efficiency_ratio'][1],  # lower eff = bull
        mode=0.0,
        right=SCENARIOS['Bear']['efficiency_ratio'][1] - base['efficiency_ratio'][1],
        size=n_trials)
    ke_draw = rng.triangular(left=SCENARIOS['Bull']['cost_of_equity'], mode=base['cost_of_equity'],
                              right=SCENARIOS['Bear']['cost_of_equity'], size=n_trials)
    tg_draw = rng.triangular(left=SCENARIOS['Bear']['terminal_growth'], mode=base['terminal_growth'],
                              right=SCENARIOS['Bull']['terminal_growth'], size=n_trials)

    prices = np.empty(n_trials)
    rev0 = GS_HISTORY[2025]['net_revenues']
    eq0 = GS_HISTORY[2025]['bvps'] * GS_SHARES_HISTORY[2025]

    for i in range(n_trials):
        g_shift = growth_shock[i]
        eff_shift = eff_shock[i]
        ke = ke_draw[i]
        tg = tg_draw[i]
        # avoid Ke <= tg degenerate draws
        if ke <= tg + 0.01:
            ke = tg + 0.01

        rev = rev0
        equity = eq0
        fcfes = []
        for t in range(5):
            g = base['rev_growth'][t] + g_shift
            eff = base['efficiency_ratio'][t] + eff_shift
            eff = min(max(eff, 0.45), 0.85)
            rev = rev * (1 + g)
            opex = rev * eff
            pretax = rev - opex
            tax = pretax * base['tax_rate']
            ne = pretax - tax
            nic = ne - PREFERRED_DIV_PROJECTED
            required_retention = equity * base['bs_growth']
            fcfe = nic - required_retention
            equity = equity + required_retention
            fcfes.append(fcfe)

        pv = sum(cf / (1 + ke) ** (t + 1) for t, cf in enumerate(fcfes))
        tv = fcfes[-1] * (1 + tg) / (ke - tg)
        pv_tv = tv / (1 + ke) ** 5
        eqval = pv + pv_tv
        prices[i] = eqval / GS_SHARES_CURRENT_MM

    return prices


def summarize(prices):
    pcts = np.percentile(prices, [5, 10, 25, 50, 75, 90, 95])
    print(f"Trials: {len(prices):,}")
    print(f"Mean implied price/share:   ${prices.mean():,.2f}")
    print(f"Std dev:                    ${prices.std():,.2f}")
    print("\nPercentile distribution of implied price/share:")
    for p, v in zip([5, 10, 25, 50, 75, 90, 95], pcts):
        print(f"  P{p:>2}: ${v:,.2f}")
    prob_above_current = (prices > GS_PRICE).mean()
    print(f"\nCurrent market price: ${GS_PRICE:,.2f}")
    print(f"P(DCF-implied price > current market price): {prob_above_current*100:.1f}%")
    print(f"P(DCF-implied price < current market price): {(1-prob_above_current)*100:.1f}%")

    # simple ASCII histogram
    print("\nDistribution (ASCII histogram, 20 bins):")
    counts, edges = np.histogram(prices, bins=20)
    max_count = counts.max()
    for c, lo, hi in zip(counts, edges[:-1], edges[1:]):
        bar = '#' * int(60 * c / max_count)
        print(f"  ${lo:7,.0f}-${hi:7,.0f} | {bar} ({c})")


if __name__ == '__main__':
    n = int(sys.argv[1]) if len(sys.argv) > 1 else 20000
    print(f"Running Monte Carlo DCF simulation, {n:,} trials...\n")
    prices = simulate(n_trials=n)
    summarize(prices)
