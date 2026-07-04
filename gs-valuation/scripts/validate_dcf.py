"""
validate_dcf.py
================
Two jobs:

1. Cross-check the Excel DCF Valuation output (implied value/share) against an
   independent Python recomputation of the same FCFE path and discounting.

2. Go beyond the Excel workbook's single-stage (closed-form) reverse DCF and
   solve the MULTI-STAGE reverse DCF numerically: what constant revenue-growth
   rate, applied across the full 5-year explicit forecast (holding every other
   driver at Base Case), makes the DCF's implied equity value equal today's
   market equity value? This is solved by bisection since there is no closed
   form once a multi-year explicit period + terminal value are both in play.

Usage:
    python validate_dcf.py
"""
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'data'))
from assumptions import *
from openpyxl import load_workbook

TOL_PCT = 0.005


def fcfe_path(rev_growth_list, scenario_overrides=None):
    """Compute the 5-yr FCFE path given a list of 5 revenue-growth rates,
    holding efficiency ratio / tax / bs_growth at Base Case (or overrides)."""
    s = dict(SCENARIOS['Base'])
    if scenario_overrides:
        s.update(scenario_overrides)

    rev = GS_HISTORY[2025]['net_revenues']
    equity = GS_HISTORY[2025]['bvps'] * GS_SHARES_HISTORY[2025]
    fcfes = []
    for i in range(5):
        rev = rev * (1 + rev_growth_list[i])
        opex = rev * s['efficiency_ratio'][i]
        pretax = rev - opex
        tax = pretax * s['tax_rate']
        net_earnings = pretax - tax
        nic = net_earnings - PREFERRED_DIV_PROJECTED

        beg_equity = equity
        required_retention = beg_equity * s['bs_growth']
        fcfe = nic - required_retention
        end_equity = beg_equity + required_retention

        fcfes.append(fcfe)
        equity = end_equity
    return fcfes


def dcf_equity_value(fcfes, ke, tg):
    pv = sum(cf / (1 + ke) ** (t + 1) for t, cf in enumerate(fcfes))
    tv = fcfes[-1] * (1 + tg) / (ke - tg)
    pv_tv = tv / (1 + ke) ** 5
    return pv + pv_tv


def part1_crosscheck(model_path):
    print("=" * 78)
    print("PART 1: Cross-check Excel DCF Base Case vs. independent Python DCF")
    print("=" * 78)
    wb = load_workbook(model_path, data_only=True)
    active_scenario = wb['Assumptions']['B3'].value
    if active_scenario != 'Base':
        print(f"NOTE: workbook active scenario is '{active_scenario}'; this cross-check assumes 'Base'. "
              f"Set Assumptions!B3='Base', save, re-run recalc.py, then re-run this script for a clean check.")

    excel_price = wb['DCF Valuation']['B17'].value
    excel_eqval = wb['DCF Valuation']['B15'].value

    base = SCENARIOS['Base']
    fcfes = fcfe_path(base['rev_growth'])
    py_eqval = dcf_equity_value(fcfes, base['cost_of_equity'], base['terminal_growth'])
    py_price = py_eqval / GS_SHARES_CURRENT_MM

    print(f"Python FCFE path ($mm): {[round(x, 1) for x in fcfes]}")
    print(f"Python Implied Equity Value: ${py_eqval:,.1f}mm   |   Excel: ${excel_eqval:,.1f}mm")
    print(f"Python Implied Price/Share:  ${py_price:,.2f}     |   Excel: ${excel_price:,.2f}")
    if excel_eqval:
        diff = abs(py_eqval - excel_eqval) / excel_eqval
        status = "PASS" if diff < TOL_PCT else "FAIL"
        print(f"Relative difference: {diff*100:.3f}%  -> {status}")
        return status == "PASS"
    return None


def part2_multistage_reverse_dcf():
    print("\n" + "=" * 78)
    print("PART 2: Multi-stage reverse DCF (numeric solve via bisection)")
    print("=" * 78)
    base = SCENARIOS['Base']
    ke, tg = base['cost_of_equity'], base['terminal_growth']
    target_equity_value = GS_PRICE * GS_SHARES_CURRENT_MM

    def value_at_uniform_growth(g):
        fcfes = fcfe_path([g] * 5)
        return dcf_equity_value(fcfes, ke, tg)

    lo, hi = -0.10, 0.40
    v_lo, v_hi = value_at_uniform_growth(lo), value_at_uniform_growth(hi)
    if not (v_lo < target_equity_value < v_hi):
        print(f"Target ${target_equity_value:,.0f}mm outside bracket [{v_lo:,.0f}, {v_hi:,.0f}]mm; widen bounds.")
        return None

    for _ in range(100):
        mid = (lo + hi) / 2
        v_mid = value_at_uniform_growth(mid)
        if abs(v_mid - target_equity_value) / target_equity_value < 1e-6:
            break
        if v_mid < target_equity_value:
            lo = mid
        else:
            hi = mid
    g_star_multistage = mid

    # Compare to the workbook's closed-form single-stage estimate
    fcfe1 = fcfe_path(base['rev_growth'])[0]
    g_star_singlestage = ke - fcfe1 / target_equity_value

    print(f"Target equity value (current price x shares): ${target_equity_value:,.0f}mm")
    print(f"Multi-stage implied uniform revenue growth g* (5-yr explicit + TV): {g_star_multistage*100:.2f}%")
    print(f"Single-stage closed-form implied FCFE growth g* (Excel Reverse DCF tab): {g_star_singlestage*100:.2f}%")
    print("\nInterpretation: the multi-stage solve answers 'what revenue growth, sustained for 5 years")
    print("and then leveled off at the 3.0% terminal rate, justifies today's price?' -- a more intuitive")
    print("number for a memo than the single-stage 'perpetual FCFE growth' figure, because revenue growth")
    print("is what analysts and management actually discuss in earnings calls and guidance.")
    return g_star_multistage


if __name__ == '__main__':
    model_path = os.path.join(os.path.dirname(__file__), '..', 'models', '03_DCF_and_Reverse_DCF.xlsx')
    ok = part1_crosscheck(model_path)
    g_star = part2_multistage_reverse_dcf()

    print("\n" + "=" * 78)
    if ok:
        print("Overall: DCF cross-check PASS.")
    elif ok is False:
        print("Overall: DCF cross-check FAIL -- review formulas.")
    else:
        print("Overall: DCF cross-check SKIPPED (workbook not on Base case).")
    sys.exit(0 if ok in (True, None) else 1)
