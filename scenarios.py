#!/usr/bin/env python3
"""
End-to-end scenario sweep — runs the full model through many mock deals and toggle
combinations and checks a battery of invariants on every one.

This exercises the model engine (which mirrors the Excel formulas cell-for-cell) across
realistic and edge-case scenarios: each business plan, pure-MF and retail-heavy mixes,
single vs separate loans, every debt constraint binding, all three tax methods, both
exit methods, refi on/off, LP present/absent, fee on/off, lease structures, interest-
rate and growth stress (recession/boom), and short/long holds.

For each scenario it asserts (where applicable):
  - all cash flows finite; no unexpected NaN
  - Sources = Uses (equity + loan = total cost, recomputed independently)
  - component NOI: MF + Retail == total NOI every month
  - waterfall conserves cash: SUM(LP CF)+SUM(GP CF) == project levered CF
  - LP dist + GP dist == total distributions
  - loan within [0, price + capital]; binding flag valid
  - exit value > 0; stabilized NOI > 0 for going concerns
  - LP-absent => LP dist 0 and GP IRR == project IRR
  - promote >= 0
  - separate loans: property loan == MF leg + Retail leg
A scenario passes only if every applicable invariant holds.
"""

import math
from dataclasses import replace
import model_inputs as MI
from solve_and_test import Inputs, engine, solve_max_bid, edate, xirr

VALID_BINDING = {"LTV", "LTC", "DSCR", "Debt Yield", "Value cap"}


def finite(x):
    return isinstance(x, (int, float)) and not math.isnan(x) and not math.isinf(x)


def caps_budget(inp):
    reno = inp.reno_cost_unit*len(inp.mf_units) if inp.business_plan == "Value-Add" else 0.0
    return inp.cap_repairs + reno + inp.cap_tilc + inp.cap_contingency


def check_scenario(name, inp, price=None, solve_target=None):
    """Run one scenario; return (ok, headline dict, list of failures)."""
    fails = []
    if solve_target is not None:
        price = solve_max_bid(inp, solve_target)
        if price is None:
            return False, {}, [f"solver could not bracket target {solve_target:.0%}"]
    if price is None:
        price = inp.input_price
    e = engine(inp, price)

    # --- invariants ---
    # 1. all monthly flows finite
    if not all(finite(x) for x in e["lcf"]):
        fails.append("non-finite values in levered CF")
    if not finite(e["loan"]) or not finite(e["equity"]):
        fails.append("non-finite loan/equity")

    # 2. Sources = Uses (independently recomputed)
    uses = price + price*inp.closing_pct + caps_budget(inp) + e["fin_cost"]
    sources = e["loan"] + e["equity"]
    if abs(sources - uses) > 1.0:
        fails.append(f"sources({sources:,.0f}) != uses({uses:,.0f})")

    # 3. component NOI ties every month
    if not all(abs(e["mf_noi"][i] + e["rt_noi"][i] - e["noi"][i]) < 1e-6
               for i in range(len(e["noi"]))):
        fails.append("MF NOI + Retail NOI != total NOI")

    # 4. waterfall conservation
    wf = e["wf"]
    proj = sum(e["lcf"])
    split = sum(wf["lp_cf"]) + sum(wf["gp_cf"])
    if abs(split - proj) > 1.0:
        fails.append(f"waterfall not conserved: split({split:,.0f}) != proj({proj:,.0f})")
    # 5. LP dist + GP dist == total distributions
    tot_dist = sum(x for x in e["lcf"] if x > 0)
    if abs((wf["lp_dist"] + wf["gp_dist"]) - tot_dist) > 1.0:
        fails.append("LP+GP distributions != total distributions")

    # 6. loan bounds + binding flag
    if e["loan"] < -1:
        fails.append(f"negative loan {e['loan']:,.0f}")
    if e["loan"] > price + caps_budget(inp) + 1:
        fails.append(f"loan {e['loan']:,.0f} exceeds price+capital")
    sep = (inp.debt_structure == "Separate MF + Retail Loans")
    if not sep and e["binding"] not in VALID_BINDING:
        fails.append(f"invalid binding flag '{e['binding']}'")

    # 7. exit value & stabilized NOI
    if e["exit_val"] <= 0:
        fails.append("exit value <= 0")
    has_income = (len(inp.mf_units) > 0 or len(inp.retail_suites) > 0)
    if has_income and e["stab"] <= 0:
        fails.append(f"stabilized NOI <= 0 ({e['stab']:,.0f})")

    # 8. LP absent => LP dist 0 and GP IRR == project IRR
    if inp.lp_present == "No":
        if wf["lp_dist"] != 0:
            fails.append("LP present=No but LP dist != 0")
        if finite(wf["gp_irr"]) and finite(e["irr"]) and abs(wf["gp_irr"] - e["irr"]) > 1e-6:
            fails.append("LP absent but GP IRR != project IRR")
    # 9. promote non-negative
    if wf["gp_promote"] < -1:
        fails.append(f"negative promote {wf['gp_promote']:,.0f}")

    # 10. separate loans: property loan == legs
    if sep and abs(e["property_loan"] - (e["mf_loan"] + e["rt_loan"])) > 1:
        fails.append("separate: property loan != MF + RT legs")

    head = dict(price=price, goingin_cap=(e["going_in"]/price if price else 0),
                irr=e["irr"], u_irr=e["u_irr"], em=e["em"], loan=e["loan"],
                binding=e["binding"], lp_irr=wf["lp_irr"], gp_irr=wf["gp_irr"],
                promote=wf["gp_promote"], stab=e["stab"])
    return (len(fails) == 0), head, fails


def big_retail_suites():
    """Retail-heavy mix: ~90k SF anchored center."""
    return [("BigBox A", 45000, 0, 84, 22, 0.015, "NNN", 0, 0.04),
            ("BigBox B", 25000, 0, 72, 24, 0.02, "NNN", 0, 0.05),
            ("Inline 1", 6000, 0, 48, 38, 0.025, "Modified Gross", 0, 0.06),
            ("Inline 2", 6000, 0, 60, 40, 0.025, "Gross", 0, 0.06),
            ("Pad", 4000, 0, 120, 55, 0.01, "NNN", 0, 0.04)]


def small_mf():
    return MI.gen_mf_units(24)


# ------------------------------------------------------------------------------------
# Scenario catalogue
# ------------------------------------------------------------------------------------
def scenarios():
    B = Inputs()
    S = []
    add = lambda name, inp, **kw: S.append((name, inp, kw))

    add("01 Base (value-add, mixed, single loan)", B)
    add("02 Stabilized plan", replace(B, business_plan="Stabilized"))
    add("03 Core-plus lease-up (start occ 65%)",
        replace(B, business_plan="Core-Plus Lease-Up", leaseup_start_occ=0.65, leaseup_absorption=4))
    add("04 Aggressive value-add (reno $400, pace 8)",
        replace(B, reno_premium=400, reno_pace=8, reno_cost_unit=28000))
    add("05 Pure multifamily (retail GLA = 0)", replace(B, retail_suites=tuple()))
    add("06 Retail-heavy (90k SF, 24 units)",
        replace(B, mf_units=tuple(small_mf()), retail_suites=tuple(big_retail_suites())))
    add("07 Separate loans (fixed/fixed)",
        replace(B, debt_structure="Separate MF + Retail Loans"))
    add("08 Separate loans (MF floating, retail fixed)",
        replace(B, debt_structure="Separate MF + Retail Loans"))  # rate type per-leg lives in Excel; replica uses leg rates
    add("09 High leverage (LTV 80, DY 6%)",
        replace(B, max_ltv=0.80, min_dy=0.06, min_dscr=1.05))
    add("10 DSCR-binding (DSCR 1.60)", replace(B, min_dscr=1.60))
    add("11 Debt-yield binding (DY 11%)", replace(B, min_dy=0.11))
    add("12 LTC binding (LTC 55%)", replace(B, max_ltc=0.55, max_ltv=0.75))
    add("13 Only LTV active (others off)",
        replace(B, use_ltc=False, use_dscr=False, use_dy=False))
    add("14 All constraints off (uncapped)",
        replace(B, use_ltv=False, use_ltc=False, use_dscr=False, use_dy=False))
    add("15 Refi ON", replace(B, refi_on="On"))
    add("16 Tax: Flat growth", replace(B, tax_method="Flat Growth"))
    add("17 Tax: PILOT schedule", replace(B, tax_method="PILOT / Abatement"))
    add("18 Exit: component caps", replace(B, exit_val_method="Component Caps (MF + Retail summed)"))
    add("19 Exit NOI: trailing 12-mo", replace(B, exit_noi_basis="Trailing 12-mo"))
    add("20 LP absent (sponsor only)", replace(B, lp_present="No"))
    add("21 No fees", replace(B, acq_fee_on=False, am_fee_on=False, disp_fee_on=False))
    add("22 AM fee basis = % of Cost", replace(B, am_fee_basis="% of Cost"))
    add("23 AM fee basis = % of NOI", replace(B, am_fee_basis="% of NOI"))
    add("24 No catch-up", replace(B, gp_catchup_on=False))
    add("25 GP co-invest 50%, LP 80/20 split", replace(B, lp_split=0.80, gp_coinvest=0.50))
    add("26 Gross-up OFF", replace(B, gross_up_on=False))
    add("27 Admin fee OFF", replace(B, admin_fee_on=False))
    add("28 Short hold (24 mo)", replace(B, disp=edate(B.acq, 24)))
    add("29 Long hold (120 mo)", replace(B, disp=edate(B.acq, 120)))
    add("30 High rate (8.5% fixed)", replace(B, fixed_rate=0.085))
    add("31 Floating rate (SOFR+spread)", replace(B, loan_rate_type="Floating"))
    add("32 RECESSION (neg growth, cap +100bp, vac 12%)",
        replace(B, vec_mfrent=(-0.02,-0.01,0.0,0.01,0.02,0.03,0.03,0.03,0.03,0.03,0.03),
                vec_retailrent=(-0.03,-0.01,0.0,0.01,0.02,0.03,0.03,0.03,0.03,0.03,0.03),
                exit_cap_blended=0.065, mf_vacancy=0.12, retail_genvac=0.12))
    add("33 BOOM (high growth, cap -75bp)",
        replace(B, vec_mfrent=(0.07,0.06,0.05,0.04,0.04,0.04,0.04,0.04,0.04,0.04,0.04),
                exit_cap_blended=0.0475))
    add("34 Zero growth everywhere",
        replace(B, vec_mfrent=(0,)*11, vec_retailrent=(0,)*11, vec_opex=(0,)*11,
                vec_tilc=(0,)*11, vec_other=(0,)*11, tax_flat_growth=0.0))
    # solve-mode scenarios (exercise the max-bid solver end to end)
    add("35 SOLVE 12% target IRR", B, solve_target=0.12)
    add("36 SOLVE 18% target IRR", B, solve_target=0.18)
    add("37 SOLVE 15% on separate loans", replace(B, debt_structure="Separate MF + Retail Loans"), solve_target=0.15)
    add("38 SOLVE 8% pure-MF stabilized (feasible)",
        replace(B, retail_suites=tuple(), business_plan="Stabilized"), solve_target=0.08)
    return S


def infeasible_solve_is_graceful():
    """Targeting an IRR above what's achievable in the search range must return None
    (not a wrong number). IRR is highest at the floor price and falls with price, so any
    target above the floor-price IRR is unreachable."""
    B = Inputs()
    floor_irr = engine(B, 5_000_000.0)["irr"]
    target = floor_irr + 0.50            # guaranteed above anything in [5M, hi]
    return solve_max_bid(B, target) is None


def main():
    rows = scenarios()
    print("=" * 110)
    print(f"{'SCENARIO':<46}{'price':>11}{'GI cap':>8}{'IRR':>8}{'EM':>7}{'LP IRR':>8}{'GP IRR':>8}{'  result'}")
    print("-" * 110)
    npass = 0
    all_fail = []
    for name, inp, kw in rows:
        ok, h, fails = check_scenario(name, inp, price=kw.get("price"), solve_target=kw.get("solve_target"))
        if ok:
            npass += 1
        else:
            all_fail.append((name, fails))
        if h:
            print(f"{name:<46}{h['price']/1e6:>9.1f}M{h['goingin_cap']:>8.2%}{h['irr']:>8.1%}"
                  f"{h['em']:>6.2f}x{h['lp_irr']:>8.1%}{h['gp_irr']:>8.1%}   {'PASS' if ok else 'FAIL'}")
        else:
            print(f"{name:<46}{'—':>40}   FAIL")
    print("-" * 110)
    print(f"{npass}/{len(rows)} scenarios passed all invariants.")
    graceful = infeasible_solve_is_graceful()
    print(f"Infeasible-target solve returns None (graceful): {'PASS' if graceful else 'FAIL'}")
    if not graceful:
        all_fail.append(("infeasible-solve", ["did not return None for unreachable target"]))
    if all_fail:
        print("\nFAILURES:")
        for name, fails in all_fail:
            for f in fails:
                print(f"  [{name}] {f}")
    else:
        print("Every scenario satisfied every applicable invariant.")
    return (npass == len(rows)) and graceful


if __name__ == "__main__":
    import sys
    sys.exit(0 if main() else 1)
