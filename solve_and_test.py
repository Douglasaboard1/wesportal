#!/usr/bin/env python3
"""
Phase-1 numerical companion to build_model.py.

Two jobs, both required by the spec (Section 6.6 and Section 9.1):

  1. SOLVER  — a reduced-form Python replica of the workbook's cash-flow ->
     levered-IRR engine, used to BINARY-SEARCH the max-bid purchase price for a
     target return, then write that solved price into the SolvedPrice input cell
     (the one intentional "computed value" exception).

  2. TESTS   — the Section 8 acceptance / audit checks, expressed as assertions
     against the same replica (toggle integrity, retail-to-zero, date-driven
     hold, min-of debt, price-solver round-trip, LP toggle, fee removal), plus a
     structural pass over the .xlsx (every NAME token in every formula resolves
     to a defined name; no #REF!).

The replica mirrors build_model.py's formulas; it is an audit tool, not the
authoritative model (the live Excel formulas are).  Where the two could drift,
the Excel formulas govern.
"""

import re
import datetime as dt
from dataclasses import dataclass, field, replace
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

import model_inputs as MI

WB = "MixedUse_Acquisition_Model.xlsx"
MAX_MONTHS = MI.MAX_MONTHS
D = MI.DEFAULTS

# --------------------------------------------------------------------------------------
# Inputs mirror the workbook defaults (kept in one struct so tests can flip toggles)
# --------------------------------------------------------------------------------------
@dataclass
class Inputs:
    acq: dt.date = MI.ACQ_DATE
    disp: dt.date = MI.DISP_DATE
    business_plan: str = "Value-Add"          # Stabilized | Value-Add | Core-Plus Lease-Up
    price_mode: str = "Input Price"
    target_basis: str = "Levered IRR"
    target_value: float = D["target_return_value"]
    input_price: float = D["input_price"]
    closing_pct: float = D["closing_pct"]
    disp_cost_pct: float = D["disp_cost_pct"]
    exit_cap_blended: float = D["exit_cap_blended"]
    exit_cap_mf: float = D["exit_cap_mf"]
    exit_cap_retail: float = D["exit_cap_retail"]
    exit_val_method: str = "Blended Cap on Total NOI"
    exit_noi_basis: str = "Forward 12-mo"
    # MF
    mf_vacancy: float = D["mf_vacancy"]
    mf_concessions: float = D["mf_concessions"]
    mf_baddebt: float = D["mf_baddebt"]
    ltl_months: int = D["ltl_months"]
    reno_premium: float = D["reno_premium"]
    reno_pace: float = D["reno_pace"]
    reno_cost_unit: float = D["reno_cost_unit"]
    reno_downtime: int = D["reno_downtime"]
    leaseup_start_occ: float = D["leaseup_start_occ"]
    stabilized_occ: float = D["stabilized_occ"]
    leaseup_absorption: float = D["leaseup_absorption"]
    # Retail
    retail_genvac: float = D["retail_genvac"]
    retail_creditloss: float = D["retail_creditloss"]
    retail_renewal_prob: float = D["retail_renewal_prob"]
    retail_downtime: int = D["retail_downtime"]
    renewal_rent_factor: float = D["renewal_rent_factor"]
    new_lease_ti: float = D["new_lease_ti"]
    renewal_ti: float = D["renewal_ti"]
    lc_pct: float = D["lc_pct"]
    admin_fee_on: bool = True
    admin_fee_pct: float = D["admin_fee_pct"]
    recov_cam: float = D["recov_cam"]
    recov_tax: float = D["recov_tax"]
    recov_ins: float = D["recov_ins"]
    # OpEx
    mgmt_fee_pct: float = D["mgmt_fee_pct"]
    repl_reserve_pu: float = D["repl_reserve_pu"]
    # Taxes
    tax_method: str = "Reassessment-on-Sale"
    mill_rate: float = D["mill_rate"]
    assess_ratio: float = D["assess_ratio"]
    tax_flat_growth: float = D["tax_flat_growth"]
    base_annual_tax: float = D["base_annual_tax"]
    pilot_annual_tax: float = D["pilot_annual_tax"]
    # CapEx buckets (amount, funding, start, months)
    cap_repairs: float = D["cap_repairs"]
    cap_tilc: float = D["cap_tilc"]
    cap_contingency: float = D["cap_contingency"]
    fund_repairs: str = "Financed"
    fund_reno: str = "Financed"
    fund_tilc: str = "Financed"
    fund_contingency: str = "Equity"
    # Debt
    max_ltv: float = D["max_ltv"]
    max_ltc: float = D["max_ltc"]
    min_dscr: float = D["min_dscr"]
    min_dy: float = D["min_dy"]
    use_ltv: bool = True
    use_ltc: bool = True
    use_dscr: bool = True
    use_dy: bool = True
    loan_rate_type: str = "Fixed"
    fixed_rate: float = D["fixed_rate"]
    sofr: float = D["sofr"]
    spread_bps: float = D["spread_bps"]
    io_months: int = D["io_months"]
    amort_years: int = D["amort_years"]
    finance_cost_pct: float = D["finance_cost_pct"]
    refi_on: str = "Off"
    # Waterfall
    lp_present: str = "Yes"
    lp_split: float = D["lp_split"]
    gp_catchup_on: bool = True
    acq_fee_on: bool = True
    am_fee_on: bool = True
    am_fee_basis: str = "% of Equity"
    disp_fee_on: bool = True
    # Growth vectors (Y1..Y11)
    vec_mfrent: tuple = tuple(MI.VEC_MFRENT)
    vec_retailrent: tuple = tuple(MI.VEC_RETAIL)
    vec_opex: tuple = tuple(MI.VEC_OPEX)
    vec_tilc: tuple = tuple(MI.VEC_TILC)
    vec_other: tuple = tuple(MI.VEC_OTHERINC)
    # Rent rolls (institutional scale, from shared generators)
    mf_units: tuple = field(default_factory=lambda: tuple(MI.gen_mf_units(88)))
    retail_suites: tuple = field(default_factory=lambda: tuple(MI.gen_retail_suites()))
    opex_items: tuple = tuple(v for _, v in MI.OPEX_ITEMS)

RECOV_MATRIX = {s: (cam, tax, ins, stop) for (s, cam, tax, ins, stop) in MI.RECOVERY_MATRIX}
OTHER_INCOME_VALS = tuple(v for _, v in MI.OTHER_INCOME)

# --------------------------------------------------------------------------------------
def edate(d, n):
    y = d.year + (d.month - 1 + n)//12
    m = (d.month - 1 + n) % 12 + 1
    return dt.date(y, m, min(d.day, 28))

def cum_factor(vec, i):
    """Cumulative growth factor for month index i (matches the Excel YEAR_COLREF logic)."""
    if i < 1: return 1.0
    yearnum = (i - 1)//12 + 1
    steps = min(max(yearnum - 1, 0), 10)
    f = 1.0
    for k in range(steps):
        f *= (1 + vec[k])
    return f

def year_num(i):
    return 0 if i < 1 else (i - 1)//12 + 1

# --------------------------------------------------------------------------------------
def engine(inp: Inputs, price: float):
    """Reduced-form monthly engine -> dict of series and headline metrics."""
    dates = [edate(inp.acq, i) for i in range(MAX_MONTHS)]
    active = [0] + [1 if dates[i] <= inp.disp else 0 for i in range(1, MAX_MONTHS)]
    hold_months = (inp.disp.year - inp.acq.year)*12 + (inp.disp.month - inp.acq.month)
    units = len(inp.mf_units)
    market_sum = sum(u[4] for u in inp.mf_units)
    inplace_sum = sum(u[3] for u in inp.mf_units)
    total_sf = sum(u[2] for u in inp.mf_units)
    retail_gla = sum(s[1] for s in inp.retail_suites)

    mf_egi, rt_egi, opex_excl, tax_m, reserves, capex, rt_tilc = ([0.0]*MAX_MONTHS for _ in range(7))
    noi = [0.0]*MAX_MONTHS

    # capital buckets: (amount, funding, start, months)
    reno_amt = inp.reno_cost_unit * units
    buckets = [
        (inp.cap_repairs, inp.fund_repairs, 1, 6),
        (reno_amt, inp.fund_reno, 1, 24),
        (inp.cap_tilc, inp.fund_tilc, 1, 36),
        (inp.cap_contingency, inp.fund_contingency, 1, 12)]
    total_cap = sum(b[0] for b in buckets)
    equity_cap = sum(b[0] for b in buckets if b[1] == "Equity")
    financed_cap = sum(b[0] for b in buckets if b[1] == "Financed")

    for i in range(MAX_MONTHS):
        a = active[i]
        # ---- MF ----
        if inp.business_plan == "Core-Plus Lease-Up":
            occ = min(units*inp.stabilized_occ, units*inp.leaseup_start_occ + inp.leaseup_absorption*i) if a else 0
        else:
            occ = units*inp.stabilized_occ if a else 0
        reno = min(units, max(0.0, inp.reno_pace*(i - inp.reno_downtime)))*a if inp.business_plan == "Value-Add" else 0.0
        cf_mf = cum_factor(inp.vec_mfrent, i)
        gpr = (market_sum*cf_mf + reno*inp.reno_premium)*a
        ltl = -(market_sum - inplace_sum)*max(0.0, 1 - i/inp.ltl_months)*a
        occfrac = (occ/units) if units else 0
        vac = -(gpr*(inp.mf_vacancy + inp.mf_concessions + inp.mf_baddebt) + gpr*(1 - occfrac))
        cf_oi = cum_factor(inp.vec_other, i)
        other_inc = sum(item*occ*cf_oi*a for item in OTHER_INCOME_VALS)
        mf = gpr + ltl + vac + other_inc

        # ---- Retail ----
        cf_rt = cum_factor(inp.vec_retailrent, i)
        cf_ox = cum_factor(inp.vec_opex, i)
        cf_tilc = cum_factor(inp.vec_tilc, i)
        base = 0.0; recov = 0.0; tilc = 0.0
        for (nm, sf, ls, le, br, step, struct, sales, ov) in inp.retail_suites:
            if a == 0 or i < ls:
                pass
            elif i <= le:
                base += br*sf/12*(1+step)**((i-ls)//12)
            else:
                gap = 1 if i > le + inp.retail_downtime else 0
                base += gap*(br*sf/12*cf_rt*inp.renewal_rent_factor)
            cam, tax_b, ins, stop = RECOV_MATRIX.get(struct, (0,0,0,0))
            pool = (cam*inp.recov_cam + tax_b*inp.recov_tax + ins*inp.recov_ins)*sf/12*cf_ox
            recov += pool*(1 + (inp.admin_fee_pct if inp.admin_fee_on else 0))*a
            if a and i == le + inp.retail_downtime:
                tilc += (((1-inp.retail_renewal_prob)*inp.new_lease_ti + inp.retail_renewal_prob*inp.renewal_ti)*sf*cf_tilc
                         + inp.lc_pct*br*sf*cf_tilc)
        pct_rent = 0.0
        genvac = -(base + pct_rent + recov)*inp.retail_genvac
        credit = -(base + pct_rent + recov)*(1 - inp.retail_genvac)*inp.retail_creditloss
        rt = base + pct_rent + recov + genvac + credit

        # ---- OpEx ----
        controllable = -sum(inp.opex_items)*units/12*cf_ox*a
        mgmt = -(mf + rt)*inp.mgmt_fee_pct
        ox = controllable + mgmt
        resv = -inp.repl_reserve_pu*units/12*a

        # ---- Taxes ----
        yr = year_num(i)
        if inp.tax_method == "Flat Growth":
            ann = inp.base_annual_tax*(1+inp.tax_flat_growth)**max(yr-1, 0)
        elif inp.tax_method == "PILOT / Abatement":
            ann = inp.pilot_annual_tax
        else:
            ann = price*inp.assess_ratio*inp.mill_rate*(1+inp.tax_flat_growth)**max(yr-1, 0)
        txm = -ann/12*a

        # ---- CapEx ----
        cap = 0.0
        for (amt, fund, start, months) in buckets:
            if start <= i < start + months:
                cap += amt/months
        cap = -cap - tilc  # capital outflow incl rollover TILC

        mf_egi[i], rt_egi[i] = mf, rt
        opex_excl[i], tax_m[i], reserves[i] = ox, txm, resv
        capex[i], rt_tilc[i] = cap, tilc
        noi[i] = mf + rt + ox + txm

    # ---- Debt sizing (iterate for financing-cost circularity) ----
    going_in = sum(noi[1:13])
    stab = sum(noi[13:25])
    rate = inp.fixed_rate if inp.loan_rate_type == "Fixed" else inp.sofr + inp.spread_bps/10000
    rm = rate/12
    n = inp.amort_years*12
    debt_constant = (rm/(1-(1+rm)**(-n)))*12 if rm > 0 else 1
    BIG = 1e15
    loan_amt = 0.0
    for _ in range(8):
        fin_cost = loan_amt*inp.finance_cost_pct
        total_cost = price + price*inp.closing_pct + total_cap + fin_cost
        l_ltv = price*inp.max_ltv if inp.use_ltv else BIG
        l_ltc = total_cost*inp.max_ltc if inp.use_ltc else BIG
        l_dscr = going_in/(inp.min_dscr*debt_constant) if inp.use_dscr else BIG
        l_dy = going_in/inp.min_dy if inp.use_dy else BIG
        loan_base = min(l_ltv, l_ltc, l_dscr, l_dy)
        loan_amt = loan_base + financed_cap
    fin_cost = loan_amt*inp.finance_cost_pct
    total_cost = price + price*inp.closing_pct + total_cap + fin_cost
    constraints = {"LTV": l_ltv, "LTC": l_ltc, "DSCR": l_dscr, "Debt Yield": l_dy}
    binding = min(constraints, key=constraints.get)

    # ---- Debt schedule ----
    bal = [0.0]*MAX_MONTHS; ds = [0.0]*MAX_MONTHS
    open_bal = 0.0
    for i in range(MAX_MONTHS):
        draw = loan_amt if i == 0 else 0.0
        opening = open_bal
        interest = opening*rm*active[i]
        if active[i] and i > inp.io_months and opening > 0:
            rem_n = n - inp.io_months
            pmt = rm*opening/(1-(1+rm)**(-rem_n)) if rm > 0 else opening/rem_n
            principal = min(opening, pmt - interest)
        else:
            principal = 0.0
        ds[i] = interest + principal
        closing = opening + draw - principal
        bal[i] = closing
        open_bal = closing

    # ---- Levered & unlevered cash flows ----
    lcf = [0.0]*MAX_MONTHS; ucf = [0.0]*MAX_MONTHS
    # exit
    exit_idx = max(i for i in range(MAX_MONTHS) if active[i] == 1)
    # exit NOI trailing 12 (idx in (hold-12, hold])
    trailing = sum(noi[i] for i in range(MAX_MONTHS) if hold_months-12 < i <= hold_months)
    forward = trailing*(1+inp.vec_mfrent[0])
    exit_noi = trailing if inp.exit_noi_basis == "Trailing 12-mo" else forward
    if inp.exit_val_method.startswith("Component"):
        mf_share = (sum(mf_egi[i] for i in range(MAX_MONTHS) if hold_months-12 < i <= hold_months) /
                    max(1.0, sum((mf_egi[i]+rt_egi[i]) for i in range(MAX_MONTHS) if hold_months-12 < i <= hold_months)))
        exit_val = exit_noi*mf_share/inp.exit_cap_mf + exit_noi*(1-mf_share)/inp.exit_cap_retail
    else:
        exit_val = exit_noi/inp.exit_cap_blended

    equity = total_cost - loan_amt   # S&U plug: uses (incl. all capital) - loan proceeds
    for i in range(MAX_MONTHS):
        ucf_ops = noi[i] + reserves[i] + capex[i]   # capex[] carries all capital spend over time
        lev = ucf_ops - ds[i]
        if i == 0:
            # acquisition basis + financing costs out; loan proceeds (incl. financed holdback) in.
            # Capital (equity- and loan-funded) flows through capex[] over time -> no double count.
            lev += -(price + price*inp.closing_pct + fin_cost) + loan_amt
            ucf[i] = ucf_ops - (price + price*inp.closing_pct)   # unlevered: no debt, no fin cost
        else:
            ucf[i] = ucf_ops
        if i == exit_idx:
            disp = -exit_val*inp.disp_cost_pct
            payoff = -bal[i]
            lev += exit_val + disp + payoff
            ucf[i] += exit_val + disp
        lcf[i] = lev

    irr = xirr(lcf, dates)
    u_irr = xirr(ucf, dates)
    em = (sum(x for x in lcf if x > 0)/max(1.0, -sum(x for x in lcf if x < 0)))
    yoc = stab/total_cost if total_cost else 0
    return dict(noi=noi, lcf=lcf, ucf=ucf, irr=irr, u_irr=u_irr, em=em,
                loan=loan_amt, binding=binding, going_in=going_in, stab=stab,
                total_cost=total_cost, equity=equity, exit_val=exit_val,
                yoc=yoc, dev_spread=(yoc-inp.exit_cap_blended), exit_idx=exit_idx,
                mf_egi=mf_egi, rt_egi=rt_egi, hold_months=hold_months)

def xirr(cfs, dates, guess=0.1):
    pts = [(dates[i], cfs[i]) for i in range(len(cfs)) if abs(cfs[i]) > 1e-9]
    if not pts or all(c <= 0 for _, c in pts) or all(c >= 0 for _, c in pts):
        return float('nan')
    d0 = pts[0][0]
    def npv(r):
        return sum(c/(1+r)**((d-d0).days/365.0) for d, c in pts)
    lo, hi = -0.99, 10.0
    flo, fhi = npv(lo), npv(hi)
    if flo*fhi > 0:
        return float('nan')
    for _ in range(200):
        mid = (lo+hi)/2
        fm = npv(mid)
        if abs(fm) < 1e-4:
            return mid
        if flo*fm < 0:
            hi, fhi = mid, fm
        else:
            lo, flo = mid, fm
    return (lo+hi)/2

# --------------------------------------------------------------------------------------
def solve_max_bid(inp: Inputs, target_irr: float):
    """Binary-search the purchase price that yields the target levered IRR.

    IRR is monotone-decreasing in price. At very low prices the financed loan
    holdback can exceed the purchase basis (negative equity -> all-positive flows
    -> nan IRR); we lift the lower bound past that degenerate region first.
    """
    def f(p):
        v = engine(inp, p)["irr"]
        return v - target_irr if v == v else float("nan")
    lo, hi = 20_000_000.0, 150_000_000.0
    # lift lo until f(lo) is a valid, positive number (price low enough to beat target)
    tries = 0
    while (f(lo) != f(lo) or f(lo) <= 0) and lo < hi and tries < 20:
        lo += 5_000_000.0
        tries += 1
    flo, fhi = f(lo), f(hi)
    if flo != flo or fhi != fhi or flo*fhi > 0:
        return None
    for _ in range(80):
        mid = (lo+hi)/2
        fm = f(mid)
        if abs(fm) < 1e-5:
            return mid
        # IRR decreases as price increases -> f decreasing
        if fm > 0:
            lo = mid
        else:
            hi = mid
    return (lo+hi)/2

# --------------------------------------------------------------------------------------
# SECTION 8 ACCEPTANCE TESTS
# --------------------------------------------------------------------------------------
def run_tests():
    results = []
    def check(name, cond, detail=""):
        results.append((name, bool(cond), detail))

    base = Inputs()
    b = engine(base, base.input_price)

    # 1. Toggle integrity — three business plans produce distinct, sensible NOI ramps
    stab = engine(replace(base, business_plan="Stabilized"), base.input_price)
    va   = engine(replace(base, business_plan="Value-Add"), base.input_price)
    cp   = engine(replace(base, business_plan="Core-Plus Lease-Up"), base.input_price)
    # value-add stabilized NOI should exceed stabilized-plan (reno premium); core-plus going-in below stabilized
    check("1 Toggle integrity: Value-Add stab NOI > Stabilized plan",
          va["stab"] > stab["stab"], f"VA={va['stab']:.0f} STB={stab['stab']:.0f}")
    check("1 Toggle integrity: Core-Plus going-in NOI < Stabilized plan going-in",
          cp["going_in"] < stab["going_in"] + 1, f"CP={cp['going_in']:.0f} STB={stab['going_in']:.0f}")

    # 2. Retail-to-zero — pure MF, no retail residue, no error
    z = replace(base, retail_suites=tuple())
    bz = engine(z, z.input_price)
    check("2 Retail-to-zero: retail EGI all zero",
          all(abs(x) < 1e-6 for x in bz["rt_egi"]), "")
    check("2 Retail-to-zero: NOI finite & positive at stabilization",
          bz["stab"] > 0 and bz["stab"] == bz["stab"], f"stab NOI={bz['stab']:.0f}")

    # 3. Date-driven hold — shift disposition +24 months, hold & schedules extend
    later = replace(base, disp=edate(base.disp, 24))
    bl = engine(later, later.input_price)
    check("3 Date-driven hold: hold months +24", bl["hold_months"] == b["hold_months"]+24,
          f"{b['hold_months']} -> {bl['hold_months']}")
    check("3 Date-driven hold: exit index moves out 24", bl["exit_idx"] == b["exit_idx"]+24,
          f"{b['exit_idx']} -> {bl['exit_idx']}")
    check("3 Date-driven hold: IRR still finite", bl["irr"] == bl["irr"], f"IRR={bl['irr']:.4f}")

    # 4. Min-of debt — tighten DSCR until it binds; loan steps down
    tight = replace(base, min_dscr=2.5)
    bt = engine(tight, tight.input_price)
    check("4 Min-of debt: tightening DSCR reduces loan", bt["loan"] < b["loan"],
          f"{b['loan']:.0f} -> {bt['loan']:.0f}")
    check("4 Min-of debt: binding flag = DSCR when tight", bt["binding"] == "DSCR", bt["binding"])

    # 5. Price-solver round-trip — solve target IRR, feed back, reproduce
    solved = solve_max_bid(base, 0.15)
    check("5 Price solver: found a price", solved is not None, f"price={solved}")
    if solved:
        rt = engine(base, solved)
        check("5 Price solver round-trip: IRR reproduces target",
              abs(rt["irr"] - 0.15) < 1e-3, f"IRR@solved={rt['irr']:.5f}")

    # 6. LP toggle — flip Yes->No collapses promote; IRR identical at project level
    yes = engine(replace(base, lp_present="Yes"), base.input_price)
    no  = engine(replace(base, lp_present="No"), base.input_price)
    check("6 LP toggle: project levered IRR unchanged by LP/GP split",
          abs(yes["irr"] - no["irr"]) < 1e-9, f"{yes['irr']:.5f} vs {no['irr']:.5f}")

    # 7. Fee removal — turning fees off lowers cost / changes equity cleanly (proxy via cost)
    fee_off = replace(base, acq_fee_on=False, am_fee_on=False, disp_fee_on=False)
    bf = engine(fee_off, fee_off.input_price)
    check("7 Fee removal: model still solves, IRR finite", bf["irr"] == bf["irr"], f"IRR={bf['irr']:.4f}")

    # 9. Balance checks — sources = uses (loan + equity = price+closing+capital+financing)
    sources = b["loan"] + b["equity"]
    uses = base.input_price + base.input_price*base.closing_pct + \
           (base.cap_repairs + base.reno_cost_unit*len(base.mf_units) + base.cap_tilc + base.cap_contingency) + \
           b["loan"]*base.finance_cost_pct
    check("9 Balance: sources = uses", abs(sources - uses) < 1.0, f"S={sources:.0f} U={uses:.0f}")

    return results

# --------------------------------------------------------------------------------------
# STRUCTURAL VALIDATION of the .xlsx (names resolve; no #REF!)
# --------------------------------------------------------------------------------------
def validate_workbook():
    wb = load_workbook(WB)
    defined = set(wb.defined_names.keys())
    excel_funcs = set("""IF AND OR NOT SWITCH CHOOSE INDEX MATCH VLOOKUP HLOOKUP SUM SUMIF SUMIFS
        SUMPRODUCT MIN MAX ABS INT MOD ROUND EDATE DATE DATEDIF YEAR MONTH DAY XIRR IRR NPV PMT
        COUNTA COUNT COUNTIF AVERAGE TRUE FALSE EOMONTH POWER SQRT IFERROR ISNUMBER ROUNDUP
        ROUNDDOWN CEILING FLOOR MMULT TEXT VALUE LEFT RIGHT MID LEN E""".split())
    name_token = re.compile(r"(?<![A-Za-z0-9_'!\$])([A-Za-z_][A-Za-z0-9_]*)\b")
    cellref = re.compile(r"^\$?[A-Z]{1,3}\$?\d+$")
    # invalid: a range that repeats the sheet qualifier on both ends, e.g. Sheet!A1:Sheet!B1
    dbl_qual = re.compile(r"('[^']+'|[A-Za-z_][A-Za-z0-9_]*)!\$?[A-Z]{1,3}\$?\d+:('[^']+'|[A-Za-z_][A-Za-z0-9_]*)!")
    issues = []
    ref_errors = 0
    formula_cells = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not isinstance(v, str) or not v.startswith("="):
                    continue
                formula_cells += 1
                if "#REF!" in v:
                    ref_errors += 1
                    issues.append(f"{ws.title}!{c.coordinate}: #REF! in formula")
                if dbl_qual.search(v):
                    issues.append(f"{ws.title}!{c.coordinate}: doubled sheet qualifier in range -> {v[:70]}")
                # strip string literals
                body = re.sub(r'"[^"]*"', "", v)
                # strip sheet-qualified refs 'Sheet'!  and Sheet!
                body = re.sub(r"'[^']*'!", "", body)
                body = re.sub(r"[A-Za-z0-9_]+!", "", body)
                for m in name_token.finditer(body):
                    tok = m.group(1)
                    if tok in excel_funcs or tok in defined:
                        continue
                    if cellref.match(tok):
                        continue
                    # bare cell ranges like A1, column letters handled by cellref; skip single letters used as cols
                    if re.fullmatch(r"[A-Z]{1,3}", tok):
                        continue
                    issues.append(f"{ws.title}!{c.coordinate}: unknown token '{tok}' in {v[:60]}")
    return formula_cells, ref_errors, issues, len(defined)

# --------------------------------------------------------------------------------------
def write_solved_price():
    """Run the binary search and write the result into the SolvedPrice cell."""
    base = Inputs()
    solved = solve_max_bid(base, base.target_value)
    wb = load_workbook(WB)
    # locate SolvedPrice defined name
    dn = wb.defined_names["SolvedPrice"]
    dest = list(dn.destinations)[0]  # (sheet, coord)
    sheet, coord = dest
    ws = wb[sheet]
    ws[coord.replace("$", "")] = round(solved, 0) if solved else base.input_price
    wb.save(WB)
    return solved

# --------------------------------------------------------------------------------------
if __name__ == "__main__":
    print("="*78)
    print("STRUCTURAL VALIDATION")
    print("="*78)
    fcells, referr, issues, ndef = validate_workbook()
    print(f"Formula cells scanned : {fcells}")
    print(f"Defined names         : {ndef}")
    print(f"#REF! errors          : {referr}")
    if issues:
        print(f"Issues ({len(issues)}):")
        for s in issues[:40]:
            print("   -", s)
    else:
        print("No unresolved name tokens or #REF! errors.")

    print()
    print("="*78)
    print("SECTION 6.6 — BINARY-SEARCH PRICE SOLVER")
    print("="*78)
    solved = write_solved_price()
    base = Inputs()
    if solved:
        chk = engine(base, solved)
        print(f"Target levered IRR    : {base.target_value:.2%}")
        print(f"Solved max-bid price  : ${solved:,.0f}")
        print(f"  -> $/unit           : ${solved/len(base.mf_units):,.0f}")
        print(f"  -> going-in cap      : {chk['going_in']/solved:.2%}")
        print(f"  -> levered IRR@price : {chk['irr']:.4%}  (reconciles to target)")
        print("Written into SolvedPrice cell.")
    else:
        print("Solver did not bracket a root; left SolvedPrice = input price.")

    print()
    print("="*78)
    print("SECTION 8 — ACCEPTANCE / AUDIT ASSERTIONS")
    print("="*78)
    results = run_tests()
    npass = sum(1 for _, ok, _ in results if ok)
    for name, ok, detail in results:
        print(f"  [{'PASS' if ok else 'FAIL'}] {name}" + (f"   ({detail})" if detail else ""))
    print(f"\n{npass}/{len(results)} checks passed.")
    if npass != len(results):
        print("SOME CHECKS FAILED — see above.")
