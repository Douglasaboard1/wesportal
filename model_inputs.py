#!/usr/bin/env python3
"""
Single source of truth for the model's sample data and default assumptions.

Both build_model.py (writes the .xlsx) and solve_and_test.py (numeric replica +
solver + Section 8 tests) import from here, so the live Excel model and the audit
replica can never drift apart.

Everything here is an EDITABLE starting point ([DEFAULT — editable] in the spec).
The asset is institutional-scale (88 MF units + ~40k SF ground-floor retail) so the
default price pencils to a realistic ~5.3% going-in cap.
"""

import datetime as dt

# ---- horizon / layout ----
MAX_MONTHS = 120          # built monthly horizon (hold gated <= DispositionDate)
M0 = 5                    # first monthly column (E); month index 0 = acquisition

# ---- dates / identity ----
ASSET_NAME = "Maplewood Commons"
RUN_DATE = dt.date(2026, 6, 16)
ACQ_DATE = dt.date(2026, 7, 1)
DISP_DATE = dt.date(2031, 7, 1)     # 60-month default hold (date-driven)

# ---- growth vectors (Y1..Y11; last repeats) ----
VEC_MFRENT   = [0.04, 0.04, 0.035, 0.03, 0.03, 0.03, 0.03, 0.03, 0.03, 0.03, 0.03]
VEC_RETAIL   = [0.025, 0.0275, 0.03, 0.03, 0.03, 0.03, 0.03, 0.03, 0.03, 0.03, 0.03]
VEC_OPEX     = [0.03] * 11
VEC_TILC     = [0.03] * 11
VEC_OTHERINC = [0.03] * 11

# ---- per-unit other income ($/occupied unit/month), itemized ----
OTHER_INCOME = [
    ("RUBS / utility reimbursement", 45),
    ("Parking", 30),
    ("Storage", 12),
    ("Pet", 8),
    ("Application / admin fees", 5),
    ("Other", 6),
]

# ---- OpEx line items ($/unit/yr, base) ----
OPEX_ITEMS = [
    ("Payroll", 1400), ("R&M", 900), ("Utilities", 1100), ("G&A", 500),
    ("Marketing", 250), ("Insurance", 600), ("Other", 300),
]

# ---- recovery matrix: structure -> (CAM, Tax, Insurance, base-year stop) ----
RECOVERY_MATRIX = [
    ("NNN", 1, 1, 1, 0),
    ("Gross", 0, 0, 0, 0),
    ("Modified Gross", 1, 1, 1, 1),
]


def gen_mf_units(n=88):
    """Institutional rent roll: (id, type, SF, in-place $/mo, market $/mo, expiry mo)."""
    archetypes = [
        ("Studio", 520, 1650), ("1BR", 720, 2150), ("1BR", 760, 2250),
        ("2BR", 1050, 2800), ("2BR", 1080, 2850), ("3BR", 1320, 3500),
    ]
    units = []
    for i in range(n):
        typ, sf, base_mkt = archetypes[i % len(archetypes)]
        market = base_mkt + (i % 5) * 15           # mild dispersion
        inplace = round(market * 0.88)             # ~12% loss-to-lease
        expiry = (i % 24) + 1                       # staggered expiries within 2 yrs
        units.append((f"U{i+1:03d}", typ, sf, inplace, market, expiry))
    return units


def gen_retail_suites():
    """(tenant, SF, lease start mo, expiry mo, base $/SF/yr, step %, structure, sales PSF, overage %)."""
    return [
        ("Grocer Anchor", 16000, 0, 60, 30, 0.020, "NNN", 0, 0.06),
        ("Pharmacy",       9000, 0, 84, 33, 0.015, "NNN", 0, 0.05),
        ("Cafe",           2400, 0, 36, 46, 0.030, "NNN", 0, 0.07),
        ("Fitness",        6000, 0, 48, 31, 0.025, "Modified Gross", 0, 0.06),
        ("Bank",           3200, 0, 72, 40, 0.020, "NNN", 0, 0.05),
        ("Restaurant",     3600, 0, 54, 42, 0.025, "NNN", 0, 0.07),
    ]


# ---- scalar defaults (mirrored into Excel input cells and the Python Inputs struct) ----
DEFAULTS = dict(
    # pricing & cost  (base case ~13% levered IRR; max-bid solver lands ~$46.9M for 15%)
    input_price=48_000_000.0,
    closing_pct=0.018,
    disp_cost_pct=0.015,
    # exit
    exit_cap_blended=0.0550,
    exit_cap_mf=0.0525,
    exit_cap_retail=0.0675,
    # MF ops
    mf_vacancy=0.05, mf_concessions=0.01, mf_baddebt=0.005, ltl_months=12,
    reno_premium=250.0, reno_pace=4.0, reno_cost_unit=18_000.0, reno_downtime=1,
    leaseup_start_occ=0.70, stabilized_occ=0.95, leaseup_absorption=6.0,
    # retail
    retail_genvac=0.05, retail_creditloss=0.01, retail_renewal_prob=0.70,
    retail_downtime=6, renewal_rent_factor=1.00, new_lease_ti=50.0, renewal_ti=15.0,
    lc_pct=0.05, admin_fee_pct=0.15, recov_cam=6.50, recov_tax=4.00, recov_ins=1.25,
    # opex
    mgmt_fee_pct=0.030, repl_reserve_pu=300.0,
    # taxes
    mill_rate=0.0185, assess_ratio=0.90, tax_flat_growth=0.03,
    base_annual_tax=720_000.0, pilot_annual_tax=350_000.0,
    # capex buckets
    cap_repairs=1_500_000.0, cap_tilc=1_200_000.0, cap_contingency=600_000.0,
    # debt
    max_ltv=0.65, max_ltc=0.70, min_dscr=1.25, min_dy=0.085,
    fixed_rate=0.0575, sofr=0.043, spread_bps=250.0, io_months=36,
    amort_years=30, finance_cost_pct=0.01,
    refi_month=36, refi_cap=0.055, refi_ltv=0.65, refi_cost_pct=0.01,
    # waterfall
    lp_split=0.90, gp_coinvest=1.0, pref_rate=0.08, catchup_pct=0.50,
    acq_fee_pct=0.01, am_fee_rate=0.0125, disp_fee_pct=0.01,
    target_return_value=0.15,
)
