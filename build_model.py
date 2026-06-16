#!/usr/bin/env python3
"""
Phase 1 builder for the Institutional Mixed-Use (Multifamily + Ground-Floor Retail)
Acquisition Model.

This script writes the full .xlsx scaffold with openpyxl:
  - 13 tabs (Control Panel ... IC One-Pager)
  - All toggles and global drivers as named ranges
  - Real Excel formula strings (no pre-computed values) for the whole engine,
    EXCEPT the two intentional exceptions in the spec:
        (1) the Phase-1 binary-search solved max-bid price (SolvedPrice cell)
        (2) the Sensitivities grid stubs (rebuilt live in Phase 2)
  - Role-based fonts (blue input / black formula / green cross-tab link)
  - Iterative calculation enabled for circularity (price->loan->equity->IRR)

Re-running this script reproduces the workbook exactly; it is the audit trail.

A reduced-form Python replica of the core cash-flow -> levered-IRR is used only to
(a) binary-search the max-bid price for Solve mode, and (b) run the Section 8
acceptance assertions.  The live Excel model carries the authoritative formulas.
"""

import datetime as dt
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.properties import CalcProperties
from openpyxl.comments import Comment

import model_inputs as MI

OUT = "MixedUse_Acquisition_Model.xlsx"

# --------------------------------------------------------------------------------------
# Global layout constants (shared with the numeric replica via model_inputs)
# --------------------------------------------------------------------------------------
MAX_MONTHS = MI.MAX_MONTHS    # built horizon; hold is gated <= DispositionDate
M0 = MI.M0                    # first monthly column = E (month index 0 = acquisition)
LAST_COL = M0 + MAX_MONTHS - 1
D = MI.DEFAULTS
sample_units = MI.gen_mf_units(88)
suites = MI.gen_retail_suites()
N_UNITS = len(sample_units)
N_SUITES = len(suites)

def mcol(i):
    """Excel column letter for month index i (0-based)."""
    return get_column_letter(M0 + i)

# --------------------------------------------------------------------------------------
# Styling helpers
# --------------------------------------------------------------------------------------
BLUE  = "FF0000FF"   # hardcoded inputs
BLACK = "FF000000"   # formulas
GREEN = "FF008000"   # cross-tab links
WHITE = "FFFFFFFF"

F_INPUT  = Font(color=BLUE)
F_FORM   = Font(color=BLACK)
F_LINK   = Font(color=GREEN)
F_HDR    = Font(color=WHITE, bold=True, size=12)
F_SUB    = Font(bold=True, color=BLACK)
F_TITLE  = Font(bold=True, size=14, color=BLACK)
F_UNIT   = Font(italic=True, color="FF808080", size=9)

FILL_HDR   = PatternFill("solid", fgColor="FF1F3864")   # dark navy banner
FILL_SUB   = PatternFill("solid", fgColor="FFD9E1F2")   # light blue subhead
FILL_INPUT = PatternFill("solid", fgColor="FFFFF2CC")   # pale yellow input cells
FILL_TOG   = PatternFill("solid", fgColor="FFFCE4D6")   # toggle cells
FILL_OK    = PatternFill("solid", fgColor="FFC6EFCE")
FILL_BAD   = PatternFill("solid", fgColor="FFFFC7CE")

THIN = Side(style="thin", color="FFBFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# number formats
NF_USD   = '#,##0;(#,##0)'
NF_USD0  = '$#,##0;($#,##0)'
NF_USD2  = '$#,##0.00'
NF_PSF   = '$#,##0.00'
NF_PCT   = '0.0%'
NF_PCT2  = '0.00%'
NF_MULT  = '0.00"x"'
NF_DATE  = 'mmm-yyyy'
NF_NUM   = '#,##0'
NF_MO    = '0" mo"'
NF_BPS   = '0" bps"'

def q(sheet):
    return f"'{sheet}'" if (" " in sheet or "-" in sheet) else sheet

# Excel rejects a cross-sheet range that repeats the sheet qualifier on both ends
# (e.g. 'Monthly CF'!E5:'Monthly CF'!DT5). The qualifier belongs only on the first
# cell: 'Monthly CF'!E5:DT5. Collapse any same-sheet doubled qualifier centrally so
# every formula written through put() is valid regardless of how it was assembled.
_DBL_QUAL = re.compile(r"('[^']+'|[A-Za-z_][A-Za-z0-9_]*)!(\$?[A-Z]{1,3}\$?\d+):\1!(\$?[A-Z]{1,3}\$?\d+)")
def fix_ranges(formula):
    prev = None
    while prev != formula:
        prev = formula
        formula = _DBL_QUAL.sub(r"\1!\2:\3", formula)
    return formula

# --------------------------------------------------------------------------------------
# Workbook / cell-writing utilities
# --------------------------------------------------------------------------------------
wb = Workbook()
named = {}   # registry: name -> "'Sheet'!$X$Y"

def add_name(name, sheet, coord):
    ref = f"{q(sheet)}!{coord}"
    named[name] = ref
    wb.defined_names[name] = DefinedName(name, attr_text=ref)
    return name

def put(ws, row, col, value=None, role=None, nf=None, bold=False,
        fill=None, align=None, border=False, wrap=False, size=None, name=None):
    """Write a cell with role-based font + formatting. Returns the cell."""
    c = ws.cell(row=row, column=col)
    if value is not None:
        if isinstance(value, str) and value.startswith("="):
            value = fix_ranges(value)
        c.value = value
    font_kwargs = {}
    if role == "input":   font_kwargs["color"] = BLUE
    elif role == "link":  font_kwargs["color"] = GREEN
    elif role == "form":  font_kwargs["color"] = BLACK
    elif role == "hdr":   font_kwargs.update(color=WHITE, bold=True, size=size or 12)
    elif role == "title": font_kwargs.update(bold=True, size=size or 14)
    elif role == "sub":   font_kwargs.update(bold=True)
    elif role == "unit":  font_kwargs.update(italic=True, color="FF808080", size=9)
    if bold: font_kwargs["bold"] = True
    if size and "size" not in font_kwargs: font_kwargs["size"] = size
    if font_kwargs:
        c.font = Font(**font_kwargs)
    if nf: c.number_format = nf
    if fill: c.fill = fill
    if align: c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    elif wrap: c.alignment = Alignment(wrap_text=True, vertical="center")
    if border: c.border = BORDER
    if name:
        add_name(name, ws.title, f"${get_column_letter(col)}${row}")
    return c

def banner(ws, title, ncols=12):
    put(ws, 1, 1, title, role="hdr", fill=FILL_HDR, size=13)
    for col in range(1, ncols + 1):
        ws.cell(row=1, column=col).fill = FILL_HDR
        if ws.cell(row=1, column=col).value is None:
            ws.cell(row=1, column=col).value = ""
    # standard header block row 2-3
    put(ws, 2, 1, "Asset:", role="sub")
    put(ws, 2, 2, "=AssetName", role="link")
    put(ws, 2, 4, "Run date:", role="sub")
    put(ws, 2, 5, "=RunDate", role="link", nf=NF_DATE)
    put(ws, 2, 7, "Business plan:", role="sub")
    put(ws, 2, 8, "=BusinessPlanMode", role="link")
    put(ws, 2, 10, "Hold (mo):", role="sub")
    put(ws, 2, 11, "=HoldMonths", role="link", nf=NF_NUM)
    put(ws, 3, 1, "Model status:", role="sub")
    put(ws, 3, 2, "=ModelStatus", role="link", bold=True)

def subhead(ws, row, text, c0=1, c1=12):
    put(ws, row, c0, text, role="sub", fill=FILL_SUB)
    for col in range(c0, c1 + 1):
        ws.cell(row=row, column=col).fill = FILL_SUB

def colhdr(ws, row, headers, start=1):
    for j, h in enumerate(headers):
        put(ws, row, start + j, h, role="sub", fill=FILL_SUB, align="center", border=True, wrap=True)

# ======================================================================================
# TAB 1 — CONTROL PANEL  (build first; all named toggles)
# ======================================================================================
cp = wb.active
cp.title = "Control Panel"
banner(cp, "CONTROL PANEL — Master Toggles", 8)
cp.column_dimensions["A"].width = 34
cp.column_dimensions["B"].width = 26
cp.column_dimensions["C"].width = 40
cp.column_dimensions["D"].width = 16

r = 5
put(cp, r, 1, "Header / identity", role="sub", fill=FILL_SUB); subhead(cp, r, "Header / identity", 1, 4); r += 1
put(cp, r, 1, "Asset Name", role="form")
put(cp, r, 2, MI.ASSET_NAME, role="input", fill=FILL_INPUT, name="AssetName"); r += 1
put(cp, r, 1, "Run Date", role="form")
put(cp, r, 2, MI.RUN_DATE, role="input", nf=NF_DATE, fill=FILL_INPUT, name="RunDate"); r += 1
put(cp, r, 1, "Acquisition Date", role="form")
put(cp, r, 2, MI.ACQ_DATE, role="input", nf=NF_DATE, fill=FILL_INPUT, name="AcquisitionDate"); r += 1
put(cp, r, 1, "Disposition Date", role="form")
put(cp, r, 2, MI.DISP_DATE, role="input", nf=NF_DATE, fill=FILL_INPUT, name="DispositionDate"); r += 1
put(cp, r, 1, "Hold (months)", role="form")
put(cp, r, 2, "=DATEDIF(AcquisitionDate,DispositionDate,\"m\")", role="form", nf=NF_NUM, name="HoldMonths"); r += 1
put(cp, r, 1, "Hold (years)", role="form")
put(cp, r, 2, "=HoldMonths/12", role="form", nf='0.0', name="HoldYears"); r += 2

def toggle(label, options_text, default, name, note="", nf=None):
    global r
    put(cp, r, 1, label, role="form")
    c = put(cp, r, 2, default, role="input", fill=FILL_TOG, name=name)
    if nf: c.number_format = nf
    put(cp, r, 3, options_text, role="unit")
    if note:
        c.comment = Comment(note, "Model")
    r += 1
    return name

subhead(cp, r, "Master toggles (referenced by name everywhere)", 1, 4); r += 1
toggle("1. Business Plan Mode", "Stabilized | Value-Add | Core-Plus Lease-Up", "Value-Add", "BusinessPlanMode")
toggle("3. Price Mode", "Input Price | Solve for Target Return", "Input Price", "PriceMode")
toggle("4. Target Return Basis", "Levered IRR | Unlevered IRR | Untrended YoC", "Levered IRR", "TargetReturnBasis")
toggle("4b. Target Return Value", "target metric value (e.g. 0.15 = 15% IRR)", 0.15, "TargetReturnValue", nf=NF_PCT)
toggle("5. Refi", "On | Off", "Off", "RefiOn")
toggle("6. Debt Structure", "Single Blended Loan | Separate MF + Retail Loans", "Single Blended Loan", "DebtStructure")
toggle("7a. Sizing: LTV constraint", "On | Off", "On", "Use_LTV")
toggle("7b. Sizing: LTC constraint", "On | Off", "On", "Use_LTC")
toggle("7c. Sizing: DSCR constraint", "On | Off", "On", "Use_DSCR")
toggle("7d. Sizing: Debt Yield constraint", "On | Off", "On", "Use_DY")
toggle("8. Loan Rate Type", "Fixed | Floating (SOFR + spread)", "Fixed", "LoanRateType")
toggle("9. Retail Lease Structure (global)", "NNN | Gross | Modified Gross", "NNN", "RetailLeaseStructure")
toggle("10. Recovery Method", "Pro-rata", "Pro-rata", "RecoveryMethod")
toggle("10a. Admin Fee", "On | Off", "On", "AdminFeeOn")
toggle("10b. Gross-Up", "On | Off", "On", "GrossUpOn")
toggle("11. Exit Valuation Method", "Blended Cap on Total NOI | Component Caps", "Blended Cap on Total NOI", "ExitValMethod")
toggle("12. Exit NOI Basis", "Trailing 12-mo | Forward 12-mo", "Forward 12-mo", "ExitNOIBasis")
toggle("13. Tax Method", "Flat Growth | Reassessment-on-Sale | PILOT / Abatement", "Reassessment-on-Sale", "TaxMethod")
toggle("15a. Acquisition Fee", "On | Off", "On", "AcqFeeOn")
toggle("15b. Asset Management Fee", "On | Off", "On", "AMFeeOn")
toggle("15c. AM Fee Basis", "% of Equity | % of Cost | % of NOI", "% of Equity", "AMFeeBasis")
toggle("15d. Disposition Fee", "On | Off", "On", "DispFeeOn")
toggle("16. GP Catch-Up", "On | Off", "On", "GPCatchUpOn")
toggle("17. LP Present", "Yes | No", "Yes", "LP_Present")
r += 1
subhead(cp, r, "Capital budget funding (per bucket): Equity | Financed", 1, 4); r += 1
toggle("14a. Immediate/Deferred Repairs funding", "Equity | Financed", "Financed", "Fund_Repairs")
toggle("14b. Value-Add Unit Renovation funding", "Equity | Financed", "Financed", "Fund_Reno")
toggle("14c. Retail TI/LC reserve funding", "Equity | Financed", "Financed", "Fund_TILC")
toggle("14d. Contingency funding", "Equity | Financed", "Equity", "Fund_Contingency")

# Model status (computed after checks defined on Summary; placeholder formula here updated later)
r += 1
subhead(cp, r, "Model status (red on any failed check)", 1, 4); r += 1
put(cp, r, 1, "Model Status", role="sub")
put(cp, r, 2, '=IF(Summary!$B$3=TRUE,"OK","CHECK FAILED")', role="form", bold=True, name="ModelStatus")
CP_STATUS_ROW = r

# ======================================================================================
# TAB 3 — ASSUMPTIONS (global drivers + Growth block + dynamic monthly date series)
# (built before revenue tabs so they can reference it; placed left->right order later)
# ======================================================================================
asm = wb.create_sheet("Assumptions")
banner(asm, "ASSUMPTIONS — Global Drivers & Growth Block", 14)
asm.column_dimensions["A"].width = 34
for col in "BCD":
    asm.column_dimensions[col].width = 14

r = 5
subhead(asm, r, "Pricing & cost", 1, 4); r += 1
put(asm, r, 1, "Input Purchase Price ($)", role="form")
put(asm, r, 2, D["input_price"], role="input", nf=NF_USD0, fill=FILL_INPUT, name="InputPrice"); r += 1
put(asm, r, 1, "Solved Max-Bid Price ($) [Python binary search]", role="form")
# This value is overwritten by the Python solver in solve_and_test.py (intentional exception).
put(asm, r, 2, D["input_price"], role="input", nf=NF_USD0, fill=FILL_INPUT, name="SolvedPrice"); r += 1
put(asm, r, 1, "Active Purchase Price ($)", role="form")
put(asm, r, 2, '=IF(PriceMode="Solve for Target Return",SolvedPrice,InputPrice)',
    role="form", nf=NF_USD0, name="PurchasePrice"); r += 1
put(asm, r, 1, "Closing / transaction cost (% of price)", role="form")
put(asm, r, 2, 0.018, role="input", nf=NF_PCT, fill=FILL_INPUT, name="ClosingCostPct"); r += 1
put(asm, r, 1, "Disposition cost (% of sale)", role="form")
put(asm, r, 2, 0.015, role="input", nf=NF_PCT, fill=FILL_INPUT, name="DispCostPct"); r += 2

subhead(asm, r, "Exit assumptions", 1, 4); r += 1
put(asm, r, 1, "Exit cap — blended (%)", role="form")
put(asm, r, 2, 0.055, role="input", nf=NF_PCT2, fill=FILL_INPUT, name="ExitCapBlended"); r += 1
put(asm, r, 1, "Exit cap — MF component (%)", role="form")
put(asm, r, 2, 0.0525, role="input", nf=NF_PCT2, fill=FILL_INPUT, name="ExitCapMF"); r += 1
put(asm, r, 1, "Exit cap — Retail component (%)", role="form")
put(asm, r, 2, 0.0675, role="input", nf=NF_PCT2, fill=FILL_INPUT, name="ExitCapRetail"); r += 2

subhead(asm, r, "MF operating assumptions", 1, 4); r += 1
put(asm, r, 1, "Economic vacancy (% of GPR)", role="form")
put(asm, r, 2, 0.05, role="input", nf=NF_PCT, fill=FILL_INPUT, name="MF_Vacancy"); r += 1
put(asm, r, 1, "Concessions (% of GPR)", role="form")
put(asm, r, 2, 0.01, role="input", nf=NF_PCT, fill=FILL_INPUT, name="MF_Concessions"); r += 1
put(asm, r, 1, "Bad debt (% of GPR)", role="form")
put(asm, r, 2, 0.005, role="input", nf=NF_PCT, fill=FILL_INPUT, name="MF_BadDebt"); r += 1
put(asm, r, 1, "Loss-to-lease burn-off (months to market)", role="form")
put(asm, r, 2, 12, role="input", nf=NF_MO, fill=FILL_INPUT, name="LTL_Months"); r += 2

subhead(asm, r, "Value-Add / Core-Plus ramps", 1, 4); r += 1
put(asm, r, 1, "Renovation premium ($/unit/mo)", role="form")
put(asm, r, 2, 250, role="input", nf=NF_USD0, fill=FILL_INPUT, name="RenoPremium"); r += 1
put(asm, r, 1, "Renovation pace (units/month)", role="form")
put(asm, r, 2, 4, role="input", nf=NF_NUM, fill=FILL_INPUT, name="RenoPace"); r += 1
put(asm, r, 1, "Renovation cost ($/unit)", role="form")
put(asm, r, 2, 18_000, role="input", nf=NF_USD0, fill=FILL_INPUT, name="RenoCostUnit"); r += 1
put(asm, r, 1, "Renovation downtime (months/unit)", role="form")
put(asm, r, 2, 1, role="input", nf=NF_MO, fill=FILL_INPUT, name="RenoDowntime"); r += 1
put(asm, r, 1, "Lease-up: start occupancy (%)", role="form")
put(asm, r, 2, 0.70, role="input", nf=NF_PCT, fill=FILL_INPUT, name="LeaseUpStartOcc"); r += 1
put(asm, r, 1, "Lease-up: stabilized occupancy (%)", role="form")
put(asm, r, 2, 0.95, role="input", nf=NF_PCT, fill=FILL_INPUT, name="StabilizedOcc"); r += 1
put(asm, r, 1, "Lease-up: absorption (units/month)", role="form")
put(asm, r, 2, 6, role="input", nf=NF_NUM, fill=FILL_INPUT, name="LeaseUpAbsorption"); r += 2

subhead(asm, r, "Retail assumptions", 1, 4); r += 1
put(asm, r, 1, "Retail general vacancy (% of retail income)", role="form")
put(asm, r, 2, 0.05, role="input", nf=NF_PCT, fill=FILL_INPUT, name="RetailGenVacancy"); r += 1
put(asm, r, 1, "Retail credit loss / bad debt (%)", role="form")
put(asm, r, 2, 0.01, role="input", nf=NF_PCT, fill=FILL_INPUT, name="RetailCreditLoss"); r += 1
put(asm, r, 1, "Retail renewal probability (%)", role="form")
put(asm, r, 2, 0.70, role="input", nf=NF_PCT, fill=FILL_INPUT, name="RetailRenewalProb"); r += 1
put(asm, r, 1, "Retail re-let downtime (months)", role="form")
put(asm, r, 2, 6, role="input", nf=NF_MO, fill=FILL_INPUT, name="RetailDowntime"); r += 1
put(asm, r, 1, "Renewal rent factor (% of market)", role="form")
put(asm, r, 2, 1.00, role="input", nf=NF_PCT, fill=FILL_INPUT, name="RenewalRentFactor"); r += 1
put(asm, r, 1, "New-lease TI ($/SF)", role="form")
put(asm, r, 2, 50, role="input", nf=NF_PSF, fill=FILL_INPUT, name="NewLeaseTI"); r += 1
put(asm, r, 1, "Renewal TI ($/SF)", role="form")
put(asm, r, 2, 15, role="input", nf=NF_PSF, fill=FILL_INPUT, name="RenewalTI"); r += 1
put(asm, r, 1, "Leasing commission (% of lease value)", role="form")
put(asm, r, 2, 0.05, role="input", nf=NF_PCT, fill=FILL_INPUT, name="LeasingCommissionPct"); r += 1
put(asm, r, 1, "Admin fee (% of recoveries)", role="form")
put(asm, r, 2, 0.15, role="input", nf=NF_PCT, fill=FILL_INPUT, name="AdminFeePct"); r += 2

subhead(asm, r, "OpEx", 1, 4); r += 1
put(asm, r, 1, "Management fee (% of EGI)", role="form")
put(asm, r, 2, 0.030, role="input", nf=NF_PCT, fill=FILL_INPUT, name="MgmtFeePct"); r += 1
put(asm, r, 1, "Replacement reserves ($/unit/yr)", role="form")
put(asm, r, 2, 300, role="input", nf=NF_USD0, fill=FILL_INPUT, name="ReplReservePU"); r += 2

subhead(asm, r, "Taxes", 1, 4); r += 1
put(asm, r, 1, "Mill rate (tax per $ assessed)", role="form")
put(asm, r, 2, 0.0185, role="input", nf=NF_PCT2, fill=FILL_INPUT, name="MillRate"); r += 1
put(asm, r, 1, "Assessment ratio (% of value)", role="form")
put(asm, r, 2, 0.90, role="input", nf=NF_PCT, fill=FILL_INPUT, name="AssessRatio"); r += 1
put(asm, r, 1, "Tax flat growth (%/yr)", role="form")
put(asm, r, 2, 0.03, role="input", nf=NF_PCT, fill=FILL_INPUT, name="TaxFlatGrowth"); r += 1
put(asm, r, 1, "Base annual tax ($) [year 1]", role="form")
put(asm, r, 2, 720_000, role="input", nf=NF_USD0, fill=FILL_INPUT, name="BaseAnnualTax"); r += 1
put(asm, r, 1, "PILOT annual tax ($) [flat schedule proxy]", role="form")
put(asm, r, 2, 350_000, role="input", nf=NF_USD0, fill=FILL_INPUT, name="PILOTAnnualTax"); r += 2

# ---- Growth block (annual vectors, year 1..11; last repeats) ----
subhead(asm, r, "GROWTH BLOCK — annual vectors (year-over-year)", 1, 14); r += 1
GROWTH_HDR_ROW = r
put(asm, r, 1, "Vector \\ Year", role="sub")
for y in range(1, 12):
    put(asm, r, 4 + y, f"Y{y}", role="sub", align="center")
r += 1
def growth_row(label, vals, name):
    global r
    put(asm, r, 1, label, role="form")
    base_col = 5
    for j, v in enumerate(vals):
        put(asm, r, base_col + j, v, role="input", nf=NF_PCT, fill=FILL_INPUT)
    add_name(name, asm.title, f"${get_column_letter(base_col)}${r}:${get_column_letter(base_col+10)}${r}")
    r += 1

growth_row("MF market rent growth", MI.VEC_MFRENT, "Vec_MFRent")
growth_row("Retail market rent growth", MI.VEC_RETAIL, "Vec_RetailRent")
growth_row("OpEx / general inflation", MI.VEC_OPEX, "Vec_OpEx")
growth_row("TI/LC growth", MI.VEC_TILC, "Vec_TILC")
growth_row("Other income growth", MI.VEC_OTHERINC, "Vec_OtherInc")
r += 1

# Cumulative annual growth factors (year 0..11): factor applied during year y = product of (1+g_k) for k<y
subhead(asm, r, "Cumulative growth factors (year 0=1.0; applied at anniversary)", 1, 14); r += 1
CUMHDR = r
put(asm, r, 1, "Factor \\ Year", role="sub")
for y in range(0, 12):
    put(asm, r, 4 + y, f"Y{y}", role="sub", align="center")
r += 1
def cum_row(label, vec_name, name):
    global r
    put(asm, r, 1, label, role="form")
    # Y0 = 1
    put(asm, r, 4, 1.0, role="form", nf='0.0000')
    add_name(name, asm.title, f"${get_column_letter(4)}${r}:${get_column_letter(15)}${r}")
    for y in range(1, 12):
        prev = get_column_letter(4 + y - 1)
        gcol = get_column_letter(5 + (y - 1))  # vector col for year y (Y1 at col E)
        put(asm, r, 4 + y, f"={prev}{r}*(1+{q('Assumptions')}!{gcol}${GROWTH_HDR_ROW + 1 + vec_idx})",
            role="form", nf='0.0000')
    r += 1

# Map vector label -> row offset for cumulative formulas
vec_rows = {  # vector name -> its row in growth block
    "Vec_MFRent": GROWTH_HDR_ROW + 1,
    "Vec_RetailRent": GROWTH_HDR_ROW + 2,
    "Vec_OpEx": GROWTH_HDR_ROW + 3,
    "Vec_TILC": GROWTH_HDR_ROW + 4,
    "Vec_OtherInc": GROWTH_HDR_ROW + 5,
}
def cum_row2(label, vec_row, name):
    global r
    put(asm, r, 1, label, role="form")
    put(asm, r, 4, 1.0, role="form", nf='0.0000')
    add_name(name, asm.title, f"${get_column_letter(4)}${r}:${get_column_letter(15)}${r}")
    for y in range(1, 12):
        prev = get_column_letter(4 + y - 1)
        gcol = get_column_letter(5 + (y - 1))
        put(asm, r, 4 + y, f"={prev}{r}*(1+{gcol}${vec_row})", role="form", nf='0.0000')
    r += 1

cum_row2("MF rent cum factor", vec_rows["Vec_MFRent"], "Cum_MFRent")
cum_row2("Retail rent cum factor", vec_rows["Vec_RetailRent"], "Cum_RetailRent")
cum_row2("OpEx cum factor", vec_rows["Vec_OpEx"], "Cum_OpEx")
cum_row2("TI/LC cum factor", vec_rows["Vec_TILC"], "Cum_TILC")
cum_row2("Other income cum factor", vec_rows["Vec_OtherInc"], "Cum_OtherInc")
CUM_FIRST_ROW = CUMHDR + 1
r += 2

# ---- Dynamic monthly period header (the spine every schedule references) ----
subhead(asm, r, "MONTHLY PERIOD HEADER (dynamic; every schedule references these rows)", 1, 14); r += 1
put(asm, r, 1, "Month index (0 = acquisition)", role="sub")
PERIOD_IDX_ROW = r
for i in range(MAX_MONTHS):
    put(asm, r, M0 + i, i, role="form", nf=NF_NUM)
r += 1
put(asm, r, 1, "Period date", role="sub")
PERIOD_DATE_ROW = r
for i in range(MAX_MONTHS):
    if i == 0:
        put(asm, r, M0 + i, "=AcquisitionDate", role="form", nf=NF_DATE)
    else:
        put(asm, r, M0 + i, f"=EDATE(AcquisitionDate,{i})", role="form", nf=NF_DATE)
r += 1
put(asm, r, 1, "Operating-month flag (1 if active)", role="sub")
OP_ACTIVE_ROW = r
for i in range(MAX_MONTHS):
    col = mcol(i)
    if i == 0:
        put(asm, r, M0 + i, 0, role="form", nf=NF_NUM)  # acquisition month: no ops
    else:
        put(asm, r, M0 + i, f"=IF({col}${PERIOD_DATE_ROW}<=DispositionDate,1,0)", role="form", nf=NF_NUM)
r += 1
put(asm, r, 1, "Hold-year number (1-based)", role="sub")
HOLD_YEAR_ROW = r
for i in range(MAX_MONTHS):
    col = mcol(i)
    put(asm, r, M0 + i, f"=IF({col}${OP_ACTIVE_ROW}=1,INT(({col}${PERIOD_IDX_ROW}-1)/12)+1,0)", role="form", nf=NF_NUM)
r += 1
put(asm, r, 1, "Exit-month flag (1 at disposition)", role="sub")
EXIT_FLAG_ROW = r
for i in range(MAX_MONTHS):
    col = mcol(i)
    nxt = mcol(i + 1) if i + 1 < MAX_MONTHS else None
    if nxt:
        put(asm, r, M0 + i, f"=IF(AND({col}${OP_ACTIVE_ROW}=1,{nxt}${OP_ACTIVE_ROW}=0),1,0)", role="form", nf=NF_NUM)
    else:
        put(asm, r, M0 + i, f"=IF({col}${OP_ACTIVE_ROW}=1,1,0)", role="form", nf=NF_NUM)
r += 1
put(asm, r, 1, "Year-fraction growth factor index (cum-factor col)", role="sub")
# helper: which cumulative-factor column (Y0..Y11 -> cols D..O i.e. 4..15) to use for this month
YEAR_COLREF_ROW = r
for i in range(MAX_MONTHS):
    col = mcol(i)
    # year number capped at 11; cum factor col = 4 + min(yearnum,11) ; but we INDEX instead
    put(asm, r, M0 + i, f"=MIN(INT(MAX({col}${HOLD_YEAR_ROW}-1,0)),10)+1", role="form", nf=NF_NUM)
r += 1

ASM = "Assumptions"
add_name("PeriodDateRow", ASM, f"${get_column_letter(M0)}${PERIOD_DATE_ROW}")  # anchor only

# Helper to reference a cumulative growth factor for month i (returns formula fragment)
def cum_factor_ref(cum_name, i):
    col = mcol(i)
    # cum vector spans D..O (cols 4..15). Use INDEX with the year-col-ref helper.
    return f"INDEX({cum_name},1,{col}${YEAR_COLREF_ROW})"

# ======================================================================================
# TAB 4 — MF REVENUE (unit-by-unit rent roll -> monthly EGI)
# ======================================================================================
mf = wb.create_sheet("MF Revenue")
banner(mf, "MF REVENUE — Unit-by-Unit Rent Roll", 14)
mf.column_dimensions["A"].width = 10
mf.column_dimensions["B"].width = 12
for col in "CD":
    mf.column_dimensions[col].width = 12

# Rent-roll block
RR_HDR = 5
colhdr(mf, RR_HDR, ["Unit ID","Type","SF","In-place rent ($/mo)","Market rent ($/mo)","Lease expiry (mo idx)","Status"], start=1)
RR0 = RR_HDR + 1
for k,(uid,typ,sf,inplace,market,exp) in enumerate(sample_units):
    rr = RR0 + k
    put(mf, rr, 1, uid, role="input", fill=FILL_INPUT, align="center")
    put(mf, rr, 2, typ, role="input", fill=FILL_INPUT, align="center")
    put(mf, rr, 3, sf, role="input", nf=NF_NUM, fill=FILL_INPUT)
    put(mf, rr, 4, inplace, role="input", nf=NF_USD0, fill=FILL_INPUT)
    put(mf, rr, 5, market, role="input", nf=NF_USD0, fill=FILL_INPUT)
    put(mf, rr, 6, exp, role="input", nf=NF_NUM, fill=FILL_INPUT)
    put(mf, rr, 7, "Occupied", role="input", fill=FILL_INPUT, align="center")
RR_LAST = RR0 + len(sample_units) - 1
# totals
TOT = RR_LAST + 1
put(mf, TOT, 1, "TOTALS", role="sub")
put(mf, TOT, 3, f"=SUM(C{RR0}:C{RR_LAST})", role="form", nf=NF_NUM, name="MF_TotalSF")
put(mf, TOT, 4, f"=SUM(D{RR0}:D{RR_LAST})", role="form", nf=NF_USD0)
put(mf, TOT, 5, f"=SUM(E{RR0}:E{RR_LAST})", role="form", nf=NF_USD0)
add_name("MF_UnitCount", mf.title, f"$H${TOT}")
put(mf, TOT, 8, f"=COUNTA(A{RR0}:A{RR_LAST})", role="form", nf=NF_NUM)
put(mf, TOT-0, 9, "<- unit count", role="unit")

# Loss-to-lease note
put(mf, TOT+1, 1, "Loss-to-lease = Market - In-place, burning off over LTL_Months to market.", role="unit")

# ---- Monthly MF build (aggregate rows referencing rent roll) ----
mrow = TOT + 3
subhead(mf, mrow, "MONTHLY MF BUILD (-> MF EGI). Cols E.. align to Assumptions period header.", 1, 14); mrow += 1
put(mf, mrow, 1, "Month index", role="sub")
for i in range(MAX_MONTHS):
    put(mf, mrow, M0 + i, f"={q(ASM)}!{mcol(i)}${PERIOD_IDX_ROW}", role="link", nf=NF_NUM)
mrow += 1

# Occupancy factor (Core-Plus lease-up ramps; else stabilized)
put(mf, mrow, 1, "Occupied units (count)", role="form")
OCC_ROW = mrow
for i in range(MAX_MONTHS):
    col = mcol(i)
    act = f"{q(ASM)}!{col}${OP_ACTIVE_ROW}"
    idx = f"{q(ASM)}!{col}${PERIOD_IDX_ROW}"
    # Core-Plus: start occ ramps by absorption to stabilized; else stabilized occ * units
    f = (f'=IF({act}=0,0,'
         f'IF(BusinessPlanMode="Core-Plus Lease-Up",'
         f'MIN(MF_UnitCount*StabilizedOcc, MF_UnitCount*LeaseUpStartOcc + LeaseUpAbsorption*{idx}),'
         f'MF_UnitCount*StabilizedOcc))')
    put(mf, mrow, M0 + i, f, role="form", nf='0.0')
mrow += 1

# Renovated units cumulative (Value-Add)
put(mf, mrow, 1, "Renovated units (cumulative)", role="form")
RENO_ROW = mrow
for i in range(MAX_MONTHS):
    col = mcol(i); idx = f"{q(ASM)}!{col}${PERIOD_IDX_ROW}"
    act = f"{q(ASM)}!{col}${OP_ACTIVE_ROW}"
    f = (f'=IF(BusinessPlanMode="Value-Add",MIN(MF_UnitCount,MAX(0,RenoPace*({idx}-RenoDowntime)))*{act},0)')
    put(mf, mrow, M0 + i, f, role="form", nf='0.0')
mrow += 1

# Average in-place & market (per occupied unit) from rent roll
put(mf, mrow, 1, "Total in-place rent ($/mo, roll)", role="form")
INPLACE_ROW = mrow
for i in range(MAX_MONTHS):
    col = mcol(i)
    put(mf, mrow, M0 + i, f"=$D${TOT}", role="form", nf=NF_USD0)  # static base sum; grown below
mrow += 1
put(mf, mrow, 1, "Total market rent ($/mo, roll)", role="form")
MARKET_ROW = mrow
for i in range(MAX_MONTHS):
    put(mf, mrow, M0 + i, f"=$E${TOT}", role="form", nf=NF_USD0)
mrow += 1

# Gross potential rent (GPR), grown by MF cum factor, with reno premium on renovated units
put(mf, mrow, 1, "Gross Potential Rent ($/mo)", role="form")
GPR_ROW = mrow
for i in range(MAX_MONTHS):
    col = mcol(i)
    cf = cum_factor_ref("Cum_MFRent", i)
    market = f"{col}${MARKET_ROW}"
    reno = f"{col}${RENO_ROW}*RenoPremium"
    act = f"{q(ASM)}!{col}${OP_ACTIVE_ROW}"
    put(mf, mrow, M0 + i, f"=({market}*{cf}+{reno})*{act}", role="form", nf=NF_USD0)
mrow += 1

# Loss-to-lease (burn-off): difference market-inplace prorated by remaining burn, only when in-place < market
put(mf, mrow, 1, "Loss-to-lease ($/mo, negative)", role="form")
LTL_ROW = mrow
for i in range(MAX_MONTHS):
    col = mcol(i); idx = f"{q(ASM)}!{col}${PERIOD_IDX_ROW}"
    act = f"{q(ASM)}!{col}${OP_ACTIVE_ROW}"
    gap = f"({col}${MARKET_ROW}-{col}${INPLACE_ROW})"
    rem = f"MAX(0,1-{idx}/LTL_Months)"
    put(mf, mrow, M0 + i, f"=-{gap}*{rem}*{act}", role="form", nf=NF_USD0)
mrow += 1

# Vacancy/concession/bad debt on GPR (and occupancy ramp shortfall in Core-Plus)
put(mf, mrow, 1, "Vacancy / concessions / bad debt ($/mo, neg)", role="form")
VAC_ROW = mrow
for i in range(MAX_MONTHS):
    col = mcol(i)
    gpr = f"{col}${GPR_ROW}"
    occfrac = f"({col}${OCC_ROW}/MF_UnitCount)"
    base_loss = f"{gpr}*(MF_Vacancy+MF_Concessions+MF_BadDebt)"
    rampshort = f"{gpr}*(1-{occfrac})"
    put(mf, mrow, M0 + i, f"=-({base_loss}+{rampshort})", role="form", nf=NF_USD0)
mrow += 1

# Other income (itemized)
oi_items = MI.OTHER_INCOME
put(mf, mrow, 1, "Other income — itemized ($/unit/mo)", role="sub"); mrow += 1
OI_FIRST = mrow
for nm, val in oi_items:
    put(mf, mrow, 1, "  " + nm, role="form")
    put(mf, mrow, 2, val, role="input", nf=NF_USD0, fill=FILL_INPUT)
    for i in range(MAX_MONTHS):
        col = mcol(i); cf = cum_factor_ref("Cum_OtherInc", i)
        act = f"{q(ASM)}!{col}${OP_ACTIVE_ROW}"
        put(mf, mrow, M0 + i, f"=$B${mrow}*{col}${OCC_ROW}*{cf}*{act}", role="form", nf=NF_USD0)
    mrow += 1
OI_LAST = mrow - 1
put(mf, mrow, 1, "Total other income ($/mo)", role="form")
OI_TOT_ROW = mrow
for i in range(MAX_MONTHS):
    col = mcol(i)
    put(mf, mrow, M0 + i, f"=SUM({col}{OI_FIRST}:{col}{OI_LAST})", role="form", nf=NF_USD0)
mrow += 1

# MF EGI
put(mf, mrow, 1, "MF EGI ($/mo)", role="sub")
MF_EGI_ROW = mrow
for i in range(MAX_MONTHS):
    col = mcol(i)
    put(mf, mrow, M0 + i,
        f"={col}${GPR_ROW}+{col}${LTL_ROW}+{col}${VAC_ROW}+{col}${OI_TOT_ROW}",
        role="form", nf=NF_USD0, bold=True)
mrow += 1
add_name("MF_EGI_Row", mf.title, f"${mcol(0)}${MF_EGI_ROW}")

# ======================================================================================
# TAB 5 — RETAIL REVENUE (tenant-by-tenant + recoveries + rollover + %rent + credit loss)
# ======================================================================================
rt = wb.create_sheet("Retail Revenue")
banner(rt, "RETAIL REVENUE — Tenant-by-Tenant", 14)
rt.column_dimensions["A"].width = 18

# Recovery matrix (editable: structure -> reimbursed buckets)
rm = 5
subhead(rt, rm, "Recovery matrix (editable: 1=tenant reimburses, 0=landlord absorbs)", 1, 8); rm += 1
colhdr(rt, rm, ["Structure","CAM","Tax","Insurance","Base-year stop?"], start=1)
RM0 = rm + 1
recov = MI.RECOVERY_MATRIX
for k,(s,cam,tax,ins,stop) in enumerate(recov):
    rr = RM0 + k
    put(rt, rr, 1, s, role="input", fill=FILL_INPUT)
    put(rt, rr, 2, cam, role="input", fill=FILL_INPUT, align="center")
    put(rt, rr, 3, tax, role="input", fill=FILL_INPUT, align="center")
    put(rt, rr, 4, ins, role="input", fill=FILL_INPUT, align="center")
    put(rt, rr, 5, stop, role="input", fill=FILL_INPUT, align="center")
RM_LAST = RM0 + len(recov) - 1
add_name("RecovMatrix", rt.title, f"$A${RM0}:$E${RM_LAST}")

# Recoverable expense pool ($/SF/yr) input
rm = RM_LAST + 2
put(rt, rm, 1, "Recoverable CAM ($/SF/yr)", role="form")
put(rt, rm, 2, 6.50, role="input", nf=NF_PSF, fill=FILL_INPUT, name="RecovCAM"); rm += 1
put(rt, rm, 1, "Recoverable Tax ($/SF/yr)", role="form")
put(rt, rm, 2, 4.00, role="input", nf=NF_PSF, fill=FILL_INPUT, name="RecovTax"); rm += 1
put(rt, rm, 1, "Recoverable Insurance ($/SF/yr)", role="form")
put(rt, rm, 2, 1.25, role="input", nf=NF_PSF, fill=FILL_INPUT, name="RecovIns"); rm += 1

# Tenant rent roll
trh = rm + 2
colhdr(rt, trh, ["Tenant","Suite SF","Lease start (mo)","Lease expiry (mo)","Base rent ($/SF/yr)",
                 "Annual step (%)","Structure (override)","%rent: sales ($/SF/yr)","%rent: overage %"], start=1)
TR0 = trh + 1
for k,(nm,sf,ls,le,br,step,struct,sales,ov) in enumerate(suites):
    rr = TR0 + k
    put(rt, rr, 1, nm, role="input", fill=FILL_INPUT)
    put(rt, rr, 2, sf, role="input", nf=NF_NUM, fill=FILL_INPUT)
    put(rt, rr, 3, ls, role="input", nf=NF_NUM, fill=FILL_INPUT)
    put(rt, rr, 4, le, role="input", nf=NF_NUM, fill=FILL_INPUT)
    put(rt, rr, 5, br, role="input", nf=NF_PSF, fill=FILL_INPUT)
    put(rt, rr, 6, step, role="input", nf=NF_PCT, fill=FILL_INPUT)
    put(rt, rr, 7, struct, role="input", fill=FILL_INPUT)
    put(rt, rr, 8, sales, role="input", nf=NF_PSF, fill=FILL_INPUT)
    put(rt, rr, 9, ov, role="input", nf=NF_PCT, fill=FILL_INPUT)
TR_LAST = TR0 + len(suites) - 1
TRTOT = TR_LAST + 1
put(rt, TRTOT, 1, "TOTAL GLA", role="sub")
put(rt, TRTOT, 2, f"=SUM(B{TR0}:B{TR_LAST})", role="form", nf=NF_NUM, name="Retail_GLA")
add_name("Retail_SuiteCount", rt.title, f"$C${TRTOT}")
put(rt, TRTOT, 3, f"=COUNTA(A{TR0}:A{TR_LAST})", role="form", nf=NF_NUM)
put(rt, TRTOT, 9, "Set GLA rows to 0 for pure-MF.", role="unit")

# ---- Monthly retail build ----
mrow = TRTOT + 3
subhead(mrow := mrow, None) if False else None
subhead(rt, mrow, "MONTHLY RETAIL BUILD (-> Retail EGI). Per-tenant base rent + steps, then pool lines.", 1, 14); mrow += 1
put(rt, mrow, 1, "Month index", role="sub")
for i in range(MAX_MONTHS):
    put(rt, mrow, M0 + i, f"={q(ASM)}!{mcol(i)}${PERIOD_IDX_ROW}", role="link", nf=NF_NUM)
mrow += 1

# Per-tenant base+steps with rollover at expiry
put(rt, mrow, 1, "Per-tenant in-place + rollover base rent ($/mo)", role="sub"); mrow += 1
TEN_FIRST = mrow
for k,(nm,sf,ls,le,br,step,struct,sales,ov) in enumerate(suites):
    rr = TR0 + k
    put(rt, mrow, 1, "  " + nm, role="form")
    for i in range(MAX_MONTHS):
        col = mcol(i); idx = f"{q(ASM)}!{col}${PERIOD_IDX_ROW}"
        act = f"{q(ASM)}!{col}${OP_ACTIVE_ROW}"
        sf_ref = f"$B${rr}"; br_ref = f"$E${rr}"; step_ref = f"$F${rr}"
        ls_ref = f"$C${rr}"; le_ref = f"$D${rr}"
        # in-place: base rent grown by annual step, while within original term
        inplace = f"{br_ref}*{sf_ref}/12*(1+{step_ref})^INT(({idx}-{ls_ref})/12)"
        # rollover: after expiry, renew at market (=base grown by retail cum) * renewal factor, with downtime gap
        cf = cum_factor_ref("Cum_RetailRent", i)
        post = (f"IF(MOD({idx}-{le_ref}-1,12)<0,0,1)")  # placeholder; downtime handled via flag below
        gap = f"IF({idx}<={le_ref}+RetailDowntime,0,1)"
        renew = f"{br_ref}*{sf_ref}/12*{cf}*RenewalRentFactor"
        f = (f"=IF({act}=0,0,IF({idx}<{ls_ref},0,"
             f"IF({idx}<={le_ref},{inplace},{gap}*{renew})))")
        put(rt, mrow, M0 + i, f, role="form", nf=NF_USD0)
    mrow += 1
TEN_LAST = mrow - 1
put(rt, mrow, 1, "Total base + steps + rollover ($/mo)", role="form")
RT_BASE_ROW = mrow
for i in range(MAX_MONTHS):
    col = mcol(i)
    put(rt, mrow, M0 + i, f"=SUM({col}{TEN_FIRST}:{col}{TEN_LAST})", role="form", nf=NF_USD0)
mrow += 1

# Percentage rent (defaulted inert: sales=0)
put(rt, mrow, 1, "Percentage rent ($/mo)", role="form")
PCT_RENT_ROW = mrow
for i in range(MAX_MONTHS):
    col = mcol(i)
    parts = []
    for k,(nm,sf,ls,le,br,step,struct,sales,ov) in enumerate(suites):
        rr = TR0 + k
        sales_ref = f"$H${rr}"; sf_ref=f"$B${rr}"; ov_ref=f"$I${rr}"; br_ref=f"$E${rr}"
        bp = f"({br_ref}*{sf_ref}/{ov_ref})"   # natural breakpoint annual
        overage = f"MAX(0,{sales_ref}*{sf_ref}-{bp})*{ov_ref}/12"
        parts.append(f"IF({ov_ref}=0,0,{overage})")
    act = f"{q(ASM)}!{col}${OP_ACTIVE_ROW}"
    put(rt, mrow, M0 + i, f"=({'+'.join(parts)})*{act}", role="form", nf=NF_USD0)
mrow += 1

# Recoveries (pool * recoverable buckets per structure, +admin fee, no gross-up detail)
put(rt, mrow, 1, "Recoveries ($/mo)", role="form")
RECOV_ROW = mrow
for i in range(MAX_MONTHS):
    col = mcol(i); cf = cum_factor_ref("Cum_OpEx", i)
    act = f"{q(ASM)}!{col}${OP_ACTIVE_ROW}"
    parts = []
    for k,(nm,sf,ls,le,br,step,struct,sales,ov) in enumerate(suites):
        rr = TR0 + k
        sf_ref=f"$B${rr}"; struct_ref=f"$G${rr}"
        cam = f"VLOOKUP({struct_ref},RecovMatrix,2,FALSE)*RecovCAM"
        tax = f"VLOOKUP({struct_ref},RecovMatrix,3,FALSE)*RecovTax"
        ins = f"VLOOKUP({struct_ref},RecovMatrix,4,FALSE)*RecovIns"
        pool = f"({cam}+{tax}+{ins})*{sf_ref}/12*{cf}"
        adminmult = f"(1+IF(AdminFeeOn=\"On\",AdminFeePct,0))"
        parts.append(f"{pool}*{adminmult}")
    put(rt, mrow, M0 + i, f"=({'+'.join(parts)})*{act}", role="form", nf=NF_USD0)
mrow += 1

# General vacancy + credit loss (applied after rollover; structured to not double-count)
put(rt, mrow, 1, "General vacancy ($/mo, neg)", role="form")
RT_GENVAC_ROW = mrow
for i in range(MAX_MONTHS):
    col = mcol(i)
    income = f"({col}${RT_BASE_ROW}+{col}${PCT_RENT_ROW}+{col}${RECOV_ROW})"
    put(rt, mrow, M0 + i, f"=-{income}*RetailGenVacancy", role="form", nf=NF_USD0)
mrow += 1
put(rt, mrow, 1, "Credit loss / bad debt ($/mo, neg)", role="form")
RT_CREDIT_ROW = mrow
for i in range(MAX_MONTHS):
    col = mcol(i)
    income = f"({col}${RT_BASE_ROW}+{col}${PCT_RENT_ROW}+{col}${RECOV_ROW})"
    put(rt, mrow, M0 + i, f"=-{income}*(1-RetailGenVacancy)*RetailCreditLoss", role="form", nf=NF_USD0)
mrow += 1

# Retail EGI
put(rt, mrow, 1, "Retail EGI ($/mo)", role="sub")
RT_EGI_ROW = mrow
for i in range(MAX_MONTHS):
    col = mcol(i)
    put(rt, mrow, M0 + i,
        f"={col}${RT_BASE_ROW}+{col}${PCT_RENT_ROW}+{col}${RECOV_ROW}+{col}${RT_GENVAC_ROW}+{col}${RT_CREDIT_ROW}",
        role="form", nf=NF_USD0, bold=True)
mrow += 1

# Rollover TI/LC cost row (to Capital Budget) - cost at re-let month
put(rt, mrow, 1, "Rollover TI/LC ($, at re-let)", role="form")
RT_TILC_ROW = mrow
for i in range(MAX_MONTHS):
    col = mcol(i); idx = f"{q(ASM)}!{col}${PERIOD_IDX_ROW}"
    act = f"{q(ASM)}!{col}${OP_ACTIVE_ROW}"
    cf = cum_factor_ref("Cum_TILC", i)
    parts = []
    for k,(nm,sf,ls,le,br,step,struct,sales,ov) in enumerate(suites):
        rr = TR0 + k
        sf_ref=f"$B${rr}"; le_ref=f"$D${rr}"; br_ref=f"$E${rr}"
        relet = f"{le_ref}+RetailDowntime"
        # expected TI/LC at re-let: prob-weighted renewal vs new
        ti = (f"((1-RetailRenewalProb)*NewLeaseTI+RetailRenewalProb*RenewalTI)*{sf_ref}*{cf}")
        lc = (f"LeasingCommissionPct*{br_ref}*{sf_ref}*{cf}")
        parts.append(f"IF({idx}={relet},{ti}+{lc},0)")
    put(rt, mrow, M0 + i, f"=({'+'.join(parts)})*{act}", role="form", nf=NF_USD0)
mrow += 1
add_name("RT_EGI_Row", rt.title, f"${mcol(0)}${RT_EGI_ROW}")
add_name("RT_TILC_Row", rt.title, f"${mcol(0)}${RT_TILC_ROW}")

# ======================================================================================
# TAB 6 — OPEX
# ======================================================================================
ox = wb.create_sheet("OpEx")
banner(ox, "OPERATING EXPENSES", 14)
ox.column_dimensions["A"].width = 30
oxh = 5
colhdr(ox, oxh, ["Line item","$/unit/yr (base)"], start=1)
OX0 = oxh + 1
opex_items = MI.OPEX_ITEMS
for k,(nm,val) in enumerate(opex_items):
    rr = OX0 + k
    put(ox, rr, 1, nm, role="input", fill=FILL_INPUT)
    put(ox, rr, 2, val, role="input", nf=NF_USD0, fill=FILL_INPUT)
OX_LAST = OX0 + len(opex_items) - 1

mrow = OX_LAST + 2
subhead(ox, mrow, "MONTHLY OPEX BUILD (negative outflows)", 1, 14); mrow += 1
put(ox, mrow, 1, "Month index", role="sub")
for i in range(MAX_MONTHS):
    put(ox, mrow, M0 + i, f"={q(ASM)}!{mcol(i)}${PERIOD_IDX_ROW}", role="link", nf=NF_NUM)
mrow += 1

# Controllable opex (per-unit items grown by OpEx vector)
put(ox, mrow, 1, "Controllable OpEx ($/mo, neg)", role="form")
OX_CTRL_ROW = mrow
for i in range(MAX_MONTHS):
    col = mcol(i); cf = cum_factor_ref("Cum_OpEx", i)
    act = f"{q(ASM)}!{col}${OP_ACTIVE_ROW}"
    put(ox, mrow, M0 + i, f"=-SUM($B${OX0}:$B${OX_LAST})*MF_UnitCount/12*{cf}*{act}", role="form", nf=NF_USD0)
mrow += 1

# Management fee = % of total EGI
put(ox, mrow, 1, "Management fee ($/mo, neg)", role="form")
OX_MGMT_ROW = mrow
for i in range(MAX_MONTHS):
    col = mcol(i)
    egi = f"({q('MF Revenue')}!{col}${MF_EGI_ROW}+{q('Retail Revenue')}!{col}${RT_EGI_ROW})"
    put(ox, mrow, M0 + i, f"=-{egi}*MgmtFeePct", role="link", nf=NF_USD0)
mrow += 1

# Replacement reserves
put(ox, mrow, 1, "Replacement reserves ($/mo, neg)", role="form")
OX_RESV_ROW = mrow
for i in range(MAX_MONTHS):
    col = mcol(i)
    act = f"{q(ASM)}!{col}${OP_ACTIVE_ROW}"
    put(ox, mrow, M0 + i, f"=-ReplReservePU*MF_UnitCount/12*{act}", role="form", nf=NF_USD0)
mrow += 1

# Total OpEx (excl. taxes & reserves -> reserves shown separately in CF as capital)
put(ox, mrow, 1, "Total OpEx excl. taxes & reserves ($/mo, neg)", role="sub")
OX_TOT_ROW = mrow
for i in range(MAX_MONTHS):
    col = mcol(i)
    put(ox, mrow, M0 + i, f"={col}${OX_CTRL_ROW}+{col}${OX_MGMT_ROW}", role="form", nf=NF_USD0, bold=True)
mrow += 1
add_name("OX_Tot_Row", ox.title, f"${mcol(0)}${OX_TOT_ROW}")
add_name("OX_Resv_Row", ox.title, f"${mcol(0)}${OX_RESV_ROW}")

# ======================================================================================
# TAB 7 — TAXES
# ======================================================================================
tx = wb.create_sheet("Taxes")
banner(tx, "REAL ESTATE TAXES", 14)
tx.column_dimensions["A"].width = 34
mrow = 5
subhead(tx, mrow, "MONTHLY TAX BUILD (method per toggle #13)", 1, 14); mrow += 1
put(tx, mrow, 1, "Month index", role="sub")
for i in range(MAX_MONTHS):
    put(tx, mrow, M0 + i, f"={q(ASM)}!{mcol(i)}${PERIOD_IDX_ROW}", role="link", nf=NF_NUM)
mrow += 1
put(tx, mrow, 1, "Annual tax (method-driven $)", role="form")
TX_ANNUAL_ROW = mrow
for i in range(MAX_MONTHS):
    col = mcol(i)
    yr = f"{q(ASM)}!{col}${HOLD_YEAR_ROW}"
    flat = f"BaseAnnualTax*(1+TaxFlatGrowth)^(MAX({yr}-1,0))"
    reass = f"PurchasePrice*AssessRatio*MillRate*(1+TaxFlatGrowth)^(MAX({yr}-1,0))"
    pilot = f"PILOTAnnualTax"
    put(tx, mrow, M0 + i,
        f'=IF(TaxMethod="Flat Growth",{flat},IF(TaxMethod="PILOT / Abatement",{pilot},{reass}))',
        role="form", nf=NF_USD0)
mrow += 1
put(tx, mrow, 1, "Monthly tax ($/mo, neg)", role="sub")
TX_MONTHLY_ROW = mrow
for i in range(MAX_MONTHS):
    col = mcol(i)
    act = f"{q(ASM)}!{col}${OP_ACTIVE_ROW}"
    put(tx, mrow, M0 + i, f"=-{col}${TX_ANNUAL_ROW}/12*{act}", role="form", nf=NF_USD0, bold=True)
mrow += 1
add_name("TX_Monthly_Row", tx.title, f"${mcol(0)}${TX_MONTHLY_ROW}")

# ======================================================================================
# TAB 8 — CAPEX / CAPITAL BUDGET (Sources & Uses + draws)
# ======================================================================================
cx = wb.create_sheet("CapEx")
banner(cx, "CAPEX / CAPITAL BUDGET", 14)
cx.column_dimensions["A"].width = 34
for col in "BCD":
    cx.column_dimensions[col].width = 16

r = 5
subhead(cx, r, "Capital budget buckets", 1, 6); r += 1
colhdr(cx, r, ["Bucket","Amount ($)","Funding","Spend start (mo)","Spend months"], start=1)
CB0 = r + 1
buckets = [
    ("Immediate/deferred repairs", 1_500_000, "=Fund_Repairs", 1, 6),
    ("Value-add unit renovation", None, "=Fund_Reno", 1, 24),   # auto-linked below
    ("Retail TI/LC reserve", 1_200_000, "=Fund_TILC", 1, 36),
    ("Contingency", 600_000, "=Fund_Contingency", 1, 12),
]
for k,(nm,amt,fund,start,months) in enumerate(buckets):
    rr = CB0 + k
    put(cx, rr, 1, nm, role="input", fill=FILL_INPUT)
    if amt is None:
        put(cx, rr, 2, "=RenoCostUnit*MF_UnitCount", role="form", nf=NF_USD0)
    else:
        put(cx, rr, 2, amt, role="input", nf=NF_USD0, fill=FILL_INPUT)
    put(cx, rr, 3, fund, role="form")
    put(cx, rr, 4, start, role="input", nf=NF_NUM, fill=FILL_INPUT)
    put(cx, rr, 5, months, role="input", nf=NF_NUM, fill=FILL_INPUT)
CB_LAST = CB0 + len(buckets) - 1
CBTOT = CB_LAST + 1
put(cx, CBTOT, 1, "Total capital budget", role="sub")
put(cx, CBTOT, 2, f"=SUM(B{CB0}:B{CB_LAST})", role="form", nf=NF_USD0, name="TotalCapBudget")
put(cx, CBTOT+1, 1, "Equity-funded capital", role="form")
put(cx, CBTOT+1, 2, f'=SUMIF(C{CB0}:C{CB_LAST},"Equity",B{CB0}:B{CB_LAST})', role="form", nf=NF_USD0, name="EquityCapBudget")
put(cx, CBTOT+2, 1, "Financed capital (loan holdback)", role="form")
put(cx, CBTOT+2, 2, f'=SUMIF(C{CB0}:C{CB_LAST},"Financed",B{CB0}:B{CB_LAST})', role="form", nf=NF_USD0, name="FinancedCapBudget")

# Monthly capital spend (straight-line over spend window) + TI/LC from retail rollover
mrow = CBTOT + 4
subhead(cx, mrow, "MONTHLY CAPITAL SPEND ($/mo, neg)", 1, 14); mrow += 1
put(cx, mrow, 1, "Month index", role="sub")
for i in range(MAX_MONTHS):
    put(cx, mrow, M0 + i, f"={q(ASM)}!{mcol(i)}${PERIOD_IDX_ROW}", role="link", nf=NF_NUM)
mrow += 1
put(cx, mrow, 1, "Capital budget spend ($/mo, neg)", role="form")
CX_SPEND_ROW = mrow
for i in range(MAX_MONTHS):
    col = mcol(i); idx = f"{q(ASM)}!{col}${PERIOD_IDX_ROW}"
    parts = []
    for k in range(len(buckets)):
        rr = CB0 + k
        parts.append(f"IF(AND({idx}>=$D${rr},{idx}<$D${rr}+$E${rr}),$B${rr}/$E${rr},0)")
    put(cx, mrow, M0 + i, f"=-({'+'.join(parts)})", role="form", nf=NF_USD0)
mrow += 1
put(cx, mrow, 1, "Retail rollover TI/LC ($/mo, neg)", role="form")
CX_TILC_ROW = mrow
for i in range(MAX_MONTHS):
    col = mcol(i)
    put(cx, mrow, M0 + i, f"=-{q('Retail Revenue')}!{col}${RT_TILC_ROW}", role="link", nf=NF_USD0)
mrow += 1
put(cx, mrow, 1, "Total capital outflow ($/mo, neg)", role="sub")
CX_TOT_ROW = mrow
for i in range(MAX_MONTHS):
    col = mcol(i)
    put(cx, mrow, M0 + i, f"={col}${CX_SPEND_ROW}+{col}${CX_TILC_ROW}", role="form", nf=NF_USD0, bold=True)
mrow += 1
add_name("CX_Tot_Row", cx.title, f"${mcol(0)}${CX_TOT_ROW}")

# ======================================================================================
# TAB 9 — DEBT (acquisition sizing + min-of + refi)  -- needs NOI; place NOI on Monthly CF
# We'll compute stabilized NOI proxy here referencing Monthly CF NOI row (built next),
# so define Debt after Monthly CF? To avoid forward refs, we compute a Year-1 NOI here
# directly from component EGI - opex - taxes.
# ======================================================================================
db = wb.create_sheet("Debt")
banner(db, "DEBT — Acquisition Sizing, Min-of Constraints, Refi", 14)
db.column_dimensions["A"].width = 36
for col in "BCD":
    db.column_dimensions[col].width = 16

r = 5
subhead(db, r, "Loan terms (inputs)", 1, 4); r += 1
put(db, r, 1, "Max LTV (%)", role="form"); put(db, r, 2, 0.65, role="input", nf=NF_PCT, fill=FILL_INPUT, name="MaxLTV"); r += 1
put(db, r, 1, "Max LTC (%)", role="form"); put(db, r, 2, 0.70, role="input", nf=NF_PCT, fill=FILL_INPUT, name="MaxLTC"); r += 1
put(db, r, 1, "Min DSCR (x)", role="form"); put(db, r, 2, 1.25, role="input", nf=NF_MULT, fill=FILL_INPUT, name="MinDSCR"); r += 1
put(db, r, 1, "Min Debt Yield (%)", role="form"); put(db, r, 2, 0.085, role="input", nf=NF_PCT, fill=FILL_INPUT, name="MinDebtYield"); r += 1
put(db, r, 1, "Fixed rate (%)", role="form"); put(db, r, 2, D["fixed_rate"], role="input", nf=NF_PCT2, fill=FILL_INPUT, name="FixedRate"); r += 1
put(db, r, 1, "SOFR (%)", role="form"); put(db, r, 2, 0.043, role="input", nf=NF_PCT2, fill=FILL_INPUT, name="SOFR"); r += 1
put(db, r, 1, "Spread (bps)", role="form"); put(db, r, 2, 250, role="input", nf=NF_BPS, fill=FILL_INPUT, name="SpreadBps"); r += 1
put(db, r, 1, "All-in rate (%)", role="form")
put(db, r, 2, '=IF(LoanRateType="Fixed",FixedRate,SOFR+SpreadBps/10000)', role="form", nf=NF_PCT2, name="LoanRate"); r += 1
put(db, r, 1, "IO period (months)", role="form"); put(db, r, 2, 36, role="input", nf=NF_MO, fill=FILL_INPUT, name="IOMonths"); r += 1
put(db, r, 1, "Amortization (years)", role="form"); put(db, r, 2, 30, role="input", nf=NF_NUM, fill=FILL_INPUT, name="AmortYears"); r += 1
put(db, r, 1, "Financing cost (% of loan)", role="form"); put(db, r, 2, 0.01, role="input", nf=NF_PCT, fill=FILL_INPUT, name="FinanceCostPct"); r += 2

# Stabilized NOI proxy (forward 12mo near stabilization) for sizing
subhead(db, r, "Sizing inputs (NOI & value)", 1, 4); r += 1
put(db, r, 1, "Stabilized annual NOI ($) [from CF]", role="form")
# reference Monthly CF NOI row (defined later); we will set after building CF. Use SUM of months 13-24 as proxy.
DB_STAB_NOI = r
put(db, r, 2, 0, role="form", nf=NF_USD0, name="StabNOI"); r += 1   # filled after CF built
put(db, r, 1, "Going-in annual NOI ($) [yr1 from CF]", role="form")
DB_GOINGIN_NOI = r
put(db, r, 2, 0, role="form", nf=NF_USD0, name="GoingInNOI"); r += 1
put(db, r, 1, "Annual debt constant (amortizing)", role="form")
put(db, r, 2, "=LoanRate/12/(1-(1+LoanRate/12)^(-AmortYears*12))*12", role="form", nf=NF_PCT2, name="DebtConstant"); r += 2

subhead(db, r, "Max loan under each active constraint (min-of binds)", 1, 6); r += 1
put(db, r, 1, "Max loan — LTV", role="form")
put(db, r, 2, '=IF(Use_LTV="On",PurchasePrice*MaxLTV,1E15)', role="form", nf=NF_USD0, name="Loan_LTV"); r += 1
put(db, r, 1, "Max loan — LTC", role="form")
put(db, r, 2, '=IF(Use_LTC="On",TotalCost*MaxLTC,1E15)', role="form", nf=NF_USD0, name="Loan_LTC"); r += 1
put(db, r, 1, "Max loan — DSCR", role="form")
put(db, r, 2, '=IF(Use_DSCR="On",GoingInNOI/(MinDSCR*DebtConstant),1E15)', role="form", nf=NF_USD0, name="Loan_DSCR"); r += 1
put(db, r, 1, "Max loan — Debt Yield", role="form")
put(db, r, 2, '=IF(Use_DY="On",GoingInNOI/MinDebtYield,1E15)', role="form", nf=NF_USD0, name="Loan_DY"); r += 1
put(db, r, 1, "Loan amount (MIN of active)", role="sub")
put(db, r, 2, '=MIN(Loan_LTV,Loan_LTC,Loan_DSCR,Loan_DY)+FinancedCapBudget', role="form", nf=NF_USD0, bold=True, name="LoanAmount"); r += 1
put(db, r, 1, "Binding constraint", role="form")
put(db, r, 2, '=IF(LoanAmount-FinancedCapBudget=Loan_LTV,"LTV",IF(LoanAmount-FinancedCapBudget=Loan_LTC,"LTC",IF(LoanAmount-FinancedCapBudget=Loan_DSCR,"DSCR","Debt Yield")))',
    role="form", bold=True, name="BindingConstraint"); r += 1
put(db, r, 1, "Financing costs ($)", role="form")
put(db, r, 2, "=LoanAmount*FinanceCostPct", role="form", nf=NF_USD0, name="FinancingCosts"); r += 2

# Refi block
subhead(db, r, "Refi (toggle #5)", 1, 4); r += 1
put(db, r, 1, "Refi month", role="form"); put(db, r, 2, 36, role="input", nf=NF_MO, fill=FILL_INPUT, name="RefiMonth"); r += 1
put(db, r, 1, "Refi cap on stabilized value (%)", role="form"); put(db, r, 2, 0.055, role="input", nf=NF_PCT2, fill=FILL_INPUT, name="RefiCap"); r += 1
put(db, r, 1, "Refi LTV (%)", role="form"); put(db, r, 2, 0.65, role="input", nf=NF_PCT, fill=FILL_INPUT, name="RefiLTV"); r += 1
put(db, r, 1, "Refi cost (% of new loan)", role="form"); put(db, r, 2, 0.01, role="input", nf=NF_PCT, fill=FILL_INPUT, name="RefiCostPct"); r += 1
put(db, r, 1, "Stabilized value (for refi)", role="form")
put(db, r, 2, "=StabNOI/RefiCap", role="form", nf=NF_USD0, name="StabValue"); r += 1
put(db, r, 1, "New refi loan", role="form")
put(db, r, 2, '=IF(RefiOn="On",MIN(StabValue*RefiLTV,StabNOI/(MinDSCR*DebtConstant)),0)', role="form", nf=NF_USD0, name="RefiLoan"); r += 1

# Monthly debt schedule
mrow = r + 2
subhead(db, mrow, "MONTHLY DEBT SCHEDULE", 1, 14); mrow += 1
put(db, mrow, 1, "Month index", role="sub")
for i in range(MAX_MONTHS):
    put(db, mrow, M0 + i, f"={q(ASM)}!{mcol(i)}${PERIOD_IDX_ROW}", role="link", nf=NF_NUM)
mrow += 1
# Opening balance
put(db, mrow, 1, "Loan balance — opening ($)", role="form")
DB_BAL_OPEN = mrow
# Closing balance
DB_BAL_CLOSE = mrow + 1
DB_INT = mrow + 2
DB_PRIN = mrow + 3
DB_DS = mrow + 4
DB_DRAW = mrow + 5
DB_EVENT = mrow + 6
for i in range(MAX_MONTHS):
    col = mcol(i); idx = f"{col}${mrow-1}"  # month index row above? actually period idx row local
# Recompute with explicit references
for i in range(MAX_MONTHS):
    col = mcol(i)
    prev = mcol(i-1) if i>0 else None
    idxref = f"{q(ASM)}!{col}${PERIOD_IDX_ROW}"
    act = f"{q(ASM)}!{col}${OP_ACTIVE_ROW}"
    # opening balance
    if i == 0:
        put(db, DB_BAL_OPEN, M0 + i, "=0", role="form", nf=NF_USD0)
    else:
        put(db, DB_BAL_OPEN, M0 + i, f"={prev}${DB_BAL_CLOSE}", role="form", nf=NF_USD0)
    # draw at month 0 (initial loan, aligns with acquisition) and refi handling thereafter
    if i == 0:
        draw = "=LoanAmount"
    else:
        draw = f'=IF(AND(RefiOn="On",{idxref}=RefiMonth),RefiLoan,0)'
    put(db, DB_DRAW, M0 + i, draw, role="form", nf=NF_USD0)
    # payoff old at refi (event = negative principal reduction handled in close)
    # interest on opening balance
    put(db, DB_INT, M0 + i, f"={col}${DB_BAL_OPEN}*LoanRate/12*{act}", role="form", nf=NF_USD0)
    # principal: amortize after IO, only when loan active and not refi-payoff month
    io_done = f"{idxref}>IOMonths"
    pmt = (f"PMT(LoanRate/12,AmortYears*12-(IOMonths),-{col}${DB_BAL_OPEN})")
    put(db, DB_PRIN, M0 + i,
        f"=IF(AND({act}=1,{io_done},{col}${DB_BAL_OPEN}>0),MIN({col}${DB_BAL_OPEN},{pmt}-{col}${DB_INT}),0)",
        role="form", nf=NF_USD0)
    # debt service
    put(db, DB_DS, M0 + i, f"={col}${DB_INT}+{col}${DB_PRIN}", role="form", nf=NF_USD0)
    # refi payoff event (pay off opening balance at refi month)
    put(db, DB_EVENT, M0 + i,
        f'=IF(AND(RefiOn="On",{idxref}=RefiMonth),-{col}${DB_BAL_OPEN},0)', role="form", nf=NF_USD0)
    # closing balance = opening + draw - principal + refi payoff
    put(db, DB_BAL_CLOSE, M0 + i,
        f"={col}${DB_BAL_OPEN}+{col}${DB_DRAW}-{col}${DB_PRIN}+{col}${DB_EVENT}", role="form", nf=NF_USD0)
put(db, DB_BAL_OPEN, 1, "Loan balance — opening ($)", role="form")
put(db, DB_BAL_CLOSE, 1, "Loan balance — closing ($)", role="form")
put(db, DB_INT, 1, "Interest ($/mo)", role="form")
put(db, DB_PRIN, 1, "Principal ($/mo)", role="form")
put(db, DB_DS, 1, "Debt service ($/mo)", role="form")
put(db, DB_DRAW, 1, "Loan draw ($)", role="form")
put(db, DB_EVENT, 1, "Refi payoff / event ($)", role="form")
add_name("DB_DS_Row", db.title, f"${mcol(0)}${DB_DS}")
add_name("DB_Draw_Row", db.title, f"${mcol(0)}${DB_DRAW}")
add_name("DB_BalClose_Row", db.title, f"${mcol(0)}${DB_BAL_CLOSE}")
add_name("DB_Event_Row", db.title, f"${mcol(0)}${DB_EVENT}")

# Refi cash-out (distribution event) = new loan - old payoff - refi cost, at refi month
mrow2 = DB_EVENT + 2
put(db, mrow2, 1, "Refi cash-out (to waterfall) ($)", role="form")
DB_CASHOUT = mrow2
for i in range(MAX_MONTHS):
    col = mcol(i); idxref = f"{q(ASM)}!{col}${PERIOD_IDX_ROW}"
    put(db, mrow2, M0 + i,
        f'=IF(AND(RefiOn="On",{idxref}=RefiMonth),RefiLoan-{col}${DB_BAL_OPEN}-RefiLoan*RefiCostPct,0)',
        role="form", nf=NF_USD0)
add_name("DB_CashOut_Row", db.title, f"${mcol(0)}${DB_CASHOUT}")

# ======================================================================================
# TAB 10 — MONTHLY CASH FLOW + ANNUAL ROLL-UP
# ======================================================================================
mc = wb.create_sheet("Monthly CF")
banner(mc, "MONTHLY CASH FLOW + ANNUAL ROLL-UP", 14)
mc.column_dimensions["A"].width = 34

mrow = 5
put(mc, mrow, 1, "Month index", role="sub")
for i in range(MAX_MONTHS):
    put(mc, mrow, M0 + i, f"={q(ASM)}!{mcol(i)}${PERIOD_IDX_ROW}", role="link", nf=NF_NUM)
mrow += 1
put(mc, mrow, 1, "Period date", role="sub")
MC_DATE_ROW = mrow
for i in range(MAX_MONTHS):
    put(mc, mrow, M0 + i, f"={q(ASM)}!{mcol(i)}${PERIOD_DATE_ROW}", role="link", nf=NF_DATE)
mrow += 1

def cf_line(label, role_, builder, nf=NF_USD0, bold=False):
    global mrow
    put(mc, mrow, 1, label, role="sub" if bold else "form")
    rownum = mrow
    for i in range(MAX_MONTHS):
        col = mcol(i)
        put(mc, mrow, M0 + i, builder(col, i), role=role_, nf=nf, bold=bold)
    mrow += 1
    return rownum

MC_MFEGI = cf_line("MF EGI", "link", lambda col,i: f"={q('MF Revenue')}!{col}${MF_EGI_ROW}")
MC_RTEGI = cf_line("Retail EGI", "link", lambda col,i: f"={q('Retail Revenue')}!{col}${RT_EGI_ROW}")
MC_EGI = cf_line("Total EGI", "form", lambda col,i: f"={col}${MC_MFEGI}+{col}${MC_RTEGI}", bold=True)
MC_OPEX = cf_line("OpEx (excl. taxes)", "link", lambda col,i: f"={q('OpEx')}!{col}${OX_TOT_ROW}")
MC_TAX = cf_line("Taxes", "link", lambda col,i: f"={q('Taxes')}!{col}${TX_MONTHLY_ROW}")
MC_NOI = cf_line("NOI", "form", lambda col,i: f"={col}${MC_EGI}+{col}${MC_OPEX}+{col}${MC_TAX}", bold=True)
MC_RESV = cf_line("Replacement reserves", "link", lambda col,i: f"={q('OpEx')}!{col}${OX_RESV_ROW}")
MC_CAP = cf_line("CapEx / TI/LC", "link", lambda col,i: f"={q('CapEx')}!{col}${CX_TOT_ROW}")
MC_UCF = cf_line("Unlevered cash flow (ops)", "form",
                 lambda col,i: f"={col}${MC_NOI}+{col}${MC_RESV}+{col}${MC_CAP}", bold=True)
MC_DS = cf_line("Debt service", "link", lambda col,i: f"=-{q('Debt')}!{col}${DB_DS}")
MC_LOANEV = cf_line("Loan draws / refi events", "link",
                    lambda col,i: f"={q('Debt')}!{col}${DB_DRAW}+{q('Debt')}!{col}${DB_EVENT}")

# Acquisition (month 0) and disposition (exit month) lines
put(mc, mrow, 1, "Acquisition outflow (equity & price, mo 0)", role="form")
MC_ACQ = mrow
for i in range(MAX_MONTHS):
    col = mcol(i)
    if i == 0:
        # Acquisition basis + financing costs out. Loan proceeds (incl. financed holdback)
        # come in via MC_LOANEV (DB draw at m0). All capital flows through CapEx over time,
        # so capital is NOT lumped here (avoids double-counting).
        put(mc, mrow, M0 + i,
            "=-(PurchasePrice+PurchasePrice*ClosingCostPct+FinancingCosts)",
            role="form", nf=NF_USD0)
    else:
        put(mc, mrow, M0 + i, "=0", role="form", nf=NF_USD0)
mrow += 1

# Disposition (gross value - costs - loan payoff) at exit month
put(mc, mrow, 1, "Exit gross value ($)", role="form")
MC_EXITVAL = mrow
for i in range(MAX_MONTHS):
    col = mcol(i)
    exitflag = f"{q(ASM)}!{col}${EXIT_FLAG_ROW}"
    put(mc, mrow, M0 + i, f"=IF({exitflag}=1,ExitGrossValue,0)", role="form", nf=NF_USD0)
mrow += 1
put(mc, mrow, 1, "Disposition costs ($)", role="form")
MC_DISPCOST = mrow
for i in range(MAX_MONTHS):
    col = mcol(i); exitflag = f"{q(ASM)}!{col}${EXIT_FLAG_ROW}"
    put(mc, mrow, M0 + i, f"=IF({exitflag}=1,-ExitGrossValue*DispCostPct,0)", role="form", nf=NF_USD0)
mrow += 1
put(mc, mrow, 1, "Loan payoff at exit ($)", role="form")
MC_PAYOFF = mrow
for i in range(MAX_MONTHS):
    col = mcol(i); exitflag = f"{q(ASM)}!{col}${EXIT_FLAG_ROW}"
    put(mc, mrow, M0 + i, f"=IF({exitflag}=1,-{q('Debt')}!{col}${DB_BAL_CLOSE},0)", role="form", nf=NF_USD0)
mrow += 1

# Levered cash flow
put(mc, mrow, 1, "Levered cash flow ($/mo)", role="sub")
MC_LCF = mrow
for i in range(MAX_MONTHS):
    col = mcol(i)
    put(mc, mrow, M0 + i,
        f"={col}${MC_UCF}+{col}${MC_DS}+{col}${MC_LOANEV}+{col}${MC_ACQ}+{col}${MC_EXITVAL}+{col}${MC_DISPCOST}+{col}${MC_PAYOFF}",
        role="form", nf=NF_USD0, bold=True)
mrow += 1
# Unlevered total (with acq/disposition, no debt)
put(mc, mrow, 1, "Unlevered total cash flow ($/mo)", role="sub")
MC_UTOT = mrow
for i in range(MAX_MONTHS):
    col = mcol(i)
    # unlevered: ops (incl. capex over time) + acq(price+closing only) + exit(no payoff, no debt)
    if i == 0:
        put(mc, mrow, M0 + i,
            f"={col}${MC_UCF}-(PurchasePrice+PurchasePrice*ClosingCostPct)+{col}${MC_EXITVAL}+{col}${MC_DISPCOST}",
            role="form", nf=NF_USD0, bold=True)
    else:
        put(mc, mrow, M0 + i,
            f"={col}${MC_UCF}+{col}${MC_EXITVAL}+{col}${MC_DISPCOST}",
            role="form", nf=NF_USD0, bold=True)
mrow += 1
add_name("MC_LCF_Row", mc.title, f"${mcol(0)}${MC_LCF}")
add_name("MC_UCF_Row", mc.title, f"${mcol(0)}${MC_UTOT}")
add_name("MC_NOI_Row", mc.title, f"${mcol(0)}${MC_NOI}")
add_name("MC_Date_Row", mc.title, f"${mcol(0)}${MC_DATE_ROW}")

# Now wire Debt StabNOI / GoingInNOI to CF NOI (sum windows)
# Going-in = sum NOI months 1..12 ; Stabilized = sum NOI months 13..24 (post-ramp proxy)
goingin_formula = f"=SUM({mcol(1)}${MC_NOI}:{mcol(12)}${MC_NOI})"
stab_formula = f"=SUM({mcol(13)}${MC_NOI}:{mcol(24)}${MC_NOI})"
put(db, DB_GOINGIN_NOI, 2, f"={q('Monthly CF')}!{mcol(1)}${MC_NOI}*0+"+f"SUM({q('Monthly CF')}!{mcol(1)}${MC_NOI}:{q('Monthly CF')}!{mcol(12)}${MC_NOI})",
    role="link", nf=NF_USD0)
put(db, DB_STAB_NOI, 2, f"=SUM({q('Monthly CF')}!{mcol(13)}${MC_NOI}:{q('Monthly CF')}!{mcol(24)}${MC_NOI})",
    role="link", nf=NF_USD0)

# ---- Annual roll-up ----
ar = mrow + 2
subhead(mc, ar, "ANNUAL ROLL-UP (hold-year basis)", 1, 14); ar += 1
put(mc, ar, 1, "Hold year", role="sub")
AR_HDR = ar
MAXYEARS = MAX_MONTHS // 12
for y in range(1, MAXYEARS + 1):
    put(mc, ar, 4 + y, y, role="sub", align="center")
ar += 1
def annual_row(label, monthly_row, name=None):
    global ar
    put(mc, ar, 1, label, role="form")
    for y in range(1, MAXYEARS + 1):
        c0 = mcol((y-1)*12 + 1); c1 = mcol(y*12)
        put(mc, ar, 4 + y, f"=SUMIF({q(ASM)}!{c0}${HOLD_YEAR_ROW}:{q(ASM)}!{c1}${HOLD_YEAR_ROW},{4+y if False else y},{c0}${monthly_row}:{c1}${monthly_row})",
            role="form", nf=NF_USD0)
    if name:
        add_name(name, mc.title, f"${get_column_letter(5)}${ar}")
    ar += 1
    return ar - 1

AR_NOI = annual_row("NOI", MC_NOI)
AR_UCF = annual_row("Unlevered CF", MC_UTOT)
AR_LCF = annual_row("Levered CF", MC_LCF)

# ======================================================================================
# Exit valuation & total cost named cells (on Assumptions, reference CF NOI windows)
# ======================================================================================
# Exit NOI basis: forward or trailing 12 months around exit. Use stabilized proxy windows.
# Forward 12mo NOI at exit = sum of 12 months ending at exit (trailing) or starting at exit (forward).
# We approximate using last 12 active months (trailing) and StabNOI grown (forward).
r2 = asm.max_row + 2
subhead(asm, r2, "Derived exit & cost metrics", 1, 6); r2 += 1
put(asm, r2, 1, "Trailing 12-mo exit NOI ($)", role="form")
# trailing: sum of NOI over the 12 months up to and including exit -> use SUMPRODUCT with exit window
# Simpler robust proxy: trailing = stabilized annualized * cum growth to exit year.
put(asm, r2, 2, f"=SUMPRODUCT({q('Monthly CF')}!{mcol(0)}${MC_NOI}:{mcol(MAX_MONTHS-1)}${MC_NOI},"
                f"{q(ASM)}!{mcol(0)}${EXIT_FLAG_ROW}:{mcol(MAX_MONTHS-1)}${EXIT_FLAG_ROW})",
    role="link", nf=NF_USD0)
# above gives only exit-month NOI; build proper trailing using offset sum:
TRAIL_ROW = r2
put(asm, r2, 2,
    f"=SUMPRODUCT(({q('Monthly CF')}!{mcol(0)}${MC_NOI}:{q('Monthly CF')}!{mcol(MAX_MONTHS-1)}${MC_NOI}),"
    f"--({q(ASM)}!{mcol(0)}${PERIOD_IDX_ROW}:{q(ASM)}!{mcol(MAX_MONTHS-1)}${PERIOD_IDX_ROW}>HoldMonths-12),"
    f"--({q(ASM)}!{mcol(0)}${PERIOD_IDX_ROW}:{q(ASM)}!{mcol(MAX_MONTHS-1)}${PERIOD_IDX_ROW}<=HoldMonths))",
    role="link", nf=NF_USD0, name="ExitNOI_Trailing"); r2 += 1
put(asm, r2, 1, "Forward 12-mo exit NOI ($)", role="form")
put(asm, r2, 2, "=ExitNOI_Trailing*(1+INDEX(Vec_MFRent,1,1))", role="form", nf=NF_USD0, name="ExitNOI_Forward"); r2 += 1
put(asm, r2, 1, "Exit NOI (basis-selected) ($)", role="form")
put(asm, r2, 2, '=IF(ExitNOIBasis="Trailing 12-mo",ExitNOI_Trailing,ExitNOI_Forward)', role="form", nf=NF_USD0, name="ExitNOI"); r2 += 1
# MF/Retail split for component caps (approx via EGI share applied to NOI)
put(asm, r2, 1, "MF NOI share (exit, approx)", role="form")
put(asm, r2, 2,
    f"=SUMPRODUCT({q('Monthly CF')}!{mcol(0)}${MC_MFEGI}:{q('Monthly CF')}!{mcol(MAX_MONTHS-1)}${MC_MFEGI},"
    f"--({q(ASM)}!{mcol(0)}${PERIOD_IDX_ROW}:{q(ASM)}!{mcol(MAX_MONTHS-1)}${PERIOD_IDX_ROW}>HoldMonths-12),"
    f"--({q(ASM)}!{mcol(0)}${PERIOD_IDX_ROW}:{q(ASM)}!{mcol(MAX_MONTHS-1)}${PERIOD_IDX_ROW}<=HoldMonths))/"
    f"MAX(1,SUMPRODUCT(({q('Monthly CF')}!{mcol(0)}${MC_MFEGI}:{q('Monthly CF')}!{mcol(MAX_MONTHS-1)}${MC_MFEGI}+{q('Monthly CF')}!{mcol(0)}${MC_RTEGI}:{q('Monthly CF')}!{mcol(MAX_MONTHS-1)}${MC_RTEGI}),"
    f"--({q(ASM)}!{mcol(0)}${PERIOD_IDX_ROW}:{q(ASM)}!{mcol(MAX_MONTHS-1)}${PERIOD_IDX_ROW}>HoldMonths-12),"
    f"--({q(ASM)}!{mcol(0)}${PERIOD_IDX_ROW}:{q(ASM)}!{mcol(MAX_MONTHS-1)}${PERIOD_IDX_ROW}<=HoldMonths)))",
    role="link", nf=NF_PCT, name="MF_NOI_Share"); r2 += 1
put(asm, r2, 1, "Exit gross value ($)", role="form")
put(asm, r2, 2,
    '=IF(ExitValMethod="Component Caps (MF + Retail summed)",'
    'ExitNOI*MF_NOI_Share/ExitCapMF+ExitNOI*(1-MF_NOI_Share)/ExitCapRetail,'
    'ExitNOI/ExitCapBlended)',
    role="form", nf=NF_USD0, name="ExitGrossValue"); r2 += 1
put(asm, r2, 1, "Total cost ($)", role="form")
put(asm, r2, 2, "=PurchasePrice+PurchasePrice*ClosingCostPct+TotalCapBudget+FinancingCosts",
    role="form", nf=NF_USD0, name="TotalCost"); r2 += 1
put(asm, r2, 1, "Yield-on-cost (stabilized) (%)", role="form")
put(asm, r2, 2, "=StabNOI/TotalCost", role="form", nf=NF_PCT2, name="YieldOnCost"); r2 += 1
put(asm, r2, 1, "Development spread (YoC - exit cap) (bps)", role="form")
put(asm, r2, 2, "=(YieldOnCost-ExitCapBlended)*10000", role="form", nf=NF_BPS, name="DevSpread"); r2 += 1
put(asm, r2, 1, "Going-in cap (%)", role="form")
put(asm, r2, 2, "=GoingInNOI/PurchasePrice", role="form", nf=NF_PCT2, name="GoingInCap"); r2 += 1
put(asm, r2, 1, "Stabilized cap (%)", role="form")
put(asm, r2, 2, "=StabNOI/PurchasePrice", role="form", nf=NF_PCT2, name="StabCap"); r2 += 1
put(asm, r2, 1, "Initial equity ($) [= total cost - loan]", role="form")
put(asm, r2, 2, "=TotalCost-LoanAmount",
    role="form", nf=NF_USD0, name="InitialEquity"); r2 += 1

# ======================================================================================
# TAB 11 — WATERFALL (American; gated by LP_Present)
# ======================================================================================
wf = wb.create_sheet("Waterfall")
banner(wf, "EQUITY WATERFALL — American (gated by LP Present)", 14)
wf.column_dimensions["A"].width = 36
for col in "BCDE":
    wf.column_dimensions[col].width = 16

r = 5
subhead(wf, r, "Equity split & promote inputs", 1, 4); r += 1
put(wf, r, 1, "LP capital split (%)", role="form"); put(wf, r, 2, 0.90, role="input", nf=NF_PCT, fill=FILL_INPUT, name="LP_Split"); r += 1
put(wf, r, 1, "GP capital split (%)", role="form"); put(wf, r, 2, "=1-LP_Split", role="form", nf=NF_PCT, name="GP_Split"); r += 1
put(wf, r, 1, "GP co-invest (% of GP share)", role="form"); put(wf, r, 2, 1.0, role="input", nf=NF_PCT, fill=FILL_INPUT, name="GP_CoInvest"); r += 1
put(wf, r, 1, "Preferred return (%, compounded)", role="form"); put(wf, r, 2, 0.08, role="input", nf=NF_PCT, fill=FILL_INPUT, name="PrefRate"); r += 1
put(wf, r, 1, "GP catch-up (%)", role="form"); put(wf, r, 2, 0.50, role="input", nf=NF_PCT, fill=FILL_INPUT, name="CatchUpPct"); r += 1
put(wf, r, 1, "ROC vs Pref order", role="form"); put(wf, r, 2, "Sequential", role="input", fill=FILL_TOG, name="ROCPrefOrder"); r += 2

subhead(wf, r, "Promote / carry tiers (IRR hurdles -> LP/GP split)", 1, 6); r += 1
colhdr(wf, r, ["Tier hurdle IRR (upper)","LP %","GP %"], start=1)
WT0 = r + 1
tiers = [(0.08,0.90,0.10),(0.12,0.80,0.20),(0.15,0.70,0.30),(9.99,0.60,0.40)]
for k,(h,lp,gp) in enumerate(tiers):
    rr = WT0 + k
    put(wf, rr, 1, h, role="input", nf=NF_PCT, fill=FILL_INPUT)
    put(wf, rr, 2, lp, role="input", nf=NF_PCT, fill=FILL_INPUT)
    put(wf, rr, 3, gp, role="input", nf=NF_PCT, fill=FILL_INPUT)
WT_LAST = WT0 + len(tiers) - 1
add_name("CarryTiers", wf.title, f"$A${WT0}:$C${WT_LAST}")

r = WT_LAST + 2
subhead(wf, r, "Fees (#15)", 1, 4); r += 1
put(wf, r, 1, "Acquisition fee (% of price)", role="form"); put(wf, r, 2, 0.01, role="input", nf=NF_PCT, fill=FILL_INPUT, name="AcqFeePct"); r += 1
put(wf, r, 1, "AM fee rate (%)", role="form"); put(wf, r, 2, 0.0125, role="input", nf=NF_PCT, fill=FILL_INPUT, name="AMFeeRate"); r += 1
put(wf, r, 1, "Disposition fee (% of sale)", role="form"); put(wf, r, 2, 0.01, role="input", nf=NF_PCT, fill=FILL_INPUT, name="DispFeePct"); r += 1
put(wf, r, 1, "Acquisition fee ($)", role="form")
put(wf, r, 2, '=IF(AcqFeeOn="On",PurchasePrice*AcqFeePct,0)', role="form", nf=NF_USD0, name="AcqFee"); r += 1
put(wf, r, 1, "Disposition fee ($)", role="form")
put(wf, r, 2, '=IF(DispFeeOn="On",ExitGrossValue*DispFeePct,0)', role="form", nf=NF_USD0, name="DispFee"); r += 1
put(wf, r, 1, "AM fee basis amount ($/yr)", role="form")
put(wf, r, 2, '=IF(AMFeeBasis="% of Cost",TotalCost,IF(AMFeeBasis="% of NOI",StabNOI,InitialEquity))',
    role="form", nf=NF_USD0, name="AMFeeBaseAmt"); r += 1
put(wf, r, 1, "AM fee ($/yr)", role="form")
put(wf, r, 2, '=IF(AMFeeOn="On",AMFeeBaseAmt*AMFeeRate,0)', role="form", nf=NF_USD0, name="AMFeeAnnual"); r += 2

# Waterfall summary results (simplified IRR-tier promote on aggregate levered CF)
subhead(wf, r, "Distribution results (LP / GP)", 1, 6); r += 1
put(wf, r, 1, "Total levered CF to equity ($)", role="form")
put(wf, r, 2, f"=SUM({q('Monthly CF')}!{mcol(0)}${MC_LCF}:{q('Monthly CF')}!{mcol(MAX_MONTHS-1)}${MC_LCF})"
              f"+SUM({q('Debt')}!{mcol(0)}${DB_CASHOUT}:{q('Debt')}!{mcol(MAX_MONTHS-1)}${DB_CASHOUT})",
    role="link", nf=NF_USD0, name="TotalLeveredCF"); r += 1
WF_PROFIT = r
put(wf, r, 1, "Total profit (CF + initial equity) ($)", role="form")
put(wf, r, 2, "=TotalLeveredCF", role="form", nf=NF_USD0); r += 1
put(wf, r, 1, "Project levered IRR (%)", role="form")
put(wf, r, 2, f"=XIRR({q('Monthly CF')}!{mcol(0)}${MC_LCF}:{q('Monthly CF')}!{mcol(MAX_MONTHS-1)}${MC_LCF},"
              f"{q('Monthly CF')}!{mcol(0)}${MC_DATE_ROW}:{q('Monthly CF')}!{mcol(MAX_MONTHS-1)}${MC_DATE_ROW})",
    role="link", nf=NF_PCT2, name="LeveredIRR"); r += 1
put(wf, r, 1, "Equity multiple (x)", role="form")
put(wf, r, 2, f"=SUMIF({q('Monthly CF')}!{mcol(0)}${MC_LCF}:{q('Monthly CF')}!{mcol(MAX_MONTHS-1)}${MC_LCF},\">0\")/"
              f"MAX(1,-SUMIF({q('Monthly CF')}!{mcol(0)}${MC_LCF}:{q('Monthly CF')}!{mcol(MAX_MONTHS-1)}${MC_LCF},\"<0\"))",
    role="link", nf=NF_MULT, name="EquityMultiple"); r += 1

# Simplified promote: GP carry on profit using blended carry from tier table by IRR
put(wf, r, 1, "GP carry % (by IRR tier)", role="form")
# IFERROR falls back to the lowest tier's GP% when IRR is below the first hurdle.
put(wf, r, 2, '=IF(LP_Present="No",0,IFERROR(INDEX(CarryTiers,MATCH(LeveredIRR,INDEX(CarryTiers,0,1),1)+1,3),INDEX(CarryTiers,1,3)))',
    role="form", nf=NF_PCT, name="GPCarryPct"); r += 1
put(wf, r, 1, "Distributable profit (>0) ($)", role="form")
put(wf, r, 2, "=MAX(0,TotalLeveredCF)", role="form", nf=NF_USD0, name="DistributableProfit"); r += 1
# Promote is a reallocation FROM LP TO GP on the LP's profit share, scaled by the
# catch-up toggle. By construction LPDist+GPDist == TotalLeveredCF (conserves cash).
put(wf, r, 1, "GP promote ($) [carry on LP profit share]", role="form")
put(wf, r, 2, '=IF(LP_Present="No",0,DistributableProfit*LP_Split*GPCarryPct*(1+IF(GPCatchUpOn="On",CatchUpPct,0)))',
    role="form", nf=NF_USD0, name="GPPromote"); r += 1
put(wf, r, 1, "LP distribution ($)", role="form")
put(wf, r, 2, '=IF(LP_Present="No",0,TotalLeveredCF*LP_Split-GPPromote)',
    role="form", nf=NF_USD0, name="LPDist"); r += 1
put(wf, r, 1, "GP distribution ($)", role="form")
put(wf, r, 2, '=IF(LP_Present="No",TotalLeveredCF,TotalLeveredCF*GP_Split+GPPromote)',
    role="form", nf=NF_USD0, name="GPDist"); r += 1
put(wf, r, 1, "Sponsor / single-sponsor levered return (%)", role="form")
put(wf, r, 2, "=LeveredIRR", role="form", nf=NF_PCT2, name="SponsorIRR"); r += 1
put(wf, r, 1, "Check: LP+GP = total levered CF", role="form")
put(wf, r, 2, '=IF(LP_Present="No",GPDist,LPDist+GPDist)', role="form", nf=NF_USD0, name="WF_Check"); r += 1

# ======================================================================================
# TAB 2 — SUMMARY / RETURNS (built last, placed second)
# ======================================================================================
sm = wb.create_sheet("Summary")
# move Summary to position 2
wb.move_sheet("Summary", -(wb.sheetnames.index("Summary")-1))
banner(sm, "SUMMARY / RETURNS", 14)
sm.column_dimensions["A"].width = 38
sm.column_dimensions["B"].width = 18

# Model status check cell (B3) — TRUE if all checks pass
put(sm, 3, 1, "All-checks-pass flag", role="form")
# We'll compute a combined check: sources=uses, WF check ties, no negative loan
SM_CHECK = 3
put(sm, 3, 2,
    "=AND(ABS(SourcesTotal-UsesTotal)<1,ABS(WF_Check-TotalLeveredCF)<MAX(1,ABS(TotalLeveredCF))*0.001,LoanAmount>=0)",
    role="form", bold=True)

r = 5
subhead(sm, r, "Deal snapshot", 1, 4); r += 1
def kv(label, formula, nf=NF_USD0, name=None, role="link"):
    global r
    put(sm, r, 1, label, role="form")
    put(sm, r, 2, formula, role=role, nf=nf, name=name)
    r += 1
kv("Purchase price ($)", "=PurchasePrice")
kv("$/unit", "=PurchasePrice/MF_UnitCount")
kv("$/SF (total)", "=PurchasePrice/(MF_TotalSF+Retail_GLA)", nf=NF_USD2)
kv("MF units", "=MF_UnitCount", nf=NF_NUM)
kv("Retail GLA (SF)", "=Retail_GLA", nf=NF_NUM)
kv("Going-in cap (%)", "=GoingInCap", nf=NF_PCT2)
kv("Stabilized cap (%)", "=StabCap", nf=NF_PCT2)
kv("Business plan", "=BusinessPlanMode", nf=None)
r += 1

subhead(sm, r, "Sources & Uses", 1, 4); r += 1
put(sm, r, 1, "USES", role="sub"); r += 1
kv("Purchase price", "=PurchasePrice")
kv("Closing / transaction costs", "=PurchasePrice*ClosingCostPct")
kv("Capital budget", "=TotalCapBudget")
kv("Financing costs", "=FinancingCosts")
put(sm, r, 1, "Total uses", role="sub")
put(sm, r, 2, f"=SUM(B{r-4}:B{r-1})", role="form", nf=NF_USD0, name="UsesTotal"); r += 2
put(sm, r, 1, "SOURCES", role="sub"); r += 1
kv("Loan proceeds", "=LoanAmount")
kv("Equity", "=InitialEquity")
put(sm, r, 1, "Total sources", role="sub")
put(sm, r, 2, f"=SUM(B{r-2}:B{r-1})", role="form", nf=NF_USD0, name="SourcesTotal"); r += 2

subhead(sm, r, "Returns", 1, 4); r += 1
kv("Unlevered IRR (%)",
   f"=XIRR({q('Monthly CF')}!{mcol(0)}${MC_UTOT}:{q('Monthly CF')}!{mcol(MAX_MONTHS-1)}${MC_UTOT},"
   f"{q('Monthly CF')}!{mcol(0)}${MC_DATE_ROW}:{q('Monthly CF')}!{mcol(MAX_MONTHS-1)}${MC_DATE_ROW})",
   nf=NF_PCT2, name="UnleveredIRR")
kv("Levered IRR (%)", "=LeveredIRR", nf=NF_PCT2)
kv("Equity multiple (x)", "=EquityMultiple", nf=NF_MULT)
kv("Yield-on-cost (%)", "=YieldOnCost", nf=NF_PCT2)
kv("Exit cap (blended) (%)", "=ExitCapBlended", nf=NF_PCT2)
kv("Development spread (bps)", "=DevSpread", nf=NF_BPS)
kv("Loan amount ($)", "=LoanAmount")
kv("Binding constraint", "=BindingConstraint", nf=None)
kv("Initial equity ($)", "=InitialEquity")
r += 1
subhead(sm, r, "Acceptance checks", 1, 4); r += 1
kv("Sources = Uses?", "=ABS(SourcesTotal-UsesTotal)<1", nf=None, name="Chk_SU")
kv("Waterfall ties to levered CF?", "=ABS(WF_Check-TotalLeveredCF)<MAX(1,ABS(TotalLeveredCF))*0.001", nf=None, name="Chk_WF")
kv("Loan non-negative?", "=LoanAmount>=0", nf=None, name="Chk_Loan")

# ======================================================================================
# TAB 12 — SENSITIVITIES (Phase 1 stub layout + anchors)
# ======================================================================================
se = wb.create_sheet("Sensitivities")
banner(se, "SENSITIVITIES — Phase 1 stub (live data tables built in Phase 2)", 14)
se.column_dimensions["A"].width = 30
put(se, 4, 1, "Output anchors (Phase 2 sets row/col input cells via Data Table):", role="sub")
put(se, 5, 1, "Levered IRR anchor", role="form"); put(se, 5, 2, "=LeveredIRR", role="link", nf=NF_PCT2, name="Sens_LIRR")
put(se, 6, 1, "Equity multiple anchor", role="form"); put(se, 6, 2, "=EquityMultiple", role="link", nf=NF_MULT, name="Sens_EM")
put(se, 7, 1, "Unlevered IRR anchor", role="form"); put(se, 7, 2, "=UnleveredIRR", role="link", nf=NF_PCT2)

# Grid 1: exit cap (rows) x rent growth (cols) -> Levered IRR
put(se, 9, 1, "Grid 1: Exit cap (rows) x MF rent growth Y1 (cols) -> Levered IRR", role="sub")
put(se, 10, 1, "=Sens_LIRR", role="link", nf=NF_PCT2)   # data-table corner anchor
growth_axis = [0.02,0.03,0.04,0.05,0.06]
cap_axis = [0.045,0.05,0.055,0.06,0.065]
for j,g in enumerate(growth_axis):
    put(se, 10, 2 + j, g, role="input", nf=NF_PCT, fill=FILL_INPUT, align="center")
for i2,cp_ in enumerate(cap_axis):
    put(se, 11 + i2, 1, cp_, role="input", nf=NF_PCT2, fill=FILL_INPUT)
    for j in range(len(growth_axis)):
        put(se, 11 + i2, 2 + j, "", role="form")  # Phase 2 fills via Data Table
put(se, 17, 1, "Row input cell -> ExitCapBlended ; Column input cell -> Vec_MFRent Y1", role="unit")

# Grid 2: purchase price x exit cap -> IRR & EM
put(se, 19, 1, "Grid 2: Purchase price (rows) x exit cap (cols) -> Levered IRR", role="sub")
put(se, 20, 1, "=Sens_LIRR", role="link", nf=NF_PCT2)
price_axis = [55e6,58e6,62e6,66e6,70e6]
for j,cp_ in enumerate(cap_axis):
    put(se, 20, 2 + j, cp_, role="input", nf=NF_PCT2, fill=FILL_INPUT, align="center")
for i2,p in enumerate(price_axis):
    put(se, 21 + i2, 1, p, role="input", nf=NF_USD0, fill=FILL_INPUT)
put(se, 27, 1, "Row input -> InputPrice ; Column input -> ExitCapBlended", role="unit")

# Grid 3: one-way disposition date -> IRR/EM
put(se, 29, 1, "Grid 3 (one-way): Disposition date -> Levered IRR & EM", role="sub")
put(se, 30, 1, "Disp date", role="sub"); put(se, 30, 2, "=Sens_LIRR", role="link", nf=NF_PCT2); put(se,30,3,"=Sens_EM", role="link", nf=NF_MULT)
for k in range(5):
    put(se, 31 + k, 1, f"=EDATE(DispositionDate,{(k-2)*12})", role="form", nf=NF_DATE)
put(se, 37, 1, "Column input -> DispositionDate", role="unit")

# ======================================================================================
# TAB 13 — IC ONE-PAGER
# ======================================================================================
ic = wb.create_sheet("IC One-Pager")
banner(ic, "INVESTMENT COMMITTEE — One-Pager", 8)
ic.column_dimensions["A"].width = 32
ic.column_dimensions["B"].width = 18
ic.column_dimensions["C"].width = 32
ic.column_dimensions["D"].width = 18
r = 5
def icrow(rw, l1, f1, nf1, l2=None, f2=None, nf2=None):
    put(ic, rw, 1, l1, role="form"); put(ic, rw, 2, f1, role="link", nf=nf1)
    if l2:
        put(ic, rw, 3, l2, role="form"); put(ic, rw, 4, f2, role="link", nf=nf2)
subhead(ic, r, "Deal snapshot", 1, 4); r += 1
icrow(r, "Asset", "=AssetName", None, "Business plan", "=BusinessPlanMode", None); r += 1
icrow(r, "Units", "=MF_UnitCount", NF_NUM, "Retail GLA", "=Retail_GLA", NF_NUM); r += 1
icrow(r, "Purchase price", "=PurchasePrice", NF_USD0, "$/unit", "=PurchasePrice/MF_UnitCount", NF_USD0); r += 1
icrow(r, "Going-in cap", "=GoingInCap", NF_PCT2, "Stabilized cap", "=StabCap", NF_PCT2); r += 2
subhead(ic, r, "Sources & Uses", 1, 4); r += 1
icrow(r, "Loan", "=LoanAmount", NF_USD0, "Equity", "=InitialEquity", NF_USD0); r += 1
icrow(r, "Total cost", "=TotalCost", NF_USD0, "Binding constraint", "=BindingConstraint", None); r += 2
subhead(ic, r, "Returns", 1, 4); r += 1
icrow(r, "Unlevered IRR", "=UnleveredIRR", NF_PCT2, "Levered IRR", "=LeveredIRR", NF_PCT2); r += 1
icrow(r, "Equity multiple", "=EquityMultiple", NF_MULT, "Yield-on-cost", "=YieldOnCost", NF_PCT2); r += 1
icrow(r, "Exit cap", "=ExitCapBlended", NF_PCT2, "Development spread (bps)", "=DevSpread", NF_BPS); r += 2
subhead(ic, r, "Debt terms", 1, 4); r += 1
icrow(r, "Loan rate", "=LoanRate", NF_PCT2, "Amort (yrs)", "=AmortYears", NF_NUM); r += 1
icrow(r, "IO (mo)", "=IOMonths", NF_MO, "Refi", "=RefiOn", None); r += 2
subhead(ic, r, "Key assumptions / risks", 1, 4); r += 1
put(ic, r, 1, "Exit NOI basis", role="form"); put(ic, r, 2, "=ExitNOIBasis", role="link"); r += 1
put(ic, r, 1, "Tax method", role="form"); put(ic, r, 2, "=TaxMethod", role="link"); r += 1
put(ic, r, 1, "LP present", role="form"); put(ic, r, 2, "=LP_Present", role="link"); r += 1
ic.print_area = f"A1:D{r}"

# ======================================================================================
# Workbook calc properties: iterative calculation ON (circularity)
# ======================================================================================
wb.calculation = CalcProperties(calcId=124519, fullCalcOnLoad=True,
                                iterate=True, iterateCount=100, iterateDelta=0.001)

# Order sheets left->right per spec
order = ["Control Panel","Summary","Assumptions","MF Revenue","Retail Revenue","OpEx",
         "Taxes","CapEx","Debt","Monthly CF","Waterfall","Sensitivities","IC One-Pager"]
for idx, name in enumerate(order):
    cur = wb.sheetnames.index(name)
    wb.move_sheet(name, idx - cur)

wb.save(OUT)
print(f"Wrote {OUT} with {len(wb.sheetnames)} tabs and {len(named)} named ranges.")
print("Tabs:", wb.sheetnames)
