# Institutional Mixed-Use (Multifamily + Ground-Floor Retail) Acquisition Model — Phase 1

Phase 1 (Claude Code / openpyxl) of the hybrid build spec. This repo generates the
full `.xlsx` scaffold with real Excel formulas, named ranges, role-based formatting,
iterative-calc properties, a binary-search max-bid price solver, and the Section 8
acceptance tests. Phase 2 (Claude for Excel) finishes the live data tables, the
interactive Goal-Seek solver, conditional formatting, and print polish.

## Files

| File | Purpose |
|------|---------|
| `model_inputs.py` | **Single source of truth** — sample rent rolls, growth vectors, recovery matrix, and all default assumptions. Imported by both scripts so the live Excel model and the numeric replica can never drift. |
| `build_model.py` | Writes `MixedUse_Acquisition_Model.xlsx`: 13 tabs, every formula as a real cell-formula string, 178 named ranges, fonts by role, iterative calc on. |
| `solve_and_test.py` | A reduced-form numeric replica of the engine used to (1) binary-search the max-bid price and write it into the `SolvedPrice` cell, and (2) run the Section 8 assertions + a structural validation of the workbook. |
| `run_build.sh` | Runs the two scripts in order (build, then solve+test). |
| `MixedUse_Acquisition_Model.xlsx` | The generated workbook (the Phase-1 deliverable). |

## Run

```bash
pip install openpyxl
./run_build.sh
```

Re-running reproduces the workbook exactly — the scripts are the audit trail.

## What Phase 1 delivers (maps to the spec)

- **13 tabs** (§4) left→right: Control Panel · Summary · Assumptions · MF Revenue ·
  Retail Revenue · OpEx · Taxes · CapEx · Debt · Monthly CF · Waterfall ·
  Sensitivities · IC One-Pager.
- **17 master toggles** (§3) as named ranges, referenced by name everywhere
  (`BusinessPlanMode`, `PriceMode`, `LP_Present`, …). Toggles re-wire logic via
  `IF`/`SWITCH`-style formulas, not cosmetics.
- **Monthly engine → annual roll-up** (§5.7) across a 120-month built horizon. The
  hold is **date-driven**: every period is gated by `date <= DispositionDate`, so
  shifting Acquisition/Disposition dates extends or contracts every schedule, the
  debt amortization, the waterfall, and the IRRs automatically (Section 8 test #3).
- **MF rent roll** (§5.1, 88 units) → GPR, loss-to-lease burn-off, vacancy/
  concessions/bad-debt, itemized other income → MF EGI. Value-add reno ramp and
  core-plus lease-up ramp are mode-gated.
- **Retail rent roll** (§5.2, 6 suites) → base + contractual steps + rollover at
  expiry + recoveries (editable recovery matrix, §6.3) + percentage rent (built but
  defaulted inert) − general vacancy − credit loss → Retail EGI; rollover TI/LC to
  CapEx.
- **OpEx / Taxes / CapEx** (§5.3–5.5): itemized OpEx with mgmt fee % of EGI; three
  tax methods behind toggle #13 (Reassessment-on-Sale default); Sources & Uses with
  per-bucket Equity/Financed funding.
- **Debt** (§6.4, 6.6): min-of sizing across LTV / LTC / DSCR / Debt Yield (each
  independently on/off) with a binding-constraint flag; fixed or floating rate; IO
  then amortization; monthly balance schedule; refi block with cash-out to the
  waterfall. **Two structures** (toggle #6): a single blended loan, or **separate MF +
  Retail loans** — each leg sized against its own component value/cost/NOI with its
  own rate type, rate, IO, amortization, and sizing constraints (Loan A + Loan B run as
  parallel schedules that sum). **Financed capital** is drawn **pro-rata as spent**, not
  lumped at close, so interest accrues on the actual outstanding balance.
- **Component NOI** (Monthly CF): OpEx and taxes allocated to MF vs. Retail by EGI
  share (MF NOI + Retail NOI = total NOI), driving the separate-loan sizing.
- **Waterfall** (§6.8): a true **period-by-period American waterfall** — Return of
  Capital → compounded Preferred → GP Catch-up (toggle) → residual split by IRR-hurdle
  tiers — computed monthly with running balances, gated by `LP_Present`. Emits LP/GP
  cash flows, LP/GP IRR & equity multiple, and GP promote. Conserves by construction:
  `SUM(LP CF) + SUM(GP CF) = project levered CF`.
- **Growth block** (§6.10): separate annual growth **vectors** (not single rates)
  with cumulative-factor rows that every revenue/expense line references.
- **Retail recoveries** (§6.3): NNN/Gross reimburse the grown pool; **Modified Gross**
  reimburses only the expense growth above a base-year stop; the **gross-up** toggle
  grosses recoverable expenses to full occupancy; admin fee applies on top.
- **Taxes** (§5.4): Flat Growth · Reassessment-on-Sale (default) · **PILOT / Abatement
  as a user year-by-year schedule** (§5.4), all behind toggle #13.
- **Exit & returns** (§6.12–6.13): blended or component-cap exit, trailing/forward
  exit-NOI basis, yield-on-cost, development spread, going-in/stabilized caps,
  unlevered & levered XIRR, equity multiple.
- **Price solver** (§6.6): `solve_and_test.py` binary-searches the purchase price for
  the target levered IRR and writes the result into `SolvedPrice` (the one
  intentional computed value). `PurchasePrice` = `IF(PriceMode="Solve…", SolvedPrice,
  InputPrice)`, so the live formula model stays intact.
- **Iterative calculation** is enabled (`iterate=True`, count 100, delta 0.001,
  `fullCalcOnLoad`) to resolve the price→loan→equity→IRR and
  financing-cost→cost→LTC→loan circularities.

### Conventions enforced (§2)
- Blue font = hardcoded input · Black = formula · Green = cross-tab link.
- Named ranges for all toggles/global drivers; toggles referenced by name.
- Every tab carries a header block (asset, run date, plan mode, hold, model-status).
- A central `fix_ranges()` pass guarantees no cross-sheet range repeats its sheet
  qualifier (`'Monthly CF'!E5:DT5`, never `…:'Monthly CF'!DT5`).

## Acceptance tests (Section 8) — all 24 pass

Toggle integrity · retail-to-zero · date-driven hold · min-of debt binding ·
**separate MF+Retail loan sizing** · price-solver round-trip · LP-present toggle ·
**waterfall cash conservation (LP & GP)** · fee removal · sources = uses ·
**component NOI reconciliation**. Run `python3 solve_and_test.py` to see the report; it
also scans all ~14,300 formula cells for `#REF!`, unresolved name tokens, and doubled
sheet qualifiers.

## Waterfall interpretation (documented choice)

The spec's tier table is ambiguous in the pref region. The model implements the
standard institutional sequence and applies the table to the **residual** above the
pref: Return of Capital → 8% compounded Preferred (100% to investors) → GP Catch-up
(toggle; GP share = `CatchUpPct`) → residual split by investor IRR hurdle
(8–12% → 80/20, 12–15% → 70/30, >15% → 60/40). Hurdle tiers use the accreted-balance
(lookback) method, so they are exact and non-iterative. Edit any input on the
Waterfall tab to change the structure.

## Known simplifications → Phase 2 (live calc engine required)

1. **Loss-to-lease, lease-up, and reno ramps** operate at the rent-roll aggregate on
   the monthly build rather than per-unit-per-month; per-unit detail lives on the rent
   roll. Setting retail GLA to 0 still yields a clean pure-MF model.
2. **Sensitivities** tab is the Phase-1 stub (axes + output anchors). The live two-way
   data tables and Goal-Seek are Phase-2 (native Excel calc engine required).

## Phase 2 handoff (§9.2)

Open the workbook, confirm iterative calc is on, then: build the native two-way data
tables (row/column input cells are documented on the Sensitivities tab), add the
Goal-Seek price solver (Set `LeveredIRR` → To target → By changing `InputPrice`),
apply conditional formatting to the model-status and binding-constraint flags, set the
IC one-pager print area, and re-run the Section 8 checks visually.
