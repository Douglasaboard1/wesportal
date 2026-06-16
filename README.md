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
- **Debt** (§6.4): min-of sizing across LTV / LTC / DSCR / Debt Yield (each
  independently on/off) with a binding-constraint flag; fixed or floating rate; IO
  then amortization; monthly balance schedule; refi block with cash-out to the
  waterfall.
- **Growth block** (§6.10): separate annual growth **vectors** (not single rates)
  with cumulative-factor rows that every revenue/expense line references.
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

## Acceptance tests (Section 8) — all 14 pass

Toggle integrity · retail-to-zero · date-driven hold · min-of debt binding ·
price-solver round-trip · LP-present toggle · fee removal · no-hardcode / name-
resolution scan · sources = uses. Run `python3 solve_and_test.py` to see the report;
it also scans all 9,272 formula cells for `#REF!`, unresolved name tokens, and
doubled sheet qualifiers.

## Known simplifications → deepen in Phase 2 (or a Phase 1.5)

These are deliberate, documented scaffolding choices, not silent shortcuts:

1. **Waterfall** is a conserving, IRR-tier promote (return-of-capital + 8% pref +
   catch-up are represented at the aggregate level, gated by `LP_Present`), not yet a
   full period-by-period American distribution with sequential ROC/pref tracking. The
   LP/GP split conserves cash by construction (`LPDist + GPDist = TotalLeveredCF`).
2. **Initial loan draw** is taken in full at month 0; the financed-capital holdback is
   not drawn pro-rata as capital is spent (a mild, conservative timing simplification
   for interest).
3. **Loss-to-lease, lease-up, and reno ramps** operate at the rent-roll aggregate on
   the monthly build rather than per-unit-per-month; per-unit detail lives on the rent
   roll. Setting retail GLA to 0 still yields a clean pure-MF model.
4. **Sensitivities** tab is the Phase-1 stub (axes + output anchors). The live two-way
   data tables and Goal-Seek are Phase-2 (native Excel calc engine required).

## Phase 2 handoff (§9.2)

Open the workbook, confirm iterative calc is on, then: build the native two-way data
tables (row/column input cells are documented on the Sensitivities tab), add the
Goal-Seek price solver (Set `LeveredIRR` → To target → By changing `InputPrice`),
apply conditional formatting to the model-status and binding-constraint flags, set the
IC one-pager print area, and re-run the Section 8 checks visually.
