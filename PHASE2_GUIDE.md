# Phase 2 in Claude for Excel — Step-by-Step Guide

You now have **`MixedUse_Acquisition_Model.xlsx`** (Phase 1 complete). Phase 2 finishes
the four things that need Excel's live calculation engine: confirm recalculation, build
the sensitivity data tables, add the Goal-Seek price solver, add conditional
formatting, and set print areas. None of this rewrites the model — it sits on top of
what's already there.

---

## 0. Open the file and turn on iterative calculation (do this first)

The model is intentionally circular (price → loan → equity → return; DSCR → loan →
interest). It is already flagged for iterative calc, but confirm it:

- **Excel (Windows):** File ▸ Options ▸ Formulas ▸ check **Enable iterative
  calculation**, Maximum iterations **100**, Maximum change **0.001**.
- **Excel (Mac):** Excel ▸ Preferences ▸ Calculation ▸ same settings.
- Then press **F9** (or Cmd-=) to recalculate. If you ever see a circular-reference
  warning, the fix is to confirm this setting — **do not** break any links.

**Sanity check after recalc:** go to the **Summary** tab. The **Model status** cell
(top, "All-checks-pass flag") should read **TRUE**. On **Control Panel**, the **Model
Status** cell should read **OK**. If so, the live model ties out.

### A prompt you can paste into Claude for Excel
> "Open the workbook. Confirm iterative calculation is enabled (100 iterations, 0.001
> max change) and recalculate. Then tell me the value of the named cells `LeveredIRR`,
> `EquityMultiple`, `LP_IRR`, `GP_IRR`, and the `Summary` model-status flag in B3."

---

## 1. Build the sensitivity data tables (Sensitivities tab)

The tab already has three grids laid out, each with the output formula in the
**top-left corner** and the axis values along the top row / left column. You just need
to apply Excel's Data Table to each. **The exact input cells are printed in the note
under each grid.**

### Grid 1 — Exit cap (rows) × MF rent growth (cols) → Levered IRR
1. Select the whole block **A10:F15** (corner formula + both axes).
2. **Data ▸ What-If Analysis ▸ Data Table…**
3. **Row input cell:** `Assumptions!$E$56`  (MF rent growth, Year 1)
4. **Column input cell:** `Assumptions!$B$13`  (`ExitCapBlended`)
5. OK. The interior fills with Levered IRR at each combination.

### Grid 2 — Purchase price (rows) × exit cap (cols) → Levered IRR
1. Select **A20:F25**.
2. **Data ▸ What-If Analysis ▸ Data Table…**
3. **Row input cell:** `Assumptions!$B$13`  (`ExitCapBlended`)
4. **Column input cell:** `Assumptions!$B$6`  (`InputPrice`)
5. OK.

### Grid 3 — Disposition date → Levered IRR & Equity Multiple (one-variable)
1. Select **A30:C35** (dates run **down** column A; two output columns B and C).
2. **Data ▸ What-If Analysis ▸ Data Table…**
3. Leave **Row input cell blank**; **Column input cell:** `Control Panel!$B$9`
   (`DispositionDate`).
4. OK. This demonstrates the date-driven hold — IRR/EM re-extend as the exit moves.

> ⚠️ Data Tables only work in real Excel (desktop or the web app with calc on). They
> won't compute in a pure viewer. After building, press **F9** if values show 0.

### Prompt for Claude for Excel
> "On the Sensitivities tab, build three What-If Data Tables exactly as described in the
> note cell under each grid: Grid 1 over A10:F15 (row input Assumptions!$E$56, column
> input Assumptions!$B$13), Grid 2 over A20:F25 (row input Assumptions!$B$13, column
> input Assumptions!$B$6), and Grid 3 over A30:C35 (one-variable, column input
> 'Control Panel'!$B$9). Recalculate and confirm the interiors populate."

---

## 2. Add the interactive Goal-Seek price solver

Phase 1 already wrote a solved max-bid price into `SolvedPrice` (Assumptions!B7). Goal
Seek lets you re-solve interactively for any target.

1. First set the model to read your solved price: on **Control Panel**, set **Price
   Mode** (cell B15) to **`Solve for Target Return`**. Now `PurchasePrice` follows the
   price you solve for. (Set it back to `Input Price` to type a price directly.)
2. **Data ▸ What-If Analysis ▸ Goal Seek…**
   - **Set cell:** `Waterfall!$B$31`  (`LeveredIRR`)
   - **To value:** your target, e.g. **0.15**
   - **By changing cell:** `Assumptions!$B$6`  (`InputPrice`)
3. OK. Excel iterates `InputPrice` until Levered IRR hits the target. That's your
   **maximum supportable bid**.

> Because the model is circular, Goal Seek runs on top of iterative calc — that's
> expected. The Python binary-search already put a reference answer in `SolvedPrice`;
> Goal Seek should land within a rounding of it.

### Prompt for Claude for Excel
> "Run Goal Seek: set `Waterfall!$B$31` to 0.15 by changing `Assumptions!$B$6`. Report
> the resulting purchase price, the going-in cap, and confirm Levered IRR ≈ 15%."

---

## 3. Conditional formatting on the status / flag cells

Make failures impossible to miss.

1. **Summary!B3** (all-checks flag) and **Control Panel** Model Status, and the **Debt**
   tab **Binding constraint** cell: select each, **Home ▸ Conditional Formatting ▸
   New Rule**.
2. Rule: *Format only cells that contain* → text **does not contain** `OK` (or
   *equal to* `FALSE`) → fill **red**. Add a companion rule for `OK`/`TRUE` → **green**.

### Prompt for Claude for Excel
> "Add conditional formatting: turn `Summary!B3` and the Control Panel model-status
> cell **red** when not OK/TRUE and **green** when OK/TRUE."

---

## 4. Print area / IC one-pager polish

1. Go to **IC One-Pager**. The print area is already set (A1:D~30).
2. **Page Layout ▸ Print Area ▸ Set Print Area** if you adjust it; **Fit Sheet on One
   Page** under Print scaling.
3. Optional: hide the helper monthly columns on schedule tabs before printing (select
   columns E onward ▸ Hide) so printed pages stay tidy.

---

## 5. Re-run the Section 8 checks visually

Flip a few toggles on **Control Panel** and watch results move (the model recalcs live):

| Toggle (Control Panel) | What to expect |
|---|---|
| **Business Plan Mode** → Stabilized / Value-Add / Core-Plus | NOI ramp and returns change |
| **Retail GLA → 0** (set suite SF to 0 on Retail Revenue) | clean pure-MF, no errors |
| **Disposition Date** +24 mo | every schedule and IRR re-extend |
| **Debt Structure** → Separate MF + Retail Loans | two loans size; binding flag shows "MF:.. / RT:.." |
| **LP Present** → No | promote disappears, GP gets 100%, Summary still ties |
| Turn each **fee** off | equity, cash flow, and promote update |

Each should behave as described and the **Summary status stays TRUE**.

---

### What NOT to do
- Don't retype formulas or "fix" circular-reference warnings by deleting links — confirm
  iterative calc instead.
- Don't move the period columns (E onward) — every tab is aligned to that monthly grid.
- Inputs are **blue**; formulas are **black**; cross-tab links are **green**. Only change
  blue cells.
