#!/usr/bin/env bash
# Full pipeline:
#   1. build the workbook (formulas, named ranges, formatting)
#   2. solve max-bid + run the Section 8 acceptance tests + structural validation
#   3. Phase 2 finishing pass (sensitivity grids, conditional formatting, print setup)
#   4. scenario sweep (38 mock deals x toggle combos, invariant checks)
set -e
python3 build_model.py
python3 solve_and_test.py
python3 phase2_finish.py
python3 scenarios.py
echo
echo "Optional independent check (slow, ~7 min): python3 validate_excel.py"
