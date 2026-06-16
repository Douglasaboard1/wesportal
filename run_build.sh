#!/usr/bin/env bash
# Full pipeline: build scaffold -> solve max-bid + run Section 8 tests -> Phase 2 finish.
set -e
python3 build_model.py
python3 solve_and_test.py
python3 phase2_finish.py
