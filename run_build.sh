#!/usr/bin/env bash
# Phase-1 build pipeline: write the workbook, then run the Python solver + Section 8 tests.
set -e
python3 build_model.py
python3 solve_and_test.py
