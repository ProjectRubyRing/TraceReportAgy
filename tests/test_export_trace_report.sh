#!/usr/bin/env bash
# ==============================================================================
# Test Script: test_export_trace_report.sh
# End-to-end and regression tests for export_trace_report.sh on RHEL 9.8 / Bash.
# ==============================================================================

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
BIN_PATH="${SCRIPT_DIR}/bin/export_trace_report.sh"
TEST_OUTPUT_DIR="${SCRIPT_DIR}/output/test_run_$$"

# Color helpers
GREEN="\033[32m"
RED="\033[31m"
RESET="\033[0m"

pass() { echo -e "${GREEN}[PASS]${RESET} $*"; }
fail() { echo -e "${RED}[FAIL]${RESET} $*" >&2; exit 1; }

cleanup() {
    rm -rf "$TEST_OUTPUT_DIR"
}
trap cleanup EXIT INT TERM

echo "================================================="
echo " Running Bash Test Suite for export_trace_report"
echo "================================================="

# Test 1: Syntax validation
echo "Running Test 1: Bash syntax check (bash -n)..."
if bash -n "$BIN_PATH"; then
    pass "Syntax validation succeeded."
else
    fail "Syntax error in $BIN_PATH."
fi

# Test 2: Help option
echo "Running Test 2: Help option (--help)..."
HELP_OUT="$(bash "$BIN_PATH" --help)"
if echo "$HELP_OUT" | grep -q "export_trace_report.sh" && echo "$HELP_OUT" | grep -q "aws/spans"; then
    pass "Help message contains expected usage and log group."
else
    fail "Help output missing expected keywords."
fi

# Test 3: Dry run mode
echo "Running Test 3: Dry-run option (--dry-run)..."
DRY_OUT="$(bash "$BIN_PATH" --dry-run -s 1h -t "trace-test-123")"
if echo "$DRY_OUT" | grep -q "filter traceId = 'trace-test-123'" && echo "$DRY_OUT" | grep -q "aws logs start-query"; then
    pass "Dry run correctly constructed query and command preview."
else
    fail "Dry run output does not match expectation."
fi

# Test 4: Negative testing - Invalid start >= end time
echo "Running Test 4: Negative test - start_time >= end_time..."
set +e
ERR_OUT="$(bash "$BIN_PATH" --start-time "2026-10-26 10:00:00" --end-time "2026-10-26 09:00:00" 2>&1)"
RET_CODE=$?
set -e
if [[ $RET_CODE -ne 0 ]] && echo "$ERR_OUT" | grep -q "must be before end time"; then
    pass "Rejected invalid time range as expected (exit code $RET_CODE)."
else
    fail "Did not reject invalid time range as expected."
fi

# Test 5: Negative testing - Invalid date format
echo "Running Test 5: Negative test - invalid date format..."
set +e
ERR_DATE="$(bash "$BIN_PATH" --start-time "this_is_not_a_date" 2>&1)"
RET_DATE=$?
set -e
if [[ $RET_DATE -ne 0 ]] && echo "$ERR_DATE" | grep -q "Unable to parse date"; then
    pass "Rejected malformed date as expected (exit code $RET_DATE)."
else
    fail "Did not reject malformed date as expected."
fi

# Test 6: Mock execution end-to-end
echo "Running Test 6: Mock execution (--mock)..."
mkdir -p "$TEST_OUTPUT_DIR"
MOCK_LOG="$(bash "$BIN_PATH" --mock -s 1h -o "$TEST_OUTPUT_DIR")"
if echo "$MOCK_LOG" | grep -q "TRACE REPORT COMPLETED SUCCESSFULLY"; then
    pass "Mock mode ran successfully to completion."
else
    fail "Mock execution failed. Log: $MOCK_LOG"
fi

# Verify generated artifacts
CSV_COUNT=$(find "$TEST_OUTPUT_DIR" -maxdepth 1 -name "trace_spans_*.csv" | wc -l)
XLSX_COUNT=$(find "$TEST_OUTPUT_DIR" -maxdepth 1 -name "trace_report_*.xlsx" | wc -l)
JSON_COUNT=$(find "$TEST_OUTPUT_DIR" -maxdepth 1 -name "query_results_*.json" | wc -l)

if [[ "$CSV_COUNT" -ge 1 && "$XLSX_COUNT" -ge 1 && "$JSON_COUNT" -ge 1 ]]; then
    pass "All output artifacts (JSON, CSV, XLSX) exist in output directory."
else
    fail "Artifact count mismatch: CSV=$CSV_COUNT, XLSX=$XLSX_COUNT, JSON=$JSON_COUNT"
fi

# Test 7: Verify generated CSV content
CSV_FILE=$(find "$TEST_OUTPUT_DIR" -maxdepth 1 -name "trace_spans_*.csv" | head -n 1)
LINE_COUNT=$(wc -l < "$CSV_FILE")
if [[ "$LINE_COUNT" -eq 18 ]]; then # 1 header + 17 spans
    pass "Exported CSV contains exactly 18 lines (header + 17 spans)."
else
    fail "CSV line count expected 18, got $LINE_COUNT."
fi

echo "================================================="
echo -e "${GREEN} ALL BASH TESTS PASSED SUCCESSFULLY! ${RESET}"
echo "================================================="
