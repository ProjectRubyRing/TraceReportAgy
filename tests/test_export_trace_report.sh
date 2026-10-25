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

# ==============================================================================
# Switchback and Permission Testing with Mock AWS CLI
# ==============================================================================
echo ""
echo "--- Setting up Mock AWS CLI environment for Switchback Testing ---"
MOCK_BIN="${TEST_OUTPUT_DIR}/mock_bin"
mkdir -p "$MOCK_BIN"
SAMPLE_JSON="${SCRIPT_DIR}/sample_data/spans_query_result.json"
MOCK_SW_SCRIPT="${TEST_OUTPUT_DIR}/test_switchback.sh"

cat << 'EOF' > "$MOCK_SW_SCRIPT"
#!/usr/bin/env bash
export MOCK_SWITCHBACK_EXECUTED=1
export AWS_PROFILE="switched-ops-profile"
echo "[MOCK_SWITCHBACK] Dedicated switchback shell sourced successfully."
EOF
chmod +x "$MOCK_SW_SCRIPT"

cat << EOF > "${MOCK_BIN}/aws"
#!/usr/bin/env bash
if [ "\$1" = "sts" ] && [ "\$2" = "get-caller-identity" ]; then
    if [ "\${MOCK_SCENARIO:-}" = "unauthenticated" ]; then
        echo "An error occurred (ExpiredToken): The security token included in the request is expired" >&2
        exit 254
    elif [ "\${MOCK_SCENARIO:-}" = "codecommit_role" ]; then
        echo '{"UserId": "AROAEXAMPLE", "Account": "123456789012", "Arn": "arn:aws:sts::123456789012:assumed-role/CodeCommit-Dev-Role/session"}'
        exit 0
    elif [ "\${MOCK_SCENARIO:-}" = "switchback_success" ]; then
        if [ "\${MOCK_SWITCHBACK_EXECUTED:-0}" = "1" ]; then
            echo '{"UserId": "AROAEXAMPLE", "Account": "123456789012", "Arn": "arn:aws:sts::123456789012:assumed-role/Operations-Role/session"}'
            exit 0
        else
            echo '{"UserId": "AROAEXAMPLE", "Account": "123456789012", "Arn": "arn:aws:sts::123456789012:assumed-role/CodeCommit-Dev-Role/session"}'
            exit 0
        fi
    else
        echo '{"UserId": "AROAEXAMPLE", "Account": "123456789012", "Arn": "arn:aws:sts::123456789012:assumed-role/AppSignals-Operator/session"}'
        exit 0
    fi
fi

if [ "\$1" = "logs" ] && [ "\$2" = "describe-log-groups" ]; then
    if [ "\${MOCK_SCENARIO:-}" = "permission_denied" ] || [ "\${MOCK_SCENARIO:-}" = "codecommit_role" ]; then
        echo "An error occurred (AccessDeniedException): User is not authorized to perform: logs:DescribeLogGroups on resource: *" >&2
        exit 254
    elif [ "\${MOCK_SCENARIO:-}" = "switchback_success" ]; then
        if [ "\${MOCK_SWITCHBACK_EXECUTED:-0}" = "1" ]; then
            echo '{"logGroups": [{"logGroupName": "aws/spans"}]}'
            exit 0
        else
            echo "An error occurred (AccessDeniedException): User is not authorized to perform: logs:DescribeLogGroups" >&2
            exit 254
        fi
    elif [ "\${MOCK_SCENARIO:-}" = "switchback_still_denied" ]; then
        echo "An error occurred (AccessDeniedException): Still denied after switchback" >&2
        exit 254
    else
        echo '{"logGroups": [{"logGroupName": "aws/spans"}]}'
        exit 0
    fi
fi

if [ "\$1" = "logs" ] && [ "\$2" = "start-query" ]; then
    echo '{"queryId": "mock-query-id-test-12345"}'
    exit 0
fi

if [ "\$1" = "logs" ] && [ "\$2" = "get-query-results" ]; then
    cat "$SAMPLE_JSON"
    exit 0
fi

echo "Mock aws unhandled invocation: \$*" >&2
exit 1
EOF
chmod +x "${MOCK_BIN}/aws"

# Test 8: CLI options check in help message
echo "Running Test 8: CLI options for switchback present in --help..."
HELP_FULL="$(bash "$BIN_PATH" --help)"
if echo "$HELP_FULL" | grep -q -- "--auto-switchback" && echo "$HELP_FULL" | grep -q -- "--switchback-script"; then
    pass "Help message contains --auto-switchback and --switchback-script options."
else
    fail "Help message missing switchback CLI options."
fi

# Test 9: Unauthenticated scenario -> Exit 1
echo "Running Test 9: Unauthenticated session detection..."
set +e
UNAUTH_OUT="$(PATH="${MOCK_BIN}:${PATH}" MOCK_SCENARIO="unauthenticated" bash "$BIN_PATH" -s 1h -o "$TEST_OUTPUT_DIR" 2>&1)"
UNAUTH_RET=$?
set -e
if [[ $UNAUTH_RET -eq 1 ]] && echo "$UNAUTH_OUT" | grep -q "AWS にログインされていないか"; then
    pass "Unauthenticated scenario correctly exited with code 1 and login prompt."
else
    fail "Unauthenticated scenario did not exit with code 1. Code: $UNAUTH_RET, Out: $UNAUTH_OUT"
fi

# Test 10: Permission denied, auto-switchback disabled -> Exit 2, warning message to switch back
echo "Running Test 10: Permission denied with switchback warning (auto-switchback disabled)..."
set +e
DENIED_OUT="$(PATH="${MOCK_BIN}:${PATH}" MOCK_SCENARIO="codecommit_role" bash "$BIN_PATH" -s 1h -S "$MOCK_SW_SCRIPT" -o "$TEST_OUTPUT_DIR" 2>&1)"
DENIED_RET=$?
set -e
if [[ $DENIED_RET -eq 2 ]] && echo "$DENIED_OUT" | grep -q "スイッチバック警告" && echo "$DENIED_OUT" | grep -q "codecommit"; then
    pass "Permission denied with switchback warning correctly warned and exited with code 2."
else
    fail "Permission denied test failed. Code: $DENIED_RET, Out: $DENIED_OUT"
fi

# Test 11: Auto-switchback enabled -> Auto source switchback script, Exit 0 and complete
echo "Running Test 11: Auto-switchback execution and recovery (--auto-switchback)..."
set +e
SW_SUCCESS_OUT="$(PATH="${MOCK_BIN}:${PATH}" MOCK_SCENARIO="switchback_success" MOCK_SAMPLE_JSON="$SAMPLE_JSON" bash "$BIN_PATH" -s 1h -a -S "$MOCK_SW_SCRIPT" -o "$TEST_OUTPUT_DIR" 2>&1)"
SW_SUCCESS_RET=$?
set -e
if [[ $SW_SUCCESS_RET -eq 0 ]] && echo "$SW_SUCCESS_OUT" | grep -q "スイッチバックに成功しました" && echo "$SW_SUCCESS_OUT" | grep -q "TRACE REPORT COMPLETED SUCCESSFULLY"; then
    pass "Auto-switchback automatically sourced script, recovered permissions, and generated report (exit code 0)."
else
    fail "Auto-switchback execution failed. Code: $SW_SUCCESS_RET, Out: $SW_SUCCESS_OUT"
fi

# Test 12: Auto-switchback enabled but script does not exist -> Exit 3
echo "Running Test 12: Switchback script missing detection..."
set +e
MISSING_SW_OUT="$(PATH="${MOCK_BIN}:${PATH}" MOCK_SCENARIO="switchback_success" bash "$BIN_PATH" -s 1h -a -S "${TEST_OUTPUT_DIR}/non_existent_switchback.sh" -o "$TEST_OUTPUT_DIR" 2>&1)"
MISSING_SW_RET=$?
set -e
if [[ $MISSING_SW_RET -eq 3 ]] && echo "$MISSING_SW_OUT" | grep -q "スイッチバック用シェルが存在しません"; then
    pass "Missing switchback script correctly detected with exit code 3."
else
    fail "Missing switchback script test failed. Code: $MISSING_SW_RET, Out: $MISSING_SW_OUT"
fi

# Test 13: Auto-switchback executed but permissions still denied -> Exit 4
echo "Running Test 13: Auto-switchback executed but permissions still denied..."
set +e
STILL_DENIED_OUT="$(PATH="${MOCK_BIN}:${PATH}" MOCK_SCENARIO="switchback_still_denied" bash "$BIN_PATH" -s 1h -a -S "$MOCK_SW_SCRIPT" -o "$TEST_OUTPUT_DIR" 2>&1)"
STILL_DENIED_RET=$?
set -e
if [[ $STILL_DENIED_RET -eq 4 ]] && echo "$STILL_DENIED_OUT" | grep -q "依然として AWS 操作権限が不足しています"; then
    pass "Permission still denied after switchback correctly exited with code 4."
else
    fail "Still denied test failed. Code: $STILL_DENIED_RET, Out: $STILL_DENIED_OUT"
fi

# Test 14: Environment variable AWS_SWITCHBACK_SCRIPT fallback
echo "Running Test 14: Environment variable AWS_SWITCHBACK_SCRIPT resolution..."
set +e
ENV_SW_OUT="$(PATH="${MOCK_BIN}:${PATH}" MOCK_SCENARIO="codecommit_role" AWS_SWITCHBACK_SCRIPT="$MOCK_SW_SCRIPT" bash "$BIN_PATH" -s 1h -o "$TEST_OUTPUT_DIR" 2>&1)"
ENV_SW_RET=$?
set -e
if [[ $ENV_SW_RET -eq 2 ]] && echo "$ENV_SW_OUT" | grep -q "$MOCK_SW_SCRIPT"; then
    pass "Switchback script path correctly resolved from AWS_SWITCHBACK_SCRIPT environment variable."
else
    fail "AWS_SWITCHBACK_SCRIPT environment variable test failed. Code: $ENV_SW_RET, Out: $ENV_SW_OUT"
fi

echo "================================================="
echo -e "${GREEN} ALL BASH TESTS PASSED SUCCESSFULLY! ${RESET}"
echo "================================================="
