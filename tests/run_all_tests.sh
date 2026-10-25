#!/usr/bin/env bash
# ==============================================================================
# Master Test Runner: run_all_tests.sh
# Runs Python unit tests and Bash integration tests for the trace report suite.
# ==============================================================================

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

echo "=========================================================="
echo " 1/3 Running Python Unit Tests for csv_transformer.py"
echo "=========================================================="
PY_CMD=""
if [[ -n "${PYTHON_BIN:-}" ]] && "$PYTHON_BIN" -c "import openpyxl" 2>/dev/null; then
    PY_CMD="$PYTHON_BIN"
elif command -v python3 >/dev/null 2>&1 && python3 -c "import openpyxl" 2>/dev/null; then
    PY_CMD="python3"
elif command -v python >/dev/null 2>&1 && python -c "import openpyxl" 2>/dev/null; then
    PY_CMD="python"
elif command -v python.exe >/dev/null 2>&1 && python.exe -c "import openpyxl" 2>/dev/null; then
    PY_CMD="python.exe"
fi

if [[ -z "$PY_CMD" ]]; then
    echo "Python 3 with openpyxl not found. Please install python3-openpyxl." >&2
    exit 1
fi

cd "$SCRIPT_DIR"
"$PY_CMD" -m unittest tests/test_csv_transformer.py

echo ""
echo "=========================================================="
echo " 2/3 Running Python Unit Tests for inject_to_excel.py"
echo "=========================================================="
"$PY_CMD" -m unittest tests/test_inject_to_excel.py

echo ""
echo "=========================================================="
echo " 3/3 Running Bash End-to-End Tests for export_trace_report.sh"
echo "=========================================================="
bash tests/test_export_trace_report.sh

echo ""
echo "=========================================================="
echo " ALL TEST SUITES PASSED! SUCCESS! "
echo "=========================================================="
