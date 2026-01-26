#!/bin/bash
# Generate test results summary for production test suite
# 
# Usage: bash tests/generate_summary.sh
# Output: tests/artifacts/TEST_SUMMARY.txt

cd "$(dirname "$0")/.."

echo "Running production test suite..."
pytest tests/test_scalar_replacement.py \
       tests/test_block_markdown.py \
       tests/test_lists.py \
       tests/test_tables.py \
       tests/test_textboxes.py \
       tests/test_documented_limitations.py \
       tests/test_negative_cases.py \
       tests/test_golden_regression.py \
       -v --tb=line > tests/artifacts/pytest_output.txt 2>&1

# Create summary file
SUMMARY_FILE="tests/artifacts/TEST_SUMMARY.txt"
echo "=== TEST RESULTS SUMMARY ===" > "$SUMMARY_FILE"
echo "Generated: $(date)" >> "$SUMMARY_FILE"
echo "" >> "$SUMMARY_FILE"

# Extract test results
echo "## Test Results by File" >> "$SUMMARY_FILE"
echo "" >> "$SUMMARY_FILE"

# Group by test file
for test_file in test_scalar_replacement test_block_markdown test_lists test_tables test_textboxes test_documented_limitations test_negative_cases test_golden_regression; do
    echo "### ${test_file}.py" >> "$SUMMARY_FILE"
    echo "" >> "$SUMMARY_FILE"
    echo "| Test Function | Status |" >> "$SUMMARY_FILE"
    echo "|---------------|--------|" >> "$SUMMARY_FILE"
    
    # Extract test functions and their status
    grep "tests/${test_file}.py::" tests/artifacts/pytest_output.txt | \
        grep -E "PASSED|FAILED|SKIPPED" | \
        while IFS= read -r line; do
            # Extract test function name (everything after :: and before PASSED/FAILED/SKIPPED)
            test_func=$(echo "$line" | sed -n 's/.*::\([^ ]*\) .*/\1/p')
            # Extract status
            if echo "$line" | grep -q "PASSED"; then
                status="✅ PASSED"
            elif echo "$line" | grep -q "FAILED"; then
                status="❌ FAILED"
            elif echo "$line" | grep -q "SKIPPED"; then
                status="⏭️ SKIPPED"
            else
                status="❓ UNKNOWN"
            fi
            
            if [ -n "$test_func" ]; then
                echo "| \`${test_func}\` | ${status} |" >> "$SUMMARY_FILE"
            fi
        done
    
    echo "" >> "$SUMMARY_FILE"
done

# Add summary statistics
echo "## Summary Statistics" >> "$SUMMARY_FILE"
echo "" >> "$SUMMARY_FILE"
TOTAL=$(grep -cE "PASSED|FAILED|SKIPPED" tests/artifacts/pytest_output.txt 2>/dev/null || echo "0")
PASSED=$(grep -c "PASSED" tests/artifacts/pytest_output.txt 2>/dev/null || echo "0")
FAILED=$(grep -c "FAILED" tests/artifacts/pytest_output.txt 2>/dev/null || echo "0")
SKIPPED=$(grep -c "SKIPPED" tests/artifacts/pytest_output.txt 2>/dev/null || echo "0")

# Remove extra newlines from counts
TOTAL=$(echo "$TOTAL" | tr -d '\n')
PASSED=$(echo "$PASSED" | tr -d '\n')
FAILED=$(echo "$FAILED" | tr -d '\n')
SKIPPED=$(echo "$SKIPPED" | tr -d '\n')

echo "| Metric | Count |" >> "$SUMMARY_FILE"
echo "|--------|-------|" >> "$SUMMARY_FILE"
echo "| **Total Tests** | ${TOTAL} |" >> "$SUMMARY_FILE"
echo "| ✅ **Passed** | ${PASSED} |" >> "$SUMMARY_FILE"
echo "| ❌ **Failed** | ${FAILED} |" >> "$SUMMARY_FILE"
echo "| ⏭️ **Skipped** | ${SKIPPED} |" >> "$SUMMARY_FILE"
echo "" >> "$SUMMARY_FILE"

# Add final status
if [ "$FAILED" -eq 0 ]; then
    echo "## ✅ All Tests Passed!" >> "$SUMMARY_FILE"
else
    echo "## ❌ Some Tests Failed" >> "$SUMMARY_FILE"
    echo "" >> "$SUMMARY_FILE"
    echo "Check individual test artifacts in:" >> "$SUMMARY_FILE"
    echo "\`\`\`" >> "$SUMMARY_FILE"
    echo "tests/artifacts/<test_name>/" >> "$SUMMARY_FILE"
    echo "\`\`\`" >> "$SUMMARY_FILE"
fi

echo "" >> "$SUMMARY_FILE"
echo "---" >> "$SUMMARY_FILE"
echo "" >> "$SUMMARY_FILE"
echo "For detailed output, see: \`tests/artifacts/pytest_output.txt\`" >> "$SUMMARY_FILE"

# Display summary
cat "$SUMMARY_FILE"
echo ""
echo "✅ Summary saved to: $SUMMARY_FILE"
