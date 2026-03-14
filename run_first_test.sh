#!/bin/bash
# CSV Automation Tool - First Test Runner

# Activate virtual environment
source venv/bin/activate

# Check if API key is set
if [ -z "$ANTHROPIC_API_KEY" ]; then
    echo "❌ Error: ANTHROPIC_API_KEY not set"
    echo ""
    echo "Please set your API key first:"
    echo "  export ANTHROPIC_API_KEY='your-key-here'"
    echo ""
    exit 1
fi

echo "🚀 CSV Automation Tool - First Test"
echo "=================================="
echo ""
echo "URL: https://example.com"
echo "Workbook: scenarios/sample-test.xlsx"
echo "Mode: Visible browser (--no-headless)"
echo ""
echo "Press Ctrl+C to cancel, or wait 3 seconds to start..."
sleep 3

# Run the test
csvtool run \
  --url "https://example.com" \
  --workbook "scenarios/sample-test.xlsx" \
  --no-headless

echo ""
echo "✅ Test Complete!"
echo ""
echo "Check results:"
echo "  Evidence document: ls -lh output/*.docx"
echo "  Screenshots: ls output/screenshots/"
echo "  Cache: cat cache/*.json"
