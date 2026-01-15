#!/bin/bash
#
# Quick test script - runs batch processing for a single manager
# Usage: ./run_batch_test.sh YOUR_APIFY_TOKEN
#

set -e

if [ -z "$1" ]; then
    echo "Error: Apify token required"
    echo "Usage: $0 YOUR_APIFY_TOKEN"
    exit 1
fi

APIFY_TOKEN="$1"
TEST_MANAGER="סיגמא"  # Test with Sigma first

echo "========================================"
echo "BATCH PROCESSOR TEST"
echo "========================================"
echo "Testing with manager: $TEST_MANAGER"
echo "Output will be in: ./test_batch_output"
echo "========================================"
echo ""

# Activate venv if it exists
if [ -d "/root/mizrahi-venv" ]; then
    echo "Activating virtual environment..."
    source /root/mizrahi-venv/bin/activate
fi

# Run batch processor for single manager
python /root/batch_special_transactions.py \
    --apify-token "$APIFY_TOKEN" \
    --managers "$TEST_MANAGER" \
    --output-dir ./test_batch_output \
    --skip-tase-prices \
    --email elay.g@82labs.io

echo ""
echo "========================================"
echo "TEST COMPLETE"
echo "========================================"
echo "Check the output in: ./test_batch_output"
echo ""
echo "If successful, run all managers with:"
echo "  ./run_batch_all.sh $APIFY_TOKEN"
echo "========================================"
