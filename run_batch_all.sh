#!/bin/bash
#
# Run batch processing for ALL managers
# Emails are sent automatically per-manager during processing
# Usage: ./run_batch_all.sh YOUR_APIFY_TOKEN
#
# Requirements:
#   - RESEND_API_KEY environment variable must be set for email sending
#

set -e

if [ -z "$1" ]; then
    echo "Error: Apify token required"
    echo "Usage: $0 YOUR_APIFY_TOKEN"
    echo ""
    echo "Examples:"
    echo "  $0 apify_token_here"
    echo ""
    echo "Note: Set RESEND_API_KEY environment variable for email sending"
    echo "  export RESEND_API_KEY=re_xxxxxxxxxxxxxxxxxxxxxxxxxxxx"
    exit 1
fi

APIFY_TOKEN="$1"

echo "=========================================="
echo "BATCH PROCESSOR - ALL MANAGERS"
echo "=========================================="
echo "Processing all 10 fund managers"
echo "Output will be in: ./batch_output"
echo "=========================================="
echo ""

# Activate venv if it exists
if [ -d "/root/mizrahi-venv" ]; then
    echo "Activating virtual environment..."
    source /root/mizrahi-venv/bin/activate
fi

# Run batch processor for all managers
echo "Starting batch processing..."
python /root/batch_special_transactions.py \
    --apify-token "$APIFY_TOKEN" \
    --output-dir ./batch_output \
    --skip-tase-prices \
    --price-threshold 5.0 \
    --email elay.g@82labs.io

# Get the latest batch directory
LATEST_BATCH=$(ls -td ./batch_output/*/ 2>/dev/null | head -1)

if [ -z "$LATEST_BATCH" ]; then
    echo "Error: No batch output directory found"
    exit 1
fi

echo ""
echo "=========================================="
echo "BATCH PROCESSING COMPLETE"
echo "=========================================="
echo "Output directory: $LATEST_BATCH"
echo ""
echo "Note: Emails were sent automatically during processing"
echo "      (2 emails per manager if RESEND_API_KEY is set)"
echo "=========================================="
echo ""
echo "ALL DONE!"
