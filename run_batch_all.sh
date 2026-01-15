#!/bin/bash
#
# Run batch processing for ALL managers and send email
# Usage: ./run_batch_all.sh YOUR_APIFY_TOKEN [GMAIL_USER] [GMAIL_APP_PASSWORD]
#

set -e

if [ -z "$1" ]; then
    echo "Error: Apify token required"
    echo "Usage: $0 YOUR_APIFY_TOKEN [GMAIL_USER] [GMAIL_APP_PASSWORD]"
    echo ""
    echo "Examples:"
    echo "  $0 apify_token_here"
    echo "  $0 apify_token_here your.email@gmail.com app_password_here"
    exit 1
fi

APIFY_TOKEN="$1"
GMAIL_USER="${2:-}"
GMAIL_PASSWORD="${3:-}"

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

# Send email if credentials provided
if [ -n "$GMAIL_USER" ] && [ -n "$GMAIL_PASSWORD" ]; then
    echo "=========================================="
    echo "SENDING RESULTS EMAIL"
    echo "=========================================="
    python /root/send_batch_email.py \
        --batch-dir "$LATEST_BATCH" \
        --gmail-user "$GMAIL_USER" \
        --gmail-app-password "$GMAIL_PASSWORD" \
        --recipient elay.g@82labs.io

    echo ""
    echo "=========================================="
    echo "EMAIL SENT SUCCESSFULLY"
    echo "=========================================="
else
    echo "=========================================="
    echo "Email credentials not provided"
    echo "=========================================="
    echo "To send results via email, run:"
    echo ""
    echo "  python /root/send_batch_email.py \\"
    echo "    --batch-dir \"$LATEST_BATCH\" \\"
    echo "    --gmail-user your.email@gmail.com \\"
    echo "    --gmail-app-password your_app_password \\"
    echo "    --recipient elay.g@82labs.io"
    echo "=========================================="
fi

echo ""
echo "ALL DONE!"
