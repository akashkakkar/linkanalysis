#!/bin/bash
# Link Analyzer — Quick runner
# Usage: ./run.sh <chat_export.txt> [output.xlsx]

set -e

FILE="${1:?Usage: ./run.sh <chat_file.txt> [output.xlsx]}"
OUTPUT="${2:-output.xlsx}"

if [ ! -f "$FILE" ]; then
    echo "❌ File not found: $FILE"
    exit 1
fi

# Check dependencies
if ! python3 -c "import openpyxl" 2>/dev/null; then
    echo "📦 Installing openpyxl..."
    pip install openpyxl -q
fi

python3 "$(dirname "$0")/analyze.py" "$FILE" "$OUTPUT"
