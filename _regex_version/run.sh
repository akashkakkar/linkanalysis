#!/bin/bash
set -e
FILE="${1:?Usage: ./run.sh <chat_file.txt> [output.xlsx]}"
OUTPUT="${2:-output.xlsx}"
[ ! -f "$FILE" ] && echo "❌ File not found: $FILE" && exit 1
python3 -c "import openpyxl" 2>/dev/null || pip install openpyxl -q
python3 "$(dirname "$0")/analyze.py" "$FILE" "$OUTPUT"
