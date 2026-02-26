#!/bin/bash
# Link Analyzer Agent — AI-powered link analysis
# Usage: ./run.sh <chat_export.txt> [output.xlsx]

set -e

FILE="${1:?Usage: ./run.sh <chat_file.txt> [output.xlsx]}"
OUTPUT="${2:-output.xlsx}"

if [ ! -f "$FILE" ]; then
    echo "❌ File not found: $FILE"
    exit 1
fi

if [ -z "$ANTHROPIC_API_KEY" ]; then
    echo "❌ ANTHROPIC_API_KEY not set."
    echo "   export ANTHROPIC_API_KEY='sk-ant-...'"
    exit 1
fi

# Check dependencies
for pkg in anthropic openpyxl; do
    if ! python3 -c "import $pkg" 2>/dev/null; then
        echo "📦 Installing $pkg..."
        pip install $pkg -q
    fi
done

python3 "$(dirname "$0")/analyze.py" "$FILE" "$OUTPUT"
