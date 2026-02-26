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

# Load .env file if it exists (so Python dotenv isn't the only path)
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
if [ -f "$SCRIPT_DIR/.env" ]; then
    export $(grep -v '^#' "$SCRIPT_DIR/.env" | xargs)
fi

if [ -z "$ANTHROPIC_API_KEY" ]; then
    echo "❌ ANTHROPIC_API_KEY not set."
    echo ""
    echo "  Option 1 — .env file (recommended):"
    echo "    Create a .env file in this directory with:"
    echo "    ANTHROPIC_API_KEY=sk-ant-your-key-here"
    echo ""
    echo "  Option 2 — Environment variable:"
    echo "    export ANTHROPIC_API_KEY='sk-ant-your-key-here'"
    echo ""
    echo "  Get your key at: https://console.anthropic.com/settings/keys"
    exit 1
fi

# Check dependencies
for pkg in anthropic openpyxl; do
    if ! python3 -c "import $pkg" 2>/dev/null; then
        echo "📦 Installing $pkg..."
        pip install $pkg -q
    fi
done

python3 "$SCRIPT_DIR/analyze.py" "$FILE" "$OUTPUT"
