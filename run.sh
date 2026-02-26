#!/bin/bash
# Link Analyzer — unified runner
# Usage:
#   ./run.sh <chat_file.txt> [output.xlsx]          ← uses AI (default)
#   ./run.sh --regex <chat_file.txt> [output.xlsx]   ← uses regex (offline)

set -e
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
MODE="ai"

# Parse flags
if [ "$1" = "--regex" ] || [ "$1" = "-r" ]; then
    MODE="regex"
    shift
fi

FILE="${1:?Usage: ./run.sh [--regex] <chat_file.txt> [output.xlsx]}"
OUTPUT="${2:-output.xlsx}"

[ ! -f "$FILE" ] && echo "❌ File not found: $FILE" && exit 1

if [ "$MODE" = "ai" ]; then
    # Load .env if present
    if [ -f "$SCRIPT_DIR/.env" ]; then
        export $(grep -v '^#' "$SCRIPT_DIR/.env" | xargs)
    fi

    if [ -z "$ANTHROPIC_API_KEY" ]; then
        echo "❌ ANTHROPIC_API_KEY not set."
        echo ""
        echo "  Option A — Add your key:  echo 'ANTHROPIC_API_KEY=sk-ant-...' > .env"
        echo "  Option B — Use offline:   ./run.sh --regex $FILE"
        echo ""
        echo "  Get a key at: https://console.anthropic.com/settings/keys"
        exit 1
    fi

    for pkg in anthropic openpyxl; do
        python3 -c "import $pkg" 2>/dev/null || pip install $pkg -q
    done
    python3 "$SCRIPT_DIR/analyze_ai.py" "$FILE" "$OUTPUT"
else
    python3 -c "import openpyxl" 2>/dev/null || pip install openpyxl -q
    python3 "$SCRIPT_DIR/analyze_regex.py" "$FILE" "$OUTPUT"
fi
