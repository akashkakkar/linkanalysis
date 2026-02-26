# 🔗 Link Analyzer (Regex Version)

Fast, offline WhatsApp link analysis using keyword matching. **No API key needed.**

For the AI-powered version that uses Claude API, switch to the `main` branch.

## Quick Start

```bash
pip install openpyxl
python analyze.py chat_export.txt output.xlsx
```

## How It Works

Uses URL keyword matching to categorize links. Fast and free, but less accurate than the AI version on ambiguous URLs.

## Testing

```bash
pip install pytest openpyxl
python -m pytest test_analyze.py -v  # 68 tests
```
