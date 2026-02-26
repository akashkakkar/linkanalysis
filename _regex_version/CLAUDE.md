# Link Analyzer Agent (Regex Version)

Keyword-based link analysis. No API key needed. Fast and offline.

## How to Run

```bash
python analyze.py <chat_export.txt> [output.xlsx]
./run.sh <chat_export.txt>
```

## Key Difference from main branch

This version uses regex/keyword matching on URL slugs instead of Claude API calls.
Faster and free, but less accurate on ambiguous URLs.
