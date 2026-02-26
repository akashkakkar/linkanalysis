# Link Analyzer Agent

AI-powered analysis of WhatsApp/chat link exports using Claude API.

## Your Job

1. Parse all URLs with timestamps from the input file
2. Send URLs in batches to Claude API for intelligent categorization
3. Claude analyzes URL structure, path, slugs to determine:
   - Category (from 23 predefined categories)
   - Whether it's Claude/Anthropic related
   - Relevance score 1-5
   - Accuracy/verification notes for Claude content
   - Brief topic summary
4. Output a professional Excel workbook with 3 sheets

## How to Run

```bash
export ANTHROPIC_API_KEY='sk-ant-...'
python analyze.py <chat_export.txt> [output.xlsx]
./run.sh <chat_export.txt>
```

## Architecture

```
chat.txt → [Regex Parser] → raw URLs with dates
         → [Claude API batches] → categories, scores, summaries, notes
         → [openpyxl] → formatted Excel
```

- Parsing is deterministic (regex) — no AI needed for extracting URLs
- Categorization is AI-powered (Claude API) — understands context, not just keywords
- Excel generation is templated with professional formatting

## Key Rules

- Requires `ANTHROPIC_API_KEY` environment variable
- Uses `claude-sonnet-4-20250514` model
- Batches 30 links per API call to balance cost and quality
- Retries up to 3x on failures with exponential backoff
- Community vs. Official distinction: openclaw, clawdbot = community, NOT Anthropic official

## Testing

```bash
pip install pytest anthropic openpyxl
python -m pytest test_analyze.py -v
```

Tests use mocked API calls — no real API key needed for testing.
