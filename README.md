# 🔗 Link Analyzer Agent

AI-powered WhatsApp chat link analysis using **Claude API**. Parses shared links, sends them to Claude for intelligent categorization, and generates a professional Excel report.

## Why AI?

Unlike keyword-matching scripts, this agent uses Claude to **understand context**:
- A URL with "code" in a Claude context → Claude/Anthropic, not Coding
- A Stanford link about AI safety → Education, not just "stanford keyword match"
- A LinkedIn post about "hiring AI agents" → Career/Jobs, not AI Agents
- Generates human-readable topic summaries for every link

## Quick Start

```bash
# Install
pip install -r requirements.txt

# Set API key
export ANTHROPIC_API_KEY='sk-ant-...'

# Run
python analyze.py chat_export.txt output.xlsx

# Or use the runner
chmod +x run.sh
./run.sh chat_export.txt
```

## How It Works

```
WhatsApp .txt → [Regex Parser] → URLs + dates
              → [Claude API ×N batches] → categories, scores, summaries
              → [openpyxl] → formatted Excel workbook
```

1. **Parse** — Regex extracts URLs and timestamps (deterministic, fast)
2. **Analyze** — URLs sent to Claude Sonnet in batches of 30 for categorization
3. **Generate** — Results written to a 3-sheet Excel workbook

## Output

| Sheet | Contents |
|-------|----------|
| **All Links** | S.No, Date, Link, Category, **AI Topic Summary**, Claude flag, Relevance |
| **Claude Topics** | Claude-specific links with accuracy/verification notes |
| **Summary** | Stats, model used, category distribution |

## Cost

~459 links = ~16 API calls × ~2K tokens each ≈ **$0.10-0.15** per run (Sonnet pricing).

## Testing

```bash
python -m pytest test_analyze.py -v
```

All tests use **mocked API calls** — no real API key or charges needed for testing.

## Configuration

| Variable | Default | Description |
|----------|---------|-------------|
| `ANTHROPIC_API_KEY` | (required) | Your Anthropic API key |
| `MODEL` | `claude-sonnet-4-20250514` | Model to use (edit in analyze.py) |
| `BATCH_SIZE` | 30 | Links per API call |
| `MAX_RETRIES` | 3 | Retry attempts on failure |

## Supported Chat Formats

- `DD/MM/YYYY, HH:MM` — standard WhatsApp
- `[DD/MM/YYYY, HH:MM:SS]` — WhatsApp with brackets
- `YYYY-MM-DD HH:MM` — ISO format
