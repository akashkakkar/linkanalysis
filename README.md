# 🔗 Link Analyzer

Parse WhatsApp/chat exports, categorize every shared link, and generate a professional Excel report — with special attention to Claude/Anthropic content.

## Quick Start

```bash
# Install
pip install -r requirements.txt

# Run
python analyze.py chat_export.txt output.xlsx

# Or use the runner
chmod +x run.sh
./run.sh chat_export.txt
```

**Pipe from stdin:**
```bash
cat chat.txt | python analyze.py - output.xlsx
```

## What It Does

1. **Parses** all URLs with timestamps from WhatsApp chat exports
2. **Categorizes** each link (Claude/Anthropic, AI/ML, Voice AI, Business, Education, etc.)
3. **Flags** Claude-related content with accuracy notes (official vs. community projects)
4. **Scores** relevance 1–5
5. **Generates** a formatted Excel workbook with 3 sheets:

| Sheet | Contents |
|-------|----------|
| **All Links** | Every link with date, category, Claude flag, relevance score |
| **Claude Topics** | Claude-specific links with verification/accuracy notes |
| **Summary** | Stats: total links, date range, category distribution |

## Use with Claude Code

Drop the `CLAUDE.md` file in your project and run:

```bash
cat chat.txt | claude "Analyze all links per CLAUDE.md instructions"
```

## Categories Detected

- Claude/Anthropic (cowork, openclaw, clawdbot, claude code, etc.)
- Voice AI / TTS
- AI Prompting, RAG, AI Agents
- Coding / Dev Tools
- Education / Learning
- Business / Sales, Career / Jobs
- Product Management
- Tech Companies (OpenAI, DeepSeek, Google, etc.)
- And more...

## Testing

```bash
pip install pytest
python -m pytest test_analyze.py -v
```

68 tests covering parsing, categorization, relevance scoring, Excel output, and edge cases.

## Supported Chat Formats

- `DD/MM/YYYY, HH:MM` — standard WhatsApp
- `[DD/MM/YYYY, HH:MM:SS]` — WhatsApp with brackets
- `YYYY-MM-DD HH:MM` — ISO format
