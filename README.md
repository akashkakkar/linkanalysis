# 🔗 Link Analyzer

Parse WhatsApp chat exports, categorize every shared link, and generate a professional Excel report. Comes in two modes — pick what fits.

## Which Mode Should I Use?

| | **AI Mode** (`analyze_ai.py`) | **Regex Mode** (`analyze_regex.py`) |
|---|---|---|
| **How it works** | Sends URLs to Claude API in batches — Claude reads the URL structure and uses judgment to categorize | Matches keywords in URL slugs against a rule list |
| **Accuracy** | Higher — understands context (e.g. "hiring AI agents" → Career, not AI Agents) | Good for clear-cut URLs, misses ambiguous ones |
| **Extra output** | AI-generated topic summary per link | No summaries |
| **Speed** | ~30 sec for 450 links | Instant |
| **Cost** | ~$0.10–0.15 per run (Claude Sonnet) | Free |
| **Requires** | Anthropic API key | Nothing beyond Python + openpyxl |
| **Works offline** | No | Yes |

**Rule of thumb:** Use AI mode when accuracy matters. Use regex mode for quick-and-dirty runs, offline use, or if you don't have an API key.

## Quick Start

### AI Mode (default)

```bash
pip install -r requirements.txt

# Add your API key (one-time)
echo 'ANTHROPIC_API_KEY=sk-ant-your-key-here' > .env

# Run
./run.sh chat.txt
```

Get your API key at: https://console.anthropic.com/settings/keys

### Regex Mode (offline, free)

```bash
pip install openpyxl

./run.sh --regex chat.txt
```

### Direct Python

```bash
python analyze_ai.py chat.txt output.xlsx        # AI mode
python analyze_regex.py chat.txt output.xlsx      # regex mode
cat chat.txt | python analyze_ai.py - output.xlsx # pipe from stdin
```

## Output

Both modes produce a 3-sheet Excel workbook:

| Sheet | Contents |
|-------|----------|
| **All Links** | S.No, Date, Link, Category, Claude flag, Relevance (+ Topic Summary in AI mode) |
| **Claude Topics** | Claude-specific links with accuracy/verification notes |
| **Summary** | Stats: total links, date range, category distribution |

## Categories Detected

Claude/Anthropic, Voice AI / TTS, AI Prompting, RAG / Retrieval, Coding / Dev Tools, AI Agents, Education / Learning, Product Management, Business / Sales, Career / Jobs, Tech Companies, Open Source, Health / Wellness, Finance, India / Culture, AI / ML General, Robotics, YouTube, Facebook, Instagram, Twitter/X

## Testing

```bash
pip install pytest
python -m pytest test_analyze_ai.py -v     # 45 tests (mocked, no API key needed)
python -m pytest test_analyze_regex.py -v   # 40 tests
```

## Supported Chat Formats

- `DD/MM/YYYY, HH:MM` — standard WhatsApp
- `[DD/MM/YYYY, HH:MM:SS]` — WhatsApp with brackets
- `YYYY-MM-DD HH:MM` — ISO format

## File Structure

```
├── run.sh                 # Unified runner (./run.sh or ./run.sh --regex)
├── analyze_ai.py          # AI mode — Claude API
├── analyze_regex.py       # Regex mode — keyword matching
├── test_analyze_ai.py     # AI tests (mocked)
├── test_analyze_regex.py  # Regex tests
├── .env.example           # API key template
├── requirements.txt
├── CLAUDE.md              # Claude Code agent instructions
└── README.md
```
