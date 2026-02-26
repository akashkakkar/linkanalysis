# Link Analyzer Agent

Two modes for analyzing WhatsApp/chat link exports:

- `analyze_ai.py` — Claude API (smarter, needs API key)
- `analyze_regex.py` — keyword matching (instant, offline)

## How to Run

```bash
./run.sh chat.txt              # AI mode (default)
./run.sh --regex chat.txt      # Regex mode (offline)
```

## Testing

```bash
python -m pytest test_analyze_ai.py test_analyze_regex.py -v
```
