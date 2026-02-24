# Link Analyzer Agent

You analyze WhatsApp/chat exports containing shared links and produce a structured Excel report.

## Your Job

1. Parse all URLs with timestamps from the input file
2. Categorize each link by topic (AI, Claude, Voice AI, Business, Education, Health, etc.)
3. Flag Claude/Anthropic-related links specifically
4. Score relevance 1-5 (5=Claude, 4=AI/Tech, 3=Business, 2=Social, 1=Utility)
5. Output a professional Excel workbook with:
   - **Sheet 1 — All Links**: S.No, Date, Link (hyperlinked), Topic, Claude (Yes/No), Relevance
   - **Sheet 2 — Claude Topics**: Claude-specific links with accuracy/verification notes
   - **Sheet 3 — Summary**: Date range, total links, category distribution

## How to Run

```bash
python analyze.py <chat_export.txt> [output.xlsx]
cat chat.txt | python analyze.py - [output.xlsx]
./run.sh <chat_export.txt>
```

## Key Rules

- Use `openpyxl` for Excel (install via `pip install openpyxl`)
- Claude keywords: claude, anthropic, cowork, openclaw, clawdbot, moltbot, sonnet, opus, haiku
- **Community vs. Official**: openclaw, clawdbot, zeroclaw, nanoclaw, picoclaw are community projects — NOT official Anthropic products. Always flag this.
- Links must be clickable hyperlinks in Excel
- Claude rows highlighted green (#E2EFDA)
- Headers in dark blue (#2F5496) with white text
- Always add auto-filter and freeze panes

## Testing

```bash
pip install pytest openpyxl
python -m pytest test_analyze.py -v
```
