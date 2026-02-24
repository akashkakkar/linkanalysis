#!/usr/bin/env python3
"""
Tests for Link Analyzer.
Run: python -m pytest test_analyze.py -v
"""

import os
import tempfile
import pytest
from openpyxl import load_workbook

from analyze import (
    parse_date,
    extract_links,
    categorize_url,
    is_claude_related,
    get_relevance_score,
    get_claude_accuracy_note,
    analyze,
    generate_excel,
)


# ===========================================================================
# Test Data
# ===========================================================================

SAMPLE_CHAT = """2/1/2026, 08:15 - Akash: Check this out https://www.linkedin.com/posts/anthropic_claude-code-tips-activity-123456
2/1/2026, 09:30 - Ravi: https://www.linkedin.com/posts/johndoe_ai-agents-activity-789012
2/1/2026, 10:00 - Akash: Great article https://www.linkedin.com/posts/janedoe_voice-tts-elevenlabs-activity-345678
3/1/2026, 11:45 - Priya: https://www.youtube.com/watch?v=abc123
3/1/2026, 14:20 - Akash: https://www.linkedin.com/posts/techguy_openclaw-claude-community-activity-456789
4/1/2026, 08:00 - Ravi: https://www.linkedin.com/posts/someone_deepseek-qwen-activity-112233
4/1/2026, 09:15 - Akash: https://www.linkedin.com/posts/dev_cursor-vscode-coding-activity-223344
5/1/2026, 10:30 - Priya: https://www.linkedin.com/posts/founder_startup-saas-revenue-activity-334455
5/1/2026, 12:00 - Akash: https://www.linkedin.com/posts/pm_productmanagement-roadmap-activity-445566
5/1/2026, 16:00 - Ravi: https://www.facebook.com/reel/some-video
6/1/2026, 07:00 - Akash: https://www.linkedin.com/posts/aidev_cowork-anthropic-activity-556677
6/1/2026, 08:30 - Priya: https://www.linkedin.com/posts/scholar_stanford-course-tutorial-activity-667788
6/1/2026, 11:00 - Akash: https://www.linkedin.com/posts/hr_hiring-remote-interview-activity-778899
7/1/2026, 09:00 - Ravi: https://www.linkedin.com/posts/coder_moltbot-automation-activity-889900
7/1/2026, 14:00 - Akash: https://mahadiscom.in/bill-payment
7/1/2026, 15:30 - Priya: https://www.linkedin.com/posts/datascientist_rag-vectordb-embedding-activity-990011
"""

WHATSAPP_BRACKET_FORMAT = """[02/01/2026, 08:15:30] Akash: Check this https://www.linkedin.com/posts/user_claude-sonnet-activity-111
[02/01/2026, 09:00:15] Ravi: https://www.linkedin.com/posts/user_openai-gpt-activity-222
"""


# ===========================================================================
# Parsing Tests
# ===========================================================================

class TestParseDate:
    def test_dd_mm_yyyy(self):
        assert parse_date("02/01/2026") == "2026-01-02"

    def test_dd_mm_yy(self):
        assert parse_date("02/01/26") == "2026-01-02"

    def test_yyyy_mm_dd(self):
        assert parse_date("2026-01-02") == "2026-01-02"

    def test_with_trailing_comma(self):
        assert parse_date("02/01/2026,") == "2026-01-02"

    def test_with_brackets(self):
        assert parse_date("[02/01/2026]") == "2026-01-02"

    def test_unparseable_returns_as_is(self):
        assert parse_date("not-a-date") == "not-a-date"


class TestExtractLinks:
    def test_basic_extraction(self):
        links = extract_links(SAMPLE_CHAT)
        assert len(links) == 16

    def test_dates_extracted(self):
        links = extract_links(SAMPLE_CHAT)
        # First link should have date 2026-01-02 (DD/MM/YYYY -> YYYY-MM-DD)
        assert links[0]["date"] == "2026-01-02"

    def test_urls_are_valid(self):
        links = extract_links(SAMPLE_CHAT)
        for link in links:
            assert link["url"].startswith("http")

    def test_no_trailing_punctuation(self):
        text = "2/1/2026, 08:00 - User: visit https://example.com/page. And also this!"
        links = extract_links(text)
        assert links[0]["url"] == "https://example.com/page"

    def test_empty_input(self):
        assert extract_links("") == []

    def test_no_urls(self):
        assert extract_links("2/1/2026, 08:00 - User: Hello world!") == []

    def test_bracket_whatsapp_format(self):
        links = extract_links(WHATSAPP_BRACKET_FORMAT)
        assert len(links) == 2
        assert links[0]["date"] == "2026-01-02"

    def test_multiple_urls_on_one_line(self):
        text = "2/1/2026, 08:00 - User: https://a.com https://b.com"
        links = extract_links(text)
        assert len(links) == 2

    def test_unknown_date_when_none_found(self):
        text = "No date here: https://example.com"
        links = extract_links(text)
        assert links[0]["date"] == "Unknown"


# ===========================================================================
# Categorization Tests
# ===========================================================================

class TestCategorizeUrl:
    def test_claude_anthropic(self):
        assert categorize_url("https://linkedin.com/posts/user_claude-code-tips-activity-123") == "Claude/Anthropic"

    def test_anthropic_keyword(self):
        assert categorize_url("https://linkedin.com/posts/user_anthropic-release-activity-123") == "Claude/Anthropic"

    def test_cowork(self):
        assert categorize_url("https://linkedin.com/posts/user_cowork-anthropic-activity-123") == "Claude/Anthropic"

    def test_openclaw_community(self):
        assert categorize_url("https://linkedin.com/posts/user_openclaw-tool-activity-123") == "Claude/Anthropic"

    def test_voice_ai(self):
        assert categorize_url("https://linkedin.com/posts/user_voice-tts-elevenlabs-activity-123") == "Voice AI / TTS"

    def test_youtube(self):
        assert categorize_url("https://www.youtube.com/watch?v=abc123") == "YouTube"

    def test_facebook(self):
        assert categorize_url("https://www.facebook.com/reel/123") == "Facebook"

    def test_coding_not_claude(self):
        """Coding URL should NOT be categorized as Claude even if it has 'code'."""
        assert categorize_url("https://linkedin.com/posts/user_cursor-vscode-coding-activity-123") == "Coding / Dev Tools"

    def test_claude_code_is_claude(self):
        """'claude-code' should be Claude, not Coding."""
        assert categorize_url("https://linkedin.com/posts/user_claude-code-tips-activity-123") == "Claude/Anthropic"

    def test_deepseek(self):
        assert categorize_url("https://linkedin.com/posts/user_deepseek-qwen-activity-123") == "Tech Companies"

    def test_business(self):
        assert categorize_url("https://linkedin.com/posts/user_startup-saas-revenue-activity-123") == "Business / Sales"

    def test_product_management(self):
        assert categorize_url("https://linkedin.com/posts/user_productmanagement-roadmap-activity-123") == "Product Management"

    def test_education(self):
        assert categorize_url("https://linkedin.com/posts/user_stanford-course-activity-123") == "Education / Learning"

    def test_career(self):
        assert categorize_url("https://linkedin.com/posts/user_hiring-remote-interview-activity-123") == "Career / Jobs"

    def test_rag(self):
        assert categorize_url("https://linkedin.com/posts/user_rag-vectordb-embedding-activity-123") == "RAG / Retrieval"

    def test_generic_linkedin(self):
        assert categorize_url("https://linkedin.com/posts/user_family-vacation-photos-activity-123") == "LinkedIn Post - General"

    def test_moltbot_is_claude(self):
        assert categorize_url("https://linkedin.com/posts/user_moltbot-automation-activity-123") == "Claude/Anthropic"

    def test_other_domain(self):
        assert categorize_url("https://mahadiscom.in/bill-payment") == "Other"


class TestIsClaudeRelated:
    def test_claude_category(self):
        assert is_claude_related("https://example.com", "Claude/Anthropic") is True

    def test_non_claude_category(self):
        assert is_claude_related("https://example.com/something", "Coding / Dev Tools") is False

    def test_claude_keyword_in_url(self):
        assert is_claude_related("https://example.com/claude-tips", "Other") is True


class TestRelevanceScore:
    def test_claude_is_5(self):
        assert get_relevance_score("Claude/Anthropic", "https://example.com") == 5

    def test_ai_is_4(self):
        assert get_relevance_score("AI / ML General", "https://example.com") == 4

    def test_business_is_3(self):
        assert get_relevance_score("Business / Sales", "https://example.com") == 3

    def test_youtube_is_2(self):
        assert get_relevance_score("YouTube", "https://youtube.com") == 2

    def test_bill_payment_is_1(self):
        assert get_relevance_score("Other", "https://mahadiscom.in/bill-payment") == 1

    def test_referral_is_1(self):
        assert get_relevance_score("Other", "https://porter.in/referral/abc") == 1


class TestClaudeAccuracyNote:
    def test_cowork_note(self):
        note = get_claude_accuracy_note("https://linkedin.com/posts/user_cowork-tips")
        assert "Anthropic product" in note

    def test_openclaw_note(self):
        note = get_claude_accuracy_note("https://linkedin.com/posts/user_openclaw-tool")
        assert "NOT official" in note

    def test_moltbot_note(self):
        note = get_claude_accuracy_note("https://linkedin.com/posts/user_moltbot-tool")
        assert "NOT" in note or "Third-party" in note

    def test_claude_code_note(self):
        note = get_claude_accuracy_note("https://linkedin.com/posts/user_claudecode-tips")
        assert "Official" in note or "CLI" in note

    def test_default_note(self):
        note = get_claude_accuracy_note("https://linkedin.com/posts/user_claude-general")
        assert "docs.anthropic.com" in note


# ===========================================================================
# Pipeline Tests
# ===========================================================================

class TestAnalyzePipeline:
    def test_full_pipeline(self):
        results = analyze(SAMPLE_CHAT)
        assert len(results) == 16

    def test_results_have_all_fields(self):
        results = analyze(SAMPLE_CHAT)
        required_fields = {"date", "url", "category", "claude", "relevance", "accuracy_note"}
        for r in results:
            assert set(r.keys()) == required_fields

    def test_claude_links_detected(self):
        results = analyze(SAMPLE_CHAT)
        claude_links = [r for r in results if r["claude"]]
        # Expected Claude links: claude-code, openclaw-claude, cowork-anthropic, moltbot
        assert len(claude_links) >= 4

    def test_relevance_range(self):
        results = analyze(SAMPLE_CHAT)
        for r in results:
            assert 1 <= r["relevance"] <= 5

    def test_accuracy_notes_only_for_claude(self):
        results = analyze(SAMPLE_CHAT)
        for r in results:
            if not r["claude"]:
                assert r["accuracy_note"] == ""
            else:
                assert r["accuracy_note"] != ""


# ===========================================================================
# Excel Output Tests
# ===========================================================================

class TestExcelGeneration:
    @pytest.fixture
    def excel_output(self):
        """Generate Excel from sample data and return path."""
        results = analyze(SAMPLE_CHAT)
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            path = f.name
        generate_excel(results, path)
        yield path
        os.unlink(path)

    def test_file_created(self, excel_output):
        assert os.path.isfile(excel_output)
        assert os.path.getsize(excel_output) > 0

    def test_three_sheets(self, excel_output):
        wb = load_workbook(excel_output)
        assert len(wb.sheetnames) == 3
        assert wb.sheetnames == ["All Links", "Claude Topics", "Summary"]

    def test_all_links_sheet_row_count(self, excel_output):
        wb = load_workbook(excel_output)
        ws = wb["All Links"]
        # 16 data rows + 1 header
        assert ws.max_row == 17

    def test_all_links_headers(self, excel_output):
        wb = load_workbook(excel_output)
        ws = wb["All Links"]
        headers = [ws.cell(row=1, column=c).value for c in range(1, 7)]
        assert headers == ["S.No.", "Date", "Link", "Topic/Category", "Claude Related", "Relevance (1-5)"]

    def test_claude_sheet_has_rows(self, excel_output):
        wb = load_workbook(excel_output)
        ws = wb["Claude Topics"]
        assert ws.max_row > 1  # at least header + 1 data row

    def test_claude_sheet_headers(self, excel_output):
        wb = load_workbook(excel_output)
        ws = wb["Claude Topics"]
        headers = [ws.cell(row=1, column=c).value for c in range(1, 6)]
        assert headers == ["S.No.", "Date", "Link", "Sub-Topic", "Accuracy Note"]

    def test_summary_sheet_has_stats(self, excel_output):
        wb = load_workbook(excel_output)
        ws = wb["Summary"]
        assert ws.cell(row=1, column=1).value == "Metric"
        assert ws.cell(row=2, column=1).value == "Total Links"
        assert ws.cell(row=2, column=2).value == 16

    def test_hyperlinks_present(self, excel_output):
        wb = load_workbook(excel_output)
        ws = wb["All Links"]
        # Check first data row has a hyperlink
        cell = ws.cell(row=2, column=3)
        assert cell.hyperlink is not None or cell.value.startswith("http")

    def test_serial_numbers_sequential(self, excel_output):
        wb = load_workbook(excel_output)
        ws = wb["All Links"]
        for i in range(1, 17):
            assert ws.cell(row=i + 1, column=1).value == i

    def test_freeze_panes(self, excel_output):
        wb = load_workbook(excel_output)
        ws = wb["All Links"]
        assert ws.freeze_panes == "A2"


# ===========================================================================
# Edge Cases
# ===========================================================================

class TestEdgeCases:
    def test_empty_input(self):
        results = analyze("")
        assert results == []

    def test_no_urls_in_text(self):
        results = analyze("Just some text without URLs\nAnother line")
        assert results == []

    def test_single_url(self):
        text = "2/1/2026, 08:00 - User: https://linkedin.com/posts/user_claude-tips-activity-1"
        results = analyze(text)
        assert len(results) == 1
        assert results[0]["claude"] is True

    def test_non_linkedin_urls(self):
        text = "2/1/2026, 08:00 - User: https://github.com/anthropics/claude-code"
        results = analyze(text)
        assert len(results) == 1
        assert results[0]["claude"] is True

    def test_unicode_in_text(self):
        text = "2/1/2026, 08:00 - User: \U0001f680 https://linkedin.com/posts/user_ai-activity-1"
        results = analyze(text)
        assert len(results) == 1

    def test_excel_with_empty_results(self):
        """Excel should still generate with no data."""
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            path = f.name
        generate_excel([], path)
        assert os.path.isfile(path)
        wb = load_workbook(path)
        assert len(wb.sheetnames) == 3
        os.unlink(path)


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
