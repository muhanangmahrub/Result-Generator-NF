"""Unit tests for header/value normalisation."""

from __future__ import annotations

import datetime as dt

import pytest

from rtg.normalize import cell_text, normalize_header, slugify, tokens_of


@pytest.mark.parametrize(
    "raw, expected",
    [
        ("Test Case ID", "test case id"),
        ("  Test   Case  ID  ", "test case id"),
        ("Test Case\nID", "test case id"),
        ("TEST CASE ID", "test case id"),
        ("Test_Case_ID", "test case id"),
        ("Test-Case-ID", "test case id"),
        ("Test Case ID:", "test case id"),
        ("Pass/Fail (P/F)", "pass fail pf"),
        ("Pass / Fail", "pass fail"),
        ("Positive/Negative (P/N)", "positive negative pn"),
        ("Pre-Condition", "pre condition"),
        ("Test Executed by", "test executed by"),
        ("TC ID", "tc id"),
        ("Test Case", "test case"),
        ("", ""),
        (None, ""),
        (123, "123"),
    ],
)
def test_normalize_header(raw, expected):
    assert normalize_header(raw) == expected


def test_normalize_header_merges_single_letter_runs():
    # Abbreviations split by punctuation must survive as one token.
    assert normalize_header("(P/F)") == "pf"
    assert normalize_header("U.A.T Result") == "uat result"


def test_tokens_of():
    assert tokens_of("test case id") == ("test", "case", "id")
    assert tokens_of("") == ()


@pytest.mark.parametrize(
    "raw, expected",
    [
        (None, ""),
        ("  text  ", "text"),
        ("line1\r\nline2", "line1\nline2"),
        (1.0, "1"),
        (1.5, "1.5"),
        (7, "7"),
        (True, "TRUE"),
        (False, "FALSE"),
        (dt.datetime(2026, 10, 1), "2026-10-01"),
        (dt.datetime(2026, 10, 1, 14, 30), "2026-10-01 14:30"),
        (dt.date(2026, 10, 1), "2026-10-01"),
        (dt.time(9, 5), "09:05"),
    ],
)
def test_cell_text(raw, expected):
    assert cell_text(raw) == expected


@pytest.mark.parametrize(
    "raw, expected",
    [
        ("Move Partner Onboarding File via Sharefolder",
         "move-partner-onboarding-file-via-sharefolder"),
        ("AMF Segmentation", "amf-segmentation"),
        ("Validasi  Data/Input", "validasi-data-input"),
        ("1.1", "1-1"),
        ("", "untitled"),
        ("###", "untitled"),
    ],
)
def test_slugify(raw, expected):
    assert slugify(raw) == expected


def test_slugify_truncates():
    assert len(slugify("a " * 200, max_length=40)) <= 40
