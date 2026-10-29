"""Text normalisation helpers shared by the parser and the config loader.

The parser never looks at column letters or positions; it matches *normalised*
header text against an alias list. All normalisation rules live here so that the
config file and the workbook are normalised exactly the same way.
"""

from __future__ import annotations

import datetime as _dt
import re
import unicodedata
from typing import Any, List, Tuple

__all__ = [
    "cell_text",
    "normalize_header",
    "normalize_label",
    "tokens_of",
    "slugify",
]

_WS_RE = re.compile(r"\s+")
_SLUG_SEP_RE = re.compile(r"[^a-z0-9]+")


def cell_text(value: Any) -> str:
    """Convert a raw openpyxl cell value into clean display text.

    Keeps internal newlines (test steps are frequently multi-line) but strips
    surrounding whitespace, renders dates predictably and avoids the ``1.0``
    float artefacts openpyxl produces for integer-looking cells.
    """
    if value is None:
        return ""
    if isinstance(value, str):
        return value.replace("\r\n", "\n").replace("\xa0", " ").strip()
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, _dt.datetime):
        if value.time() == _dt.time(0, 0):
            return value.strftime("%Y-%m-%d")
        return value.strftime("%Y-%m-%d %H:%M")
    if isinstance(value, _dt.date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, _dt.time):
        return value.strftime("%H:%M")
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return repr(value).rstrip("0").rstrip(".")
    return str(value).strip()


def normalize_header(value: Any) -> str:
    """Normalise a header cell so it can be compared against an alias.

    Rules: lowercase, Unicode-normalise (so non-breaking spaces become spaces),
    replace every non-alphanumeric character with a space, collapse whitespace,
    then glue runs of single characters back together so that abbreviations
    survive.

    >>> normalize_header("Pass/Fail (P/F)")
    'pass fail pf'
    >>> normalize_header("Test  Case\\nID")
    'test case id'
    """
    text = cell_text(value)
    if not text:
        return ""
    text = unicodedata.normalize("NFKC", text).lower()
    text = "".join(ch if ch.isalnum() else " " for ch in text)
    parts = text.split()

    merged: List[str] = []
    run: List[str] = []
    for part in parts:
        if len(part) == 1:
            run.append(part)
            continue
        if run:
            merged.append("".join(run))
            run = []
        merged.append(part)
    if run:
        merged.append("".join(run))
    return " ".join(merged)


def normalize_label(value: Any) -> str:
    """Normalise a metadata label.

    Same as :func:`normalize_header`; kept as a separate name because metadata
    labels are matched against a different alias namespace and the rules may
    diverge later.
    """
    return normalize_header(value)


def tokens_of(normalized: str) -> Tuple[str, ...]:
    """Split an already-normalised string into its tokens."""
    return tuple(normalized.split())


def slugify(value: Any, *, fallback: str = "untitled", max_length: int = 60) -> str:
    """Build a filesystem-safe slug (used for screenshot folder names)."""
    text = cell_text(value)
    text = unicodedata.normalize("NFKD", text)
    text = text.encode("ascii", "ignore").decode("ascii").lower()
    slug = _SLUG_SEP_RE.sub("-", text).strip("-")
    if max_length and len(slug) > max_length:
        slug = slug[:max_length].rstrip("-")
    return slug or fallback
