"""
Agenda Parser — Extracts session titles from a DOCX workshop agenda.
"""

import re
from io import BytesIO
from typing import List, Dict

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False


# Keywords that typically indicate non-session content
SKIP_KEYWORDS = [
    "break", "lunch", "tea", "coffee", "registration", "arrival",
    "departure", "travel", "networking", "welcome", "open", "close",
    "prayer", "recitation", "photo", "dinner", "breakfast", "recess",
    "refreshment", "logistics", "housekeeping", "introduction of",
    "vote of thanks", "end of", "programme ends", "programme end",
    "closing remarks", "opening remarks", "inaugural", "valedictory",
    "cultural", "entertainment", "visit to", "field visit", "tour",
]

# Minimum character length for a session title to be considered valid
MIN_TITLE_LENGTH = 8


def _is_skip_line(text: str) -> bool:
    """Return True if the line should be skipped (logistics, breaks, etc.)."""
    lower = text.strip().lower()
    if len(lower) < MIN_TITLE_LENGTH:
        return True
    for kw in SKIP_KEYWORDS:
        if kw in lower:
            return True
    return False


def _clean_title(text: str) -> str:
    """Remove leading numbers, bullets, colons from a session title."""
    # Remove leading numbers like "1.", "1)", "Session 1:", etc.
    text = re.sub(r"^(session\s*\d+\s*[:\-–—]?\s*)", "", text, flags=re.IGNORECASE)
    text = re.sub(r"^\d+[\.\):\-–—]\s*", "", text)
    text = re.sub(r"^[•\-–—►▪◆●○]\s*", "", text)
    return text.strip()


def parse_agenda_docx(file_bytes: bytes) -> List[Dict]:
    """
    Parse a DOCX file and extract agenda sessions.

    Args:
        file_bytes: Raw bytes of the uploaded DOCX file.

    Returns:
        List of dicts: [{"session_no": int, "title": str}, ...]
    """
    if not DOCX_AVAILABLE:
        raise ImportError("python-docx is not installed. Run: pip install python-docx")

    doc = Document(BytesIO(file_bytes))
    sessions = []
    seen_titles = set()
    session_counter = 0

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        style_name = para.style.name.lower() if para.style and para.style.name else ""

        # Prioritize heading-level paragraphs
        is_heading = "heading" in style_name
        is_bold = any(run.bold for run in para.runs if run.text.strip())
        is_long_enough = len(text) >= MIN_TITLE_LENGTH

        # Skip if it's clearly a logistics/break line
        if _is_skip_line(text):
            continue

        # Accept headings, bold items, or decently long non-skipped lines
        if is_heading or is_bold or is_long_enough:
            cleaned = _clean_title(text)
            lower_cleaned = cleaned.lower()

            # Skip duplicates
            if lower_cleaned in seen_titles:
                continue

            # Skip single-word generic titles
            words = cleaned.split()
            if len(words) < 2 and not is_heading:
                continue

            session_counter += 1
            seen_titles.add(lower_cleaned)
            sessions.append({
                "session_no": session_counter,
                "title": cleaned,
            })

    # If no structured sessions found, fall back to all non-empty paragraphs
    if not sessions:
        session_counter = 0
        for para in doc.paragraphs:
            text = para.text.strip()
            if not text or _is_skip_line(text):
                continue
            if len(text) < MIN_TITLE_LENGTH:
                continue
            cleaned = _clean_title(text)
            lower_cleaned = cleaned.lower()
            if lower_cleaned not in seen_titles:
                session_counter += 1
                seen_titles.add(lower_cleaned)
                sessions.append({
                    "session_no": session_counter,
                    "title": cleaned,
                })

    return sessions


def parse_agenda_text(raw_text: str) -> List[Dict]:
    """
    Parse plain text agenda (fallback or manual entry).

    Args:
        raw_text: Newline-separated agenda text.

    Returns:
        List of session dicts.
    """
    sessions = []
    seen_titles = set()
    session_counter = 0

    for line in raw_text.splitlines():
        text = line.strip()
        if not text or _is_skip_line(text):
            continue
        cleaned = _clean_title(text)
        if len(cleaned) < MIN_TITLE_LENGTH:
            continue
        lower_cleaned = cleaned.lower()
        if lower_cleaned not in seen_titles:
            session_counter += 1
            seen_titles.add(lower_cleaned)
            sessions.append({
                "session_no": session_counter,
                "title": cleaned,
            })

    return sessions
