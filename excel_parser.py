"""
Excel parsing for the Name Board Generator — simplified.

Rule of thumb: find the person's NAME, find the rest of their descriptive
text (title / designation / organization — whatever it's called), and
ignore everything else on the sheet (section dividers, blank rows, serial
number columns, etc).

Specifically:
  - One column is identified as the Name column (by header word if a
    header row exists, otherwise the first column).
  - ALL other populated columns for that row are concatenated together
    (in sheet order) to form a single "details" block. This becomes the
    Title+Company text the board generator lays out — it decides on its
    own whether that fits on one line or needs to split into two visually,
    no comma-splitting or guessing about which part is "title" vs
    "company" is attempted here.
  - If a row's Name cell looks like a section header (e.g. "DIGNITARIES",
    "SPEAKERS", "QUIZ") and there's no other descriptive text in that row,
    it's skipped rather than turned into a blank name board.
  - Fully blank rows are skipped.

Returns a list of dicts: [{"name": ..., "title": ..., "company": ...}, ...]
("title" is always left empty; the full details text goes into "company",
since the board generator displays Title and Company as one combined
visual block whenever only one of them is present.) Plus a human-readable
note describing what was found, for display in the app.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

import pandas as pd

NAME_SYNONYMS = {
    "name", "names", "fullname", "full name", "dignitary", "dignitary name",
    "speaker", "speaker name", "guest", "guest name", "person", "attendee",
}

# If exactly two detail columns remain (besides Name) and one of them has a
# header matching this list, we trust that the sheet genuinely separates
# Title from Company/Organization, and keep them as two distinct fields
# (rendered as two stacked lines) instead of merging into one block.
TITLE_SYNONYMS = {
    "title", "designation", "position", "role", "post",
}

# Columns that are clearly NOT part of the person's details and should be
# ignored entirely (not folded into the details text), e.g. serial numbers
# or contact info that isn't shown on a name board.
IGNORE_SYNONYMS = {
    "sr", "sr no", "sr no.", "s no", "s no.", "sl no", "sl no.", "no",
    "no.", "#", "serial number", "index", "id",
    "email", "e-mail", "email id", "e-mail id", "email address",
    "phone", "phone no", "phone number", "mobile", "mobile no",
    "mobile number", "contact", "contact no", "contact number",
}


def _normalize(value) -> str:
    return re.sub(r"\s+", " ", str(value).strip().lower())


def _clean_text(value) -> str:
    """Stringify a cell, collapse embedded newlines/extra whitespace, strip."""
    text = str(value).strip()
    if text.lower() == "nan":
        return ""
    text = re.sub(r"\s*\n\s*", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


_SECTION_HEADER_WORDS = {
    "dignitaries", "dignitary", "speakers", "speaker", "guests", "guest",
    "quiz", "workshop conductors", "workshop", "panelists", "panelist",
    "chief guests", "chief guest", "moderators", "moderator", "judges",
    "judge", "participants", "delegates", "vips", "vip", "session",
    "organizers", "organisers", "committee", "members",
}


def _looks_like_section_header(name: str) -> bool:
    """A section divider like 'DIGNITARIES' or 'SPEAKERS' sitting in the
    Name column, with no details in the same row. Deliberately conservative
    so a real person's name (e.g. short initials like 'CP' or 'GV', or a
    name typed in all caps) is never mistaken for one:
    - exact/near match against a known list of section-label words, OR
    - long (3+ words) all-caps text with no digits/periods, which is a much
      stronger signal of a label than a short name/initials ever would be.
    """
    n = name.strip()
    if not n:
        return False
    norm = re.sub(r"\s+", " ", n.lower()).strip(" :-")
    if norm in _SECTION_HEADER_WORDS:
        return True
    has_lower = any(c.islower() for c in n)
    has_digit_or_period = any(c.isdigit() or c == "." for c in n)
    word_count = len(n.split())
    return (not has_lower) and (not has_digit_or_period) and word_count >= 3


@dataclass
class ParseResult:
    rows: list[dict]
    note: str  # human-readable explanation of what was found


def parse_dignitaries(file) -> ParseResult:
    """Parse an uploaded Excel file: find Name + everything else (as a
    single combined details block), ignore the rest of the sheet."""

    raw = pd.read_excel(file, header=None)
    raw = raw.dropna(how="all")
    if raw.empty:
        return ParseResult([], "The uploaded file appears to be empty.")

    # Decide whether row 0 is a header row: does any cell match a known
    # header word (name synonym or ignore synonym)?
    first_row_norms = [_normalize(v) for v in raw.iloc[0].tolist()]
    has_header = any(n in NAME_SYNONYMS or n in IGNORE_SYNONYMS for n in first_row_norms)

    if has_header:
        df = raw.copy()
        df.columns = [str(c).strip() for c in df.iloc[0]]
        df = df.iloc[1:].reset_index(drop=True)
    else:
        df = raw.copy()
        df.columns = [f"col_{i}" for i in range(df.shape[1])]

    df = df.fillna("")

    # --- Find the Name column -------------------------------------------
    name_col = None
    if has_header:
        for col in df.columns:
            if _normalize(col) in NAME_SYNONYMS:
                name_col = col
                break
    if name_col is None:
        name_col = df.columns[0]  # universal convention: name is first

    # --- Identify columns to ignore entirely (serial numbers etc.) ------
    ignore_cols = set()
    if has_header:
        for col in df.columns:
            if col != name_col and _normalize(col) in IGNORE_SYNONYMS:
                ignore_cols.add(col)

    detail_cols = [c for c in df.columns if c != name_col and c not in ignore_cols]

    # --- Decide whether to keep Title/Company as two distinct fields ----
    # Only when there are EXACTLY two detail columns left AND at least one
    # of them has a header that clearly means "Title" (Designation,
    # Position, etc.) do we trust the sheet's own structure enough to keep
    # Title and Company separate (rendered as two stacked lines). Any other
    # shape (one column, or 2+ columns with no clear Title header) falls
    # back to combining everything into a single details block, since
    # guessing which part is "title" vs "company" otherwise is unreliable.
    title_col = None
    company_col = None
    if has_header and len(detail_cols) == 2:
        for c in detail_cols:
            if _normalize(c) in TITLE_SYNONYMS:
                title_col = c
        if title_col is not None:
            company_col = next(c for c in detail_cols if c != title_col)

    # --- Build rows -------------------------------------------------------
    rows = []
    skipped_section_headers = []
    skipped_blank = 0

    for _, r in df.iterrows():
        name = _clean_text(r[name_col])
        if not name:
            skipped_blank += 1
            continue

        if title_col is not None:
            title = _clean_text(r[title_col])
            company = _clean_text(r[company_col])
            details_for_skip_check = (title + company)
        else:
            detail_parts = [_clean_text(r[c]) for c in detail_cols]
            detail_parts = [p for p in detail_parts if p]
            # Join multiple distinct columns with ", " so structure survives
            # as a comma-separated line, rather than mashed with no
            # punctuation at all.
            title = ""
            company = ", ".join(detail_parts).strip()
            details_for_skip_check = company

        if not details_for_skip_check and _looks_like_section_header(name):
            skipped_section_headers.append(name)
            continue

        rows.append({"name": name, "title": title, "company": company})

    if title_col is not None:
        note_parts = [f"Used '{name_col}' as Name, '{title_col}' as Title, '{company_col}' as Company."]
    else:
        note_parts = [
            f"Used '{name_col}' as Name; "
            + (f"combined {', '.join(repr(c) for c in detail_cols)} as Title/Company."
               if detail_cols else "no other column with details was found.")
        ]
    if not has_header:
        note_parts.append("No header row was detected; the first column was used as Name.")
    if skipped_section_headers:
        preview = ", ".join(repr(s) for s in skipped_section_headers[:3])
        more = "…" if len(skipped_section_headers) > 3 else ""
        note_parts.append(
            f"Skipped {len(skipped_section_headers)} likely section-header row(s) "
            f"with no details (e.g. {preview}{more})."
        )
    if skipped_blank:
        note_parts.append(f"Skipped {skipped_blank} blank row(s).")

    return ParseResult(rows, " ".join(note_parts))
