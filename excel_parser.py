"""
Flexible Excel parsing for the Name Board Generator.

Goal: regardless of how the uploaded Excel is structured, reliably recover
(name, title, company) for each row. Three layers, in order of preference:

1. Header synonym matching — recognize common header names for each role
   (Name / Title / Company), case-insensitively, even if the workbook
   doesn't use the exact words "Name"/"Title"/"Company".
2. Combined-column splitting — if there's a Name column plus exactly ONE
   other text column (no separate Title and Company), treat that single
   column as a combined "Title + Company" field and split it into the two
   parts (split on the last comma, since that's how this app joins them
   when generating boards in the first place).
3. No-header fallback — if the first row doesn't look like a header row
   at all (i.e. none of the cells match known header words), assume the
   file has no header row and treat columns positionally: 1st = Name,
   2nd = Title, 3rd = Company (if present).

Returns a list of dicts: [{"name": ..., "title": ..., "company": ...}, ...]
plus a human-readable note describing which interpretation was used, so the
app can show the user what happened.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

import pandas as pd

# ---------------------------------------------------------------------------
# Header synonym tables
# ---------------------------------------------------------------------------

NAME_SYNONYMS = {
    "name", "names", "fullname", "full name", "dignitary", "dignitary name",
    "speaker", "speaker name", "guest", "guest name", "person", "attendee",
}

TITLE_SYNONYMS = {
    "title", "designation", "position", "role", "post",
    "designation/title", "title/designation",
}

COMPANY_SYNONYMS = {
    "company", "organisation", "organization", "org", "department",
    "affiliation", "institute", "institution", "ministry", "office",
}

# Columns that plausibly hold "title + company combined in one field"
COMBINED_SYNONYMS = {
    "organisation", "organization", "designation", "title", "details",
    "title/organisation", "title/organization", "designation/organisation",
    "designation/organization", "position", "role",
} | TITLE_SYNONYMS | COMPANY_SYNONYMS


def _normalize(col: str) -> str:
    return re.sub(r"\s+", " ", str(col).strip().lower())


@dataclass
class ParseResult:
    rows: list[dict]
    note: str  # human-readable explanation of what interpretation was used


def _looks_like_header(value) -> bool:
    """Heuristic: does this cell look like a column header word rather
    than actual data (a person's name, title, etc.)?"""
    norm = _normalize(value)
    if not norm:
        return False
    all_known = NAME_SYNONYMS | TITLE_SYNONYMS | COMPANY_SYNONYMS | COMBINED_SYNONYMS
    return norm in all_known


def _split_combined(text: str) -> tuple[str, str]:
    """Split a combined 'Title, Company' style string into (title, company).
    Mirrors the app's own merge rule (title + ", " + company), so splitting
    on the LAST comma is the correct inverse for the common case of a
    single embedded comma. If there are multiple commas, we still split on
    the last one, which best matches 'long descriptive title, Org Name'."""
    text = (text or "").strip()
    if not text:
        return "", ""
    if "," in text:
        title_part, company_part = text.rsplit(",", 1)
        return title_part.strip(), company_part.strip()
    # No comma at all — can't split further; treat whole thing as title.
    return text, ""


def parse_dignitaries(raw_bytes_or_buffer) -> ParseResult:
    """Parse an uploaded Excel file into a list of {"name","title","company"}
    dicts, using header synonyms, combined-column splitting, or positional
    fallback as needed."""

    # First, peek at the raw sheet with no header assumption, to test
    # whether row 0 actually looks like a header row.
    raw = pd.read_excel(raw_bytes_or_buffer, header=None)
    raw = raw.dropna(how="all")
    if raw.empty:
        return ParseResult([], "The uploaded file appears to be empty.")

    first_row_values = [str(v) for v in raw.iloc[0].tolist()]
    header_like_count = sum(1 for v in first_row_values if _looks_like_header(v))
    has_header = header_like_count >= 1  # at least one recognizable header word

    if has_header:
        df = raw.copy()
        df.columns = [str(c).strip() for c in df.iloc[0]]
        df = df.iloc[1:].reset_index(drop=True)
    else:
        df = raw.copy()
        df.columns = [f"col_{i}" for i in range(df.shape[1])]

    df = df.fillna("")

    # --- Map columns to roles -------------------------------------------------
    name_col = None
    title_col = None
    company_col = None

    if has_header:
        norm_cols = {col: _normalize(col) for col in df.columns}
        for col, norm in norm_cols.items():
            if name_col is None and norm in NAME_SYNONYMS:
                name_col = col
        for col, norm in norm_cols.items():
            if title_col is None and norm in TITLE_SYNONYMS and col != name_col:
                title_col = col
        for col, norm in norm_cols.items():
            if company_col is None and norm in COMPANY_SYNONYMS and col not in (name_col, title_col):
                company_col = col

        # If we couldn't even find a name column by synonym, fall back to
        # "first column is name" — nearly universal convention.
        if name_col is None:
            name_col = df.columns[0]

    else:
        # No header row detected at all: go purely positional.
        cols = list(df.columns)
        name_col = cols[0] if len(cols) > 0 else None
        title_col = cols[1] if len(cols) > 1 else None
        company_col = cols[2] if len(cols) > 2 else None

    # --- Build rows -------------------------------------------------------
    other_text_cols = [
        c for c in df.columns
        if c not in (name_col, title_col, company_col)
        and df[c].astype(str).str.strip().ne("").any()
    ]

    rows = []
    note_parts = []

    if name_col is None:
        return ParseResult([], "Could not find a column that looks like a Name column.")

    if title_col is not None and company_col is not None:
        # Clean case: separate Title and Company columns both found.
        for _, r in df.iterrows():
            rows.append({
                "name": str(r[name_col]).strip(),
                "title": str(r[title_col]).strip(),
                "company": str(r[company_col]).strip(),
            })
        note_parts.append(
            f"Used '{name_col}' as Name, '{title_col}' as Title, '{company_col}' as Company."
        )

    elif title_col is not None and company_col is None:
        # One title-ish column found, no separate company column.
        # If there's exactly one other populated text column, treat THAT
        # as company instead of splitting (more reliable than guessing a
        # comma split when real structure is available).
        candidate_company_cols = [c for c in other_text_cols]
        if len(candidate_company_cols) == 1:
            company_col = candidate_company_cols[0]
            for _, r in df.iterrows():
                rows.append({
                    "name": str(r[name_col]).strip(),
                    "title": str(r[title_col]).strip(),
                    "company": str(r[company_col]).strip(),
                })
            note_parts.append(
                f"Used '{name_col}' as Name, '{title_col}' as Title, "
                f"and '{company_col}' (unlabeled/extra column) as Company."
            )
        else:
            # Truly only Name + one combined column: split on last comma.
            for _, r in df.iterrows():
                combined = str(r[title_col]).strip()
                t, c = _split_combined(combined)
                rows.append({"name": str(r[name_col]).strip(), "title": t, "company": c})
            note_parts.append(
                f"'{title_col}' contained combined title + company text; "
                f"split automatically on the last comma into Title and Company."
            )

    else:
        # No title-ish column found by name at all. If there's at least one
        # other populated column, treat it as combined title+company.
        if other_text_cols:
            combined_col = other_text_cols[0]
            for _, r in df.iterrows():
                combined = str(r[combined_col]).strip()
                t, c = _split_combined(combined)
                rows.append({"name": str(r[name_col]).strip(), "title": t, "company": c})
            note_parts.append(
                f"Used '{name_col}' as Name; '{combined_col}' contained combined "
                f"title + company text and was split automatically on the last comma."
            )
        else:
            # Only a name column exists, nothing else.
            for _, r in df.iterrows():
                rows.append({"name": str(r[name_col]).strip(), "title": "", "company": ""})
            note_parts.append(f"Only a Name column ('{name_col}') was found; Title/Company left blank.")

    if not has_header:
        note_parts.append("No header row was detected, so columns were matched by position (1st=Name, 2nd=Title, 3rd=Company).")

    # Drop rows with no name
    rows = [r for r in rows if r["name"].strip() != ""]

    return ParseResult(rows, " ".join(note_parts))
