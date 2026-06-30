# Name Board Generator

Generates dignitary/speaker name boards (fold-over tent cards) from an Excel
list, as editable PowerPoint (.pptx) and optional PDF.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL Streamlit prints (usually http://localhost:8501).

PDF export requires LibreOffice (`soffice`) installed on the machine running
the app. If it isn't available, PPTX download still works — open it in
PowerPoint and use File > Export to get a PDF.

## Font — important, read before deploying

**No font file is included in this repo.** AlternateGothic2 BT is a
commercially licensed font; redistributing it in a public (or shared
private) repository would violate that license.

Instead, the app has a **font upload option in the sidebar** at runtime:
each user supplies their own already-licensed copy of the font, which is
used only for that session (held in temp storage, never committed or
persisted to the repo). This keeps the codebase itself license-clean.

The same font (AlternateGothic2 BT) is used for both the Name (bold,
90pt) and Title/Company (regular, 50pt) text — one upload covers both.

Two things worth knowing:
1. **You (or whoever runs this) still need a valid license** to use
   AlternateGothic2 BT — the upload option avoids the *app* being a
   distribution channel, but doesn't grant usage rights on its own.
2. **Font is not embedded inside the generated PPTX file.** The app sets the
   font *name* in the slide XML, but python-pptx does not embed the actual
   font binary into the .pptx. This means:
   - The in-app PDF preview/export will look correct (the uploaded font is
     installed server-side for rendering during that session).
   - If someone opens the downloaded .pptx on a different machine that
     doesn't have AlternateGothic2 BT installed, PowerPoint will silently
     substitute a default font for the name/title text.
   - For guaranteed visual fidelity outside this app, install the actual
     font file on whichever machine will ultimately open/print the PPTX,
     or export to PDF (PDF rasterizes/embeds appearance, so it's safe to
     share as-is).

If no font is uploaded, the app falls back to a generic system font for
on-screen text-fitting calculations only; PowerPoint will still try to use
the configured font names and substitute if they aren't installed.

## Files

- `app.py` — Streamlit UI: upload Excel, preview, generate, download.
- `excel_parser.py` — Flexible Excel reader: detects the Name column,
  ignores serial-number/email/phone columns, skips section-header and
  blank rows, and combines the remaining columns into Title/Company.
- `board_generator.py` — Core layout engine (font fitting, wrapping, slide
  building). No Streamlit dependency; can be reused or unit tested standalone.
- `fonts/` — Empty by design (see Fonts section above). Drop a locally
  licensed .ttf/.otf here for local runs, or use the in-app uploader.

## Excel format

Just needs a column with each person's name — any common header works
(`Name`, `Full Name`, `Dignitary`, `Speaker`, `Guest`, etc.), or if there's
no header row at all, the first column is used.

Any other populated columns on that row (Title, Designation, Company,
Organization, etc.) are automatically picked up and shown on the board.
Columns that are clearly not part of a person's details — serial numbers,
email, phone — are detected and ignored automatically.

A few extra conveniences:
- Section-divider rows with no other details (e.g. a row that just says
  "DIGNITARIES" or "SPEAKERS") are skipped instead of becoming a blank
  board.
- Fully blank rows are skipped.
- If exactly two extra columns remain and one is clearly labeled Title /
  Designation / Position / Role / Post, Title and Company are kept as two
  separate stacked lines; otherwise all extra columns are combined into a
  single details block (comma-separated), and the board generator decides
  whether that fits on one line or needs to wrap.

Download a starter template from the app sidebar.

## Layout rules implemented

- Each dignitary gets one A4-landscape slide, split into two halves:
  top half rotated 180°, bottom half upright (fold-over tent card).
- Name: ALL CAPS, auto-shrinks to fit on one line.
- Title + Company: Title Case (minor words like "of", "and", "for" stay
  lowercase; apostrophe-s as in "Hon'ble" is not capitalized).
  - If both fit on their own line, they're stacked with a tight gap between
    them, and a larger gap above (between Name and Title).
  - If either is too long to fit on one line, Title and Company are merged
    into a single comma-separated block instead, which may wrap to 2 lines.

