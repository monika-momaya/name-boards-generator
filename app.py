import io
import os
import tempfile
import zipfile

import pandas as pd
import streamlit as st

from board_generator import Dignitary, build_presentation, register_fonts
from excel_parser import parse_dignitaries

st.set_page_config(page_title="Name Board Generator", page_icon="🪧", layout="centered")

APP_DIR = os.path.dirname(__file__)
DEFAULT_FONT = os.path.join(APP_DIR, "fonts", "ALTGOT2N.TTF")

st.title("🪧 Name Board Generator")
st.caption(
    "Upload an Excel sheet of dignitaries → get back an editable PowerPoint "
    "(and optional PDF) with fold-over tent-card name boards, one slide per person."
)

# ---------------------------------------------------------------------------
# Sidebar: font status + template download
# ---------------------------------------------------------------------------

with st.sidebar:
    st.header("Font")
    st.write("**Font:** Alternate Gothic No.2 BT *(used for both Name and Title/Company)*")
    st.caption(
        "No font file is bundled with this app (font files are licensed "
        "and not stored in the repo). Upload your licensed copy below for "
        "accurate sizing in the in-app preview/PDF. Without an upload, "
        "text-fitting falls back to a generic system font for its "
        "calculations, and the generated .pptx will still reference the "
        "correct font name — but PowerPoint will substitute a default "
        "font on any machine that doesn't have Alternate Gothic No.2 BT "
        "installed."
    )

    font_upload = st.file_uploader("Upload Alternate Gothic No.2 BT (.ttf/.otf)", type=["ttf", "otf"], key="font")

    font_path = None
    if font_upload is not None:
        font_path = os.path.join(tempfile.gettempdir(), "font_" + font_upload.name)
        with open(font_path, "wb") as f:
            f.write(font_upload.getbuffer())

    register_fonts(font_path)

    st.divider()
    st.header("Excel template")
    template_df = pd.DataFrame(
        {
            "Name": ["Siddaramaiah", "D K Shivakumar"],
            "Title": ["Hon'ble Chief Minister of Karnataka", "Hon'ble Deputy Chief Minister"],
            "Company": ["", "Government of Karnataka"],
        }
    )
    buf = io.BytesIO()
    template_df.to_excel(buf, index=False)
    st.download_button(
        "Download blank template (.xlsx)",
        data=buf.getvalue(),
        file_name="nameboard_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ---------------------------------------------------------------------------
# Main: upload + preview + generate
# ---------------------------------------------------------------------------

st.subheader("1. Upload your Excel file")
st.write(
    "Any layout works — the app tries to automatically detect Name, Title, "
    "and Company columns, even with different header names, a single "
    "combined title+company column, or no header row at all."
)

uploaded = st.file_uploader("Excel file (.xlsx)", type=["xlsx"])

if uploaded is not None:
    try:
        result = parse_dignitaries(uploaded)
    except Exception as e:
        st.error(f"Could not read the Excel file: {e}")
        st.stop()

    if not result.rows:
        st.warning("No usable rows were found in the uploaded file. " + result.note)
        st.stop()

    df = pd.DataFrame(result.rows)

    st.success(f"Loaded {len(df)} dignitary record(s).")
    st.caption(f"ℹ️ {result.note}")
    st.dataframe(df[["name", "title", "company"]], use_container_width=True)
    st.caption("If this looks wrong, you can edit the table directly below before generating.")
    df = st.data_editor(df[["name", "title", "company"]], use_container_width=True, num_rows="dynamic")

    st.subheader("2. Generate")
    col1, col2 = st.columns(2)
    with col1:
        also_pdf = st.checkbox("Also generate PDF", value=False)
    with col2:
        st.write("")

    if st.button("🪧 Generate Name Boards", type="primary"):
        dignitaries = [
            Dignitary(
                name=str(row["name"]).strip(),
                title=str(row["title"]).strip(),
                company=str(row["company"]).strip(),
            )
            for _, row in df.iterrows()
            if str(row["name"]).strip() != ""
        ]

        with st.spinner("Building presentation..."):
            prs = build_presentation(dignitaries)
            pptx_buf = io.BytesIO()
            prs.save(pptx_buf)
            pptx_buf.seek(0)

        st.success("Name boards generated!")
        st.download_button(
            "⬇️ Download PowerPoint (.pptx)",
            data=pptx_buf.getvalue(),
            file_name="name_boards.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

        if also_pdf:
            with st.spinner("Converting to PDF (this may take a moment)..."):
                with tempfile.TemporaryDirectory() as tmpdir:
                    pptx_path = os.path.join(tmpdir, "name_boards.pptx")
                    with open(pptx_path, "wb") as f:
                        f.write(pptx_buf.getvalue())
                    ret = os.system(
                        f'soffice --headless --convert-to pdf --outdir "{tmpdir}" "{pptx_path}" >/dev/null 2>&1'
                    )
                    pdf_path = os.path.join(tmpdir, "name_boards.pdf")
                    if os.path.isfile(pdf_path):
                        with open(pdf_path, "rb") as f:
                            pdf_bytes = f.read()
                        st.download_button(
                            "⬇️ Download PDF",
                            data=pdf_bytes,
                            file_name="name_boards.pdf",
                            mime="application/pdf",
                        )
                    else:
                        st.warning(
                            "PDF conversion isn't available in this environment. "
                            "You can open the PPTX in PowerPoint and export to PDF from there."
                        )
else:
    st.info("⬆️ Upload an Excel file to get started, or download the template from the sidebar first.")
