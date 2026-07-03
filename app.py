import io
import os
import tempfile
import zipfile

import pandas as pd
import streamlit as st

from board_generator import Dignitary, build_presentation, register_fonts, embed_font_in_pptx
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
    font_upload = st.file_uploader("Upload font (.ttf/.otf)", type=["ttf", "otf"], key="font")

    font_path = None
    if font_upload is not None:
        font_path = os.path.join(tempfile.gettempdir(), "altgothic2bt_" + font_upload.name)
        with open(font_path, "wb") as f:
            f.write(font_upload.getbuffer())

    register_fonts(font_path, font_path)

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
    "Just needs a column with each person's **name**. Any other columns "
    "(title, designation, organization, etc.) are automatically picked up "
    "and shown on the board — section divider rows, blank rows, and "
    "serial-number/email/phone columns are detected and ignored."
)

uploaded = st.file_uploader("Excel file (.xlsx)", type=["xlsx"])

if uploaded is not None:
    try:
        result = parse_dignitaries(uploaded)
    except Exception as e:
        st.error(f"Could not read the Excel file: {e}")
        st.stop()

    rows = result.rows
    if not rows:
        st.warning("No usable rows (with a Name) were found in the uploaded file.")
        st.stop()

    st.success(f"Loaded {len(rows)} dignitary record(s).")
    st.caption(result.note)
    preview_df = pd.DataFrame(rows, columns=["name", "title", "company"])
    preview_df.columns = ["Name", "Title", "Company"]
    st.dataframe(preview_df, use_container_width=True)

    st.subheader("2. Generate")
    col1, col2 = st.columns(2)
    with col1:
        also_pdf = st.checkbox("Also generate PDF", value=False)
    with col2:
        st.write("")

    if st.button("🪧 Generate Name Boards", type="primary"):
        dignitaries = [
            Dignitary(
                name=row["name"],
                title=row["title"],
                company=row["company"],
            )
            for row in rows
        ]


        with st.spinner("Building presentation..."):
            prs = build_presentation(dignitaries)
            pptx_buf = io.BytesIO()
            prs.save(pptx_buf)
            pptx_bytes = pptx_buf.getvalue()

        # Embed the font directly into the PPTX so it renders correctly
        # on any machine, even without the font installed locally.
        if font_path:
            with st.spinner("Embedding font..."):
                try:
                    pptx_bytes = embed_font_in_pptx(
                        pptx_bytes, font_path, "AlternateGothic2 BT"
                    )
                except Exception as e:
                    st.warning(f"Font could not be embedded ({e}). The PPTX will still reference the font by name.")
        else:
            st.info(
                "💡 Upload your AlternateGothic2 BT font file in the sidebar to embed it "
                "into the PPTX — this ensures the correct font renders on every machine, "
                "including those without the font installed."
            )

        pptx_buf = io.BytesIO(pptx_bytes)
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
