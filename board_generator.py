"""
Name Board Generator — core engine.

Generates A4-landscape PPTX where each slide contains a fold-over tent card:
  - Top half: NAME + TITLE/COMPANY, rotated 180°
  - Bottom half: NAME + TITLE/COMPANY, upright

Fonts:
  - Name:           AlternateGothic2 BT, bold  (ALL CAPS, 90pt)
  - Title/Company:  AlternateGothic2 BT (Title Case, allowed to wrap to 2 lines)

Layout rule: if Title and Company each fit on their own line (within the
max line count budget), they are stacked tightly (own lines). If either
would overflow its line budget, Title and Company are merged into a single
comma-separated line instead.
"""

from __future__ import annotations

import copy
import io
import uuid
import zipfile
from dataclasses import dataclass
from typing import Optional

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from PIL import ImageFont
import os

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# A4 landscape
SLIDE_W_IN = 11.69
SLIDE_H_IN = 8.27

FONT_NAME_BOLD = "AlternateGothic2 BT"
FONT_NAME_MEDIUM = "AlternateGothic2 BT"

FALLBACK_FONT_REGULAR = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"

NAME_COLOR = RGBColor(0x00, 0x00, 0x00)
TITLE_COLOR = RGBColor(0x00, 0x00, 0x00)

TEXTBOX_W_IN = 29 / 2.54
MARGIN_X     = (SLIDE_W_IN - TEXTBOX_W_IN) / 2

NAME_MAX_PT  = 95
NAME_MIN_PT  = 95
TITLE_MAX_PT = 55
TITLE_MIN_PT = 55

NAME_TITLE_GAP_IN = 0.04
TITLE_COMPANY_GAP_IN = 0.01

HALF_H_IN = SLIDE_H_IN / 2

_DEFAULT_FONT_PATH = os.path.join(os.path.dirname(__file__), "fonts", "ALTGOT2N.TTF")
_CUSTOM_FONT_PATHS: dict[str, str] = {}
if os.path.isfile(_DEFAULT_FONT_PATH):
    _CUSTOM_FONT_PATHS["demi"] = _DEFAULT_FONT_PATH
    _CUSTOM_FONT_PATHS["medium"] = _DEFAULT_FONT_PATH


def register_fonts(demi_path: Optional[str], medium_path: Optional[str]) -> None:
    global _CUSTOM_FONT_PATHS
    if demi_path and os.path.isfile(demi_path):
        _CUSTOM_FONT_PATHS["demi"] = demi_path
    if medium_path and os.path.isfile(medium_path):
        _CUSTOM_FONT_PATHS["medium"] = medium_path


def _get_measure_font(weight: str, size_pt: int) -> ImageFont.FreeTypeFont:
    path = _CUSTOM_FONT_PATHS.get(weight)
    if path:
        try:
            return ImageFont.truetype(path, size_pt * 4)
        except Exception:
            pass
    try:
        return ImageFont.truetype(FALLBACK_FONT_REGULAR, size_pt * 4)
    except Exception:
        return ImageFont.load_default()


def _measure_width_in(text: str, weight: str, size_pt: float) -> float:
    font = _get_measure_font(weight, max(1, int(size_pt)))
    bbox = font.getbbox(text)
    width_px = bbox[2] - bbox[0]
    width_pt = width_px / 4
    return width_pt / 72.0


def fit_font_size(text: str, weight: str, max_pt: float, min_pt: float, max_width_in: float) -> float:
    if not text:
        return max_pt
    lo, hi = min_pt, max_pt
    best = min_pt
    for _ in range(20):
        mid = (lo + hi) / 2
        w = _measure_width_in(text, weight, mid)
        if w <= max_width_in:
            best = mid
            lo = mid
        else:
            hi = mid
        if hi - lo < 0.25:
            break
    return round(best, 1)


def wrap_text_to_width(text: str, weight: str, size_pt: float, max_width_in: float, max_lines: int = 2) -> Optional[list[str]]:
    words = text.split()
    if not words:
        return [""]
    lines: list[str] = []
    current = ""
    for word in words:
        trial = (current + " " + word).strip()
        if _measure_width_in(trial, weight, size_pt) <= max_width_in:
            current = trial
        else:
            if current:
                lines.append(current)
            current = word
            if len(lines) >= max_lines:
                return None
    if current:
        lines.append(current)
    if len(lines) > max_lines:
        return None
    return lines


_MINOR_WORDS = {
    "a", "an", "the", "of", "and", "or", "for", "to", "in", "on", "at",
    "by", "with", "from", "as", "nor", "but", "is",
}


def smart_title_case(text: str) -> str:
    if not text:
        return text
    words = text.split(" ")
    out_words = []
    for i, word in enumerate(words):
        core = word.strip(",.;:")
        if core.isupper() and len(core) <= 4 and core.isalpha():
            out_words.append(word)
            continue
        lower_core = core.lower()
        if lower_core in _MINOR_WORDS and i != 0:
            out_words.append(word.lower())
            continue
        chars = list(word.lower())
        for j, ch in enumerate(chars):
            if ch.isalpha():
                chars[j] = ch.upper()
                break
        out_words.append("".join(chars))
    return " ".join(out_words)


@dataclass
class Dignitary:
    name: str
    title: str = ""
    company: str = ""


# ---------------------------------------------------------------------------
# Slide building
# ---------------------------------------------------------------------------

def _set_run(run, text, font_name, size_pt, bold=False, color=NAME_COLOR, caps=False):
    run.text = text.upper() if caps else text
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    for tag in ("latin", "ea", "cs"):
        el = rPr.find(qn(f"a:{tag}"))
        if el is None:
            el = rPr.makeelement(qn(f"a:{tag}"), {})
            rPr.append(el)
        el.set("typeface", font_name)


def _add_textbox(slide, left_in, top_in, width_in, height_in, rotation=0):
    box = slide.shapes.add_textbox(Inches(left_in), Inches(top_in), Inches(width_in), Inches(height_in))
    box.rotation = rotation
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    return box


def _build_title_company_lines(title: str, company: str, max_width_in: float,
                               max_total_lines: int = 3,
                               title_max_pt: float = TITLE_MAX_PT,
                               title_min_pt: float = TITLE_MIN_PT):
    title = smart_title_case((title or "").strip())
    company = smart_title_case((company or "").strip())

    if not title and not company:
        return [], title_max_pt
    if title and not company:
        size = fit_font_size(title, "medium", title_max_pt, title_min_pt, max_width_in)
        if _measure_width_in(title, "medium", size) <= max_width_in:
            return [title], size
        wrapped = wrap_text_to_width(title, "medium", title_min_pt, max_width_in, max_lines=3)
        return (wrapped or [title]), title_min_pt
    if company and not title:
        size = fit_font_size(company, "medium", title_max_pt, title_min_pt, max_width_in)
        if _measure_width_in(company, "medium", size) <= max_width_in:
            return [company], size
        wrapped = wrap_text_to_width(company, "medium", title_min_pt, max_width_in, max_lines=3)
        return (wrapped or [company]), title_min_pt

    for size in [title_max_pt - i * 1.0 for i in range(int((title_max_pt - title_min_pt)) + 1)]:
        title_fits = _measure_width_in(title, "medium", size) <= max_width_in
        company_fits = _measure_width_in(company, "medium", size) <= max_width_in
        if title_fits and company_fits:
            return [title, company], size

    merged = f"{title}, {company}"
    size = fit_font_size(merged, "medium", title_max_pt, title_min_pt, max_width_in)
    if _measure_width_in(merged, "medium", size) <= max_width_in:
        return [merged], size
    wrapped = wrap_text_to_width(merged, "medium", title_min_pt, max_width_in, max_lines=3)
    return (wrapped or [merged]), title_min_pt


def _render_half(slide, dignitary: Dignitary, top_in: float, rotation: int,
                 slide_w: float = SLIDE_W_IN, half_h: float = HALF_H_IN,
                 scale: float = 1.0):
    textbox_w      = TEXTBOX_W_IN        * scale
    margin_x       = (slide_w - textbox_w) / 2
    name_max_pt    = NAME_MAX_PT         * scale
    name_min_pt    = NAME_MIN_PT         * scale
    title_max_pt   = TITLE_MAX_PT        * scale
    title_min_pt   = TITLE_MIN_PT        * scale
    name_title_gap = 0.6 / 2.54 * scale
    tc_gap         = 0.35 / 2.54 * scale

    max_width_in = textbox_w

    name_text = dignitary.name.strip()
    name_size = fit_font_size(name_text, "demi", name_max_pt, name_min_pt, max_width_in)

    title_lines, title_size = _build_title_company_lines(
        dignitary.title, dignitary.company, max_width_in,
        title_max_pt=title_max_pt, title_min_pt=title_min_pt,
    )

    def line_h_in(pt_size):
        return (pt_size * 1.2) / 72.0

    name_h = line_h_in(name_size)
    title_block_h = 0.0
    if title_lines:
        title_block_h = (len(title_lines) * line_h_in(title_size) * 0.86) + (len(title_lines) - 1) * tc_gap

    total_h = name_h + (name_title_gap if title_lines else 0) + title_block_h

    d = 0.0

    if rotation == 180:
        name_y  = top_in + half_h - name_h
        title_y = name_y - name_title_gap - title_block_h
    else:
        name_y  = top_in
        title_y = name_y + name_h + name_title_gap

    name_box = _add_textbox(slide, margin_x, name_y, max_width_in, name_h, rotation=rotation)
    name_box.text_frame.margin_top = 0
    name_box.text_frame.margin_bottom = 0
    p = name_box.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    _set_run(run, name_text, FONT_NAME_BOLD, name_size, bold=True, color=NAME_COLOR, caps=True)

    if title_lines:
        title_box = _add_textbox(slide, margin_x, title_y, max_width_in, title_block_h, rotation=rotation)
        title_box.text_frame.margin_top = 0
        title_box.text_frame.margin_bottom = 0
        tf = title_box.text_frame
        for i, line in enumerate(title_lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.CENTER
            p.space_before = Pt(0)
            p.space_after = Pt(0)
            p.line_spacing = 1.0
            run = p.add_run()
            _set_run(run, line, FONT_NAME_MEDIUM, title_size, bold=False, color=TITLE_COLOR, caps=False)


# ---------------------------------------------------------------------------
# Paper size presets — all measurements scale proportionally from A4
# ---------------------------------------------------------------------------

PAPER_SIZES = {
    'A4 Landscape': {'w_in': 11.69, 'h_in': 8.27},
    'A5 Landscape': {'w_in': 8.27, 'h_in': 5.83},
}

_A4_W = 11.69


def build_presentation(dignitaries: list[Dignitary], paper_size: str = 'A4 Landscape') -> Presentation:
    size   = PAPER_SIZES.get(paper_size, PAPER_SIZES['A4 Landscape'])
    sw     = size['w_in']
    sh     = size['h_in']
    scale  = sw / _A4_W

    prs = Presentation()
    prs.slide_width  = Inches(sw)
    prs.slide_height = Inches(sh)
    blank_layout = prs.slide_layouts[6]

    half_h = sh / 2

    for dig in dignitaries:
        slide = prs.slides.add_slide(blank_layout)

        line = slide.shapes.add_connector(
            1,
            Inches(0.3 * scale), Inches(half_h),
            Inches(sw - 0.3 * scale), Inches(half_h)
        )
        line.line.color.rgb = RGBColor(0xE5, 0xE5, 0xE5)
        line.line.width = Pt(0.25)

        _render_half(slide, dig, top_in=0.0,   rotation=180, slide_w=sw, half_h=half_h, scale=scale)
        _render_half(slide, dig, top_in=half_h, rotation=0,   slide_w=sw, half_h=half_h, scale=scale)

    return prs


# ---------------------------------------------------------------------------
# Font embedding
# ---------------------------------------------------------------------------

def _obfuscate_font(font_data: bytes, guid: str) -> bytes:
    g = guid.strip('{}').replace('-', '')
    c1 = bytes(reversed(bytes.fromhex(g[0:8])))
    c2 = bytes(reversed(bytes.fromhex(g[8:12])))
    c3 = bytes(reversed(bytes.fromhex(g[12:16])))
    c4 = bytes.fromhex(g[16:20])
    c5 = bytes.fromhex(g[20:32])
    key = c1 + c2 + c3 + c4 + c5
    result = bytearray(font_data)
    for i in range(min(32, len(result))):
        result[i] ^= key[i % 16]
    return bytes(result)


def embed_font_in_pptx(pptx_bytes: bytes, font_path: str, font_name: str) -> bytes:
    with open(font_path, 'rb') as fh:
        font_data = fh.read()

    guid = '{' + str(uuid.uuid4()).upper() + '}'
    rel_id = guid
    font_part_name = 'ppt/fonts/font1.fntdata'
    obfuscated = _obfuscate_font(font_data, guid)

    in_buf = io.BytesIO(pptx_bytes)
    out_buf = io.BytesIO()

    with zipfile.ZipFile(in_buf, 'r') as zin:
        with zipfile.ZipFile(out_buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                text = None

                if item.filename == '[Content_Types].xml':
                    text = data.decode('utf-8')
                    if 'fntdata' not in text:
                        text = text.replace(
                            '</Types>',
                            '<Default Extension="fntdata" ContentType="application/x-fontdata"/></Types>'
                        )

                elif item.filename == 'ppt/_rels/presentation.xml.rels':
                    text = data.decode('utf-8')
                    text = text.replace(
                        '</Relationships>',
                        f'<Relationship Id="{rel_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font1.fntdata"/></Relationships>'
                    )

                elif item.filename == 'ppt/presentation.xml':
                    text = data.decode('utf-8')
                    font_block = (
                        f'<p:embeddedFontLst>'
                        f'<p:embeddedFont>'
                        f'<p:font typeface="{font_name}"/>'
                        f'<p:regular r:id="{rel_id}"/>'
                        f'</p:embeddedFont>'
                        f'</p:embeddedFontLst>'
                    )

                    if '<p:defaultTextStyle' in text:
                        text = text.replace(
                            '<p:defaultTextStyle',
                            font_block + '<p:defaultTextStyle',
                            1
                        )
                    else:
                        import re
                        text, n = re.subn(
                            r'(<p:notesSz[^/]*/\s*>)',
                            r'\1' + font_block,
                            text, count=1
                        )
                        if not n:
                            text = text.replace(
                                '</p:presentation>',
                                font_block + '</p:presentation>',
                                1
                            )

                zout.writestr(item, text.encode('utf-8') if text is not None else data)

        zout.writestr(font_part_name, obfuscated)

    out_buf.seek(0)
    return out_buf.getvalue()
