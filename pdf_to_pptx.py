"""
pdf_to_pptx.py
==============
Convert a PDF of presentation slides into a fully editable PPTX.

Pipeline:
  1. Rasterise each PDF page to a PNG (via PyMuPDF — no system deps required)
  2. Send each page image to Claude vision with a structured extraction prompt
  3. Claude returns a JSON description of every element (text boxes, tables,
     shapes, colors, fonts, positions)
  4. Reconstruct each slide in python-pptx using real shapes, text boxes,
     and native Office tables — fully editable in PowerPoint / Google Slides

Usage (standalone):
    from pdf_to_pptx import pdf_to_editable_pptx
    pptx_bytes = pdf_to_editable_pptx(pdf_bytes, api_key="sk-ant-...")

Usage (inside Streamlit):
    import streamlit as st
    from pdf_to_pptx import pdf_to_editable_pptx
    pptx_bytes = pdf_to_editable_pptx(pdf_bytes,
                                       api_key=st.secrets["ANTHROPIC_API_KEY"],
                                       progress_cb=st.progress)
"""

from __future__ import annotations

import base64
import io
import json
import re
from typing import Callable

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt
from lxml import etree


# ── Slide canvas (widescreen 16:9) ───────────────────────────────────────────
SLIDE_W_IN = 13.33
SLIDE_H_IN = 7.5

_ALIGN_MAP = {
    "left":   PP_ALIGN.LEFT,
    "center": PP_ALIGN.CENTER,
    "right":  PP_ALIGN.RIGHT,
    "justify": PP_ALIGN.JUSTIFY,
}

EXTRACTION_PROMPT = """Analyze this presentation slide image carefully and extract ALL content into a structured JSON object.

Return ONLY valid JSON — no markdown fences, no explanation, no preamble.

The slide canvas is 13.33 inches wide × 7.5 inches tall. Express all positions and sizes in inches from the top-left corner.

Use this exact schema:

{
  "background_color": "hex6 (e.g. FFFFFF)",
  "elements": [
    // BACKGROUND / DECORATIVE SHAPES — colored rectangles, bars, accent panels
    {
      "type": "shape",
      "x": 0.0, "y": 0.0, "w": 13.33, "h": 7.5,
      "fill": "hex6 or null",
      "border_color": "hex6 or null",
      "border_pt": 1.0
    },

    // TEXT BOXES — titles, body text, labels, bullets, footnotes
    {
      "type": "text",
      "x": 0.0, "y": 0.0, "w": 6.0, "h": 0.6,
      "text": "exact text content",
      "font_size": 24.0,
      "bold": true,
      "italic": false,
      "color": "hex6",
      "bg_color": null,
      "align": "left",
      "font": "Calibri"
    },

    // TABLES — any grid of data
    {
      "type": "table",
      "x": 0.0, "y": 1.5, "w": 12.0, "h": 4.0,
      "col_widths_in": [3.0, 2.5, 2.5, 2.5],
      "row_heights_in": [0.7, 1.0, 1.0, 1.0],
      "header_row": {
        "cells": ["Column 1", "Column 2", "Column 3"],
        "bg": "hex6",
        "fg": "hex6",
        "bold": true,
        "font_size": 12.0
      },
      "data_rows": [
        {
          "cells": ["Row 1 Col 1", "Row 1 Col 2", "Row 1 Col 3"],
          "bg": "hex6",
          "fg": "hex6",
          "bold": false,
          "font_size": 12.0,
          "cell_overrides": {
            "2": {"bg": "1565C0", "fg": "FFFFFF", "bold": true}
          }
        }
      ]
    }
  ]
}

Rules:
- Include EVERY visible element: backgrounds, colored bars, decorative shapes, all text, tables, footers, page numbers, logos (as text placeholders)
- Extract exact text — preserve capitalization, hyphens, line breaks (use \\n)
- For tables: capture EVERY cell. Use cell_overrides for cells with different colors than the row default (e.g. a blue "Age Today" column)
- For col_widths_in: values must sum to approximately the table width w
- For shapes: include the footer bar, accent panels, colored backgrounds
- Estimate positions carefully — a title starting near top-left might be x=0.35, y=0.1
- Colors: 6-digit hex without # symbol
- Order elements back-to-front (backgrounds first, text on top)
"""


def _rgb(hex6: str) -> RGBColor:
    h = hex6.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _extract_slide_json(image_b64: str, api_key: str, model: str) -> dict:
    """Call Claude vision API and return parsed JSON for one slide."""
    import anthropic
    client = anthropic.Anthropic(api_key=api_key)

    msg = client.messages.create(
        model=model,
        max_tokens=4096,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": "image/png",
                        "data": image_b64,
                    },
                },
                {"type": "text", "text": EXTRACTION_PROMPT},
            ],
        }],
    )

    raw = msg.content[0].text.strip()

    # Strip markdown fences if Claude added them
    if raw.startswith("```"):
        raw = re.sub(r"^```[a-z]*\n?", "", raw)
        raw = re.sub(r"\n?```$", "", raw)
        raw = raw.strip()

    return json.loads(raw)


def _render_shape(slide, el: dict) -> None:
    """Add a colored rectangle / decorative shape."""
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(el["x"]), Inches(el["y"]),
        Inches(el["w"]), Inches(el["h"]),
    )
    fill = el.get("fill")
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = _rgb(fill)
    else:
        shape.fill.background()

    border = el.get("border_color")
    if border:
        shape.line.color.rgb = _rgb(border)
        shape.line.width = Pt(float(el.get("border_pt", 1.0)))
    else:
        shape.line.fill.background()


def _render_text(slide, el: dict) -> None:
    """Add a text box."""
    txb = slide.shapes.add_textbox(
        Inches(el["x"]), Inches(el["y"]),
        Inches(el["w"]), Inches(el["h"]),
    )
    tf = txb.text_frame
    tf.word_wrap = True

    bg = el.get("bg_color")
    if bg:
        txb.fill.solid()
        txb.fill.fore_color.rgb = _rgb(bg)
    else:
        txb.fill.background()

    raw_text = el.get("text", "")
    lines = raw_text.split("\n") if "\n" in raw_text else [raw_text]

    for li, line in enumerate(lines):
        p = tf.paragraphs[0] if li == 0 else tf.add_paragraph()
        p.alignment = _ALIGN_MAP.get(el.get("align", "left"), PP_ALIGN.LEFT)
        run = p.add_run()
        run.text = line
        run.font.name = el.get("font", "Calibri")
        run.font.bold = bool(el.get("bold", False))
        run.font.italic = bool(el.get("italic", False))
        run.font.size = Pt(float(el.get("font_size", 12)))
        run.font.color.rgb = _rgb(el.get("color", "000000"))


def _set_cell_fill(cell, hex6: str) -> None:
    """Apply a solid fill color to a table cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    # Remove any existing fills
    for existing in tcPr.findall(qn("a:solidFill")):
        tcPr.remove(existing)
    sf = etree.SubElement(tcPr, qn("a:solidFill"))
    sc = etree.SubElement(sf, qn("a:srgbClr"))
    sc.set("val", hex6.lstrip("#"))


def _render_table(slide, el: dict) -> None:
    """Add a fully styled native Office table."""
    header = el.get("header_row", {})
    data_rows = el.get("data_rows", [])
    n_rows = (1 if header else 0) + len(data_rows)
    header_cells = header.get("cells", [])
    n_cols = len(header_cells) or (len(data_rows[0]["cells"]) if data_rows else 1)

    tbl_shape = slide.shapes.add_table(
        n_rows, n_cols,
        Inches(el["x"]), Inches(el["y"]),
        Inches(el["w"]), Inches(el["h"]),
    )
    tbl = tbl_shape.table

    # Column widths
    col_ws = el.get("col_widths_in", [])
    if col_ws:
        for i, cw in enumerate(col_ws[:n_cols]):
            tbl.columns[i].width = Inches(float(cw))

    # Row heights
    row_hs = el.get("row_heights_in", [])
    if row_hs:
        for i, rh in enumerate(row_hs[:n_rows]):
            tbl.rows[i].height = Inches(float(rh))

    def fill_cell(r, c, text, bg=None, fg="000000", bold=False,
                  italic=False, font_size=12.0, align=PP_ALIGN.CENTER):
        cell = tbl.cell(r, c)
        cell.text = ""
        tf = cell.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = align

        # Handle multi-line cell text
        lines = str(text).split("\n")
        for li, line in enumerate(lines):
            if li > 0:
                p = tf.add_paragraph()
                p.alignment = align
            run = p.add_run()
            run.text = line
            run.font.name = "Calibri"
            run.font.bold = bold
            run.font.italic = italic
            run.font.size = Pt(float(font_size))
            run.font.color.rgb = _rgb(fg)

        if bg:
            _set_cell_fill(cell, bg)

    # Header row
    row_offset = 0
    if header:
        hbg  = header.get("bg", "BDD7EE")
        hfg  = header.get("fg", "000000")
        hbold = header.get("bold", True)
        hsize = float(header.get("font_size", 12))
        h_overrides = header.get("cell_overrides", {})
        for c, txt in enumerate(header_cells[:n_cols]):
            ov = h_overrides.get(str(c), {})
            fill_cell(0, c, txt,
                      bg=ov.get("bg", hbg),
                      fg=ov.get("fg", hfg),
                      bold=ov.get("bold", hbold),
                      font_size=ov.get("font_size", hsize),
                      align=PP_ALIGN.CENTER)
        row_offset = 1

    # Data rows
    for ri, row_def in enumerate(data_rows):
        cells   = row_def.get("cells", [])
        rbg     = row_def.get("bg", None)
        rfg     = row_def.get("fg", "000000")
        rbold   = row_def.get("bold", False)
        rsize   = float(row_def.get("font_size", 12))
        r_overrides = row_def.get("cell_overrides", {})
        for c, txt in enumerate(cells[:n_cols]):
            ov = r_overrides.get(str(c), {})
            fill_cell(row_offset + ri, c, txt,
                      bg=ov.get("bg", rbg),
                      fg=ov.get("fg", rfg),
                      bold=ov.get("bold", rbold),
                      font_size=ov.get("font_size", rsize),
                      align=_ALIGN_MAP.get(ov.get("align", "center"), PP_ALIGN.CENTER))


def _render_element(slide, el: dict) -> None:
    t = el.get("type", "shape")
    if t == "shape":
        _render_shape(slide, el)
    elif t == "text":
        _render_text(slide, el)
    elif t == "table":
        _render_table(slide, el)


def _page_to_b64(image) -> str:
    """PIL Image → base64 PNG string."""
    buf = io.BytesIO()
    image.save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode()


def pdf_to_editable_pptx(
    pdf_bytes: bytes,
    api_key: str,
    model: str = "claude-opus-4-5",
    dpi: int = 150,
    progress_cb: Callable | None = None,
) -> bytes:
    """
    Convert a PDF of slides to an editable PPTX.

    Args:
        pdf_bytes:   Raw bytes of the input PDF.
        api_key:     Anthropic API key.
        model:       Claude model to use for vision extraction.
        dpi:         Resolution for PDF rasterisation (150 is a good balance).
        progress_cb: Optional callable(fraction: float) for progress updates.

    Returns:
        Raw bytes of the generated .pptx file.
    """
    import fitz  # PyMuPDF — pip install pymupdf

    # ── Rasterise pages via PyMuPDF (no poppler/system deps required) ─────────
    import fitz
    from PIL import Image as _PILImage
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    images = []
    for _pg in doc:
        pix = _pg.get_pixmap(matrix=mat)
        images.append(_PILImage.frombytes("RGB", [pix.width, pix.height], pix.samples))
    n = len(images)

    # ── Build presentation ────────────────────────────────────────────────────
    prs = Presentation()
    prs.slide_width  = Inches(SLIDE_W_IN)
    prs.slide_height = Inches(SLIDE_H_IN)
    blank_layout = prs.slide_layouts[6]  # truly blank

    errors = []

    for i, img in enumerate(images):
        if progress_cb:
            progress_cb((i) / n)

        b64 = _page_to_b64(img)

        try:
            spec = _extract_slide_json(b64, api_key, model)
        except Exception as e:
            errors.append(f"Page {i+1}: vision extraction failed — {e}")
            # Add a blank slide with an error note
            slide = prs.slides.add_slide(blank_layout)
            txb = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12), Inches(1))
            txb.text_frame.paragraphs[0].add_run().text = f"[Page {i+1} could not be extracted: {e}]"
            continue

        slide = prs.slides.add_slide(blank_layout)

        # Set background
        bg = spec.get("background_color", "FFFFFF")
        bg_shape = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
        bg_shape.fill.solid()
        bg_shape.fill.fore_color.rgb = _rgb(bg)
        bg_shape.line.fill.background()

        # Render elements in order (back to front)
        for el in spec.get("elements", []):
            try:
                _render_element(slide, el)
            except Exception as e:
                errors.append(f"Page {i+1} element {el.get('type','?')}: {e}")

    if progress_cb:
        progress_cb(1.0)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue(), errors