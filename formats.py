# -*- coding: utf-8 -*-
"""WordSoft formats — open and save: docx / txt / md / html / pdf.

Same idea as a `convert` command: one document model, many file types.

Copyright (c) 2026 salim-slimani. MIT License.
"""
from __future__ import annotations
import re, html as ihtml
from engine import WordDoc, Paragraph, Run, Table

STYLE2DOCX = {"Normal": "Normal", "Title": "Title", "Subtitle": "Subtitle",
              "Heading 1": "Heading 1", "Heading 2": "Heading 2",
              "Heading 3": "Heading 3", "Quote": "Intense Quote",
              "List Bullet": "List Bullet", "List Number": "List Number"}

# ---------------- DOCX ----------------
def save_docx(doc: WordDoc, path: str):
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH as WD_ALIGN
    wd = Document()
    core = wd.core_properties
    core.title = doc.title; core.author = doc.author
    amap = {"left": WD_ALIGN.LEFT, "center": WD_ALIGN.CENTER,
            "right": WD_ALIGN.RIGHT, "justify": WD_ALIGN.JUSTIFY}
    for b in doc.blocks:
        if isinstance(b, Table):
            t = wd.add_table(rows=len(b.rows), cols=max(len(r) for r in b.rows) if b.rows else 1)
            t.style = "Table Grid"
            for i, row in enumerate(b.rows):
                for j, cell in enumerate(row):
                    if j < len(t.rows[i].cells):
                        t.rows[i].cells[j].text = cell
                        if b.header and i == 0:
                            for p in t.rows[i].cells[j].paragraphs:
                                for r in p.runs:
                                    r.bold = True
            continue
        p = wd.add_paragraph(style=STYLE2DOCX.get(b.style, "Normal"))
        p.alignment = amap.get(b.align, WD_ALIGN.LEFT)
        try:
            pf = p.paragraph_format
            pf.space_after = Pt(b.spacing_after); pf.line_spacing = b.line_spacing
        except Exception:
            pass
        for r in b.runs:
            rr = p.add_run(r.text if r.text else "")
            rr.bold = r.bold; rr.italic = r.italic
            rr.underline = r.underline or False
            try:
                rr.font.strike = r.strike
                if r.size: rr.font.size = Pt(r.size)
                if r.font: rr.font.name = r.font
                if r.color:
                    h = r.color.strip("#")
                    rr.font.color.rgb = RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
            except Exception:
                pass
    wd.save(path)


def load_docx(path: str) -> WordDoc:
    from docx import Document
    wd = Document(path)
    doc = WordDoc(title=getattr(wd.core_properties, "title", "") or "Imported document")
    rev = {v: k for k, v in STYLE2DOCX.items()}
    for p in wd.paragraphs:
        if not p.text and not p.runs:
            continue
        style = rev.get(p.style.name if p.style else "Normal", "Normal")
        try:
            al = p.alignment
            align = {0: "left", 1: "center", 2: "right", 3: "justify"}.get(int(al), "left") if al is not None else "left"
        except Exception:
            align = "left"
        runs = []
        for r in p.runs:
            try: color = "%02X%02X%02X" % (r.font.color.rgb[0], r.font.color.rgb[1], r.font.color.rgb[2]) if r.font.color and r.font.color.rgb else None
            except Exception: color = None
            try: size = int(r.font.size.pt) if r.font.size else None
            except Exception: size = None
            runs.append(Run(r.text, bold=bool(r.bold), italic=bool(r.italic),
                            underline=bool(r.underline), strike=bool(getattr(r.font, "strike", False)),
                            color=color, font=r.font.name, size=size))
        if not runs:
            runs = [Run(p.text)]
        doc.blocks.append(Paragraph(runs, style=style, align=align))
    for t in wd.tables:
        rows = [[c.text for c in row.cells] for row in t.rows]
        doc.blocks.append(Table(rows=rows))
    if not doc.blocks:
        return WordDoc.blank()
    return doc

# ---------------- TXT / MD / HTML ----------------
def save_txt(doc: WordDoc, path: str):
    with open(path, "w", encoding="utf-8") as f:
        f.write(doc.full_text())

def load_txt(path: str) -> WordDoc:
    with open(path, encoding="utf-8") as f:
        lines = f.read().splitlines()
    doc = WordDoc(title="Text document")
    for ln in lines:
        if ln.startswith("# "): doc.blocks.append(Paragraph([Run(ln[2:], bold=True, size=18)], style="Heading 1"))
        elif ln.startswith("## "): doc.blocks.append(Paragraph([Run(ln[3:], bold=True)], style="Heading 2"))
        elif ln.strip().startswith(("- ", "* ")): doc.blocks.append(Paragraph([Run(ln.strip()[2:])], style="List Bullet"))
        else: doc.blocks.append(Paragraph([Run(ln)]))
    return doc or WordDoc.blank()

def save_md(doc: WordDoc, path: str):
    out = [f"# {doc.title}\n"]
    for b in doc.blocks:
        if isinstance(b, Table):
            if b.rows:
                out.append("| " + " | ".join(b.rows[0]) + " |")
                out.append("| " + " | ".join(["---"] * len(b.rows[0])) + " |")
                for row in b.rows[1:]:
                    out.append("| " + " | ".join(row) + " |")
            out.append("")
        else:
            t = b.text()
            if b.style == "Title": out.append(f"# {t}")
            elif b.style == "Heading 1": out.append(f"## {t}")
            elif b.style == "Heading 2": out.append(f"### {t}")
            elif b.style == "Heading 3": out.append(f"#### {t}")
            elif b.style == "Quote": out.append(f"> {t}")
            elif b.style == "List Bullet": out.append(f"- {t}")
            elif b.style == "List Number": out.append(f"1. {t}")
            else:
                out.append(t)
            out.append("")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(out))

BRAND_BLUE = "#2563EB"

def save_html(doc: WordDoc, path: str):
    parts = [f"<html lang='en'><head><meta charset='utf-8'><title>{ihtml.escape(doc.title)}</title>",
             f"<style>body{{font-family:'Segoe UI',Arial,sans-serif;max-width:800px;margin:40px auto;line-height:1.8;padding:0 16px;color:#111}}"
             f"h1{{color:#111}}h2{{color:{BRAND_BLUE}}}table{{border-collapse:collapse;width:100%}}"
             f"td,th{{border:1px solid #999;padding:6px}}th{{background:#e8eefc}}"
             f"blockquote{{border-left:4px solid {BRAND_BLUE};background:#f3f6ff;padding:8px 14px;margin:8px 0}}"
             f"footer{{color:#888;font-size:12px;margin-top:32px}}</style></head><body>",
             f"<h1>{ihtml.escape(doc.title)}</h1>"]
    for b in doc.blocks:
        if isinstance(b, Table):
            parts.append("<table>")
            for i, row in enumerate(b.rows):
                tag = "th" if (b.header and i == 0) else "td"
                parts.append("<tr>" + "".join(f"<{tag}>{ihtml.escape(c)}</{tag}>" for c in row) + "</tr>")
            parts.append("</table>")
        else:
            t = "".join(f"<b>{ihtml.escape(r.text)}</b>" if r.bold else ihtml.escape(r.text) for r in b.runs)
            if b.style == "Heading 1": parts.append(f"<h2>{t}</h2>")
            elif b.style == "Heading 2": parts.append(f"<h3>{t}</h3>")
            elif b.style == "Heading 3": parts.append(f"<h4>{t}</h4>")
            elif b.style == "Quote": parts.append(f"<blockquote>{t}</blockquote>")
            elif b.style == "List Bullet": parts.append(f"<ul><li>{t}</li></ul>")
            elif b.style == "List Number": parts.append(f"<ol><li>{t}</li></ol>")
            else: parts.append(f"<p style='text-align:{b.align}'>{t}</p>")
    parts.append("<footer>Created with WordSoft — Copyright (c) 2026 salim-slimani</footer></body></html>")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(parts))

# ---------------- PDF ----------------
def _shape(text: str) -> str:
    """Reshape RTL text (Arabic/Hebrew) so ReportLab renders it correctly."""
    try:
        import arabic_reshaper
        from bidi.algorithm import get_display
        return get_display(arabic_reshaper.reshape(text))
    except Exception:
        return text

def save_pdf(doc: WordDoc, path: str):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph as P, Spacer, Table as PT, TableStyle
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
    import html as H
    amap = {"left": TA_LEFT, "center": TA_CENTER, "right": TA_RIGHT, "justify": TA_JUSTIFY}
    # Prefer a TTF with broad Unicode coverage when available, else Helvetica.
    import os
    font = "Helvetica"
    fontB = "Helvetica-Bold"
    for cand in ["C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/Amiri-Regular.ttf",
                 "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"]:
        if os.path.exists(cand):
            try:
                from reportlab.pdfbase import pdfmetrics
                from reportlab.pdfbase.ttfonts import TTFont
                pdfmetrics.registerFont(TTFont("WS", cand))
                pdfmetrics.registerFont(TTFont("WS-B", cand))
                font, fontB = "WS", "WS-B"
                break
            except Exception:
                pass
    story = []
    st_title = ParagraphStyle("t", fontName=fontB, fontSize=22, alignment=TA_CENTER, spaceAfter=12)
    st_h1 = ParagraphStyle("h1", fontName=fontB, fontSize=16, spaceBefore=14, spaceAfter=6)
    st_h2 = ParagraphStyle("h2", fontName=fontB, fontSize=13, spaceBefore=10, spaceAfter=4)
    st_n = ParagraphStyle("n", fontName=font, fontSize=11, leading=17, alignment=TA_JUSTIFY, spaceAfter=6)
    story.append(P(H.escape(_shape(doc.title)), st_title)); story.append(Spacer(1, 4 * mm))
    for b in doc.blocks:
        if isinstance(b, Table):
            data = [[H.escape(_shape(c)) for c in row] for row in b.rows]
            t = PT(data, repeatRows=1 if b.header else 0)
            t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.6, colors.grey),
                                   ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e8eefc")),
                                   ("FONTSIZE", (0, 0), (-1, -1), 9), ("FONTNAME", (0, 0), (-1, -1), font)]))
            story.append(t); story.append(Spacer(1, 3 * mm))
        else:
            txt = "".join(f"<b>{H.escape(_shape(r.text))}</b>" if r.bold else H.escape(_shape(r.text)) for r in b.runs) or " "
            if b.style == "Title": story.append(P(txt, st_title))
            elif b.style == "Heading 1": story.append(P(txt, st_h1))
            elif b.style == "Heading 2": story.append(P(txt, st_h2))
            else:
                st = ParagraphStyle(f"p{id(b)}", parent=st_n, alignment=amap.get(b.align, TA_JUSTIFY))
                story.append(P(txt, st))
    SimpleDocTemplate(path, pagesize=A4, topMargin=20 * mm, bottomMargin=20 * mm,
                      rightMargin=20 * mm, leftMargin=20 * mm,
                      title=doc.title, author=doc.author).build(story)

# ---------------- dispatcher ----------------
def save(doc: WordDoc, path: str):
    p = path.lower()
    if p.endswith(".docx"): return save_docx(doc, path)
    if p.endswith(".pdf"): return save_pdf(doc, path)
    if p.endswith(".html") or p.endswith(".htm"): return save_html(doc, path)
    if p.endswith(".md"): return save_md(doc, path)
    return save_txt(doc, path)

def load(path: str) -> WordDoc:
    p = path.lower()
    if p.endswith(".docx"): return load_docx(path)
    return load_txt(path)
