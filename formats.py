# -*- coding: utf-8 -*-
"""WordSoft formats — open and save: docx / txt / md / html / pdf.

DOCX covers: styles, runs, tables, headers/footers + page numbers,
page setup (size, orientation, margins, columns), watermark, TOC field,
footnotes section, core properties.

Copyright (c) 2026 salim-slimani. MIT License.
"""
from __future__ import annotations
import re, html as ihtml
from engine import WordDoc, Paragraph, Run, Table, PageBreak, TOCField, FootnoteRef, PAGE_SIZES

STYLE2DOCX = {"Normal": "Normal", "Title": "Title", "Subtitle": "Subtitle",
              "Heading 1": "Heading 1", "Heading 2": "Heading 2",
              "Heading 3": "Heading 3", "Heading 4": "Heading 4",
              "Quote": "Intense Quote", "Caption": "Caption",
              "List Bullet": "List Bullet", "List Number": "List Number"}

BRAND_BLUE = "#2563EB"

# ---------------- DOCX helpers ----------------
def _set_section(wd, setup):
    from docx.shared import Cm
    from docx.enum.section import WD_ORIENT
    sec = wd.sections[0]
    w, h = PAGE_SIZES.get(setup.size, PAGE_SIZES["A4"])
    if setup.orientation == "landscape":
        w, h = h, w
        sec.orientation = WD_ORIENT.LANDSCAPE
    else:
        sec.orientation = WD_ORIENT.PORTRAIT
    sec.page_width, sec.page_height = Cm(w), Cm(h)
    sec.top_margin, sec.bottom_margin = Cm(setup.margin_top), Cm(setup.margin_bottom)
    sec.left_margin, sec.right_margin = Cm(setup.margin_left), Cm(setup.margin_right)
    if setup.columns > 1:
        cols = sec._sectPr.xpath("./w:cols")[0]
        cols.set("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}num", str(setup.columns))
    return sec


def _add_page_number(paragraph):
    """Append a real PAGE field so Word shows live page numbers."""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    run = paragraph.add_run()
    f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = "PAGE"
    f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "end")
    run._r.append(f1); run._r.append(it); run._r.append(f2)


def _add_toc_field(wd, toc: TOCField, doc: WordDoc):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    p = wd.add_paragraph()
    p.add_run(toc.title).bold = True
    fld = wd.add_paragraph()
    run = fld.add_run()
    f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve")
    it.text = f'TOC \\o "1-{toc.levels}" \\h \\z \\u'
    f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "separate")
    f3 = OxmlElement("w:fldChar"); f3.set(qn("w:fldCharType"), "end")
    run._r.append(f1); run._r.append(it); run._r.append(f2)
    # Static fallback so the file looks complete before Word updates fields.
    for style, text, page in doc.build_toc(toc.levels):
        lvl = int(style.split()[-1]) if style.split()[-1].isdigit() else 1
        wd.add_paragraph(f"{'    ' * (lvl - 1)}{text}\t{page}", style="Normal")
    run2 = fld.add_run(); run2._r.append(f3)


# ---------------- DOCX ----------------
def save_docx(doc: WordDoc, path: str):
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH as WD_ALIGN
    wd = Document()
    core = wd.core_properties
    core.title, core.author = doc.title, doc.author
    core.subject, core.keywords = doc.subject, doc.keywords
    sec = _set_section(wd, doc.setup)
    if doc.hf.header or doc.watermark:
        hp = sec.header.paragraphs[0]
        if doc.watermark:
            r = hp.add_run(f"{doc.watermark}  "); r.font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)
        if doc.hf.header:
            hp.add_run(doc.hf.header)
    if doc.hf.footer or doc.hf.page_numbers:
        fp = sec.footer.paragraphs[0]
        if doc.hf.footer:
            fp.add_run(doc.hf.footer + "  ")
        if doc.hf.page_numbers:
            _add_page_number(fp)
    amap = {"left": WD_ALIGN.LEFT, "center": WD_ALIGN.CENTER,
            "right": WD_ALIGN.RIGHT, "justify": WD_ALIGN.JUSTIFY}
    for b in doc.blocks:
        if isinstance(b, PageBreak):
            wd.add_page_break()
            continue
        if isinstance(b, TOCField):
            _add_toc_field(wd, b, doc)
            continue
        if isinstance(b, Table):
            t = wd.add_table(rows=len(b.rows), cols=max((len(r) for r in b.rows), default=1))
            t.style = "Table Grid"
            for i, row in enumerate(b.rows):
                for j, cell in enumerate(row):
                    if j < len(t.rows[i].cells):
                        t.rows[i].cells[j].text = cell
                        if b.header and i == 0:
                            for p in t.rows[i].cells[j].paragraphs:
                                for r in p.runs:
                                    r.bold = True
            if b.caption:
                wd.add_paragraph(b.caption, style="Caption")
            continue
        p = wd.add_paragraph(style=STYLE2DOCX.get(b.style, "Normal"))
        p.alignment = amap.get(b.align, WD_ALIGN.LEFT)
        try:
            pf = p.paragraph_format
            pf.space_after = Pt(b.spacing_after); pf.space_before = Pt(b.spacing_before)
            pf.line_spacing = b.line_spacing
            from docx.shared import Cm as _Cm
            pf.left_indent = _Cm(b.indent_left * 0.5); pf.right_indent = _Cm(b.indent_right * 0.5)
            pf.first_line_indent = _Cm(b.first_line * 0.5)
        except Exception:
            pass
        for r in b.runs:
            if isinstance(r, FootnoteRef):
                rr = p.add_run(f"[{r.nid}]")
                rr.font.superscript = True
                continue
            rr = p.add_run(r.text if r.text else "")
            rr.bold, rr.italic, rr.underline = r.bold, r.italic, r.underline or False
            try:
                rr.font.strike = r.strike
                rr.font.superscript, rr.font.subscript = r.super, r.sub
                rr.font.small_caps = r.small_caps
                if r.size: rr.font.size = Pt(r.size)
                if r.font: rr.font.name = r.font
                if r.color:
                    h = r.color.strip("#")
                    rr.font.color.rgb = RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
            except Exception:
                pass
    if doc.footnotes or doc.endnotes:
        wd.add_paragraph("Footnotes", style="Heading 1")
        for fn in doc.footnotes + doc.endnotes:
            wd.add_paragraph(f"[{fn.nid}] {fn.text}")
    wd.save(path)


def load_docx(path: str) -> WordDoc:
    from docx import Document
    wd = Document(path)
    doc = WordDoc(title=getattr(wd.core_properties, "title", "") or "Imported document")
    doc.author = getattr(wd.core_properties, "author", "") or doc.author
    doc.subject = getattr(wd.core_properties, "subject", "") or ""
    try:
        sec = wd.sections[0]
        from docx.shared import Cm as _Cm
        doc.setup.margin_top = sec.top_margin.cm; doc.setup.margin_bottom = sec.bottom_margin.cm
        doc.setup.margin_left = sec.left_margin.cm; doc.setup.margin_right = sec.right_margin.cm
        doc.setup.orientation = "landscape" if sec.page_width > sec.page_height else "portrait"
        try:
            doc.hf.header = sec.header.paragraphs[0].text
            doc.hf.footer = sec.footer.paragraphs[0].text
        except Exception:
            pass
    except Exception:
        pass
    rev = {v: k for k, v in STYLE2DOCX.items()}
    in_notes = False
    for p in wd.paragraphs:
        if not p.text and not p.runs:
            continue
        if p.text.strip().lower() == "footnotes":
            in_notes = True
            continue
        m = re.match(r"\[(\d+)\]\s+(.*)", p.text.strip())
        if in_notes and m:
            from engine import Footnote
            doc.footnotes.append(Footnote(int(m.group(1)), m.group(2)))
            continue
        style = rev.get(p.style.name if p.style else "Normal", "Normal")
        try:
            al = p.alignment
            align = {0: "left", 1: "center", 2: "right", 3: "justify"}.get(int(al), "left") if al is not None else "left"
        except Exception:
            align = "left"
        runs = []
        for r in p.runs:
            if "TOC" in r.text or "PAGEREF" in r.text:
                continue
            try: color = "%02X%02X%02X" % (r.font.color.rgb[0], r.font.color.rgb[1], r.font.color.rgb[2]) if r.font.color and r.font.color.rgb else None
            except Exception: color = None
            try: size = int(r.font.size.pt) if r.font.size else None
            except Exception: size = None
            runs.append(Run(r.text, bold=bool(r.bold), italic=bool(r.italic),
                            underline=bool(r.underline), strike=bool(getattr(r.font, "strike", False)),
                            super=bool(getattr(r.font, "superscript", False)),
                            sub=bool(getattr(r.font, "subscript", False)),
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
        if isinstance(b, PageBreak):
            out.append("\n---\n")
        elif isinstance(b, TOCField):
            out.append(f"## {b.title}\n")
            for _, text, page in doc.build_toc(b.levels):
                out.append(f"- {text} ... {page}")
            out.append("")
        elif isinstance(b, Table):
            if b.rows:
                if b.caption: out.append(f"*{b.caption}*")
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
            else: out.append(t)
            out.append("")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(out))

def save_html(doc: WordDoc, path: str):
    parts = [f"<html lang='en'><head><meta charset='utf-8'><title>{ihtml.escape(doc.title)}</title>",
             f"<style>body{{font-family:'Segoe UI',Arial,sans-serif;max-width:800px;margin:40px auto;line-height:1.8;padding:0 16px;color:#111}}"
             f"h1{{color:#111}}h2{{color:{BRAND_BLUE}}}table{{border-collapse:collapse;width:100%}}"
             f"td,th{{border:1px solid #999;padding:6px}}th{{background:#e8eefc}}"
             f"blockquote{{border-left:4px solid {BRAND_BLUE};background:#f3f6ff;padding:8px 14px;margin:8px 0}}"
             f"header,footer{{color:#555;border:1px dashed #bbb;padding:4px 8px;margin:8px 0}}"
             f".wm{{position:fixed;top:40%;left:50%;transform:translate(-50%,-50%) rotate(-30deg);"
             f"font-size:90px;color:rgba(37,99,235,.08);pointer-events:none}}</style></head><body>"]
    if doc.watermark:
        parts.append(f"<div class='wm'>{ihtml.escape(doc.watermark)}</div>")
    if doc.hf.header:
        parts.append(f"<header>{ihtml.escape(doc.hf.header)}</header>")
    parts.append(f"<h1>{ihtml.escape(doc.title)}</h1>")
    for b in doc.blocks:
        if isinstance(b, PageBreak):
            parts.append("<hr style='page-break-after:always'>")
        elif isinstance(b, TOCField):
            parts.append(f"<h2>{ihtml.escape(b.title)}</h2><ul>")
            for _, text, page in doc.build_toc(b.levels):
                parts.append(f"<li>{ihtml.escape(text)} ... {page}</li>")
            parts.append("</ul>")
        elif isinstance(b, Table):
            if b.caption: parts.append(f"<p><i>{ihtml.escape(b.caption)}</i></p>")
            parts.append("<table>")
            for i, row in enumerate(b.rows):
                tag = "th" if (b.header and i == 0) else "td"
                parts.append("<tr>" + "".join(f"<{tag}>{ihtml.escape(c)}</{tag}>" for c in row) + "</tr>")
            parts.append("</table>")
        else:
            t = "".join(f"<b>{ihtml.escape(r.text)}</b>" if r.bold else ihtml.escape(r.text)
                        for r in b.runs if isinstance(r, Run))
            t += "".join(f"<sup>[{r.nid}]</sup>" for r in b.runs if isinstance(r, FootnoteRef))
            if b.style == "Heading 1": parts.append(f"<h2>{t}</h2>")
            elif b.style == "Heading 2": parts.append(f"<h3>{t}</h3>")
            elif b.style == "Heading 3": parts.append(f"<h4>{t}</h4>")
            elif b.style == "Quote": parts.append(f"<blockquote>{t}</blockquote>")
            elif b.style == "List Bullet": parts.append(f"<ul><li>{t}</li></ul>")
            elif b.style == "List Number": parts.append(f"<ol><li>{t}</li></ol>")
            else: parts.append(f"<p style='text-align:{b.align}'>{t}</p>")
    if doc.footnotes:
        parts.append("<h2>Footnotes</h2><ol>")
        parts += [f"<li>{ihtml.escape(fn.text)}</li>" for fn in doc.footnotes]
        parts.append("</ol>")
    if doc.hf.footer:
        parts.append(f"<footer>{ihtml.escape(doc.hf.footer)}</footer>")
    parts.append("<footer>Created with WordSoft — Copyright (c) 2026 salim-slimani</footer></body></html>")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(parts))

# ---------------- PDF ----------------
def _shape(text: str) -> str:
    try:
        import arabic_reshaper
        from bidi.algorithm import get_display
        return get_display(arabic_reshaper.reshape(text))
    except Exception:
        return text

def save_pdf(doc: WordDoc, path: str):
    from reportlab.lib.pagesizes import A4, LETTER, LEGAL, landscape
    from reportlab.lib.units import mm, cm
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph as P, Spacer, Table as PT, TableStyle, PageBreak as PDFBreak
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
    import html as H
    pagesize = {"A4": A4, "Letter": LETTER, "Legal": LEGAL}.get(doc.setup.size, A4)
    if doc.setup.orientation == "landscape":
        pagesize = landscape(pagesize)
    amap = {"left": TA_LEFT, "center": TA_CENTER, "right": TA_RIGHT, "justify": TA_JUSTIFY}
    import os
    font, fontB = "Helvetica", "Helvetica-Bold"
    for cand in ["C:/Windows/Fonts/arial.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"]:
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

    def frame(canvas, _doc):
        canvas.saveState()
        if doc.watermark:
            canvas.setFont(fontB, 72); canvas.setFillGray(0.9)
            canvas.rotate(30)
            canvas.drawCentredString(0, 0, _shape(doc.watermark))
            canvas.rotate(-30)
        canvas.setFont(font, 9); canvas.setFillGray(0.4)
        w, h = pagesize
        if doc.hf.header:
            canvas.drawString(20 * mm, h - 12 * mm, _shape(doc.hf.header))
        foot = doc.hf.footer + ("  " if doc.hf.footer else "")
        if doc.hf.page_numbers:
            foot += f"Page {canvas.getPageNumber()}"
        if foot:
            canvas.drawCentredString(w / 2, 12 * mm, _shape(foot))
        canvas.restoreState()

    st_title = ParagraphStyle("t", fontName=fontB, fontSize=22, alignment=TA_CENTER, spaceAfter=12)
    st_h1 = ParagraphStyle("h1", fontName=fontB, fontSize=16, spaceBefore=14, spaceAfter=6)
    st_h2 = ParagraphStyle("h2", fontName=fontB, fontSize=13, spaceBefore=10, spaceAfter=4)
    st_n = ParagraphStyle("n", fontName=font, fontSize=11, leading=17, alignment=TA_JUSTIFY, spaceAfter=6)
    story = [P(H.escape(_shape(doc.title)), st_title), Spacer(1, 4 * mm)]
    for b in doc.blocks:
        if isinstance(b, PageBreak):
            story.append(PDFBreak()); continue
        if isinstance(b, TOCField):
            story.append(P(H.escape(_shape(b.title)), st_h1))
            for _, text, page in doc.build_toc(b.levels):
                story.append(P(f"{H.escape(_shape(text))} ... {page}", st_n))
            continue
        if isinstance(b, Table):
            data = [[H.escape(_shape(c)) for c in row] for row in b.rows]
            t = PT(data, repeatRows=1 if b.header else 0)
            t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.6, colors.grey),
                                   ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e8eefc")),
                                   ("FONTSIZE", (0, 0), (-1, -1), 9), ("FONTNAME", (0, 0), (-1, -1), font)]))
            story.append(t); story.append(Spacer(1, 3 * mm))
            continue
        txt = "".join(f"<b>{H.escape(_shape(r.text))}</b>" if r.bold else H.escape(_shape(r.text))
                      for r in b.runs if isinstance(r, Run)) or " "
        txt += "".join(f"<super>[{r.nid}]</super>" for r in b.runs if isinstance(r, FootnoteRef))
        if b.style == "Title": story.append(P(txt, st_title))
        elif b.style == "Heading 1": story.append(P(txt, st_h1))
        elif b.style == "Heading 2": story.append(P(txt, st_h2))
        else:
            story.append(P(txt, ParagraphStyle(f"p{id(b)}", parent=st_n, alignment=amap.get(b.align, TA_JUSTIFY))))
    if doc.footnotes:
        story.append(P("Footnotes", st_h1))
        for fn in doc.footnotes:
            story.append(P(f"[{fn.nid}] {H.escape(_shape(fn.text))}", st_n))
    SimpleDocTemplate(path, pagesize=pagesize,
                      topMargin=doc.setup.margin_top * cm, bottomMargin=doc.setup.margin_bottom * cm,
                      leftMargin=doc.setup.margin_left * cm, rightMargin=doc.setup.margin_right * cm,
                      title=doc.title, author=doc.author,
                      subject=doc.subject, keywords=doc.keywords).build(story, onFirstPage=frame, onLaterPages=frame)

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
