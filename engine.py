# -*- coding: utf-8 -*-
"""WordSoft engine — UI-independent document model.

Blocks: paragraphs (styled runs), tables, page breaks, TOC fields,
footnote references. Document-level: page setup, theme, headers/footers,
watermark, footnotes, endnotes, comments, tracked changes, citations,
captions, index entries.

Copyright (c) 2026 salim-slimani. MIT License.
"""
from __future__ import annotations
import re
from dataclasses import dataclass, field


# ---------------- inline ----------------
@dataclass
class Run:
    text: str = ""
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strike: bool = False
    color: str | None = None        # hex, e.g. FF0000
    highlight: str | None = None    # hex background
    font: str | None = None
    size: int | None = None         # pt
    super: bool = False             # superscript
    sub: bool = False               # subscript
    small_caps: bool = False


# ---------------- blocks ----------------
@dataclass
class Paragraph:
    runs: list[Run] = field(default_factory=list)
    style: str = "Normal"
    align: str = "left"             # left, center, right, justify
    direction: str = "auto"
    spacing_before: int = 0
    spacing_after: int = 6
    line_spacing: float = 1.15
    indent_left: int = 0
    indent_right: int = 0
    first_line: int = 0
    shading: str | None = None      # hex background
    page_break_before: bool = False

    def text(self) -> str:
        return "".join(r.text if isinstance(r, Run) else f"[{r.nid}]" for r in self.runs)


@dataclass
class Table:
    rows: list[list[str]] = field(default_factory=list)
    style: str = "Table Grid"
    header: bool = True
    caption: str | None = None


@dataclass
class PageBreak:
    pass


@dataclass
class TOCField:
    title: str = "Contents"
    levels: int = 3


@dataclass
class FootnoteRef:
    nid: int = 0


# ---------------- document-level records ----------------
@dataclass
class PageSetup:
    size: str = "A4"                # A4, Letter, Legal
    orientation: str = "portrait"   # portrait, landscape
    margin_top: float = 2.54        # cm
    margin_bottom: float = 2.54
    margin_left: float = 2.54
    margin_right: float = 2.54
    columns: int = 1
    line_numbers: bool = False
    hyphenation: bool = False


@dataclass
class Theme:
    name: str = "Office"
    heading_font: str = "Calibri Light"
    body_font: str = "Calibri"
    accent: str = "2563EB"


THEMES = {
    "Office": Theme("Office", "Calibri Light", "Calibri", "2563EB"),
    "Blue": Theme("Blue", "Arial", "Arial", "1D4ED8"),
    "Facet": Theme("Facet", "Trebuchet MS", "Trebuchet MS", "047857"),
    "Gallery": Theme("Gallery", "Georgia", "Georgia", "7C3AED"),
    "Sunset": Theme("Sunset", "Verdana", "Verdana", "C2410C"),
    "Mono": Theme("Mono", "Consolas", "Consolas", "334155"),
}

PAGE_SIZES = {"A4": (21.0, 29.7), "Letter": (21.59, 27.94), "Legal": (21.59, 35.56)}


@dataclass
class HeaderFooter:
    header: str = ""
    footer: str = ""
    first_page_different: bool = False
    first_header: str = ""
    first_footer: str = ""
    page_numbers: bool = True


@dataclass
class Footnote:
    nid: int = 0
    text: str = ""


@dataclass
class Comment:
    cid: int = 0
    author: str = "salim-slimani"
    quote: str = ""
    text: str = ""
    resolved: bool = False


@dataclass
class Change:
    author: str = "salim-slimani"
    kind: str = "insert"            # insert, delete, format
    text: str = ""


@dataclass
class Citation:
    key: str = ""
    author: str = ""
    title: str = ""
    year: str = ""
    publisher: str = ""

    def apa(self) -> str:
        return f"{self.author} ({self.year}). {self.title}. {self.publisher}."

    def mla(self) -> str:
        return f"{self.author}. {self.title}. {self.publisher}, {self.year}."

    def chicago(self) -> str:
        return f"{self.author}. {self.year}. {self.title}. {self.publisher}."

    def ieee(self, n: int) -> str:
        return f"[{n}] {self.author}, {self.title}, {self.publisher}, {self.year}."


@dataclass
class IndexEntry:
    term: str = ""
    page: int = 1


# ---------------- document ----------------
@dataclass
class WordDoc:
    title: str = "Untitled document"
    author: str = "salim-slimani"
    subject: str = ""
    keywords: str = ""
    language: str = "en-US"
    blocks: list = field(default_factory=list)
    setup: PageSetup = field(default_factory=PageSetup)
    theme: Theme = field(default_factory=lambda: THEMES["Office"])
    hf: HeaderFooter = field(default_factory=HeaderFooter)
    watermark: str = ""
    page_color: str = "FFFFFF"
    footnotes: list[Footnote] = field(default_factory=list)
    endnotes: list[Footnote] = field(default_factory=list)
    comments: list[Comment] = field(default_factory=list)
    changes: list[Change] = field(default_factory=list)
    citations: list[Citation] = field(default_factory=list)
    index: list[IndexEntry] = field(default_factory=list)
    track_changes: bool = False
    protection: str | None = None   # password or None

    # ---------- constructors ----------
    @classmethod
    def blank(cls, title="Untitled document") -> "WordDoc":
        d = cls(title=title)
        d.blocks.append(Paragraph([Run("")], style="Normal"))
        return d

    @classmethod
    def sample(cls) -> "WordDoc":
        d = cls(title="The Open Studio Handbook — WordSoft")
        d.blocks.append(Paragraph([Run("The Open Studio Handbook", bold=True, size=28)], style="Title", align="center"))
        d.blocks.append(Paragraph([Run("Write beautifully — a fast, friendly word processor", size=13)], style="Subtitle", align="center"))
        d.blocks.append(Paragraph([Run("Introduction", bold=True, size=18)], style="Heading 1"))
        d.blocks.append(Paragraph([Run("Welcome to "), Run("WordSoft", bold=True), Run(" — a lightweight word processor that just works: ribbon tabs, styles, tables, references, mail merge, review tools, and export to DOCX and PDF. Try selecting this sentence and pressing "), Run("Ctrl+B", bold=True), Run(" for bold.")]))
        d.blocks.append(Paragraph([Run("Styles and structure", bold=True, size=18)], style="Heading 1"))
        d.blocks.append(Paragraph([Run("Use the style gallery on the Home tab: Title, Heading 1, Heading 2, Quote, captions, and bulleted or numbered lists. Headings feed the navigation pane and the automatic table of contents.")], style="Normal"))
        d.blocks.append(Paragraph([Run("Working with references", bold=True, size=15)], style="Heading 2"))
        d.blocks.append(Paragraph([Run("Insert footnotes, citations, captions and a table of contents from the References tab.", italic=True)], style="Quote"))
        d.blocks.append(Paragraph([Run("Fast typing with undo and redo")], style="List Bullet"))
        d.blocks.append(Paragraph([Run("DOCX open and save that works with Word")], style="List Bullet"))
        d.blocks.append(Paragraph([Run("Open the Insert tab and add a table")], style="List Number"))
        d.blocks.append(Paragraph([Run("Export to PDF from the File backstage")], style="List Number"))
        d.blocks.append(Paragraph([Run("Sample table", bold=True, size=18)], style="Heading 1"))
        d.blocks.append(Table(
            rows=[["Item", "Description", "Status"], ["Styles", "headings + quotes + lists", "Works"],
                  ["Tables", "insert, edit and export", "Works"], ["Mail merge", "CSV fields + preview", "Works"]],
            header=True, caption="Table 1 — Feature overview"))
        d.blocks.append(Paragraph([Run("Review and finish", bold=True, size=18)], style="Heading 1"))
        d.blocks.append(Paragraph([Run("Check spelling with F7, count words, add comments, track changes, then export a polished PDF.")]))
        d.footnotes.append(Footnote(1, "WordSoft footnotes travel with the document and export to DOCX and PDF."))
        d.blocks.append(Paragraph([Run("A word about footnotes"), FootnoteRef(1)], style="Normal"))
        return d

    @classmethod
    def template(cls, kind: str) -> "WordDoc":
        if kind == "sample":
            return cls.sample()
        if kind == "cover":
            d = cls(title="Cover page")
            d.blocks.append(Paragraph([Run("")], style="Normal"))
            d.blocks.append(Paragraph([Run("DOCUMENT TITLE", bold=True, size=36)], style="Title", align="center"))
            d.blocks.append(Paragraph([Run("Subtitle goes here", size=16)], style="Subtitle", align="center"))
            d.blocks.append(Paragraph([Run("Author name — 2026", size=12)], style="Normal", align="center"))
            d.blocks.append(PageBreak())
            d.blocks.append(Paragraph([Run("")], style="Normal"))
            return d
        d = cls(title={"letter": "Formal letter", "report": "Report", "resume": "Resume",
                       "envelope": "Envelope", "label": "Labels"}.get(kind, "Untitled document"))
        if kind == "letter":
            d.blocks.append(Paragraph([Run("Full name", bold=True, size=16)], align="left"))
            d.blocks.append(Paragraph([Run("Street address — City")], align="left"))
            d.blocks.append(Paragraph([Run("Date: .... / .... / 2026")], align="left"))
            d.blocks.append(Paragraph([Run("Subject: ....................", bold=True)], align="left"))
            d.blocks.append(Paragraph([Run("Dear Sir or Madam,")], align="left"))
            d.blocks.append(Paragraph([Run("I am writing to...")], align="left"))
            d.blocks.append(Paragraph([Run("Yours faithfully,")], align="left"))
            d.blocks.append(Paragraph([Run("Signature", italic=True)], align="left"))
        elif kind == "report":
            d.blocks.append(Paragraph([Run("Report", bold=True, size=26)], style="Title", align="center"))
            d.blocks.append(TOCField("Contents", 2))
            d.blocks.append(Paragraph([Run("Executive summary", bold=True, size=18)], style="Heading 1"))
            d.blocks.append(Paragraph([Run("Write the summary here...")]))
            d.blocks.append(Paragraph([Run("Findings", bold=True, size=18)], style="Heading 1"))
            d.blocks.append(Paragraph([Run("First finding...")], style="List Bullet"))
            d.blocks.append(Paragraph([Run("Recommendations", bold=True, size=18)], style="Heading 1"))
        elif kind == "resume":
            d.blocks.append(Paragraph([Run("Full name", bold=True, size=24)], style="Title", align="center"))
            d.blocks.append(Paragraph([Run("Email | Phone | Address")], align="center"))
            d.blocks.append(Paragraph([Run("Experience", bold=True, size=18)], style="Heading 1"))
            d.blocks.append(Paragraph([Run("Job title — Company (Year)")], style="List Bullet"))
            d.blocks.append(Paragraph([Run("Skills", bold=True, size=18)], style="Heading 1"))
            d.blocks.append(Paragraph([Run("Skill one, skill two")], style="List Bullet"))
        elif kind == "envelope":
            d.blocks.append(Paragraph([Run("Sender name\nStreet\nCity", size=11)], align="left"))
            d.blocks.append(Paragraph([Run("")], style="Normal"))
            d.blocks.append(Paragraph([Run("Recipient name\nStreet\nCity", size=12)], align="center"))
        elif kind == "label":
            d.blocks.append(Table(rows=[["Name\nStreet\nCity"] * 3] * 5, header=False))
        else:
            return cls.blank()
        return d

    # ---------- text & stats ----------
    def full_text(self) -> str:
        out = []
        for b in self.blocks:
            if isinstance(b, Paragraph):
                out.append(b.text())
            elif isinstance(b, Table):
                for row in b.rows:
                    out.append(" | ".join(row))
            elif isinstance(b, PageBreak):
                out.append("\f")
            elif isinstance(b, TOCField):
                out.append(b.title)
        return "\n".join(out)

    def stats(self) -> dict:
        text = self.full_text().replace("\f", "\n")
        words = re.findall(r"\S+", text)
        chars = len(text.replace("\n", ""))
        chars_ns = len(re.sub(r"\s", "", text))
        paras = sum(1 for b in self.blocks if isinstance(b, Paragraph) and b.text().strip())
        lines = max(1, len(text.split("\n")))
        pages = max(1, round(len(words) / 400 + 0.5))
        return {"words": len(words), "chars": chars, "chars_no_spaces": chars_ns,
                "paragraphs": paras, "lines": lines,
                "tables": sum(1 for b in self.blocks if isinstance(b, Table)),
                "pages": pages, "footnotes": len(self.footnotes),
                "comments": len(self.comments)}

    def headings(self, max_level: int = 3) -> list[tuple[str, str]]:
        out = []
        for b in self.blocks:
            if isinstance(b, Paragraph) and b.style.startswith("Heading") and b.text().strip():
                try:
                    lvl = int(b.style.split()[-1])
                except Exception:
                    lvl = 1
                if lvl <= max_level:
                    out.append((b.style, b.text().strip()))
        return out

    # ---------- references ----------
    def build_toc(self, levels: int = 3) -> list[tuple[str, str, int]]:
        """(style, text, page estimate) rows for a table of contents."""
        rows = []
        words = 0
        for b in self.blocks:
            if isinstance(b, Paragraph):
                words += len(re.findall(r"\S+", b.text()))
                if b.style.startswith("Heading") and b.text().strip():
                    try:
                        lvl = int(b.style.split()[-1])
                    except Exception:
                        lvl = 1
                    if lvl <= levels:
                        rows.append((b.style, b.text().strip(), max(1, words // 400 + 1)))
        return rows

    def bibliography(self, style: str = "APA") -> list[str]:
        out = []
        for i, c in enumerate(self.citations, 1):
            s = style.upper()
            if s == "MLA": out.append(c.mla())
            elif s == "CHICAGO": out.append(c.chicago())
            elif s == "IEEE": out.append(c.ieee(i))
            else: out.append(c.apa())
        return out

    def build_index(self) -> list[tuple[str, list[int]]]:
        by_term: dict[str, list[int]] = {}
        for e in self.index:
            by_term.setdefault(e.term, []).append(e.page)
        return sorted(by_term.items(), key=lambda kv: kv[0].lower())

    # ---------- mail merge ----------
    @staticmethod
    def load_csv(path: str) -> tuple[list[str], list[dict]]:
        import csv
        with open(path, encoding="utf-8-sig", newline="") as f:
            rows = list(csv.DictReader(f))
        fields = list(rows[0].keys()) if rows else []
        return fields, rows

    def merged(self, record: dict) -> "WordDoc":
        """Return a copy with <<Field>> placeholders replaced by record values."""
        import copy
        d = copy.deepcopy(self)
        pat = re.compile(r"<<([^<>]+)>>")
        for b in d.blocks:
            if isinstance(b, Paragraph):
                for r in b.runs:
                    if isinstance(r, Run):
                        r.text = pat.sub(lambda m: str(record.get(m.group(1).strip(), m.group(0))), r.text)
        return d

    def merge_all(self, records: list[dict]) -> "WordDoc":
        import copy
        out = WordDoc(title=self.title + " (merged)", author=self.author)
        out.setup = copy.deepcopy(self.setup)
        out.theme = copy.deepcopy(self.theme)
        out.hf = copy.deepcopy(self.hf)
        for i, rec in enumerate(records):
            m = self.merged(rec)
            out.blocks.extend(copy.deepcopy(m.blocks))
            if i < len(records) - 1:
                out.blocks.append(PageBreak())
        return out
