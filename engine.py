# -*- coding: utf-8 -*-
"""WordSoft engine — UI-independent document model.

Supports: formatted paragraphs + runs, headings, lists, tables,
images (paths), and word counts.

Copyright (c) 2026 salim-slimani. MIT License.
"""
from __future__ import annotations
import re
import os
from dataclasses import dataclass, field, asdict


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


@dataclass
class Paragraph:
    runs: list[Run] = field(default_factory=list)
    style: str = "Normal"           # Normal, Title, Heading 1/2/3, Quote, List Bullet, List Number
    align: str = "left"             # left, center, right, justify
    direction: str = "auto"         # auto, rtl, ltr
    spacing_before: int = 0
    spacing_after: int = 6
    line_spacing: float = 1.15
    indent: int = 0

    def text(self) -> str:
        return "".join(r.text for r in self.runs)


@dataclass
class Table:
    rows: list[list[str]] = field(default_factory=list)
    style: str = "Table Grid"
    header: bool = True


Block = Paragraph | Table


@dataclass
class WordDoc:
    title: str = "Untitled document"
    author: str = "salim-slimani"
    blocks: list = field(default_factory=list)
    track_changes: bool = False

    # ---------- constructors ----------
    @classmethod
    def blank(cls, title="Untitled document") -> "WordDoc":
        d = cls(title=title)
        d.blocks.append(Paragraph([Run("")], style="Normal"))
        return d

    @classmethod
    def sample(cls) -> "WordDoc":
        """Built-in English sample document: 'The Open Studio Handbook'."""
        d = cls(title="The Open Studio Handbook — WordSoft")
        d.blocks.append(Paragraph([Run("The Open Studio Handbook", bold=True, size=28)], style="Title", align="center"))
        d.blocks.append(Paragraph([Run("Write beautifully — a fast, friendly word processor", size=13)], style="Subtitle", align="center"))
        d.blocks.append(Paragraph([Run("Introduction", bold=True, size=18)], style="Heading 1"))
        d.blocks.append(Paragraph([Run("Welcome to "), Run("WordSoft", bold=True), Run(" — a lightweight word processor that just works: the ribbon tabs, styles, tables, word counts, and export to DOCX and PDF. Try selecting this sentence and pressing "), Run("Ctrl+B", bold=True), Run(" for bold.")]))
        d.blocks.append(Paragraph([Run("Styles", bold=True, size=18)], style="Heading 1"))
        d.blocks.append(Paragraph([Run("Use the style gallery on the Home tab: Title, Heading 1, Heading 2, Quote, and bulleted or numbered lists. Every style previews live in the document.")], style="Normal"))
        d.blocks.append(Paragraph([Run("A sample subheading", bold=True, size=15)], style="Heading 2"))
        d.blocks.append(Paragraph([Run("This is a sample quote showing the Quote style in WordSoft.", italic=True)], style="Quote"))
        d.blocks.append(Paragraph([Run("First point: fast typing with undo and redo")], style="List Bullet"))
        d.blocks.append(Paragraph([Run("Second point: DOCX open and save that works with Word")], style="List Bullet"))
        d.blocks.append(Paragraph([Run("First step: open the Insert tab and add a table")], style="List Number"))
        d.blocks.append(Paragraph([Run("Second step: export to PDF from the File menu")], style="List Number"))
        d.blocks.append(Paragraph([Run("Sample table", bold=True, size=18)], style="Heading 1"))
        t = Table(rows=[["Item", "Description", "Status"], ["Styles", "headings + quotes + lists", "Works"], ["Tables", "insert, edit and export", "Works"], ["PDF", "real export with selectable text", "Works"]], header=True)
        d.blocks.append(t)
        d.blocks.append(Paragraph([Run("Conclusion", bold=True, size=18)], style="Heading 1"))
        d.blocks.append(Paragraph([Run("WordSoft is written in Python so it runs instantly with no build step. Every action is also available headlessly through the CLI, so documents can be automated.")]))
        return d

    @classmethod
    def template(cls, kind: str) -> "WordDoc":
        if kind == "sample":
            return cls.sample()
        d = cls(title={"letter": "Formal letter", "report": "Report", "resume": "Resume"}.get(kind, "Untitled document"))
        if kind == "letter":
            d.blocks.append(Paragraph([Run("Full name", bold=True, size=16)], align="left"))
            d.blocks.append(Paragraph([Run("Date: .... / .... / ......")], align="left"))
            d.blocks.append(Paragraph([Run("Subject: ....................", bold=True)], align="left"))
            d.blocks.append(Paragraph([Run("Dear Sir or Madam,")], align="left"))
            d.blocks.append(Paragraph([Run("I am writing to...")], align="left"))
            d.blocks.append(Paragraph([Run("Yours faithfully,")], align="left"))
            d.blocks.append(Paragraph([Run("Signature", italic=True)], align="left"))
        elif kind == "report":
            d.blocks.append(Paragraph([Run("Report", bold=True, size=26)], style="Title", align="center"))
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
        else:
            return cls.blank()
        return d

    # ---------- stats ----------
    def full_text(self) -> str:
        out = []
        for b in self.blocks:
            if isinstance(b, Paragraph):
                out.append(b.text())
            elif isinstance(b, Table):
                for row in b.rows:
                    out.append(" | ".join(row))
        return "\n".join(out)

    def stats(self) -> dict:
        text = self.full_text()
        words = re.findall(r"\S+", text)
        chars = len(text.replace("\n", ""))
        paras = sum(1 for b in self.blocks if isinstance(b, Paragraph) and b.text().strip())
        pages = max(1, round(len(words) / 400 + 0.5))
        return {"words": len(words), "chars": chars, "paragraphs": paras,
                "tables": sum(1 for b in self.blocks if isinstance(b, Table)),
                "pages": pages}

    def headings(self) -> list[tuple[str, str]]:
        """Navigation helper: (style, text) for headings only."""
        out = []
        for b in self.blocks:
            if isinstance(b, Paragraph) and b.style.startswith("Heading") and b.text().strip():
                out.append((b.style, b.text().strip()))
        return out
