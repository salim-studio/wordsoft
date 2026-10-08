# -*- coding: utf-8 -*-
"""WordSoft command registry — every action is a command with an id.

The same ids drive the ribbon, keyboard shortcuts, the CLI batch runner
and recorded macros. Usage: COMMANDS["insert.toc"](doc).

Copyright (c) 2026 salim-slimani. MIT License.
"""
from __future__ import annotations
from engine import WordDoc, Paragraph, Run, Table, PageBreak, TOCField, FootnoteRef, Footnote, Comment, Change

COMMANDS: dict[str, dict] = {}


def command(cid: str, title: str, tab: str = "Home"):
    def deco(fn):
        COMMANDS[cid] = {"id": cid, "title": title, "tab": tab, "run": fn}
        return fn
    return deco


def list_commands(tab: str | None = None) -> list[dict]:
    return [{"id": c["id"], "title": c["title"], "tab": c["tab"]}
            for c in COMMANDS.values() if tab is None or c["tab"] == tab]


def execute(doc: WordDoc, cid: str, **kw):
    if cid not in COMMANDS:
        raise KeyError(f"Unknown command: {cid}")
    return COMMANDS[cid]["run"](doc, **kw)


def batch(doc: WordDoc, steps: list[dict]) -> WordDoc:
    for s in steps:
        execute(doc, s["cmd"], **{k: v for k, v in s.items() if k != "cmd"})
    return doc


# ---------------- file ----------------
@command("file.new", "New blank document", "File")
def _new(doc: WordDoc, template: str = "blank") -> WordDoc:
    return WordDoc.template(template)


@command("file.stats", "Document statistics", "File")
def _stats(doc: WordDoc) -> dict:
    return doc.stats()


# ---------------- home ----------------
@command("format.clear", "Clear all formatting", "Home")
def _clear(doc: WordDoc) -> WordDoc:
    for b in doc.blocks:
        if isinstance(b, Paragraph):
            for r in b.runs:
                if isinstance(r, Run):
                    r.bold = r.italic = r.underline = r.strike = False
                    r.color = r.highlight = None
    return doc


@command("format.styleAll", "Apply style to every paragraph", "Home")
def _style_all(doc: WordDoc, style: str = "Normal") -> WordDoc:
    for b in doc.blocks:
        if isinstance(b, Paragraph):
            b.style = style
    return doc


@command("format.case", "Change case", "Home")
def _case(doc: WordDoc, mode: str = "upper") -> WordDoc:
    for b in doc.blocks:
        if isinstance(b, Paragraph):
            for r in b.runs:
                if isinstance(r, Run):
                    if mode == "upper": r.text = r.text.upper()
                    elif mode == "lower": r.text = r.text.lower()
                    elif mode == "title": r.text = r.text.title()
    return doc


# ---------------- insert ----------------
@command("insert.pageBreak", "Insert page break", "Insert")
def _pb(doc: WordDoc) -> WordDoc:
    doc.blocks.append(PageBreak())
    doc.blocks.append(Paragraph([Run("")]))
    return doc


@command("insert.table", "Insert table", "Insert")
def _table(doc: WordDoc, rows: int = 3, cols: int = 3) -> WordDoc:
    data = [[f"Header {j + 1}" if i == 0 else "" for j in range(cols)] for i in range(rows)]
    doc.blocks.append(Table(rows=data, header=True))
    return doc


@command("insert.toc", "Insert table of contents", "References")
def _toc(doc: WordDoc, levels: int = 3) -> WordDoc:
    doc.blocks.insert(0, TOCField("Contents", levels))
    return doc


@command("insert.footnote", "Insert footnote", "References")
def _fn(doc: WordDoc, text: str = "Footnote text.") -> WordDoc:
    nid = len(doc.footnotes) + len(doc.endnotes) + 1
    doc.footnotes.append(Footnote(nid, text))
    doc.blocks.append(Paragraph([Run("See note"), FootnoteRef(nid)]))
    return doc


@command("insert.comment", "Insert comment", "Review")
def _comment(doc: WordDoc, quote: str = "", text: str = "Comment.") -> WordDoc:
    doc.comments.append(Comment(len(doc.comments) + 1, doc.author, quote, text))
    return doc


@command("review.track", "Toggle track changes", "Review")
def _track(doc: WordDoc, on: bool = True) -> WordDoc:
    doc.track_changes = on
    return doc


@command("review.acceptAll", "Accept all changes", "Review")
def _accept(doc: WordDoc) -> WordDoc:
    doc.changes.clear()
    for c in doc.comments:
        c.resolved = True
    return doc


@command("design.watermark", "Set watermark", "Design")
def _wm(doc: WordDoc, text: str = "DRAFT") -> WordDoc:
    doc.watermark = text
    return doc


@command("design.theme", "Apply theme", "Design")
def _theme(doc: WordDoc, name: str = "Office") -> WordDoc:
    from engine import THEMES
    if name in THEMES:
        doc.theme = THEMES[name]
    return doc


@command("layout.margins", "Set margins (cm)", "Layout")
def _margins(doc: WordDoc, top: float = 2.54, bottom: float = 2.54,
             left: float = 2.54, right: float = 2.54) -> WordDoc:
    doc.setup.margin_top, doc.setup.margin_bottom = top, bottom
    doc.setup.margin_left, doc.setup.margin_right = left, right
    return doc


@command("layout.orientation", "Set orientation", "Layout")
def _orient(doc: WordDoc, orientation: str = "portrait") -> WordDoc:
    doc.setup.orientation = orientation
    return doc


@command("mail.merge", "Finish mail merge from CSV", "Mailings")
def _merge(doc: WordDoc, csv: str = "", out: str = "merged.docx") -> WordDoc:
    fields, records = WordDoc.load_csv(csv)
    merged = doc.merge_all(records)
    import formats
    formats.save(merged, out)
    return merged


# ---------------- macros ----------------
class Macro:
    def __init__(self, name: str):
        self.name = name
        self.steps: list[dict] = []
        self.recording = False

    def record(self, cid: str, **kw):
        if self.recording:
            self.steps.append({"cmd": cid, **kw})

    def play(self, doc: WordDoc) -> WordDoc:
        return batch(doc, self.steps)

    def to_dict(self) -> dict:
        return {"name": self.name, "steps": self.steps}

    @staticmethod
    def from_dict(d: dict) -> "Macro":
        m = Macro(d.get("name", "macro"))
        m.steps = d.get("steps", [])
        return m
