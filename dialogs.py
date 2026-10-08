# -*- coding: utf-8 -*-
"""WordSoft dialogs — reusable Toplevel editors for tables, styles, page setup,
headers/footers, mail merge, citations, macros, inspector and more.

Each dialog returns plain data; the app applies it to the document.

Copyright (c) 2026 salim-slimani. MIT License.
"""
from __future__ import annotations
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from engine import Table, Citation, THEMES

SYMBOLS = list("Ω π Σ √ ∞ § © ® ™ € £ ¥ → ← ↑ ↓ • … — – ± × ÷ α β γ δ λ μ ∴ ∵ ∀ ∃ ∈ ∉ ★ ☆ ♠ ♥ ♦ ♣ ✔ ✘ § ¶")


def _dlg(parent, title, w=460, h=380):
    win = tk.Toplevel(parent)
    win.title(title)
    win.geometry(f"{w}x{h}")
    win.transient(parent)
    win.grab_set()
    return win


# ---------------- tables ----------------
def insert_table_dialog(parent) -> Table | None:
    win = _dlg(parent, "Insert table", 300, 180)
    r = tk.IntVar(value=3); c = tk.IntVar(value=3)
    ttk.Label(win, text="Rows:").pack(pady=(10, 0)); ttk.Spinbox(win, from_=1, to=20, textvariable=r, width=8).pack()
    ttk.Label(win, text="Columns:").pack(pady=(8, 0)); ttk.Spinbox(win, from_=1, to=10, textvariable=c, width=8).pack()
    out: dict = {}
    ttk.Button(win, text="Insert", command=lambda: (out.update(rows=r.get(), cols=c.get()), win.destroy())).pack(pady=12)
    win.wait_window()
    if not out: return None
    data = [[f"Header {j + 1}" if i == 0 else "" for j in range(out["cols"])] for i in range(out["rows"])]
    return Table(rows=data, header=True)


def edit_table_dialog(parent, table: Table) -> bool:
    """Spreadsheet-like cell editor. Returns True when applied."""
    win = _dlg(parent, f"Edit table ({len(table.rows)}x{len(table.rows[0]) if table.rows else 0})", 560, 420)
    frame = ttk.Frame(win); frame.pack(fill="both", expand=True, padx=8, pady=8)
    entries: list[list[tk.Entry]] = []

    def rebuild():
        for w in frame.winfo_children(): w.destroy()
        entries.clear()
        for i, row in enumerate(table.rows):
            erow = []
            for j, cell in enumerate(row):
                e = tk.Entry(frame, width=16)
                e.insert(0, cell); e.grid(row=i, column=j, padx=1, pady=1)
                erow.append(e)
            entries.append(erow)

    def apply():
        table.rows = [[e.get() for e in row] for row in entries]
        win.destroy(); win_saved.append(True)

    def add_row(): table.rows.append([""] * (len(table.rows[0]) if table.rows else 1)); rebuild()
    def add_col():
        for row in table.rows: row.append("")
        rebuild()
    def del_row():
        if len(table.rows) > 1: table.rows.pop(); rebuild()

    bar = ttk.Frame(win); bar.pack(fill="x", padx=8)
    ttk.Button(bar, text="+ Row", command=add_row).pack(side="left", padx=2)
    ttk.Button(bar, text="+ Column", command=add_col).pack(side="left", padx=2)
    ttk.Button(bar, text="− Row", command=del_row).pack(side="left", padx=2)
    hdr = tk.BooleanVar(value=table.header)
    ttk.Checkbutton(bar, text="Header row", variable=hdr,
                    command=lambda: setattr(table, "header", hdr.get())).pack(side="left", padx=6)
    ttk.Button(bar, text="Apply", command=apply).pack(side="right", padx=2)
    win_saved: list[bool] = []
    rebuild()
    win.wait_window()
    return bool(win_saved)


# ---------------- styles ----------------
BUILTIN_STYLES = ["Normal", "No Spacing", "Title", "Subtitle", "Heading 1", "Heading 2",
                  "Heading 3", "Heading 4", "Quote", "Intense Quote", "Caption",
                  "List Bullet", "List Number", "Subtitle"]


def style_manager_dialog(parent, apply_fn, custom: dict):
    win = _dlg(parent, "Styles", 380, 460)
    ttk.Label(win, text="Click a style to apply it to the current paragraph:").pack(pady=6)
    lb = tk.Listbox(win); lb.pack(fill="both", expand=True, padx=8)
    for s in BUILTIN_STYLES + sorted(k for k in custom if k not in BUILTIN_STYLES):
        lb.insert("end", s + ("  (custom)" if s in custom else ""))
    def go(_e=None):
        if lb.curselection():
            name = lb.get(lb.curselection()[0]).replace("  (custom)", "")
            apply_fn(name); win.destroy()
    lb.bind("<Double-Button-1>", go)
    ttk.Button(win, text="Apply", command=go).pack(pady=6)


def new_style_dialog(parent) -> dict | None:
    win = _dlg(parent, "New style", 320, 300)
    name = tk.StringVar(value="My Style")
    base = tk.StringVar(value="Normal")
    bold = tk.BooleanVar(); italic = tk.BooleanVar()
    size = tk.StringVar(value="12")
    ttk.Label(win, text="Name:").pack(); ttk.Entry(win, textvariable=name, width=28).pack()
    ttk.Label(win, text="Based on:").pack(pady=(6, 0))
    ttk.Combobox(win, textvariable=base, values=BUILTIN_STYLES, state="readonly", width=26).pack()
    ttk.Checkbutton(win, text="Bold", variable=bold).pack()
    ttk.Checkbutton(win, text="Italic", variable=italic).pack()
    ttk.Label(win, text="Size:").pack(); ttk.Entry(win, textvariable=size, width=8).pack()
    out: dict = {}
    ttk.Button(win, text="Create",
               command=lambda: (out.update(name=name.get(), base=base.get(), bold=bold.get(),
                                           italic=italic.get(), size=int(size.get() or 12)), win.destroy())).pack(pady=10)
    win.wait_window()
    return out or None


# ---------------- page setup / header-footer / watermark / theme ----------------
def page_setup_dialog(parent, setup) -> bool:
    win = _dlg(parent, "Page setup", 380, 460)
    size = tk.StringVar(value=setup.size); orient = tk.StringVar(value=setup.orientation)
    mt = tk.DoubleVar(value=setup.margin_top); mb = tk.DoubleVar(value=setup.margin_bottom)
    ml = tk.DoubleVar(value=setup.margin_left); mr = tk.DoubleVar(value=setup.margin_right)
    cols = tk.IntVar(value=setup.columns)
    ln = tk.BooleanVar(value=setup.line_numbers); hy = tk.BooleanVar(value=setup.hyphenation)
    ttk.Label(win, text="Paper size:").pack(); ttk.Combobox(win, textvariable=size, values=["A4", "Letter", "Legal"], state="readonly").pack()
    ttk.Label(win, text="Orientation:").pack(pady=(6, 0)); ttk.Combobox(win, textvariable=orient, values=["portrait", "landscape"], state="readonly").pack()
    for label, var in [("Top (cm):", mt), ("Bottom (cm):", mb), ("Left (cm):", ml), ("Right (cm):", mr)]:
        ttk.Label(win, text=label).pack(); ttk.Entry(win, textvariable=var, width=10).pack()
    ttk.Label(win, text="Columns:").pack(pady=(6, 0)); ttk.Spinbox(win, from_=1, to=3, textvariable=cols, width=6).pack()
    ttk.Checkbutton(win, text="Line numbers", variable=ln).pack()
    ttk.Checkbutton(win, text="Automatic hyphenation", variable=hy).pack()
    saved: list[bool] = []
    def ok():
        setup.size, setup.orientation = size.get(), orient.get()
        setup.margin_top, setup.margin_bottom = mt.get(), mb.get()
        setup.margin_left, setup.margin_right = ml.get(), mr.get()
        setup.columns, setup.line_numbers, setup.hyphenation = cols.get(), ln.get(), hy.get()
        saved.append(True); win.destroy()
    ttk.Button(win, text="OK", command=ok).pack(pady=10)
    win.wait_window()
    return bool(saved)


def header_footer_dialog(parent, hf) -> bool:
    win = _dlg(parent, "Header & footer", 420, 320)
    h = tk.StringVar(value=hf.header); f = tk.StringVar(value=hf.footer)
    pn = tk.BooleanVar(value=hf.page_numbers); fp = tk.BooleanVar(value=hf.first_page_different)
    ttk.Label(win, text="Header:").pack(); ttk.Entry(win, textvariable=h, width=48).pack()
    ttk.Label(win, text="Footer:").pack(pady=(6, 0)); ttk.Entry(win, textvariable=f, width=48).pack()
    ttk.Checkbutton(win, text="Page numbers", variable=pn).pack(pady=4)
    ttk.Checkbutton(win, text="Different first page", variable=fp).pack()
    saved: list[bool] = []
    def ok():
        hf.header, hf.footer, hf.page_numbers = h.get(), f.get(), pn.get()
        hf.first_page_different = fp.get()
        saved.append(True); win.destroy()
    ttk.Button(win, text="OK", command=ok).pack(pady=10)
    win.wait_window()
    return bool(saved)


def watermark_dialog(parent, current: str) -> str | None:
    win = _dlg(parent, "Watermark", 320, 220)
    v = tk.StringVar(value=current or "DRAFT")
    for preset in ["DRAFT", "CONFIDENTIAL", "FINAL", "DO NOT COPY", "(none)"]:
        ttk.Radiobutton(win, text=preset, value="" if preset == "(none)" else preset, variable=v).pack(anchor="w", padx=20)
    ttk.Label(win, text="Custom:").pack(pady=(6, 0)); ttk.Entry(win, textvariable=v, width=24).pack()
    out: list = []
    ttk.Button(win, text="OK", command=lambda: (out.append(v.get()), win.destroy())).pack(pady=8)
    win.wait_window()
    return out[0] if out else None


def theme_dialog(parent, current: str) -> str | None:
    win = _dlg(parent, "Themes", 340, 300)
    ttk.Label(win, text="A theme sets heading/body fonts and the accent color:").pack(pady=6)
    lb = tk.Listbox(win); lb.pack(fill="both", expand=True, padx=10)
    names = list(THEMES.keys())
    for n in names:
        t = THEMES[n]
        lb.insert("end", f"{n} — {t.heading_font} / {t.accent}")
    try: lb.selection_set(names.index(current))
    except ValueError: pass
    out: list = []
    def ok():
        if lb.curselection(): out.append(names[lb.curselection()[0]])
        win.destroy()
    ttk.Button(win, text="Apply theme", command=ok).pack(pady=8)
    win.wait_window()
    return out[0] if out else None


# ---------------- symbols / case / goto ----------------
def symbol_dialog(parent, insert_fn):
    win = _dlg(parent, "Symbols", 420, 320)
    frame = ttk.Frame(win); frame.pack(fill="both", expand=True, padx=8, pady=8)
    for i, s in enumerate(SYMBOLS):
        ttk.Button(frame, text=s, width=3,
                   command=lambda ch=s: (insert_fn(ch),)).grid(row=i // 10, column=i % 10, padx=1, pady=1)


def change_case_dialog(parent, apply_fn):
    pick = simpledialog.askstring("Change case", "Mode: UPPER / lower / Title / tOGGLE",
                                  parent=parent, initialvalue="Title")
    if pick:
        m = pick.strip().lower()
        mode = {"upper": "upper", "lower": "lower", "title": "title"}.get(m, "toggle")
        apply_fn(mode)


def goto_dialog(parent, go_fn):
    what = simpledialog.askstring("Go to", "Type: page N, heading text, or line N", parent=parent)
    if what: go_fn(what)


# ---------------- word count ----------------
def word_count_dialog(parent, stats: dict):
    messagebox.showinfo("Word count",
        f"Pages: {stats['pages']}\nWords: {stats['words']}\n"
        f"Characters (with spaces): {stats['chars']}\n"
        f"Characters (no spaces): {stats['chars_no_spaces']}\n"
        f"Paragraphs: {stats['paragraphs']}\nLines: {stats['lines']}\n"
        f"Tables: {stats['tables']}  Footnotes: {stats['footnotes']}  Comments: {stats['comments']}",
        parent=parent)


# ---------------- citations ----------------
def citation_manager_dialog(parent, doc, insert_fn, refresh_fn):
    win = _dlg(parent, "Citations & bibliography", 520, 440)
    lb = tk.Listbox(win); lb.pack(fill="both", expand=True, padx=8, pady=6)
    def reload():
        lb.delete(0, "end")
        for c in doc.citations:
            lb.insert("end", f"{c.author} ({c.year}) {c.title}")
    style = tk.StringVar(value="APA")
    ttk.Combobox(win, textvariable=style, values=["APA", "MLA", "Chicago", "IEEE"],
                 state="readonly", width=10).pack()
    bar = ttk.Frame(win); bar.pack(fill="x", padx=8, pady=4)
    def add():
        a = simpledialog.askstring("Source", "Author:", parent=win)
        if not a: return
        t = simpledialog.askstring("Source", "Title:", parent=win) or ""
        y = simpledialog.askstring("Source", "Year:", parent=win) or ""
        p = simpledialog.askstring("Source", "Publisher:", parent=win) or ""
        doc.citations.append(Citation(key=f"s{len(doc.citations)+1}", author=a, title=t, year=y, publisher=p))
        reload(); refresh_fn()
    def insert():
        if not lb.curselection(): return
        c = doc.citations[lb.curselection()[0]]
        s = style.get().upper()
        txt = {"APA": c.apa(), "MLA": c.mla(), "CHICAGO": c.chicago(), "IEEE": c.ieee(lb.curselection()[0]+1)}.get(s, c.apa())
        insert_fn(f"({txt})")
    def biblio():
        insert_fn("\nBibliography\n" + "\n".join(doc.bibliography(style.get())) + "\n")
    ttk.Button(bar, text="Add source", command=add).pack(side="left", padx=2)
    ttk.Button(bar, text="Insert citation", command=insert).pack(side="left", padx=2)
    ttk.Button(bar, text="Insert bibliography", command=biblio).pack(side="left", padx=2)
    reload()


# ---------------- mail merge ----------------
def mail_merge_dialog(parent, doc, insert_field_fn, preview_fn, finish_fn):
    win = _dlg(parent, "Mail merge", 480, 480)
    state: dict = {"fields": [], "records": [], "i": 0, "csv": ""}
    ttk.Button(win, text="1. Load recipients (CSV)...", command=lambda: load()).pack(fill="x", padx=10, pady=4)
    info = ttk.Label(win, text="No list loaded."); info.pack()
    flb = tk.Listbox(win, height=6); flb.pack(fill="x", padx=10)
    ttk.Button(win, text="2. Insert selected merge field", command=lambda: insert_field()).pack(fill="x", padx=10, pady=4)
    nav = ttk.Frame(win); nav.pack(pady=4)
    ttk.Button(nav, text="◀ Prev", command=lambda: step(-1)).pack(side="left", padx=4)
    ttk.Button(nav, text="Next ▶", command=lambda: step(1)).pack(side="left", padx=4)
    ttk.Button(win, text="3. Finish & merge to document...", command=lambda: finish()).pack(fill="x", padx=10, pady=6)

    def load():
        p = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")], parent=win)
        if not p: return
        fields, records = WordDoc.load_csv(p)
        state.update(fields=fields, records=records, i=0, csv=p)
        flb.delete(0, "end")
        for f in fields: flb.insert("end", f)
        info.config(text=f"{len(records)} recipients from {p}")
        if records: preview_fn(records[0])
    def insert_field():
        if flb.curselection(): insert_field_fn(f"<<{flb.get(flb.curselection()[0])}>>")
    def step(d):
        if not state["records"]: return
        state["i"] = (state["i"] + d) % len(state["records"])
        preview_fn(state["records"][state["i"]])
        info.config(text=f"Record {state['i']+1} of {len(state['records'])}")
    def finish():
        if not state["records"]:
            messagebox.showinfo("Mail merge", "Load a CSV list first.", parent=win); return
        finish_fn(state["records"]); win.destroy()
    # WordDoc import here to avoid circulars at module import time
    from engine import WordDoc  # noqa: F401


# ---------------- macros ----------------
def macro_dialog(parent, macros: dict, record_fn, play_fn):
    win = _dlg(parent, "Macros", 420, 380)
    lb = tk.Listbox(win); lb.pack(fill="both", expand=True, padx=8, pady=6)
    for name, m in macros.items():
        lb.insert("end", f"{name} ({len(m.steps)} steps)")
    bar = ttk.Frame(win); bar.pack(fill="x", padx=8, pady=4)
    ttk.Button(bar, text="Record", command=lambda: (record_fn(), win.destroy())).pack(side="left", padx=2)
    ttk.Button(bar, text="Run", command=lambda: run()).pack(side="left", padx=2)
    def run():
        if lb.curselection():
            name = lb.get(lb.curselection()[0]).split(" ")[0]
            play_fn(name); win.destroy()


# ---------------- compare / inspector / accessibility / print ----------------
def compare_dialog(parent, run_fn):
    a = filedialog.askopenfilename(title="Original document", filetypes=[("Word", "*.docx")])
    if not a: return
    b = filedialog.askopenfilename(title="Revised document", filetypes=[("Word", "*.docx")])
    if b: run_fn(a, b)


def inspector_dialog(parent, doc, clean_fn):
    win = _dlg(parent, "Document inspector", 400, 330)
    checks = {"Comments": len(doc.comments), "Tracked changes": len(doc.changes),
              "Hidden metadata": 1 if (doc.subject or doc.keywords) else 0,
              "Headers/footers": 1 if (doc.hf.header or doc.hf.footer) else 0,
              "Watermark": 1 if doc.watermark else 0}
    vars: dict[str, tk.BooleanVar] = {}
    for k, n in checks.items():
        v = tk.BooleanVar(value=n > 0)
        ttk.Checkbutton(win, text=f"{k} ({n} found)", variable=v).pack(anchor="w", padx=20, pady=2)
        vars[k] = v
    ttk.Button(win, text="Remove selected",
               command=lambda: (clean_fn([k for k, v in vars.items() if v.get()]), win.destroy())).pack(pady=10)


def accessibility_report(parent, doc) -> str:
    issues = []
    heads = doc.headings(9)
    if not heads:
        issues.append("No headings — add Heading 1/2 so screen readers can navigate.")
    levels = []
    for s, _ in heads:
        try: levels.append(int(s.split()[-1]))
        except Exception: levels.append(1)
    for a, b in zip(levels, levels[1:]):
        if b - a > 1:
            issues.append(f"Skipped heading level (H{a} followed by H{b}).")
            break
    for b in doc.blocks:
        from engine import Table as _T
        if isinstance(b, _T) and not b.caption:
            issues.append("A table has no caption — add one from References > Caption.")
            break
    if not issues:
        return "No issues found. Headings, tables and structure look accessible. ✓"
    return "Accessibility checker:\n• " + "\n• ".join(issues)


def print_dialog(parent, print_fn):
    win = _dlg(parent, "Print", 300, 180)
    copies = tk.IntVar(value=1)
    ttk.Label(win, text="Copies:").pack(pady=(10, 0))
    ttk.Spinbox(win, from_=1, to=99, textvariable=copies, width=8).pack()
    ttk.Button(win, text="Print", command=lambda: (print_fn(copies.get()), win.destroy())).pack(pady=12)
