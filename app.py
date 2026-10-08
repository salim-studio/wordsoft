# -*- coding: utf-8 -*-
"""WordSoft — Write beautifully. A fast, friendly word processor.

Run:  python app.py   or   python -m Wordsoft
Requires: python-docx + reportlab + pillow (tkinter is built in).

Copyright (c) 2026 salim-slimani. MIT License.
"""
import os, re, datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, colorchooser, font as tkfont

from engine import WordDoc, Paragraph, Run, Table
import formats

BRAND = {"name": "WordSoft", "tagline": "Write beautifully.",
         "blue": "#2563EB", "version": "2.0.0"}

DARK = {"bg": "#161a23", "panel": "#1f2530", "panel2": "#2a3242", "doc": "#10141b",
        "page": "#232b3a", "fg": "#e8eaf0", "muted": "#8b94a7", "accent": "#2563EB", "border": "#39435a"}
LIGHT = {"bg": "#eef1f6", "panel": "#ffffff", "panel2": "#e8edf5", "doc": "#dde3ee",
         "page": "#ffffff", "fg": "#1a2233", "muted": "#5b6478", "accent": "#2563EB", "border": "#cdd6e4"}

FONTS = ["Arial", "Segoe UI", "Times New Roman", "Courier New", "Calibri",
         "Traditional Arabic", "Simplified Arabic", "Amiri", "Tahoma"]
SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 48]
STYLES = ["Normal", "Title", "Heading 1", "Heading 2", "Heading 3", "Quote", "List Bullet", "List Number"]
STYLE_LABELS = {"Normal": "Normal", "Title": "Title", "Heading 1": "Heading 1", "Heading 2": "Heading 2",
                "Heading 3": "Heading 3", "Quote": "Quote", "List Bullet": "Bullets", "List Number": "Numbering"}

OPEN_TYPES = [("Word / Text", "*.docx *.txt *.md"), ("Word", "*.docx"), ("Text", "*.txt *.md"), ("All", "*.*")]
SAVE_DOCX = [("Word Document", "*.docx")]

AUTOCORRECT = {"(c)": "©", "(r)": "®", "...": "…", "--": "–", "teh": "the"}


class WordSoftApp:
    def __init__(self, root, path=None):
        self.root = root
        root.title(f"{BRAND['name']} — {BRAND['tagline']}")
        root.geometry("1380x860")
        self.theme = DARK
        self.dark_mode = True
        self.zoom = 1.0
        self.current_path = None
        self.doc = WordDoc.sample()
        self.tables: dict[str, Table] = {}   # line_index -> Table
        self.track = tk.BooleanVar(value=False)
        self.smart = tk.BooleanVar(value=True)
        self._images = []  # keep PhotoImage refs alive
        self._set_icon()
        self._style()
        self._menu()
        self._ribbon()
        self._body()
        self._status()
        self._config_tags()
        self.load_doc(self.doc)
        self._bind_keys()
        if path and os.path.exists(path):
            self.open_path(path)

    def _set_icon(self):
        for cand in (os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "icon.png"),
                     os.path.join("assets", "icon.png")):
            if os.path.exists(cand):
                try:
                    self._icon = tk.PhotoImage(file=cand)
                    self.root.iconphoto(True, self._icon)
                except Exception:
                    pass
                break

    # ---------- theme ----------
    def _style(self):
        s = ttk.Style()
        try: s.theme_use("clam")
        except Exception: pass
        t = self.theme
        s.configure(".", background=t["panel"], foreground=t["fg"], fieldbackground=t["panel2"])
        s.configure("TFrame", background=t["panel"])
        s.configure("TLabel", background=t["panel"], foreground=t["fg"])
        s.configure("TButton", background=t["panel2"], foreground=t["fg"], padding=3)
        s.configure("TNotebook", background=t["bg"])
        s.configure("TNotebook.Tab", background=t["panel2"], foreground=t["fg"], padding=(10, 5))
        s.configure("TCombobox", fieldbackground=t["panel2"], background=t["panel2"], foreground=t["fg"])
        self.root.configure(bg=t["bg"])

    def toggle_theme(self):
        self.dark_mode = not self.dark_mode
        self.theme = DARK if self.dark_mode else LIGHT
        self._style()
        t = self.theme
        self.text.configure(bg=t["page"], fg=t["fg"], insertbackground=t["fg"],
                            selectbackground=self.theme["accent"])
        self.nav.configure(bg=t["panel2"], fg=t["fg"])
        self.render_page_bg()

    # ---------- menu ----------
    def _menu(self):
        m = tk.Menu(self.root); self.root.config(menu=m)
        fm = tk.Menu(m, tearoff=0)
        fm.add_command(label="New  Ctrl+N", command=self.file_new)
        fm.add_command(label="Open  Ctrl+O", command=self.file_open)
        fm.add_command(label="Save  Ctrl+S", command=self.file_save)
        fm.add_command(label="Save As...", command=self.file_saveas)
        fm.add_separator()
        fm.add_command(label="Export PDF", command=self.export_pdf)
        fm.add_command(label="Export HTML", command=self.export_html)
        fm.add_command(label="Export Markdown", command=self.export_md)
        fm.add_command(label="Export TXT", command=self.export_txt)
        fm.add_separator()
        fm.add_command(label="Document properties", command=self.doc_props)
        fm.add_command(label="Exit", command=self.root.quit)
        m.add_cascade(label="File", menu=fm)
        em = tk.Menu(m, tearoff=0)
        em.add_command(label="Undo  Ctrl+Z", command=lambda: self.text.edit_undo())
        em.add_command(label="Redo  Ctrl+Y", command=lambda: self.text.edit_redo())
        em.add_separator()
        em.add_command(label="Find  Ctrl+F", command=self.find_dialog)
        em.add_command(label="Replace  Ctrl+H", command=self.replace_dialog)
        em.add_command(label="Select All  Ctrl+A", command=lambda: self.text.tag_add("sel", "1.0", "end"))
        m.add_cascade(label="Edit", menu=em)
        vm = tk.Menu(m, tearoff=0)
        vm.add_command(label="Zoom In", command=lambda: self.set_zoom(self.zoom + 0.1))
        vm.add_command(label="Zoom Out", command=lambda: self.set_zoom(self.zoom - 0.1))
        vm.add_command(label="Dark / Light mode", command=self.toggle_theme)
        vm.add_command(label="Show / Hide navigation", command=self.toggle_nav)
        m.add_cascade(label="View", menu=vm)
        m.add_command(label=f"About {BRAND['name']}", command=self.about)

    # ---------- ribbon ----------
    def _ribbon(self):
        bar = ttk.Frame(self.root); bar.pack(side="top", fill="x")
        brand = tk.Label(bar, text=" W ", font=("Segoe UI", 14, "bold"),
                         bg=BRAND["blue"], fg="white")
        brand.pack(side="left", padx=(8, 2), pady=4)
        tk.Label(bar, text=f" {BRAND['name']}  ", font=("Segoe UI", 13, "bold"),
                 bg=self.theme["panel"], fg=BRAND["blue"] if False else self.theme["fg"]).pack(side="left")
        self.brand_sub = tk.Label(bar, text=BRAND["tagline"], font=("Segoe UI", 9, "italic"),
                                  bg=self.theme["panel"], fg=self.theme["muted"])
        self.brand_sub.pack(side="left", padx=(0, 8))
        self.nb = ttk.Notebook(bar); self.nb.pack(side="left", fill="x", expand=True, padx=4, pady=2)
        self._tab_home()
        self._tab_insert()
        self._tab_layout()
        self._tab_review()
        self._tab_view()

    def _tab_home(self):
        f = ttk.Frame(self.nb); self.nb.add(f, text=" Home ")
        # Font
        self.font_var = tk.StringVar(value="Arial")
        cb = ttk.Combobox(f, textvariable=self.font_var, values=FONTS, width=16, state="readonly")
        cb.pack(side="left", padx=2); cb.bind("<<ComboboxSelected>>", lambda e: self.apply_font())
        self.size_var = tk.StringVar(value="12")
        cb2 = ttk.Combobox(f, textvariable=self.size_var, values=SIZES, width=5)
        cb2.pack(side="left", padx=2); cb2.bind("<<ComboboxSelected>>", lambda e: self.apply_font())
        for txt, tag in [("B", "bold"), ("I", "italic"), ("U", "underline"), ("S", "strike")]:
            b = ttk.Button(f, text=txt, width=3, command=lambda t=tag: self.toggle_tag(t))
            b.pack(side="left", padx=1)
        ttk.Button(f, text="Color A", width=7, command=self.pick_fg).pack(side="left", padx=2)
        ttk.Button(f, text="Highlight", width=8, command=self.pick_hl).pack(side="left", padx=1)
        ttk.Separator(f, orient="vertical").pack(side="left", fill="y", padx=4)
        for icon, al in [("Right", "right"), ("Center", "center"), ("Left", "left"), ("Justify", "justify")]:
            ttk.Button(f, text=icon, width=7, command=lambda a=al: self.set_align(a)).pack(side="left", padx=1)
        ttk.Button(f, text="Bullets", width=7, command=lambda: self.set_style("List Bullet")).pack(side="left", padx=1)
        ttk.Button(f, text="Numbering", width=9, command=lambda: self.set_style("List Number")).pack(side="left", padx=1)
        ttk.Separator(f, orient="vertical").pack(side="left", fill="y", padx=4)
        self.style_var = tk.StringVar(value="Normal")
        st = ttk.Combobox(f, textvariable=self.style_var, width=11, state="readonly",
                          values=list(STYLE_LABELS.values()))
        st.pack(side="left", padx=2); st.bind("<<ComboboxSelected>>", lambda e: self.set_style(
            [k for k, v in STYLE_LABELS.items() if v == self.style_var.get()][0]))
        ttk.Button(f, text="Clear formatting", command=self.clear_format).pack(side="left", padx=2)
        ttk.Button(f, text="Format painter", command=self.format_painter).pack(side="left", padx=1)

    def _tab_insert(self):
        f = ttk.Frame(self.nb); self.nb.add(f, text=" Insert ")
        ttk.Button(f, text="Table", command=self.insert_table_dlg).pack(side="left", padx=2)
        ttk.Button(f, text="Picture", command=self.insert_image).pack(side="left", padx=2)
        ttk.Button(f, text="Page break", command=lambda: self.insert_text("\n— — page break — —\n")).pack(side="left", padx=2)
        ttk.Button(f, text="Page number", command=self.insert_pagenum).pack(side="left", padx=2)
        ttk.Button(f, text="Date", command=lambda: self.insert_text(datetime.date.today().isoformat())).pack(side="left", padx=2)
        ttk.Button(f, text="Link", command=self.insert_link).pack(side="left", padx=2)
        ttk.Button(f, text="Symbol", command=self.insert_symbol).pack(side="left", padx=2)
        ttk.Button(f, text="Footer", command=lambda: self.insert_text("\n— — —\nPage footer")).pack(side="left", padx=2)

    def _tab_layout(self):
        f = ttk.Frame(self.nb); self.nb.add(f, text=" Layout ")
        ttk.Label(f, text="Line spacing:").pack(side="left", padx=2)
        self.ls_var = tk.StringVar(value="1.15")
        cb = ttk.Combobox(f, textvariable=self.ls_var, values=["1.0", "1.15", "1.5", "2.0"], width=5, state="readonly")
        cb.pack(side="left"); cb.bind("<<ComboboxSelected>>", lambda e: self.set_linespacing())
        ttk.Button(f, text="Indent +", command=lambda: self.indent(1)).pack(side="left", padx=2)
        ttk.Button(f, text="Indent −", command=lambda: self.indent(-1)).pack(side="left", padx=1)
        ttk.Button(f, text="Two columns (experimental)", command=self.two_columns).pack(side="left", padx=2)
        ttk.Button(f, text="Page color", command=self.page_color).pack(side="left", padx=2)

    def _tab_review(self):
        f = ttk.Frame(self.nb); self.nb.add(f, text=" Review ")
        ttk.Button(f, text="Spelling F7", command=self.spell_check).pack(side="left", padx=2)
        ttk.Button(f, text="Word count", command=self.wordcount_dialog).pack(side="left", padx=2)
        ttk.Checkbutton(f, text="Track changes", variable=self.track).pack(side="left", padx=4)
        ttk.Button(f, text="Comment", command=self.add_comment).pack(side="left", padx=2)
        ttk.Button(f, text="Find & Replace", command=self.replace_dialog).pack(side="left", padx=2)
        ttk.Checkbutton(f, text="Smart autocorrect", variable=self.smart).pack(side="left", padx=4)

    def _tab_view(self):
        f = ttk.Frame(self.nb); self.nb.add(f, text=" View ")
        ttk.Button(f, text="Zoom +", command=lambda: self.set_zoom(self.zoom + 0.1)).pack(side="left", padx=2)
        ttk.Button(f, text="Zoom −", command=lambda: self.set_zoom(self.zoom - 0.1)).pack(side="left", padx=2)
        ttk.Button(f, text="100%", command=lambda: self.set_zoom(1.0)).pack(side="left", padx=1)
        ttk.Button(f, text="Dark / Light", command=self.toggle_theme).pack(side="left", padx=2)
        ttk.Button(f, text="Focus", command=self.focus_mode).pack(side="left", padx=2)
        ttk.Button(f, text="Navigation", command=self.toggle_nav).pack(side="left", padx=2)

    # ---------- body ----------
    def _body(self):
        mid = ttk.Frame(self.root); mid.pack(fill="both", expand=True)
        # Navigation pane
        self.nav_frame = ttk.Frame(mid, width=230)
        self.nav_frame.pack(side="left", fill="y"); self.nav_frame.pack_propagate(False)
        ttk.Label(self.nav_frame, text=" Navigation — Headings ").pack(fill="x", padx=4, pady=2)
        self.nav = tk.Listbox(self.nav_frame, bg=self.theme["panel2"], fg=self.theme["fg"], relief="flat")
        self.nav.pack(fill="both", expand=True, padx=4, pady=2)
        self.nav.bind("<Double-Button-1>", self.nav_goto)
        ttk.Button(self.nav_frame, text="Refresh", command=self.refresh_nav).pack(fill="x", padx=4, pady=2)
        # Editor
        doc_wrap = tk.Frame(mid, bg=self.theme["doc"])
        doc_wrap.pack(side="left", fill="both", expand=True)
        self.canvas_bg = doc_wrap
        self.text = tk.Text(doc_wrap, wrap="word", undo=True, maxundo=200,
                            bg=self.theme["page"], fg=self.theme["fg"],
                            insertbackground=self.theme["fg"],
                            font=("Arial", 12), padx=28, pady=28,
                            selectbackground=self.theme["accent"])
        vs = ttk.Scrollbar(doc_wrap, orient="vertical", command=self.text.yview)
        hs = ttk.Scrollbar(doc_wrap, orient="horizontal", command=self.text.xview)
        self.text.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        vs.pack(side="right", fill="y"); hs.pack(side="bottom", fill="x")
        self.text.pack(side="left", fill="both", expand=True, padx=18, pady=14)
        self.text.bind("<<Modified>>", self.on_modified)
        self.text.bind("<KeyRelease>", self.on_keyrelease)

    def _status(self):
        self.status = tk.Label(self.root, text="", bg=self.theme["bg"], fg=self.theme["muted"], anchor="w")
        self.status.pack(side="bottom", fill="x")
        self.update_status()

    def _config_tags(self):
        z = self.zoom
        self.text.tag_configure("bold", font=("Arial", int(12 * z), "bold"))
        self.text.tag_configure("italic", font=("Arial", int(12 * z), "italic"))
        self.text.tag_configure("underline", underline=True)
        self.text.tag_configure("strike", overstrike=True)
        self.text.tag_configure("Title", font=("Arial", int(26 * z), "bold"), justify="center", spacing1=10, spacing3=10)
        self.text.tag_configure("Heading 1", font=("Arial", int(18 * z), "bold"), spacing1=12, spacing3=6, foreground="#2563EB")
        self.text.tag_configure("Heading 2", font=("Arial", int(15 * z), "bold"), spacing1=10, spacing3=4)
        self.text.tag_configure("Heading 3", font=("Arial", int(13 * z), "bold"), spacing1=8, spacing3=4)
        self.text.tag_configure("Quote", font=("Arial", int(12 * z), "italic"), lmargin1=30, lmargin2=30, background="#eef3ff" if not self.dark_mode else "#2c3550")
        self.text.tag_configure("List Bullet", lmargin1=36, lmargin2=36)
        self.text.tag_configure("List Number", lmargin1=36, lmargin2=36)
        for al in ("left", "center", "right", "justify"):
            self.text.tag_configure(f"al-{al}", justify=al)

    # ---------- load / save ----------
    def load_doc(self, doc: WordDoc):
        self.doc = doc
        self.text.delete("1.0", "end")
        self.tables.clear()
        for b in doc.blocks:
            if isinstance(b, Table):
                line = self.text.index("end-1c").split(".")[0]
                self.text.insert("end", f"Table ({len(b.rows)}x{len(b.rows[0]) if b.rows else 0}): {b.rows[0][0] if b.rows and b.rows[0] else ''} ...\n", ("table-ph",))
                self.tables[line] = b
            else:
                tags = []
                if b.style != "Normal": tags.append(b.style)
                if b.align != "left": tags.append(f"al-{b.align}")
                if b.style == "List Bullet": pre = "• "
                elif b.style == "List Number": pre = "1. "
                else: pre = ""
                start = self.text.index("end-1c")
                self.text.insert("end", pre + (b.text() or "") + "\n")
                end = self.text.index("end-1c")
                # Apply run formatting
                pos = len(pre)
                for r in b.runs:
                    n = len(r.text)
                    if n == 0: continue
                    rtags = list(tags)
                    if r.bold: rtags.append("bold")
                    if r.italic: rtags.append("italic")
                    if r.underline: rtags.append("underline")
                    if r.strike: rtags.append("strike")
                    if r.color:
                        tg = f"fg-{r.color}"
                        try: self.text.tag_configure(tg, foreground="#" + r.color)
                        except Exception: pass
                        rtags.append(tg)
                    if r.size:
                        tg = f"sz-{r.size}"
                        try: self.text.tag_configure(tg, font=("Arial", int(r.size * self.zoom)))
                        except Exception: pass
                        rtags.append(tg)
                    if rtags:
                        self.text.tag_add(tuple(rtags) if len(rtags) > 1 else rtags[0],
                                           f"{start}+{pos}c", f"{start}+{pos + n}c")
                    pos += n
                if tags:
                    self.text.tag_add(tuple(tags) if len(tags) > 1 else tags[0], start, end)
        self.text.edit_reset()
        self.text.edit_modified(False)
        self.root.title(f"{BRAND['name']} — {doc.title}")
        self.refresh_nav(); self.update_status()

    def text_to_doc(self) -> WordDoc:
        content = self.text.get("1.0", "end-1c").split("\n")
        doc = WordDoc(title=self.doc.title, author="salim-slimani", track_changes=self.track.get())
        for i, line in enumerate(content, start=1):
            if str(i) in self.tables:
                doc.blocks.append(self.tables[str(i)])
                continue
            if not line.strip():
                doc.blocks.append(Paragraph([Run("")]))
                continue
            # Detect style and alignment
            tags = set(self.text.tag_names(f"{i}.0"))
            style = "Normal"
            for s in STYLES:
                if s in tags: style = s; break
            align = "left"
            for a in ("left", "center", "right", "justify"):
                if f"al-{a}" in tags: align = a; break
            # Strip list prefixes
            t = line
            if style in ("List Bullet",) and t.startswith("• "): t = t[2:]
            if style in ("List Number",) and re.match(r"^\d+\.\s", t): t = re.sub(r"^\d+\.\s", "", t)
            # Build runs from formatting changes
            runs: list[Run] = []
            buf = ""; cur = None
            for j, ch in enumerate(t):
                idx = f"{i}.{j + (len(line) - len(t))}"
                try: tg = set(self.text.tag_names(idx))
                except Exception: tg = set()
                key = (("bold" in tg), ("italic" in tg), ("underline" in tg), ("strike" in tg),
                       next((x[3:] for x in tg if x.startswith("fg-")), None),
                       next((x[3:] for x in tg if x.startswith("sz-")), None))
                if cur is None: cur, buf = key, ch
                elif key == cur: buf += ch
                else:
                    runs.append(Run(buf, bold=cur[0], italic=cur[1], underline=cur[2], strike=cur[3],
                                    color=cur[4], size=int(cur[5]) if cur[5] and str(cur[5]).isdigit() else None))
                    cur, buf = key, ch
            if buf or not runs:
                runs.append(Run(buf, bold=cur[0] if cur else False, italic=cur[1] if cur else False,
                                underline=cur[2] if cur else False, strike=cur[3] if cur else False,
                                color=cur[4] if cur else None,
                                size=int(cur[5]) if cur and cur[5] and str(cur[5]).isdigit() else None))
            doc.blocks.append(Paragraph(runs, style=style, align=align))
        return doc

    # ---------- file ----------
    def file_new(self):
        win = tk.Toplevel(self.root); win.title(f"New document — {BRAND['name']} templates"); win.geometry("520x340")
        ttk.Label(win, text="Choose a template:", font=("Segoe UI", 12, "bold")).pack(pady=8)
        for label, kind in [("Blank", "blank"), ("Open Studio Handbook (sample)", "sample"),
                            ("Formal letter", "letter"), ("Report", "report"), ("Resume", "resume")]:
            ttk.Button(win, text=label, width=40,
                       command=lambda k=kind: (self.load_doc(WordDoc.template(k)), self.__setattr__("current_path", None), win.destroy())
                       ).pack(pady=3)

    def file_open(self):
        p = filedialog.askopenfilename(filetypes=OPEN_TYPES)
        if p: self.open_path(p)

    def open_path(self, p):
        try:
            doc = formats.load(p)
            self.current_path = p if p.endswith(".docx") else None
            self.load_doc(doc)
            messagebox.showinfo(BRAND["name"], f"Opened:\n{p}")
        except Exception as ex:
            messagebox.showerror("Error", str(ex))

    def file_save(self):
        if not self.current_path: return self.file_saveas()
        try:
            formats.save_docx(self.text_to_doc(), self.current_path)
            messagebox.showinfo(BRAND["name"], "Saved ✓")
        except Exception as ex: messagebox.showerror("Error", str(ex))

    def file_saveas(self):
        p = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=SAVE_DOCX)
        if not p: return
        try:
            formats.save_docx(self.text_to_doc(), p)
            self.current_path = p
            messagebox.showinfo(BRAND["name"], f"Saved:\n{p}")
        except Exception as ex: messagebox.showerror("Error", str(ex))

    def _export(self, ext, ftype, saver):
        p = filedialog.asksaveasfilename(defaultextension=ext, filetypes=[ftype])
        if not p: return
        try:
            saver(self.text_to_doc(), p)
            messagebox.showinfo(BRAND["name"], f"Exported:\n{p}")
        except Exception as ex: messagebox.showerror("Error", str(ex))

    def export_pdf(self): self._export(".pdf", ("PDF", "*.pdf"), formats.save_pdf)
    def export_html(self): self._export(".html", ("HTML", "*.html"), formats.save_html)
    def export_md(self): self._export(".md", ("Markdown", "*.md"), formats.save_md)
    def export_txt(self): self._export(".txt", ("Text", "*.txt"), formats.save_txt)

    def doc_props(self):
        d = self.text_to_doc().stats()
        messagebox.showinfo("Document properties",
            f"Title: {self.doc.title}\nWords: {d['words']}\nCharacters: {d['chars']}\n"
            f"Paragraphs: {d['paragraphs']}\nTables: {d['tables']}\nPages ≈ {d['pages']}")

    # ---------- formatting ----------
    def _sel(self):
        try: return (self.text.index("sel.first"), self.text.index("sel.last"))
        except Exception:
            line = self.text.index("insert").split(".")[0]
            return (f"{line}.0", f"{line}.end")

    def toggle_tag(self, tag):
        a, b = self._sel()
        try:
            has = all(tag in self.text.tag_names(f"{a}+{i}c") for i in range(min(5, int(self.text.count(a, b, 'chars')[0] or 1))))
        except Exception: has = False
        if has: self.text.tag_remove(tag, a, b)
        else: self.text.tag_add(tag, a, b)
        self.update_status()

    def apply_font(self):
        a, b = self._sel()
        fam, sz = self.font_var.get(), self.size_var.get()
        try: sz = int(sz)
        except Exception: sz = 12
        self.text.tag_configure(f"font-{fam}", font=(fam, int(sz * self.zoom)))
        self.text.tag_add(f"font-{fam}", a, b)
        try:
            self.text.tag_configure(f"sz-{sz}", font=(fam, int(sz * self.zoom)))
            self.text.tag_add(f"sz-{sz}", a, b)
        except Exception: pass

    def pick_fg(self):
        c, hx = colorchooser.askcolor(title="Text color")
        if c:
            a, b = self._sel()
            tg = f"fg-{hx.strip('#').upper()}"
            self.text.tag_configure(tg, foreground=hx)
            self.text.tag_add(tg, a, b)

    def pick_hl(self):
        c, hx = colorchooser.askcolor(title="Highlight color")
        if c:
            a, b = self._sel()
            tg = f"hl-{hx.strip('#').upper()}"
            self.text.tag_configure(tg, background=hx)
            self.text.tag_add(tg, a, b)

    def set_align(self, al):
        a, b = self._sel()
        for x in ("al-left", "al-center", "al-right", "al-justify"): self.text.tag_remove(x, a, b)
        line_a = self.text.index(a).split(".")[0]; line_b = self.text.index(b).split(".")[0]
        for ln in range(int(line_a), int(line_b) + 1):
            self.text.tag_add(f"al-{al}", f"{ln}.0", f"{ln}.end")

    def set_style(self, style):
        a, b = self._sel()
        for s in STYLES: self.text.tag_remove(s, a, b)
        line_a = int(self.text.index(a).split(".")[0]); line_b = int(self.text.index(b).split(".")[0])
        for ln in range(line_a, line_b + 1):
            txt = self.text.get(f"{ln}.0", f"{ln}.end")
            txt2 = re.sub(r"^[•\d\.\s]+", "", txt) if style in ("Normal", "Title") else txt
            if style == "List Bullet" and not txt.startswith("• "): txt2 = "• " + re.sub(r"^[•\d\.\s]+", "", txt)
            if style == "List Number" and not re.match(r"^\d+\.\s", txt): txt2 = "1. " + re.sub(r"^[•\d\.\s]+", "", txt)
            if txt2 != txt:
                self.text.delete(f"{ln}.0", f"{ln}.end"); self.text.insert(f"{ln}.0", txt2)
            if style != "Normal":
                self.text.tag_add(style, f"{ln}.0", f"{ln}.end")
        self.refresh_nav(); self.update_status()

    def clear_format(self):
        a, b = self._sel()
        for t in self.text.tag_names(a):
            if t not in ("sel",): self.text.tag_remove(t, a, b)

    def format_painter(self):
        try:
            tags = [t for t in self.text.tag_names("sel.first") if t != "sel"]
            self._painter = tags
            messagebox.showinfo("Format painter", "Format copied — select text, then press Ctrl+Shift+V to apply.")
        except Exception: messagebox.showinfo("Format painter", "Select some text first.")

    def paste_format(self):
        if not hasattr(self, "_painter"): return
        try:
            a, b = self._sel()
            for t in self._painter: self.text.tag_add(t, a, b)
        except Exception: pass

    def indent(self, d):
        a, b = self._sel()
        la, lb = int(self.text.index(a).split(".")[0]), int(self.text.index(b).split(".")[0])
        for ln in range(la, lb + 1):
            if d > 0: self.text.insert(f"{ln}.0", "    ")
            else:
                if self.text.get(f"{ln}.0", f"{ln}.4") == "    ": self.text.delete(f"{ln}.0", f"{ln}.4")

    def set_linespacing(self):
        messagebox.showinfo("Line spacing", f"Set to {self.ls_var.get()} (applies on DOCX/PDF export).")

    def two_columns(self):
        messagebox.showinfo("Two columns", "Experimental: write ' || ' to split the columns.")

    def page_color(self):
        c, hx = colorchooser.askcolor(title="Page color")
        if hx: self.text.configure(bg=hx)

    def render_page_bg(self):
        self.canvas_bg.configure(bg=self.theme["doc"])

    # ---------- insert ----------
    def insert_text(self, t):
        self.text.insert("insert", t); self.update_status()

    def insert_table_dlg(self):
        r = simpledialog.askinteger("Table", "Rows:", initialvalue=3, minvalue=1, maxvalue=20)
        c = simpledialog.askinteger("Table", "Columns:", initialvalue=3, minvalue=1, maxvalue=10)
        if not r or not c: return
        rows = [[f"Column {j + 1}" if i == 0 else "" for j in range(c)] for i in range(r)]
        line = self.text.index("insert").split(".")[0]
        self.text.insert("insert", f"\nTable ({r}x{c})\n")
        self.tables[line] = Table(rows=rows)
        messagebox.showinfo("Table", "Table inserted — saved into DOCX/PDF/HTML on export.")

    def insert_image(self):
        from PIL import Image, ImageTk
        p = filedialog.askopenfilename(filetypes=[("Images", "*.png *.jpg *.jpeg *.gif *.bmp *.webp"), ("All", "*.*")])
        if not p: return
        try:
            img = Image.open(p); img.thumbnail((420, 420))
            ph = ImageTk.PhotoImage(img)
            self._images.append(ph)
            self.text.image_create("insert", image=ph)
            self.text.insert("insert", f" [Image: {os.path.basename(p)}]\n")
        except Exception as ex: messagebox.showerror("Error", str(ex))

    def insert_pagenum(self):
        n = len(self.text.get("1.0", "end-1c").split("\n")) // 45 + 1
        self.insert_text(f" [Page {n}] ")

    def insert_link(self):
        url = simpledialog.askstring("Link", "Paste the URL:")
        if url: self.insert_text(f" {url} ")

    def insert_symbol(self):
        syms = "Ω π Σ √ ∞ § © ® ™ € £ ¥ → ← • … — – α β γ"
        s = simpledialog.askstring("Symbol", f"Pick a symbol and paste it:\n{syms}")
        if s: self.insert_text(s)

    def add_comment(self):
        try: sel = self.text.get("sel.first", "sel.last")
        except Exception: sel = "text"
        self.insert_text(f" [Comment: {sel[:30]}...] ")

    # ---------- find ----------
    def find_dialog(self):
        w = simpledialog.askstring("Find", "Search for:")
        if not w: return
        self.text.tag_remove("found", "1.0", "end")
        self.text.tag_configure("found", background="yellow", foreground="black")
        idx, n = "1.0", 0
        while True:
            idx = self.text.search(w, idx, nocase=True, stopindex="end")
            if not idx: break
            self.text.tag_add("found", idx, f"{idx}+{len(w)}c"); idx = f"{idx}+{len(w)}c"; n += 1
        messagebox.showinfo("Find", f"Found {n} match(es).")

    def replace_dialog(self):
        win = tk.Toplevel(self.root); win.title("Replace"); win.geometry("340x170")
        ttk.Label(win, text="Find:").pack(); e1 = ttk.Entry(win, width=36); e1.pack(pady=2)
        ttk.Label(win, text="Replace with:").pack(); e2 = ttk.Entry(win, width=36); e2.pack(pady=2)
        def go():
            a, n = "1.0", 0
            w, r = e1.get(), e2.get()
            if not w: return
            while True:
                idx = self.text.search(w, a, nocase=True, stopindex="end")
                if not idx: break
                self.text.delete(idx, f"{idx}+{len(w)}c"); self.text.insert(idx, r)
                a, n = f"{idx}+{len(r)}c", n + 1
            messagebox.showinfo("Replace", f"Replaced {n}."); win.destroy()
        ttk.Button(win, text="Replace all", command=go).pack(pady=8)

    # ---------- review ----------
    def wordcount_dialog(self):
        d = self.text_to_doc().stats()
        messagebox.showinfo("Word count",
            f"Words: {d['words']}\nCharacters: {d['chars']}\nParagraphs: {d['paragraphs']}\nPages ≈ {d['pages']}")

    def spell_check(self):
        txt = self.text.get("1.0", "end-1c")
        words = re.findall(r"[A-Za-z\u0600-\u06FF]+", txt)
        # Lightweight check: repeated words + over-repeated letters.
        issues = []
        prev = None
        for w in words:
            if prev and w.lower() == prev.lower(): issues.append(f"Repeated word: {w}")
            if re.search(r"(.)\1{3,}", w): issues.append(f"Repeated letters: {w}")
            prev = w
        self.text.tag_remove("spell", "1.0", "end")
        self.text.tag_configure("spell", underline=True)
        msg = "No obvious issues ✓" if not issues else "\n".join(issues[:15]) + (f"\n... +{len(issues) - 15} more" if len(issues) > 15 else "")
        messagebox.showinfo("Spelling (F7)", msg)

    # ---------- view ----------
    def set_zoom(self, z):
        self.zoom = max(0.5, min(3.0, z))
        self._config_tags()
        for tg in self.text.tag_names():
            if tg.startswith("sz-"):
                try: self.text.tag_configure(tg, font=("Arial", int(int(tg[3:]) * self.zoom)))
                except Exception: pass
        self.update_status()

    def focus_mode(self):
        if self.nav_frame.winfo_ismapped(): self.nav_frame.pack_forget()
        else: self.nav_frame.pack(side="left", fill="y")

    def toggle_nav(self): self.focus_mode()

    def refresh_nav(self):
        self.nav.delete(0, "end")
        self._nav_pos = []
        idx = "1.0"
        while True:
            nl = self.text.search("\n", idx, stopindex="end")
            if not nl: break
            for h in ("Title", "Heading 1", "Heading 2", "Heading 3"):
                if h in self.text.tag_names(idx):
                    t = self.text.get(idx, nl).strip()[:60]
                    if t:
                        prefix = "#" if h == "Title" else "#" * int(h[-1])
                        self.nav.insert("end", f"{prefix} {t}")
                        self._nav_pos.append(idx)
                    break
            idx = f"{nl}+1c"
            if self.text.compare(idx, ">=", "end-1c"): break

    def nav_goto(self, e=None):
        try:
            i = self.nav.curselection()[0]
            self.text.see(self._nav_pos[i]); self.text.mark_set("insert", self._nav_pos[i])
        except Exception: pass

    # ---------- events ----------
    def _bind_keys(self):
        r = self.root
        r.bind("<Control-s>", lambda e: (self.file_save(), "break")[1])
        r.bind("<Control-o>", lambda e: self.file_open())
        r.bind("<Control-n>", lambda e: self.file_new())
        r.bind("<Control-f>", lambda e: self.find_dialog())
        r.bind("<Control-h>", lambda e: self.replace_dialog())
        r.bind("<F7>", lambda e: self.spell_check())
        r.bind("<Control-b>", lambda e: (self.toggle_tag("bold"), "break")[1])
        r.bind("<Control-i>", lambda e: (self.toggle_tag("italic"), "break")[1])
        r.bind("<Control-u>", lambda e: (self.toggle_tag("underline"), "break")[1])
        self.text.bind("<Control-Shift-V>", lambda e: self.paste_format())

    def on_modified(self, e=None):
        if self.text.edit_modified():
            self.text.edit_modified(False)
            self.refresh_nav(); self.update_status()

    def on_keyrelease(self, e=None):
        # Smart autocorrect (like Word)
        if self.smart.get() and e and e.char in (" ", "\r", "."):
            try:
                end = self.text.index("insert-1c")
                for k, v in AUTOCORRECT.items():
                    s = f"{end}-{len(k)}c"
                    if self.text.get(s, end) == k:
                        self.text.delete(s, end); self.text.insert(s, v)
                        break
            except Exception: pass
        # Continue lists on Enter (like Word)
        if e and e.keysym == "Return":
            try:
                prev = int(self.text.index("insert").split(".")[0]) - 1
                t = self.text.get(f"{prev}.0", f"{prev}.end")
                if t.startswith("• "):
                    if t.strip() == "•":  # empty item ends the list
                        self.text.delete(f"{prev}.0", f"{prev}.end")
                    else:
                        self.text.insert("insert", "• ")
                        self.text.tag_add("List Bullet", "insert linestart", "insert lineend")
                elif re.match(r"^1\.\s", t):
                    self.text.insert("insert", "1. ")
                    self.text.tag_add("List Number", "insert linestart", "insert lineend")
            except Exception: pass
        self.update_status()

    def update_status(self):
        txt = self.text.get("1.0", "end-1c")
        words = len(re.findall(r"\S+", txt))
        pages = max(1, round(words / 400 + 0.5))
        track = "tracking" if self.track.get() else "not tracking"
        self.status.config(text=f" Words: {words} | Pages ≈ {pages} | Zoom: {int(self.zoom * 100)}% "
                                f"| Track changes: {track} | {BRAND['name']} {BRAND['version']}")

    def about(self):
        messagebox.showinfo(f"About {BRAND['name']}",
            f"{BRAND['name']} {BRAND['version']}\n{BRAND['tagline']}\n\n"
            "Ribbon tabs, styles, tables, word count, spelling, and DOCX/PDF/HTML/MD export.\n"
            "Written in Python so it runs instantly.\n\nCopyright (c) 2026 salim-slimani. MIT License.")


def main(path=None):
    root = tk.Tk()
    WordSoftApp(root, path=path)
    root.mainloop()


if __name__ == "__main__":
    import sys
    main(sys.argv[1] if len(sys.argv) > 1 else None)
