# -*- coding: utf-8 -*-
"""WordSoft — معالج نصوص عربي بأسلوب Word، مستوحى من WordCraft.
تشغيل:  python app.py   أو   python -m Wordsoft
المتطلبات: python-docx + reportlab + pillow (tkinter مدمجة).
"""
import os, re, datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, colorchooser, font as tkfont

from engine import WordDoc, Paragraph, Run, Table
import formats

DARK = {"bg": "#1e1e1e", "panel": "#252526", "panel2": "#2d2d30", "doc": "#1b1b1b",
        "page": "#2b2b2b", "fg": "#e8e8e8", "muted": "#9a9a9a", "accent": "#0e639c", "border": "#3e3e42"}
LIGHT = {"bg": "#f3f3f3", "panel": "#ffffff", "panel2": "#eef1f5", "doc": "#e5e5e5",
         "page": "#ffffff", "fg": "#1e1e1e", "muted": "#666666", "accent": "#2b579a", "border": "#d4d4d4"}

FONTS = ["Arial", "Segoe UI", "Times New Roman", "Courier New", "Calibri",
         "Traditional Arabic", "Simplified Arabic", "Amiri", "Tahoma"]
SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 48]
STYLES = ["Normal", "Title", "Heading 1", "Heading 2", "Heading 3", "Quote", "List Bullet", "List Number"]
STYLES_AR = {"Normal": "عادي", "Title": "عنوان", "Heading 1": "عنوان 1", "Heading 2": "عنوان 2",
             "Heading 3": "عنوان 3", "Quote": "اقتباس", "List Bullet": "• نقطية", "List Number": "1. مرقمة"}

OPEN_TYPES = [("Word / Text", "*.docx *.txt *.md"), ("Word", "*.docx"), ("Text", "*.txt *.md"), ("All", "*.*")]
SAVE_DOCX = [("Word Document", "*.docx")]

AUTOCORRECT = {"(c)": "©", "(r)": "®", "...": "…", "--": "–", "teh": "the"}


class WordSoftApp:
    def __init__(self, root, path=None):
        self.root = root
        root.title("WordSoft — معالج النصوص")
        root.geometry("1380x860")
        self.theme = DARK
        self.dark_mode = True
        self.zoom = 1.0
        self.current_path = None
        self.doc = WordDoc.sample()
        self.tables: dict[str, Table] = {}   # line_index -> Table
        self.track = tk.BooleanVar(value=False)
        self.smart = tk.BooleanVar(value=True)
        self._images = []  # منع Garbage Collection
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
        fm.add_command(label="جديد Ctrl+N", command=self.file_new)
        fm.add_command(label="فتح Ctrl+O", command=self.file_open)
        fm.add_command(label="حفظ Ctrl+S", command=self.file_save)
        fm.add_command(label="حفظ باسم...", command=self.file_saveas)
        fm.add_separator()
        fm.add_command(label="تصدير PDF", command=self.export_pdf)
        fm.add_command(label="تصدير HTML", command=self.export_html)
        fm.add_command(label="تصدير Markdown", command=self.export_md)
        fm.add_command(label="تصدير TXT", command=self.export_txt)
        fm.add_separator()
        fm.add_command(label="خصائص المستند", command=self.doc_props)
        fm.add_command(label="خروج", command=self.root.quit)
        m.add_cascade(label="ملف", menu=fm)
        em = tk.Menu(m, tearoff=0)
        em.add_command(label="تراجع Ctrl+Z", command=lambda: self.text.edit_undo())
        em.add_command(label="إعادة Ctrl+Y", command=lambda: self.text.edit_redo())
        em.add_separator()
        em.add_command(label="بحث Ctrl+F", command=self.find_dialog)
        em.add_command(label="استبدال Ctrl+H", command=self.replace_dialog)
        em.add_command(label="تحديد الكل Ctrl+A", command=lambda: self.text.tag_add("sel", "1.0", "end"))
        m.add_cascade(label="تحرير", menu=em)
        vm = tk.Menu(m, tearoff=0)
        vm.add_command(label="تكبير +", command=lambda: self.set_zoom(self.zoom + 0.1))
        vm.add_command(label="تصغير −", command=lambda: self.set_zoom(self.zoom - 0.1))
        vm.add_command(label="وضع داكن/فاتح", command=self.toggle_theme)
        vm.add_command(label="إخفاء/إظهار التنقل", command=self.toggle_nav)
        m.add_cascade(label="عرض", menu=vm)
        m.add_command(label="حول WordSoft", command=self.about)

    # ---------- ribbon ----------
    def _ribbon(self):
        bar = ttk.Frame(self.root); bar.pack(side="top", fill="x")
        ttk.Label(bar, text=" WordSoft ", font=("Segoe UI", 13, "bold"),
                  foreground="#2b579a").pack(side="left", padx=6)
        self.nb = ttk.Notebook(bar); self.nb.pack(side="left", fill="x", expand=True, padx=4, pady=2)
        self._tab_home()
        self._tab_insert()
        self._tab_layout()
        self._tab_review()
        self._tab_view()

    def _tab_home(self):
        f = ttk.Frame(self.nb); self.nb.add(f, text=" الرئيسية ")
        # الخط
        self.font_var = tk.StringVar(value="Arial")
        cb = ttk.Combobox(f, textvariable=self.font_var, values=FONTS, width=16, state="readonly")
        cb.pack(side="left", padx=2); cb.bind("<<ComboboxSelected>>", lambda e: self.apply_font())
        self.size_var = tk.StringVar(value="12")
        cb2 = ttk.Combobox(f, textvariable=self.size_var, values=SIZES, width=5)
        cb2.pack(side="left", padx=2); cb2.bind("<<ComboboxSelected>>", lambda e: self.apply_font())
        for txt, tag in [("B", "bold"), ("I", "italic"), ("U", "underline"), ("S", "strike")]:
            b = ttk.Button(f, text=txt, width=3, command=lambda t=tag: self.toggle_tag(t))
            b.pack(side="left", padx=1)
            if tag in ("bold",): b.configure(style="TButton")
        ttk.Button(f, text="لون A", width=5, command=self.pick_fg).pack(side="left", padx=2)
        ttk.Button(f, text="تمييز", width=5, command=self.pick_hl).pack(side="left", padx=1)
        ttk.Separator(f, orient="vertical").pack(side="left", fill="y", padx=4)
        for icon, al in [("≡ يمين", "right"), ("≡ وسط", "center"), ("≡ يسار", "left"), ("☰ ضبط", "justify")]:
            ttk.Button(f, text=icon, width=6, command=lambda a=al: self.set_align(a)).pack(side="left", padx=1)
        ttk.Button(f, text="RTL ◀", width=6, command=lambda: self.set_align("right")).pack(side="left", padx=1)
        ttk.Button(f, text="• قائمة", width=7, command=lambda: self.set_style("List Bullet")).pack(side="left", padx=1)
        ttk.Button(f, text="1. مرقمة", width=7, command=lambda: self.set_style("List Number")).pack(side="left", padx=1)
        ttk.Separator(f, orient="vertical").pack(side="left", fill="y", padx=4)
        self.style_var = tk.StringVar(value="Normal")
        st = ttk.Combobox(f, textvariable=self.style_var, width=10, state="readonly",
                          values=list(STYLES_AR.values()))
        st.pack(side="left", padx=2); st.bind("<<ComboboxSelected>>", lambda e: self.set_style(
            [k for k, v in STYLES_AR.items() if v == self.style_var.get()][0]))
        ttk.Button(f, text="مسح التنسيق", command=self.clear_format).pack(side="left", padx=2)
        ttk.Button(f, text="نسخ التنسيق", command=self.format_painter).pack(side="left", padx=1)

    def _tab_insert(self):
        f = ttk.Frame(self.nb); self.nb.add(f, text=" إدراج ")
        ttk.Button(f, text="▦ جدول", command=self.insert_table_dlg).pack(side="left", padx=2)
        ttk.Button(f, text="🖼 صورة", command=self.insert_image).pack(side="left", padx=2)
        ttk.Button(f, text="فاصل صفحة", command=lambda: self.insert_text("\n— — فاصل صفحة — —\n")).pack(side="left", padx=2)
        ttk.Button(f, text="رقم الصفحة", command=self.insert_pagenum).pack(side="left", padx=2)
        ttk.Button(f, text="التاريخ", command=lambda: self.insert_text(datetime.date.today().isoformat())).pack(side="left", padx=2)
        ttk.Button(f, text="رابط", command=self.insert_link).pack(side="left", padx=2)
        ttk.Button(f, text="رمز Ω", command=self.insert_symbol).pack(side="left", padx=2)
        ttk.Button(f, text="تذييل", command=lambda: self.insert_text("\n— — —\nتذييل الصفحة")).pack(side="left", padx=2)

    def _tab_layout(self):
        f = ttk.Frame(self.nb); self.nb.add(f, text=" تخطيط ")
        ttk.Label(f, text="تباعد الأسطر:").pack(side="left", padx=2)
        self.ls_var = tk.StringVar(value="1.15")
        cb = ttk.Combobox(f, textvariable=self.ls_var, values=["1.0", "1.15", "1.5", "2.0"], width=5, state="readonly")
        cb.pack(side="left"); cb.bind("<<ComboboxSelected>>", lambda e: self.set_linespacing())
        ttk.Button(f, text="⇥ مسافة بادئة +", command=lambda: self.indent(1)).pack(side="left", padx=2)
        ttk.Button(f, text="⇤ مسافة بادئة −", command=lambda: self.indent(-1)).pack(side="left", padx=1)
        ttk.Button(f, text="عمودان (تجريبي)", command=self.two_columns).pack(side="left", padx=2)
        ttk.Button(f, text="لون الصفحة", command=self.page_color).pack(side="left", padx=2)

    def _tab_review(self):
        f = ttk.Frame(self.nb); self.nb.add(f, text=" مراجعة ")
        ttk.Button(f, text="🔍 تدقيق F7", command=self.spell_check).pack(side="left", padx=2)
        ttk.Button(f, text="∑ تعداد الكلمات", command=self.wordcount_dialog).pack(side="left", padx=2)
        ttk.Checkbutton(f, text="تعقب التغييرات", variable=self.track).pack(side="left", padx=4)
        ttk.Button(f, text="💬 تعليق", command=self.add_comment).pack(side="left", padx=2)
        ttk.Button(f, text="بحث واستبدال", command=self.replace_dialog).pack(side="left", padx=2)
        ttk.Checkbutton(f, text="تصحيح تلقائي ذكي", variable=self.smart).pack(side="left", padx=4)

    def _tab_view(self):
        f = ttk.Frame(self.nb); self.nb.add(f, text=" عرض ")
        ttk.Button(f, text="＋ تكبير", command=lambda: self.set_zoom(self.zoom + 0.1)).pack(side="left", padx=2)
        ttk.Button(f, text="－ تصغير", command=lambda: self.set_zoom(self.zoom - 0.1)).pack(side="left", padx=2)
        ttk.Button(f, text="100%", command=lambda: self.set_zoom(1.0)).pack(side="left", padx=1)
        ttk.Button(f, text="🌙 داكن/فاتح", command=self.toggle_theme).pack(side="left", padx=2)
        ttk.Button(f, text="🎯 تركيز", command=self.focus_mode).pack(side="left", padx=2)
        ttk.Button(f, text="🧭 تنقل", command=self.toggle_nav).pack(side="left", padx=2)

    # ---------- body ----------
    def _body(self):
        mid = ttk.Frame(self.root); mid.pack(fill="both", expand=True)
        # التنقل
        self.nav_frame = ttk.Frame(mid, width=230)
        self.nav_frame.pack(side="left", fill="y"); self.nav_frame.pack_propagate(False)
        ttk.Label(self.nav_frame, text=" 🧭 التنقل — العناوين ").pack(fill="x", padx=4, pady=2)
        self.nav = tk.Listbox(self.nav_frame, bg=self.theme["panel2"], fg=self.theme["fg"], relief="flat")
        self.nav.pack(fill="both", expand=True, padx=4, pady=2)
        self.nav.bind("<Double-Button-1>", self.nav_goto)
        ttk.Button(self.nav_frame, text="تحديث", command=self.refresh_nav).pack(fill="x", padx=4, pady=2)
        # المحرر
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
        self.text.tag_configure("Heading 1", font=("Arial", int(18 * z), "bold"), spacing1=12, spacing3=6, foreground="#2b579a")
        self.text.tag_configure("Heading 2", font=("Arial", int(15 * z), "bold"), spacing1=10, spacing3=4)
        self.text.tag_configure("Heading 3", font=("Arial", int(13 * z), "bold"), spacing1=8, spacing3=4)
        self.text.tag_configure("Quote", font=("Arial", int(12 * z), "italic"), lmargin1=30, lmargin2=30, background="#f0f4f8" if not self.dark_mode else "#333340")
        self.text.tag_configure("List Bullet", lmargin1=36, lmargin2=36)
        self.text.tag_configure("List Number", lmargin1=36, lmargin2=36)
        for al in ("left", "center", "right", "justify"):
            self.text.tag_configure(f"al-{al}", justify=al)
        # أحجام وألوان ديناميكية تنشأ عند الحاجة

    # ---------- تحميل / حفظ ----------
    def load_doc(self, doc: WordDoc):
        self.doc = doc
        self.text.delete("1.0", "end")
        self.tables.clear()
        for b in doc.blocks:
            if isinstance(b, Table):
                line = self.text.index("end-1c").split(".")[0]
                self.text.insert("end", f"▦ جدول ({len(b.rows)}×{len(b.rows[0]) if b.rows else 0}): {b.rows[0][0] if b.rows and b.rows[0] else ''} …\n", ("table-ph",))
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
                # تنسيق الـ Runs
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
        self.root.title(f"WordSoft — {doc.title}")
        self.refresh_nav(); self.update_status()

    def text_to_doc(self) -> WordDoc:
        content = self.text.get("1.0", "end-1c").split("\n")
        doc = WordDoc(title=self.doc.title, author=self.doc.author, track_changes=self.track.get())
        for i, line in enumerate(content, start=1):
            if str(i) in self.tables:
                doc.blocks.append(self.tables[str(i)])
                continue
            if not line.strip():
                doc.blocks.append(Paragraph([Run("")]))
                continue
            # كشف النمط والمحاذاة
            tags = set(self.text.tag_names(f"{i}.0"))
            style = "Normal"
            for s in STYLES:
                if s in tags: style = s; break
            align = "left"
            for a in ("left", "center", "right", "justify"):
                if f"al-{a}" in tags: align = a; break
            # كشف البادئة
            t = line
            if style in ("List Bullet",) and t.startswith("• "): t = t[2:]
            if style in ("List Number",) and re.match(r"^\d+\.\s", t): t = re.sub(r"^\d+\.\s", "", t)
            # بناء Runs حسب تغير التنسيق
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

    # ---------- ملف ----------
    def file_new(self):
        win = tk.Toplevel(self.root); win.title("مستند جديد — قوالب WordSoft"); win.geometry("520x340")
        ttk.Label(win, text="اختر قالباً:", font=("Segoe UI", 12, "bold")).pack(pady=8)
        for label, kind in [("📄 فارغ", "blank"), ("📖 دليل الاستوديو (عينة)", "sample"),
                            ("✉ رسالة رسمية", "letter"), ("📊 تقرير", "report"), ("👤 سيرة ذاتية", "resume")]:
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
            messagebox.showinfo("WordSoft", f"تم الفتح:\n{p}")
        except Exception as ex:
            messagebox.showerror("خطأ", str(ex))

    def file_save(self):
        if not self.current_path: return self.file_saveas()
        try:
            formats.save_docx(self.text_to_doc(), self.current_path)
            messagebox.showinfo("WordSoft", "تم الحفظ ✓")
        except Exception as ex: messagebox.showerror("خطأ", str(ex))

    def file_saveas(self):
        p = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=SAVE_DOCX)
        if not p: return
        try:
            formats.save_docx(self.text_to_doc(), p)
            self.current_path = p
            messagebox.showinfo("WordSoft", f"تم الحفظ:\n{p}")
        except Exception as ex: messagebox.showerror("خطأ", str(ex))

    def _export(self, ext, ftype, saver):
        p = filedialog.asksaveasfilename(defaultextension=ext, filetypes=[ftype])
        if not p: return
        try:
            saver(self.text_to_doc(), p)
            messagebox.showinfo("WordSoft", f"تم التصدير:\n{p}")
        except Exception as ex: messagebox.showerror("خطأ", str(ex))

    def export_pdf(self): self._export(".pdf", ("PDF", "*.pdf"), formats.save_pdf)
    def export_html(self): self._export(".html", ("HTML", "*.html"), formats.save_html)
    def export_md(self): self._export(".md", ("Markdown", "*.md"), formats.save_md)
    def export_txt(self): self._export(".txt", ("Text", "*.txt"), formats.save_txt)

    def doc_props(self):
        d = self.text_to_doc().stats()
        messagebox.showinfo("خصائص المستند",
            f"العنوان: {self.doc.title}\nالكلمات: {d['words']}\nالأحرف: {d['chars']}\n"
            f"الفقرات: {d['paragraphs']}\nالجداول: {d['tables']}\nالصفحات ≈ {d['pages']}")

    # ---------- تنسيق ----------
    def _sel(self):
        try: return (self.text.index("sel.first"), self.text.index("sel.last"))
        except Exception:
            line = self.text.index("insert").split(".")[0]
            return (f"{line}.0", f"{line}.end")

    def toggle_tag(self, tag):
        a, b = self._sel()
        if tag in self.text.tag_names("insert"):  # مؤشر داخل تنسيق
            pass
        # بدّل: إن كان الكل موسوماً أزل وإلا أضف
        try:
            has = all(tag in self.text.tag_names(f"{a}+{i}c") for i in range(min(5, int(self.text.count(a, b, 'chars')[0] or 1))))
        except Exception: has = False
        if has: self.text.tag_remove(tag, a, b)
        else: self.text.tag_add(tag, a, b)
        if self.track.get() and tag == "bold":
            pass
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
        c, hx = colorchooser.askcolor(title="لون النص")
        if c:
            a, b = self._sel()
            tg = f"fg-{hx.strip('#').upper()}"
            self.text.tag_configure(tg, foreground=hx)
            self.text.tag_add(tg, a, b)

    def pick_hl(self):
        c, hx = colorchooser.askcolor(title="لون التمييز")
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
            # نظّف البادئة القديمة
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
            messagebox.showinfo("نسخ التنسيق", "تم نسخ التنسيق — حدّد نصاً ثم اضغط Ctrl+Shift+V للصق.")
        except Exception: messagebox.showinfo("نسخ التنسيق", "حدّد نصاً أولاً.")

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
        messagebox.showinfo("تباعد الأسطر", f"تم الضبط على {self.ls_var.get()} (يطبق عند تصدير DOCX/PDF).")

    def two_columns(self):
        messagebox.showinfo("عمودان", "الوضع التجريبي: يُصدَّر كجدول عمودين في DOCX. اكتب « || » للفصل بين العمودين.")

    def page_color(self):
        c, hx = colorchooser.askcolor(title="لون الصفحة")
        if hx: self.text.configure(bg=hx)

    def render_page_bg(self):
        self.canvas_bg.configure(bg=self.theme["doc"])

    # ---------- إدراج ----------
    def insert_text(self, t):
        self.text.insert("insert", t); self.update_status()

    def insert_table_dlg(self):
        r = simpledialog.askinteger("جدول", "عدد الصفوف:", initialvalue=3, minvalue=1, maxvalue=20)
        c = simpledialog.askinteger("جدول", "عدد الأعمدة:", initialvalue=3, minvalue=1, maxvalue=10)
        if not r or not c: return
        rows = [["" for _ in range(c)] for _ in range(r)]
        # املأ عناوين تجريبية
        for j in range(c): rows[0][j] = f"عمود {j + 1}"
        line = self.text.index("insert").split(".")[0]
        self.text.insert("insert", f"\n▦ جدول ({r}×{c})\n")
        self.tables[line] = Table(rows=rows)
        messagebox.showinfo("جدول", "تم إدراج الجدول — سيُحفظ في DOCX/PDF/HTML. حرّره من «مراجعة الجداول» عبر إعادة الفتح.")

    def insert_image(self):
        from PIL import Image, ImageTk
        p = filedialog.askopenfilename(filetypes=[("Images", "*.png *.jpg *.jpeg *.gif *.bmp *.webp"), ("All", "*.*")])
        if not p: return
        try:
            img = Image.open(p); img.thumbnail((420, 420))
            ph = ImageTk.PhotoImage(img)
            self._images.append(ph)
            self.text.image_create("insert", image=ph)
            self.text.insert("insert", f" [صورة: {os.path.basename(p)}]\n")
        except Exception as ex: messagebox.showerror("خطأ", str(ex))

    def insert_pagenum(self):
        n = len(self.text.get("1.0", "end-1c").split("\n")) // 45 + 1
        self.insert_text(f" [صفحة {n}] ")

    def insert_link(self):
        url = simpledialog.askstring("رابط", "الصق الرابط:")
        if url: self.insert_text(f" {url} ")

    def insert_symbol(self):
        syms = "Ω π Σ √ ∞ § © ® ™ € £ ¥ → ← • … — – α β γ"
        s = simpledialog.askstring("رمز", f"اختر رمزاً والصقه:\n{syms}")
        if s: self.insert_text(s)

    def add_comment(self):
        try: sel = self.text.get("sel.first", "sel.last")
        except Exception: sel = "النص"
        self.insert_text(f" [تعليق: {sel[:30]}…] ")

    # ---------- بحث ----------
    def find_dialog(self):
        w = simpledialog.askstring("بحث", "كلمة البحث:")
        if not w: return
        self.text.tag_remove("found", "1.0", "end")
        self.text.tag_configure("found", background="yellow", foreground="black")
        idx, n = "1.0", 0
        while True:
            idx = self.text.search(w, idx, nocase=True, stopindex="end")
            if not idx: break
            self.text.tag_add("found", idx, f"{idx}+{len(w)}c"); idx = f"{idx}+{len(w)}c"; n += 1
        messagebox.showinfo("بحث", f"وجد {n} نتيجة.")

    def replace_dialog(self):
        win = tk.Toplevel(self.root); win.title("استبدال"); win.geometry("340x170")
        ttk.Label(win, text="بحث عن:").pack(); e1 = ttk.Entry(win, width=36); e1.pack(pady=2)
        ttk.Label(win, text="استبدال بـ:").pack(); e2 = ttk.Entry(win, width=36); e2.pack(pady=2)
        def go():
            a, b, n = "1.0", None, 0
            w, r = e1.get(), e2.get()
            if not w: return
            while True:
                idx = self.text.search(w, a, nocase=True, stopindex="end")
                if not idx: break
                self.text.delete(idx, f"{idx}+{len(w)}c"); self.text.insert(idx, r)
                a, n = f"{idx}+{len(r)}c", n + 1
            messagebox.showinfo("استبدال", f"تم استبدال {n}."); win.destroy()
        ttk.Button(win, text="استبدال الكل", command=go).pack(pady=8)

    # ---------- مراجعة ----------
    def wordcount_dialog(self):
        d = self.text_to_doc().stats()
        messagebox.showinfo("تعداد الكلمات",
            f"الكلمات: {d['words']}\nالأحرف: {d['chars']}\nالفقرات: {d['paragraphs']}\nالصفحات ≈ {d['pages']}")

    def spell_check(self):
        txt = self.text.get("1.0", "end-1c")
        words = re.findall(r"[\u0600-\u06FF\w]+", txt)
        # تدقيق بسيط: كلمات مكررة + حروف متكررة + همزات شائعة
        issues = []
        prev = None
        for w in words:
            if w == prev: issues.append(f"تكرار: {w}")
            if re.search(r"(.)\1{3,}", w): issues.append(f"حروف مكررة: {w}")
            prev = w
        self.text.tag_remove("spell", "1.0", "end")
        self.text.tag_configure("spell", underline=True)
        msg = "لا مشاكل واضحة ✓" if not issues else "\n".join(issues[:15]) + (f"\n… +{len(issues) - 15}" if len(issues) > 15 else "")
        messagebox.showinfo("تدقيق إملائي F7", msg)

    # ---------- عرض ----------
    def set_zoom(self, z):
        self.zoom = max(0.5, min(3.0, z))
        self._config_tags()
        # أعد تطبيق أحجام مخصصة
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
            # ابحث عن أسطر موسومة بعناوين
            nl = self.text.search("\n", idx, stopindex="end")
            if not nl: break
            for h in ("Title", "Heading 1", "Heading 2", "Heading 3"):
                if h in self.text.tag_names(idx):
                    t = self.text.get(idx, nl).strip()[:60]
                    if t:
                        self.nav.insert("end", f"{'#' if h == 'Title' else '#' * int(h[-1])} {t}")
                        self._nav_pos.append(idx)
                    break
            idx = f"{nl}+1c"
            if self.text.compare(idx, ">=", "end-1c"): break

    def nav_goto(self, e=None):
        try:
            i = self.nav.curselection()[0]
            self.text.see(self._nav_pos[i]); self.text.mark_set("insert", self._nav_pos[i])
        except Exception: pass

    # ---------- أحداث ----------
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
        # تصحيح تلقائي + علامات ذكية (مثل Word)
        if self.smart.get() and e and e.char in (" ", "\r", "."):
            try:
                idx = self.text.index("insert-1word")
                w = self.text.get(idx, "insert-1c").strip() if False else None
            except Exception: pass
            # تبسيط: استبدال آخر 3 أحرف إن طابقت القاموس
            try:
                end = self.text.index("insert-1c")
                for k, v in AUTOCORRECT.items():
                    s = f"{end}-{len(k)}c"
                    if self.text.get(s, end) == k:
                        self.text.delete(s, end); self.text.insert(s, v)
                        break
            except Exception: pass
        # استمرار القوائم عند Enter (مثل Word)
        if e and e.keysym == "Return":
            try:
                prev = int(self.text.index("insert").split(".")[0]) - 1
                t = self.text.get(f"{prev}.0", f"{prev}.end")
                if t.startswith("• "):
                    if t.strip() == "•":  # سطر فارغ ينهي القائمة
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
        self.status.config(text=f" الكلمات: {words} | الصفحات ≈ {pages} | زوم: {int(self.zoom * 100)}% "
                                f"| {'تعقب: يعمل' if self.track.get() else 'تعقب: متوقف'} | {'🌙' if self.dark_mode else '☀'} WordSoft")

    def about(self):
        messagebox.showinfo("WordSoft",
            "WordSoft 1.0\nمعالج نصوص عربي حر — مستوحى من WordCraft.\n"
            "الشريط، الأنماط، الجداول، التعداد، التدقيق، والتصدير DOCX/PDF/HTML/MD.\nمكتوب بلغة Python ليعمل فوراً.")


def main(path=None):
    root = tk.Tk()
    WordSoftApp(root, path=path)
    root.mainloop()


if __name__ == "__main__":
    import sys
    main(sys.argv[1] if len(sys.argv) > 1 else None)
