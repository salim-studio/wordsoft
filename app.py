# -*- coding: utf-8 -*-
"""WordSoft — Write beautifully. A full Word-style word processor.

Quick Access toolbar, File backstage, ribbon (Home/Insert/Draw/Design/
Layout/References/Mailings/Review/View), ruler, navigation + comments
panes, views, live status bar, macros, and headless automation.

Run:  python app.py   or   python -m Wordsoft

Copyright (c) 2026 salim-slimani. MIT License.
"""
import os, re, json, difflib, datetime, tempfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, colorchooser

from engine import WordDoc, Paragraph, Run, Table, PageBreak, TOCField, FootnoteRef, Footnote, Comment, THEMES
import formats
import dialogs
import extras
from commands import COMMANDS, execute, Macro

BRAND = {"name": "WordSoft", "tagline": "Write beautifully.", "blue": "#2563EB", "version": "2.2.0"}
SETTINGS_PATH = os.path.join(os.path.expanduser("~"), ".wordsoft.json")

DARK = {"bg": "#161a23", "panel": "#1f2530", "panel2": "#2a3242", "doc": "#10141b",
        "page": "#232b3a", "fg": "#e8eaf0", "muted": "#8b94a7", "accent": "#2563EB", "border": "#39435a"}
# macOS light: airy panels, hairline borders, system blue accent
LIGHT = {"bg": "#E9EAEE", "panel": "#F5F5F7", "panel2": "#FFFFFF", "doc": "#D8D8DE",
         "page": "#ffffff", "fg": "#1C1C1E", "muted": "#6E6E73", "accent": "#0A84FF", "border": "#D1D1D6"}

FONTS = ["SF Pro Text", "Calibri", "Calibri Light", "Arial", "Segoe UI", "Times New Roman", "Georgia",
         "Courier New", "Verdana", "Tahoma", "Traditional Arabic", "Amiri"]
SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 48, 72]
STYLES = ["Normal", "No Spacing", "Title", "Subtitle", "Heading 1", "Heading 2", "Heading 3",
          "Heading 4", "Quote", "Intense Quote", "Caption", "List Bullet", "List Number"]
OPEN_TYPES = [("Word / Text", "*.docx *.txt *.md"), ("Word", "*.docx"), ("Text", "*.txt *.md"), ("All", "*.*")]
SAVE_DOCX = [("Word Document", "*.docx")]
AUTOCORRECT = {"(c)": "©", "(r)": "®", "(tm)": "™", "...": "…", "--": "–", "teh": "the", "adn": "and"}
_BUILTIN_AC = set(AUTOCORRECT)
THESAURUS = {"fast": ["quick", "rapid", "swift"], "good": ["fine", "great", "excellent"],
    "big": ["large", "huge", "vast"], "small": ["little", "tiny", "compact"],
    "happy": ["glad", "pleased", "delighted"], "important": ["key", "crucial", "vital"],
    "show": ["display", "present", "reveal"], "use": ["apply", "employ", "utilize"],
    "make": ["create", "build", "produce"], "help": ["assist", "aid", "support"],
    "start": ["begin", "commence", "launch"], "end": ["finish", "conclude", "close"]}


class WordSoftApp:
    # ================= init =================
    def __init__(self, root, path=None):
        self.root = root
        root.title(f"{BRAND['name']} — {BRAND['tagline']}")
        root.geometry("1440x900")
        self.settings = self._load_settings()
        self.theme = DARK if self.settings.get("dark", True) else LIGHT
        self.dark_mode = self.settings.get("dark", True)
        self.zoom = 1.0
        self.view = "print"
        self.marks = False
        self.current_path = None
        self.doc = WordDoc.sample()
        self.doc.author = self.settings.get("user", "salim-slimani")
        self.tables: dict[str, Table] = {}
        self.tocs: dict[str, TOCField] = {}
        self.custom_styles: dict = self.settings.get("styles", {})
        self.macros: dict[str, Macro] = {}
        self.recording: Macro | None = None
        self.last_action = None
        self.caption_no = {"Table": 0, "Figure": 0, "Equation": 0}
        self.track = tk.BooleanVar(value=False)
        self.smart = tk.BooleanVar(value=True)
        self.user_dict = set(self.settings.get("dictionary", []))
        self._images = []
        self._strokes: list = []
        self._pen = {"color": "#FF0000", "width": 3}
        self._set_icon()
        self._style()
        self._quick_access()
        self._ribbon()
        self._ruler()
        self._body()
        self._statusbar()
        self._backstage()
        AUTOCORRECT.update(self.settings.get("autocorrect", {}))
        self.goal = int(self.settings.get("goal", 0))
        self.typewriter = False
        self._dirty = False
        self.versions_dir = os.path.join(os.path.expanduser("~"), ".wordsoft", "versions")
        try: os.makedirs(self.versions_dir, exist_ok=True)
        except Exception: pass
        self._config_tags()
        self._apply_margins()
        self.load_doc(self.doc)
        self._bind_keys()
        if path and os.path.exists(path):
            self.open_path(path)
        elif self.settings.get("restore") and self.settings.get("recent"):
            for p in self.settings["recent"]:
                if p and os.path.exists(p):
                    self.open_path(p)
                    break
        self._autosave_tick()
        if self.settings.get("gallery", True) and not path:
            self.root.after(300, self.show_gallery)

    # ================= settings =================
    def _load_settings(self) -> dict:
        try:
            with open(SETTINGS_PATH, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {"dark": True, "recent": [], "user": "salim-slimani", "lang": "en-US",
                    "dictionary": [], "styles": {}}

    def _save_settings(self):
        try:
            self.settings.update(dark=self.dark_mode, user=self.doc.author,
                                 dictionary=sorted(self.user_dict), styles=self.custom_styles)
            with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _remember(self, path: str):
        rec = [p for p in self.settings.get("recent", []) if p != path]
        self.settings["recent"] = ([path] + rec)[:10]

    # ================= chrome =================
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
        for w in (self.text,):
            w.configure(bg=t["page"], fg=t["fg"], insertbackground=t["fg"], selectbackground=t["accent"])
        self.nav_lb.configure(bg=t["panel2"], fg=t["fg"])
        self.cmt_lb.configure(bg=t["panel2"], fg=t["fg"])
        self.chg_lb.configure(bg=t["panel2"], fg=t["fg"])
        self.canvas_bg.configure(bg=t["doc"])
        self._save_settings()

    def _quick_access(self):
        qa = ttk.Frame(self.root); qa.pack(side="top", fill="x")
        tk.Label(qa, text=" W ", font=("Segoe UI", 11, "bold"), bg=BRAND["blue"], fg="white").pack(side="left", padx=(6, 2), pady=3)
        tk.Label(qa, text=f"{BRAND['name']}", font=("Segoe UI", 11, "bold")).pack(side="left")
        for icon, tip, fn in [("＋", "New (Ctrl+N)", self.file_new), ("📂", "Open (Ctrl+O)", self.file_open),
                               ("💾", "Save (Ctrl+S)", self.file_save), ("↩", "Undo", self.undo),
                               ("↪", "Redo", self.redo), ("🖨", "Print (Ctrl+P)", self.print_dialog),
                               ("⚙", "Preferences", self.preferences), ("?", "About", self.about)]:
            b = ttk.Button(qa, text=icon, width=3, command=fn)
            b.pack(side="left", padx=1)
            self._tip(b, tip)
        ttk.Label(qa, text=BRAND["tagline"], font=("Segoe UI", 9, "italic")).pack(side="left", padx=10)

    # ================= ribbon =================
    def _ribbon(self):
        self.nb = ttk.Notebook(self.root); self.nb.pack(side="top", fill="x", padx=2)
        self._tab_file(); self._tab_home(); self._tab_insert(); self._tab_draw()
        self._tab_design(); self._tab_layout(); self._tab_refs(); self._tab_mail()
        self._tab_review(); self._tab_view()
        self.nb.bind("<<NotebookTabChanged>>", self._on_tab)

    def _tab(self, title):
        f = ttk.Frame(self.nb); self.nb.add(f, text=f" {title} ")
        return f

    def _grp(self, parent, label):
        g = ttk.Frame(parent); g.pack(side="left", fill="y", padx=4)
        body = ttk.Frame(g); body.pack(side="top")
        ttk.Label(g, text=label, font=("Segoe UI", 8)).pack(side="bottom")
        return body

    def _tab_file(self):
        self._tab("File")  # content replaced by the backstage overlay

    def _on_tab(self, _e=None):
        if self.nb.tab(self.nb.select(), "text").strip() == "File":
            self.show_backstage("New")

    def _tab_home(self):
        f = self._tab("Home")
        g = self._grp(f, "Clipboard")
        for t, fn in [("Cut", self.cut), ("Copy", self.copy), ("Paste", self.paste),
                      ("Format", self.format_painter)]:
            ttk.Button(g, text=t, width=7, command=fn).pack(side="left", padx=1)
        g = self._grp(f, "Font")
        self.font_var = tk.StringVar(value="Calibri")
        cb = ttk.Combobox(g, textvariable=self.font_var, values=FONTS, width=14, state="readonly")
        cb.pack(side="left", padx=1); cb.bind("<<ComboboxSelected>>", lambda e: self.apply_font())
        self.size_var = tk.StringVar(value="11")
        cb2 = ttk.Combobox(g, textvariable=self.size_var, values=SIZES, width=4)
        cb2.pack(side="left", padx=1); cb2.bind("<<ComboboxSelected>>", lambda e: self.apply_font())
        ttk.Button(g, text="A+", width=3, command=lambda: self.grow_font(1)).pack(side="left")
        ttk.Button(g, text="A−", width=3, command=lambda: self.grow_font(-1)).pack(side="left")
        for txt, tag in [("B", "bold"), ("I", "italic"), ("U", "underline"), ("S", "strike"),
                         ("x²", "super"), ("x₂", "sub")]:
            ttk.Button(g, text=txt, width=3, command=lambda t=tag: self.toggle_tag(t)).pack(side="left", padx=1)
        ttk.Button(g, text="Aa", width=3, command=self.change_case).pack(side="left", padx=1)
        ttk.Button(g, text="A", width=3, command=self.pick_fg).pack(side="left")
        ttk.Button(g, text="ab", width=3, command=self.pick_hl).pack(side="left")
        ttk.Button(g, text="⌫", width=3, command=self.clear_format).pack(side="left", padx=1)
        g = self._grp(f, "Paragraph")
        ttk.Button(g, text="•", width=3, command=lambda: self.set_style("List Bullet")).pack(side="left")
        ttk.Button(g, text="1.", width=3, command=lambda: self.set_style("List Number")).pack(side="left")
        ttk.Button(g, text="⇥", width=3, command=lambda: self.indent(1)).pack(side="left")
        ttk.Button(g, text="⇤", width=3, command=lambda: self.indent(-1)).pack(side="left")
        for icon, al in [("L", "left"), ("C", "center"), ("R", "right"), ("J", "justify")]:
            ttk.Button(g, text=icon, width=3, command=lambda a=al: self.set_align(a)).pack(side="left", padx=1)
        ttk.Button(g, text="↕", width=3, command=self.line_spacing_menu).pack(side="left")
        ttk.Button(g, text="▦", width=3, command=self.para_shading).pack(side="left")
        ttk.Button(g, text="▢", width=3, command=self.para_border).pack(side="left")
        ttk.Button(g, text="A↓Z", width=4, command=self.sort_paras).pack(side="left", padx=1)
        g = self._grp(f, "Styles")
        self.style_var = tk.StringVar(value="Normal")
        st = ttk.Combobox(g, textvariable=self.style_var, width=12, state="readonly", values=STYLES)
        st.pack(side="left", padx=1); st.bind("<<ComboboxSelected>>", lambda e: self.set_style(self.style_var.get()))
        ttk.Button(g, text="Styles", width=7, command=self.styles_pane).pack(side="left", padx=1)
        ttk.Button(g, text="+ New", width=6, command=self.new_style).pack(side="left")
        g = self._grp(f, "Editing")
        ttk.Button(g, text="Find", width=7, command=self.find_dialog).pack(side="left", padx=1)
        ttk.Button(g, text="Replace", width=7, command=self.replace_dialog).pack(side="left", padx=1)
        ttk.Button(g, text="Go To", width=7, command=self.goto).pack(side="left", padx=1)
        ttk.Button(g, text="Select", width=7, command=self.select_menu).pack(side="left", padx=1)

    def _tab_insert(self):
        f = self._tab("Insert")
        g = self._grp(f, "Pages")
        ttk.Button(g, text="Cover", width=7, command=self.insert_cover).pack(side="left", padx=1)
        ttk.Button(g, text="Blank", width=7, command=self.insert_blank_page).pack(side="left", padx=1)
        ttk.Button(g, text="Break", width=7, command=self.insert_page_break).pack(side="left", padx=1)
        g = self._grp(f, "Tables")
        ttk.Button(g, text="Table", width=7, command=self.insert_table).pack(side="left", padx=1)
        ttk.Button(g, text="Edit", width=6, command=self.edit_table).pack(side="left", padx=1)
        g = self._grp(f, "Illustrations")
        ttk.Button(g, text="Picture", width=8, command=self.insert_image).pack(side="left", padx=1)
        ttk.Button(g, text="Shot", width=6, command=self.insert_screenshot).pack(side="left", padx=1)
        ttk.Button(g, text="Chart", width=7, command=self.insert_chart).pack(side="left", padx=1)
        ttk.Button(g, text="WordArt", width=8, command=self.insert_wordart).pack(side="left", padx=1)
        ttk.Button(g, text="Shape", width=7, command=self.insert_shape).pack(side="left", padx=1)
        ttk.Button(g, text="Text Box", width=8, command=self.insert_textbox).pack(side="left", padx=1)
        g = self._grp(f, "Header & Footer")
        ttk.Button(g, text="Header", width=7, command=self.edit_header_footer).pack(side="left", padx=1)
        ttk.Button(g, text="Footer", width=7, command=self.edit_header_footer).pack(side="left", padx=1)
        ttk.Button(g, text="Page #", width=7, command=self.insert_pagenum).pack(side="left", padx=1)
        g = self._grp(f, "Text")
        ttk.Button(g, text="Date", width=6, command=self.insert_datetime).pack(side="left", padx=1)
        ttk.Button(g, text="Field", width=6, command=self.insert_field).pack(side="left", padx=1)
        ttk.Button(g, text="Drop Cap", width=8, command=self.drop_cap).pack(side="left", padx=1)
        ttk.Button(g, text="Symbol", width=7, command=self.symbols).pack(side="left", padx=1)
        ttk.Button(g, text="Equation", width=8, command=self.insert_equation).pack(side="left", padx=1)
        ttk.Button(g, text="Quick Parts", width=10, command=self.quick_parts).pack(side="left", padx=1)
        g = self._grp(f, "Comments")
        ttk.Button(g, text="Comment", width=8, command=self.add_comment).pack(side="left", padx=1)

    def _tab_draw(self):
        f = self._tab("Draw")
        g = self._grp(f, "Pens")
        ttk.Button(g, text="Pen", width=7, command=lambda: self.draw_mode("pen")).pack(side="left", padx=1)
        ttk.Button(g, text="Marker", width=7, command=lambda: self.draw_mode("marker")).pack(side="left", padx=1)
        ttk.Button(g, text="Eraser", width=7, command=self.draw_clear).pack(side="left", padx=1)
        ttk.Button(g, text="Color", width=7, command=self.pen_color).pack(side="left", padx=1)
        g = self._grp(f, "Ink")
        ttk.Button(g, text="Keep as picture", width=13, command=self.ink_to_picture).pack(side="left", padx=1)
        ttk.Label(g, text="Draw over the page, then keep it as a picture.").pack(side="left", padx=6)

    def _tab_design(self):
        f = self._tab("Design")
        g = self._grp(f, "Themes")
        ttk.Button(g, text="Themes", width=8, command=self.pick_theme).pack(side="left", padx=1)
        ttk.Button(g, text="Fonts", width=7, command=self.theme_fonts).pack(side="left", padx=1)
        g = self._grp(f, "Page background")
        ttk.Button(g, text="Watermark", width=10, command=self.set_watermark).pack(side="left", padx=1)
        ttk.Button(g, text="Page color", width=10, command=self.page_color).pack(side="left", padx=1)
        ttk.Button(g, text="Borders", width=8, command=self.page_borders).pack(side="left", padx=1)

    def _tab_layout(self):
        f = self._tab("Layout")
        g = self._grp(f, "Page setup")
        ttk.Button(g, text="Margins", width=8, command=self.margins_menu).pack(side="left", padx=1)
        ttk.Button(g, text="Portrait", width=8, command=lambda: self.set_orientation("portrait")).pack(side="left", padx=1)
        ttk.Button(g, text="Landscape", width=9, command=lambda: self.set_orientation("landscape")).pack(side="left", padx=1)
        ttk.Button(g, text="Size", width=6, command=self.paper_size).pack(side="left", padx=1)
        ttk.Button(g, text="Setup...", width=8, command=self.page_setup).pack(side="left", padx=1)
        g = self._grp(f, "Paragraph")
        ttk.Button(g, text="Indent +", width=8, command=lambda: self.indent(1)).pack(side="left", padx=1)
        ttk.Button(g, text="Indent −", width=8, command=lambda: self.indent(-1)).pack(side="left", padx=1)
        ttk.Button(g, text="Spacing", width=8, command=self.para_spacing).pack(side="left", padx=1)
        g = self._grp(f, "Arrange")
        ttk.Button(g, text="Columns", width=8, command=self.columns_menu).pack(side="left", padx=1)
        ttk.Button(g, text="Breaks", width=8, command=self.breaks_menu).pack(side="left", padx=1)
        ttk.Button(g, text="Line #", width=7, command=self.toggle_line_numbers).pack(side="left", padx=1)
        ttk.Button(g, text="Hyphen.", width=8, command=self.toggle_hyphenation).pack(side="left", padx=1)

    def _tab_refs(self):
        f = self._tab("References")
        g = self._grp(f, "Table of contents")
        ttk.Button(g, text="Insert TOC", width=10, command=self.insert_toc).pack(side="left", padx=1)
        ttk.Button(g, text="Update", width=7, command=self.update_toc).pack(side="left", padx=1)
        g = self._grp(f, "Footnotes")
        ttk.Button(g, text="Footnote", width=9, command=self.insert_footnote).pack(side="left", padx=1)
        ttk.Button(g, text="Endnote", width=8, command=self.insert_endnote).pack(side="left", padx=1)
        ttk.Button(g, text="Show notes", width=10, command=self.show_footnotes).pack(side="left", padx=1)
        g = self._grp(f, "Citations")
        ttk.Button(g, text="Sources", width=8, command=self.citation_manager).pack(side="left", padx=1)
        ttk.Button(g, text=".bib", width=5, command=self.import_bibtex).pack(side="left", padx=1)
        ttk.Button(g, text="Caption", width=8, command=self.insert_caption).pack(side="left", padx=1)
        ttk.Button(g, text="X-ref", width=6, command=self.cross_ref).pack(side="left", padx=1)
        g = self._grp(f, "Index")
        ttk.Button(g, text="Mark", width=6, command=self.mark_index).pack(side="left", padx=1)
        ttk.Button(g, text="Index", width=6, command=self.insert_index).pack(side="left", padx=1)

    def _tab_mail(self):
        f = self._tab("Mailings")
        g = self._grp(f, "Merge")
        ttk.Button(g, text="Start merge", width=10, command=self.mail_merge).pack(side="left", padx=1)
        ttk.Button(g, text="Field << >>", width=10, command=self.insert_merge_field).pack(side="left", padx=1)
        ttk.Button(g, text="Preview", width=8, command=self.merge_preview).pack(side="left", padx=1)
        ttk.Button(g, text="Finish", width=7, command=self.merge_finish).pack(side="left", padx=1)
        g = self._grp(f, "Create")
        ttk.Button(g, text="Envelope", width=9, command=lambda: self.load_doc(WordDoc.template("envelope"))).pack(side="left", padx=1)
        ttk.Button(g, text="Labels", width=7, command=lambda: self.load_doc(WordDoc.template("label"))).pack(side="left", padx=1)

    def _tab_review(self):
        f = self._tab("Review")
        g = self._grp(f, "Proofing")
        ttk.Button(g, text="Spelling F7", width=11, command=self.spell_dialog).pack(side="left", padx=1)
        ttk.Button(g, text="Thesaurus", width=9, command=self.thesaurus).pack(side="left", padx=1)
        ttk.Button(g, text="Word count", width=10, command=self.wordcount_dialog).pack(side="left", padx=1)
        ttk.Button(g, text="Editor ★", width=9, command=self.editor_pane).pack(side="left", padx=1)
        ttk.Button(g, text="Goal", width=6, command=self.set_goal).pack(side="left", padx=1)
        ttk.Button(g, text="AutoCap", width=8, command=self.autocorrect_mgr).pack(side="left", padx=1)
        ttk.Button(g, text="Language", width=9, command=self.set_language).pack(side="left", padx=1)
        g = self._grp(f, "Comments")
        ttk.Button(g, text="New", width=6, command=self.add_comment).pack(side="left", padx=1)
        ttk.Button(g, text="Pane", width=6, command=self.toggle_comments).pack(side="left", padx=1)
        ttk.Button(g, text="Resolve", width=8, command=self.resolve_comment).pack(side="left", padx=1)
        g = self._grp(f, "Tracking")
        ttk.Button(g, text="Track: OFF", width=10, command=self.toggle_track).pack(side="left", padx=1)
        self.track_btn = g.winfo_children()[-1]
        ttk.Button(g, text="Accept all", width=9, command=self.accept_all).pack(side="left", padx=1)
        ttk.Button(g, text="Compare", width=8, command=self.compare_docs).pack(side="left", padx=1)
        g = self._grp(f, "Protect")
        ttk.Button(g, text="Restrict", width=8, command=self.restrict_editing).pack(side="left", padx=1)
        ttk.Button(g, text="Inspect", width=8, command=self.inspect_dialog).pack(side="left", padx=1)
        ttk.Button(g, text="Access.", width=8, command=self.accessibility).pack(side="left", padx=1)

    def _tab_view(self):
        f = self._tab("View")
        g = self._grp(f, "Views")
        for v, label in [("print", "Print"), ("web", "Web"), ("draft", "Draft"),
                         ("read", "Read"), ("focus", "Focus")]:
            ttk.Button(g, text=label, width=6, command=lambda vv=v: self.set_view(vv)).pack(side="left", padx=1)
        g = self._grp(f, "Show")
        ttk.Button(g, text="Ruler", width=7, command=self.toggle_ruler).pack(side="left", padx=1)
        ttk.Button(g, text="¶ Marks", width=7, command=self.toggle_marks).pack(side="left", padx=1)
        ttk.Button(g, text="Nav pane", width=8, command=self.toggle_nav).pack(side="left", padx=1)
        ttk.Button(g, text="Typewriter", width=10, command=self.toggle_typewriter).pack(side="left", padx=1)
        ttk.Button(g, text="Full scr", width=8, command=self.toggle_fullscreen).pack(side="left", padx=1)
        g = self._grp(f, "Zoom")
        for z in ["100%", "Page width", "150%"]:
            ttk.Button(g, text=z, width=9, command=lambda zz=z: self.zoom_preset(zz)).pack(side="left", padx=1)
        g = self._grp(f, "Window & macros")
        ttk.Button(g, text="New window", width=10, command=self.new_window).pack(side="left", padx=1)
        ttk.Button(g, text="Macros", width=7, command=self.macro_dialog).pack(side="left", padx=1)
        ttk.Button(g, text="● Rec", width=6, command=self.macro_record).pack(side="left", padx=1)

    # ================= ruler =================
    def _ruler(self):
        self.ruler = tk.Canvas(self.root, height=26, bg=self.theme["panel2"], highlightthickness=0)
        self.ruler.pack(side="top", fill="x")
        self.ruler.bind("<Button-1>", self._ruler_click)
        self._draw_ruler()

    def _draw_ruler(self):
        r = self.ruler
        try: r.delete("all")
        except Exception: return
        w = r.winfo_width() or 1200
        r.create_rectangle(0, 0, w, 26, fill=self.theme["panel2"], outline="")
        pxcm = 28 * self.zoom
        ml = self.doc.setup.margin_left * pxcm + 20
        x = 20
        n = 0
        while x < w:
            h = 12 if n % 2 == 0 else 7
            r.create_line(x, 26, x, 26 - h, fill=self.theme["muted"])
            if n % 2 == 0:
                r.create_text(x + 2, 6, text=str(n // 2), anchor="w", font=("Arial", 7), fill=self.theme["muted"])
            x += pxcm / 2; n += 1
        r.create_rectangle(0, 0, ml, 26, fill=self.theme["border"], outline="")
        r.create_polygon(ml, 4, ml - 6, 14, ml + 6, 14, fill=BRAND["blue"], tags="indent")

    def _ruler_click(self, e):
        pxcm = 28 * self.zoom
        ml_cm = max(0.5, (e.x - 20) / pxcm)
        self.doc.setup.margin_left = round(ml_cm, 2)
        self._apply_margins(); self._draw_ruler()

    def toggle_ruler(self):
        if self.ruler.winfo_ismapped(): self.ruler.pack_forget()
        else: self.ruler.pack(side="top", fill="x", after=self.nb)

    def _apply_margins(self):
        pad = int(28 + self.doc.setup.margin_left * 6 * self.zoom)
        try: self.text.configure(padx=pad)
        except Exception: pass

    # ================= body =================
    def _body(self):
        self.mid = ttk.Frame(self.root); self.mid.pack(fill="both", expand=True)
        # --- navigation pane
        self.nav_frame = ttk.Frame(self.mid, width=250)
        self.nav_frame.pack(side="left", fill="y"); self.nav_frame.pack_propagate(False)
        self.nav_nb = ttk.Notebook(self.nav_frame); self.nav_nb.pack(fill="both", expand=True)
        f1 = ttk.Frame(self.nav_nb); self.nav_nb.add(f1, text="Headings")
        self.nav_lb = tk.Listbox(f1, bg=self.theme["panel2"], fg=self.theme["fg"], relief="flat")
        self.nav_lb.pack(fill="both", expand=True, padx=4, pady=2)
        self.nav_lb.bind("<Double-Button-1>", self.nav_goto)
        f2 = ttk.Frame(self.nav_nb); self.nav_nb.add(f2, text="Pages")
        self.pages_lb = tk.Listbox(f2, bg=self.theme["panel2"], fg=self.theme["fg"], relief="flat")
        self.pages_lb.pack(fill="both", expand=True, padx=4, pady=2)
        self.pages_lb.bind("<Double-Button-1>", self.page_goto)
        f3 = ttk.Frame(self.nav_nb); self.nav_nb.add(f3, text="Search")
        self.search_var = tk.StringVar()
        ttk.Entry(f3, textvariable=self.search_var, width=24).pack(fill="x", padx=4, pady=2)
        ttk.Button(f3, text="Find all", command=self.nav_search).pack(fill="x", padx=4)
        self.search_lb = tk.Listbox(f3, bg=self.theme["panel2"], fg=self.theme["fg"], relief="flat")
        self.search_lb.pack(fill="both", expand=True, padx=4, pady=2)
        self.search_lb.bind("<Double-Button-1>", self.search_goto)
        # --- editor + gutters
        center = tk.Frame(self.mid, bg=self.theme["doc"])
        center.pack(side="left", fill="both", expand=True)
        self.canvas_bg = center
        editrow = tk.Frame(center, bg=self.theme["doc"])
        editrow.pack(fill="both", expand=True)
        self.gutter = tk.Text(editrow, width=4, bg=self.theme["panel2"], fg=self.theme["muted"],
                              relief="flat", state="disabled", font=("Consolas", 10))
        vs = ttk.Scrollbar(editrow, orient="vertical")
        hs = ttk.Scrollbar(editrow, orient="horizontal")
        self.text = tk.Text(editrow, wrap="word", undo=True, maxundo=300,
                            bg=self.theme["page"], fg=self.theme["fg"],
                            insertbackground=self.theme["fg"],
                            font=("Calibri", 11), padx=40, pady=30,
                            selectbackground=self.theme["accent"],
                            yscrollcommand=self._on_scroll, xscrollcommand=hs.set)
        vs.configure(command=self._on_scrollbar)
        hs.configure(command=self.text.xview)
        vs.pack(side="right", fill="y"); hs.pack(side="bottom", fill="x")
        self.text.pack(side="left", fill="both", expand=True, padx=16, pady=12)
        self.text.bind("<<Modified>>", self.on_modified)
        self.text.bind("<KeyRelease>", self.on_keyrelease)
        # ink overlay for the Draw tab (placed over the editor on demand)
        self.ink = tk.Canvas(editrow, highlightthickness=0, bg=self.theme["doc"])
        self.ink.bind("<B1-Motion>", self._ink_draw)
        self.ink.bind("<ButtonPress-1>", self._ink_start)
        # --- comments pane
        self.cmt_frame = ttk.Frame(self.mid, width=260)
        ttk.Label(self.cmt_frame, text=" Comments ").pack()
        self.cmt_lb = tk.Listbox(self.cmt_frame, bg=self.theme["panel2"], fg=self.theme["fg"], relief="flat", height=12)
        self.cmt_lb.pack(fill="both", expand=True, padx=4)
        ttk.Label(self.cmt_frame, text=" Tracked changes ").pack()
        self.chg_lb = tk.Listbox(self.cmt_frame, bg=self.theme["panel2"], fg=self.theme["fg"], relief="flat", height=8)
        self.chg_lb.pack(fill="both", expand=True, padx=4)
        brow = ttk.Frame(self.cmt_frame); brow.pack(fill="x")
        ttk.Button(brow, text="Resolve", command=self.resolve_comment).pack(side="left", padx=2)
        ttk.Button(brow, text="Delete", command=self.delete_comment).pack(side="left", padx=2)
        ttk.Button(brow, text="Accept all", command=self.accept_all).pack(side="left", padx=2)

    def _on_scroll(self, *args):
        self.gutter.yview_moveto(args[0] if len(args) == 1 else args[1])
        if hasattr(self, "_vs"): pass
        # keep scrollbar in sync manually
        for sb in self.mid.winfo_children():
            pass
        self._sync_scrollbar(args)

    def _on_scrollbar(self, *args):
        self.text.yview(*args)
        self.gutter.yview(*args)

    def _sync_scrollbar(self, args):
        # find the vertical scrollbar and move it
        for w in self.text.master.winfo_children():
            if isinstance(w, ttk.Scrollbar) and str(w.cget("orient")) == "vertical":
                w.set(*args if len(args) == 2 else (args[0], args[0]))

    # ================= status bar =================
    def _statusbar(self):
        sb = ttk.Frame(self.root); sb.pack(side="bottom", fill="x")
        self.page_lbl = ttk.Label(sb, text="Page 1 of 1"); self.page_lbl.pack(side="left", padx=8)
        self.word_lbl = ttk.Label(sb, text="0 words"); self.word_lbl.pack(side="left", padx=8)
        self.lang_btn = ttk.Button(sb, text=self.settings.get("lang", "en-US"), width=8, command=self.set_language)
        self.lang_btn.pack(side="left", padx=4)
        self.track_lbl = ttk.Label(sb, text="Track: off"); self.track_lbl.pack(side="left", padx=8)
        self.goal_lbl = ttk.Label(sb, text=""); self.goal_lbl.pack(side="left", padx=8)
        self.zoom_lbl = ttk.Label(sb, text="100%"); self.zoom_lbl.pack(side="right", padx=4)
        self.zoom_scale = ttk.Scale(sb, from_=50, to=200, value=100, length=130,
                                    command=lambda v: self.set_zoom(float(v) / 100))
        self.zoom_scale.pack(side="right", padx=4)
        for v, label in [("print", "Print"), ("web", "Web"), ("draft", "Draft"), ("read", "Read")]:
            ttk.Button(sb, text=label, width=6, command=lambda vv=v: self.set_view(vv)).pack(side="right", padx=1)

    # ================= tags =================
    def _config_tags(self):
        z = self.zoom
        base = self.doc.theme.body_font
        head = self.doc.theme.heading_font
        self.text.tag_configure("bold", font=(base, int(11 * z), "bold"))
        self.text.tag_configure("italic", font=(base, int(11 * z), "italic"))
        self.text.tag_configure("underline", underline=True)
        self.text.tag_configure("strike", overstrike=True)
        self.text.tag_configure("super", offset=f"{int(4 * z)}p", font=(base, int(8 * z)))
        self.text.tag_configure("sub", offset=f"{-int(2 * z)}p", font=(base, int(8 * z)))
        self.text.tag_configure("Title", font=(head, int(26 * z), "bold"), justify="center", spacing1=10, spacing3=10)
        self.text.tag_configure("Subtitle", font=(head, int(14 * z), "italic"), justify="center", foreground="#5b6478")
        self.text.tag_configure("Heading 1", font=(head, int(18 * z), "bold"), spacing1=12, spacing3=6, foreground=BRAND["blue"])
        self.text.tag_configure("Heading 2", font=(head, int(15 * z), "bold"), spacing1=10, spacing3=4, foreground=BRAND["blue"])
        self.text.tag_configure("Heading 3", font=(head, int(13 * z), "bold"), spacing1=8, spacing3=4)
        self.text.tag_configure("Heading 4", font=(head, int(12 * z), "bold italic"), spacing1=6, spacing3=4)
        self.text.tag_configure("Quote", font=(base, int(11 * z), "italic"), lmargin1=30, lmargin2=30,
                                background="#eef3ff" if not self.dark_mode else "#2c3550")
        self.text.tag_configure("Intense Quote", font=(base, int(11 * z), "bold italic"), lmargin1=30,
                                borderwidth=1, relief="solid")
        self.text.tag_configure("Caption", font=(base, int(9 * z), "italic"), justify="center", foreground="#5b6478")
        self.text.tag_configure("List Bullet", lmargin1=36, lmargin2=36)
        self.text.tag_configure("List Number", lmargin1=36, lmargin2=36)
        self.text.tag_configure("pagebreak", justify="center", foreground="#5b6478", spacing1=8, spacing3=8)
        self.text.tag_configure("toc", foreground=BRAND["blue"])
        self.text.tag_configure("fnmark", offset=f"{int(4 * z)}p", foreground=BRAND["blue"])
        self.text.tag_configure("cmark", background="#fff3bf" if not self.dark_mode else "#5c4d00")
        self.text.tag_configure("tracked", underline=True)
        self.text.tag_configure("pmark", foreground="#8b94a7")
        for al in ("left", "center", "right", "justify"):
            # Tk has no "justify" rendering: display as left, export keeps justify.
            self.text.tag_configure(f"al-{al}", justify="left" if al == "justify" else al)

    # ================= doc <-> editor =================
    def load_doc(self, doc: WordDoc):
        self.doc = doc
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.tables.clear(); self.tocs.clear()
        for b in doc.blocks:
            if isinstance(b, Table):
                line = self.text.index("end-1c").split(".")[0]
                cap = f" — {b.caption}" if b.caption else ""
                self.text.insert("end", f"[Table {len(b.rows)}x{len(b.rows[0]) if b.rows else 0}{cap} — double-click to edit]\n", ("table-ph",))
                self.tables[line] = b
            elif isinstance(b, PageBreak):
                self.text.insert("end", "— — — Page Break — — —\n", ("pagebreak",))
            elif isinstance(b, TOCField):
                line = self.text.index("end-1c").split(".")[0]
                self.text.insert("end", f"{b.title}\n", ("toc", "bold"))
                for style, text, page in doc.build_toc(b.levels):
                    lvl = int(style.split()[-1]) if style.split()[-1].isdigit() else 1
                    self.text.insert("end", f"{'   ' * (lvl - 1)}{text} ... {page}\n", ("toc",))
                self.tocs[line] = b
            else:
                tags = []
                if b.style != "Normal": tags.append(b.style)
                if b.align != "left": tags.append(f"al-{b.align}")
                pre = "• " if b.style == "List Bullet" else ("1. " if b.style == "List Number" else "")
                start = self.text.index("end-1c")
                body = "".join(r.text if isinstance(r, Run) else "" for r in b.runs)
                self.text.insert("end", pre + (body or "") + "\n")
                for r in b.runs:
                    if isinstance(r, FootnoteRef):
                        self.text.insert(f"{start} lineend-1c", f"[{r.nid}]", ("fnmark",))
                end = self.text.index("end-1c")
                pos = len(pre)
                for r in b.runs:
                    if not isinstance(r, Run): continue
                    n = len(r.text)
                    if n:
                        rtags = list(tags)
                        for flag, tg in [(r.bold, "bold"), (r.italic, "italic"), (r.underline, "underline"),
                                         (r.strike, "strike"), (r.super, "super"), (r.sub, "sub")]:
                            if flag: rtags.append(tg)
                        if r.color:
                            tg = f"fg-{r.color}"
                            try: self.text.tag_configure(tg, foreground="#" + r.color)
                            except Exception: pass
                            rtags.append(tg)
                        if r.highlight:
                            tg = f"hl-{r.highlight}"
                            try: self.text.tag_configure(tg, background="#" + r.highlight)
                            except Exception: pass
                            rtags.append(tg)
                        if r.size:
                            tg = f"sz-{r.size}"
                            try: self.text.tag_configure(tg, font=(self.doc.theme.body_font, int(r.size * self.zoom)))
                            except Exception: pass
                            rtags.append(tg)
                        if rtags:
                            self.text.tag_add(tuple(rtags) if len(rtags) > 1 else rtags[0],
                                              f"{start}+{pos}c", f"{start}+{pos + n}c")
                        pos += n
                if tags:
                    self.text.tag_add(tuple(tags) if len(tags) > 1 else tags[0], start, end)
                if b.shading:
                    tg = f"sh-{b.shading}"
                    try: self.text.tag_configure(tg, background="#" + b.shading)
                    except Exception: pass
                    self.text.tag_add(tg, start, end)
                if b.indent_left or b.indent_right or b.first_line:
                    tg = f"pgind-{start}"
                    try:
                        self.text.tag_configure(tg, lmargin1=(b.indent_left + b.first_line) * 18,
                                                lmargin2=b.indent_left * 18, rmargin=b.indent_right * 18)
                        self.text.tag_add(tg, start, end)
                    except Exception: pass
        self.text.edit_reset(); self.text.edit_modified(False)
        if self.doc.protection:
            self.text.configure(state="disabled")
        self.root.title(f"{BRAND['name']} — {doc.title}")
        self.refresh_nav(); self.refresh_comments(); self.refresh_changes(); self.update_status(); self._draw_ruler()

    def _line_tags(self, line: int) -> set:
        try: return set(self.text.tag_names(f"{line}.0"))
        except Exception: return set()

    def text_to_doc(self) -> WordDoc:
        self._strip_marks()
        content = self.text.get("1.0", "end-1c").split("\n")
        doc = WordDoc(title=self.doc.title, author=self.doc.author, subject=self.doc.subject,
                      keywords=self.doc.keywords, language=self.doc.language)
        doc.setup = self.doc.setup; doc.theme = self.doc.theme; doc.hf = self.doc.hf
        doc.watermark, doc.page_color = self.doc.watermark, self.doc.page_color
        doc.footnotes, doc.endnotes = self.doc.footnotes, self.doc.endnotes
        doc.comments, doc.citations, doc.index = self.doc.comments, self.doc.citations, self.doc.index
        doc.track_changes = self.track.get()
        for i, line in enumerate(content, start=1):
            if str(i) in self.tables:
                doc.blocks.append(self.tables[str(i)]); continue
            if "pagebreak" in self._line_tags(i):
                doc.blocks.append(PageBreak()); continue
            if str(i) in self.tocs:
                doc.blocks.append(self.tocs[str(i)]); continue
            if "toc" in self._line_tags(i):
                continue
            if not line.strip():
                doc.blocks.append(Paragraph([Run("")])); continue
            tags = self._line_tags(i)
            style = next((s for s in STYLES + list(self.custom_styles) if s in tags), "Normal")
            align = next((a for a in ("left", "center", "right", "justify") if f"al-{a}" in tags), "left")
            t = line
            if style == "List Bullet" and t.startswith("• "): t = t[2:]
            if style == "List Number" and re.match(r"^\d+\.\s", t): t = re.sub(r"^\d+\.\s", "", t)
            # footnote markers [n] with fnmark tag
            runs: list = []
            # split footnote markers out
            segs = re.split(r"(\[\d+\])", t)
            off = len(line) - len(t)
            pos = 0
            for seg in segs:
                if not seg: continue
                idx = f"{i}.{off + pos}"
                tg = set(self.text.tag_names(idx)) if self.text.compare(idx, "<", f"{i}.end") else set()
                if re.fullmatch(r"\[\d+\]", seg) and "fnmark" in tg:
                    runs.append(FootnoteRef(int(seg.strip("[]"))))
                else:
                    runs.extend(self._runs_for(seg, i, off + pos))
                pos += len(seg)
            # indent / shading
            il = ir = fl = 0
            for tg in tags:
                if tg.startswith("pgind-"):
                    try:
                        il = int(int(self.text.tag_cget(tg, "lmargin2")) / 18)
                        ir = int(int(self.text.tag_cget(tg, "rmargin")) / 18)
                        fl = int(int(self.text.tag_cget(tg, "lmargin1")) / 18) - il
                    except Exception: pass
            shading = next((x[3:] for x in tags if x.startswith("sh-")), None)
            doc.blocks.append(Paragraph(runs if runs else [Run(t)], style=style, align=align,
                                        indent_left=il, indent_right=ir, first_line=fl, shading=shading))
        return doc

    def _runs_for(self, t: str, line: int, col0: int) -> list[Run]:
        runs: list[Run] = []
        buf = ""; cur = None
        for j, ch in enumerate(t):
            idx = f"{line}.{col0 + j}"
            try: tg = set(self.text.tag_names(idx))
            except Exception: tg = set()
            key = (("bold" in tg), ("italic" in tg), ("underline" in tg), ("strike" in tg),
                   ("super" in tg), ("sub" in tg),
                   next((x[3:] for x in tg if x.startswith("fg-")), None),
                   next((x[4:] for x in tg if x.startswith("hl-")), None),
                   next((x[3:] for x in tg if x.startswith("sz-")), None))
            if cur is None: cur, buf = key, ch
            elif key == cur: buf += ch
            else:
                runs.append(Run(buf, bold=cur[0], italic=cur[1], underline=cur[2], strike=cur[3],
                                super=cur[4], sub=cur[5], color=cur[6], highlight=cur[7],
                                size=int(cur[8]) if cur[8] and str(cur[8]).isdigit() else None))
                cur, buf = key, ch
        if buf or not runs:
            runs.append(Run(buf, bold=cur[0] if cur else False, italic=cur[1] if cur else False,
                            underline=cur[2] if cur else False, strike=cur[3] if cur else False,
                            super=cur[4] if cur else False, sub=cur[5] if cur else False,
                            color=cur[4] if False else (cur[6] if cur else None),
                            highlight=cur[7] if cur else None,
                            size=int(cur[8]) if cur and cur[8] and str(cur[8]).isdigit() else None))
        return runs

    # ================= backstage (File) =================
    def _backstage(self):
        self.bs = tk.Frame(self.mid, bg=self.theme["panel"])
        self.bs_menu = tk.Listbox(self.bs, bg=self.theme["panel2"], fg=self.theme["fg"], relief="flat",
                                  font=("Segoe UI", 11), width=18)
        for item in ["New", "Open", "Info", "Save As", "Export", "Print", "Close"]:
            self.bs_menu.insert("end", item)
        self.bs_menu.bind("<<ListboxSelect>>", self._bs_show)
        self.bs_body = tk.Frame(self.bs, bg=self.theme["panel"])

    def show_backstage(self, item="New"):
        self.bs.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.bs_menu.pack(side="left", fill="y", padx=8, pady=8)
        self.bs_body.pack(side="left", fill="both", expand=True, padx=8, pady=8)
        items = self.bs_menu.get(0, "end")
        if item in items:
            self.bs_menu.selection_clear(0, "end")
            self.bs_menu.selection_set(items.index(item))
        self._bs_show()

    def hide_backstage(self):
        self.bs.place_forget()
        # return to the tab the user was on (first non-File tab)
        for tab in self.nb.tabs():
            if self.nb.tab(tab, "text").strip() != "File":
                self.nb.select(tab); break

    def _bs_show(self, _e=None):
        for w in self.bs_body.winfo_children(): w.destroy()
        sel = self.bs_menu.curselection()
        item = self.bs_menu.get(sel[0]) if sel else "New"
        b = self.bs_body
        if item == "New":
            tk.Label(b, text="New document", font=("Segoe UI", 16, "bold"), bg=self.theme["panel"], fg=self.theme["fg"]).pack(pady=8)
            for label, kind in [("Blank document", "blank"), ("Cover page", "cover"),
                                ("Open Studio Handbook (sample)", "sample"), ("Formal letter", "letter"),
                                ("Report", "report"), ("Resume", "resume")]:
                ttk.Button(b, text=label, width=42,
                           command=lambda k=kind: (self.load_doc(WordDoc.template(k)), self.__setattr__("current_path", None), self.hide_backstage())).pack(pady=3)
        elif item == "Open":
            tk.Label(b, text="Open", font=("Segoe UI", 16, "bold"), bg=self.theme["panel"], fg=self.theme["fg"]).pack(pady=8)
            ttk.Button(b, text="Browse...", width=30, command=lambda: (self.file_open(), self.hide_backstage())).pack(pady=4)
            tk.Label(b, text="Recent:", bg=self.theme["panel"], fg=self.theme["fg"]).pack(pady=(10, 2))
            for p in self.settings.get("recent", [])[:8]:
                ttk.Button(b, text=p, width=60, command=lambda pp=p: (self.open_path(pp), self.hide_backstage())).pack(pady=1)
        elif item == "Info":
            d = self.text_to_doc().stats()
            tk.Label(b, text="Info", font=("Segoe UI", 16, "bold"), bg=self.theme["panel"], fg=self.theme["fg"]).pack(pady=8)
            for k, v in [("Title", self.doc.title), ("Author", self.doc.author), ("Words", d["words"]),
                         ("Pages ≈", d["pages"]), ("Tables", d["tables"]), ("Comments", d["comments"]),
                         ("Tracking", "on" if self.track.get() else "off"),
                         ("Protection", "locked" if self.doc.protection else "none")]:
                tk.Label(b, text=f"{k}: {v}", bg=self.theme["panel"], fg=self.theme["fg"]).pack(anchor="w", padx=20)
            ttk.Button(b, text="Document inspector...", command=self.inspect_dialog).pack(pady=4)
            ttk.Button(b, text="Versions / AutoRecover...", command=self.versions_browser).pack(pady=2)
            ttk.Button(b, text="Preferences...", command=self.preferences).pack(pady=2)
        elif item == "Save As":
            tk.Label(b, text="Save a copy", font=("Segoe UI", 16, "bold"), bg=self.theme["panel"], fg=self.theme["fg"]).pack(pady=8)
            ttk.Button(b, text="Word (*.docx)", width=30, command=lambda: (self.file_saveas(), self.hide_backstage())).pack(pady=2)
            for label, fn in [("PDF", self.export_pdf), ("EPUB ebook", self.export_epub),
                              ("HTML", self.export_html),
                              ("Markdown", self.export_md), ("Plain text", self.export_txt)]:
                ttk.Button(b, text=label, width=30, command=lambda f=fn: (f(), self.hide_backstage())).pack(pady=2)
        elif item == "Export":
            tk.Label(b, text="Export", font=("Segoe UI", 16, "bold"), bg=self.theme["panel"], fg=self.theme["fg"]).pack(pady=8)
            for label, fn in [("Export PDF", self.export_pdf), ("Export EPUB", self.export_epub),
                              ("Export HTML", self.export_html),
                              ("Export Markdown", self.export_md), ("Export TXT", self.export_txt)]:
                ttk.Button(b, text=label, width=30, command=fn).pack(pady=2)
        elif item == "Print":
            tk.Label(b, text="Print", font=("Segoe UI", 16, "bold"), bg=self.theme["panel"], fg=self.theme["fg"]).pack(pady=8)
            ttk.Button(b, text="Print...", width=30, command=self.print_dialog).pack(pady=4)
            ttk.Button(b, text="Print preview (PDF)...", width=30, command=self.print_preview).pack(pady=2)
        elif item == "Close":
            self.hide_backstage()

    # ================= file ops =================
    def file_new(self): self.show_backstage("New")
    def file_open(self):
        p = filedialog.askopenfilename(filetypes=OPEN_TYPES)
        if p: self.open_path(p)

    def open_path(self, p):
        try:
            doc = formats.load(p)
            self.current_path = p if p.lower().endswith(".docx") else None
            if self.current_path: self._remember(p); self._save_settings()
            self.load_doc(doc)
        except Exception as ex:
            messagebox.showerror("Error", str(ex))

    def file_save(self):
        if not self.current_path: return self.file_saveas()
        try:
            formats.save_docx(self.text_to_doc(), self.current_path)
            self.update_status("Saved ✓")
        except Exception as ex: messagebox.showerror("Error", str(ex))

    def file_saveas(self):
        p = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=SAVE_DOCX)
        if not p: return
        try:
            formats.save_docx(self.text_to_doc(), p)
            self.current_path = p; self._remember(p); self._save_settings()
            self.update_status(f"Saved: {os.path.basename(p)}")
        except Exception as ex: messagebox.showerror("Error", str(ex))

    def _export(self, ext, ftype, saver):
        p = filedialog.asksaveasfilename(defaultextension=ext, filetypes=[ftype])
        if not p: return
        try:
            saver(self.text_to_doc(), p)
            messagebox.showinfo(BRAND["name"], f"Exported:\n{p}")
        except Exception as ex: messagebox.showerror("Error", str(ex))

    def export_pdf(self): self._export(".pdf", ("PDF", "*.pdf"), formats.save_pdf)
    def export_epub(self): self._export(".epub", ("EPUB ebook", "*.epub"), formats.save_epub)
    def export_html(self): self._export(".html", ("HTML", "*.html"), formats.save_html)
    def export_md(self): self._export(".md", ("Markdown", "*.md"), formats.save_md)
    def export_txt(self): self._export(".txt", ("Text", "*.txt"), formats.save_txt)

    def doc_props(self):
        win = tk.Toplevel(self.root); win.title("Document properties"); win.geometry("360x260")
        t = tk.StringVar(value=self.doc.title); a = tk.StringVar(value=self.doc.author)
        s = tk.StringVar(value=self.doc.subject); k = tk.StringVar(value=self.doc.keywords)
        for label, var in [("Title:", t), ("Author:", a), ("Subject:", s), ("Keywords:", k)]:
            ttk.Label(win, text=label).pack(); ttk.Entry(win, textvariable=var, width=40).pack()
        def ok():
            self.doc.title, self.doc.author, self.doc.subject, self.doc.keywords = t.get(), a.get(), s.get(), k.get()
            self._save_settings(); win.destroy(); self.update_status()
        ttk.Button(win, text="OK", command=ok).pack(pady=10)

    # ---- print ----
    def _temp_pdf(self) -> str:
        fd, path = tempfile.mkstemp(suffix=".pdf"); os.close(fd)
        formats.save_pdf(self.text_to_doc(), path)
        return path

    def print_preview(self):
        p = self._temp_pdf()
        messagebox.showinfo("Print preview", f"Preview PDF saved:\n{p}")
        try:
            if os.name == "nt": os.startfile(p)
        except Exception: pass

    def print_dialog(self):
        dialogs.print_dialog(self.root, self.do_print)

    def do_print(self, copies: int):
        p = self._temp_pdf()
        try:
            if os.name == "nt":
                for _ in range(copies): os.startfile(p, "print")
                messagebox.showinfo("Print", f"Sent {copies} copie(s) to the printer.")
            else:
                messagebox.showinfo("Print", f"Print file ready:\n{p}\n(open it in your PDF viewer to print)")
        except Exception as ex:
            messagebox.showerror("Error", str(ex))

    # ================= clipboard / undo =================
    def cut(self): self.text.event_generate("<<Cut>>")
    def copy(self): self.text.event_generate("<<Copy>>")
    def paste(self): self.text.event_generate("<<Paste>>")
    def paste_plain(self):
        try: self.text.insert("insert", self.root.clipboard_get())
        except Exception: pass
    def undo(self):
        try: self.text.edit_undo()
        except Exception: pass
    def redo(self):
        try: self.text.edit_redo()
        except Exception: pass
    def repeat(self):
        if self.last_action:
            try: self.last_action()
            except Exception: pass

    # ================= formatting =================
    def _sel(self):
        try: return (self.text.index("sel.first"), self.text.index("sel.last"))
        except Exception:
            line = self.text.index("insert").split(".")[0]
            return (f"{line}.0", f"{line}.end")

    def _log_macro(self, cid, **kw):
        if self.recording: self.recording.record(cid, **kw)

    def toggle_tag(self, tag):
        a, b = self._sel()
        try:
            n = int(self.text.count(a, b, "chars")[0] or 1)
            has = all(tag in self.text.tag_names(f"{a}+{i}c") for i in range(min(5, n)))
        except Exception: has = False
        if has: self.text.tag_remove(tag, a, b)
        else: self.text.tag_add(tag, a, b)
        self.last_action = lambda t=tag: self.toggle_tag(t)
        self.update_status()

    def apply_font(self):
        a, b = self._sel()
        fam, sz = self.font_var.get(), self.size_var.get()
        try: sz = int(sz)
        except Exception: sz = 11
        try:
            self.text.tag_configure(f"sz-{sz}", font=(fam, int(sz * self.zoom)))
            self.text.tag_add(f"sz-{sz}", a, b)
        except Exception: pass

    def grow_font(self, d):
        try: sz = int(self.size_var.get()) + d
        except Exception: sz = 12
        sz = max(6, min(200, sz)); self.size_var.set(str(sz)); self.apply_font()

    def pick_fg(self):
        c, hx = colorchooser.askcolor(title="Font color")
        if c:
            a, b = self._sel(); tg = f"fg-{hx.strip('#').upper()}"
            self.text.tag_configure(tg, foreground=hx); self.text.tag_add(tg, a, b)

    def pick_hl(self):
        c, hx = colorchooser.askcolor(title="Highlight color")
        if c:
            a, b = self._sel(); tg = f"hl-{hx.strip('#').upper()}"
            self.text.tag_configure(tg, background=hx); self.text.tag_add(tg, a, b)

    def set_align(self, al):
        a, b = self._sel()
        for x in ("al-left", "al-center", "al-right", "al-justify"): self.text.tag_remove(x, a, b)
        la, lb = int(self.text.index(a).split(".")[0]), int(self.text.index(b).split(".")[0])
        for ln in range(la, lb + 1):
            self.text.tag_add(f"al-{al}", f"{ln}.0", f"{ln}.end")
        self._log_macro("format.align", align=al)

    def set_style(self, style):
        a, b = self._sel()
        for s in STYLES + list(self.custom_styles): self.text.tag_remove(s, a, b)
        la, lb = int(self.text.index(a).split(".")[0]), int(self.text.index(b).split(".")[0])
        for ln in range(la, lb + 1):
            txt = self.text.get(f"{ln}.0", f"{ln}.end")
            txt2 = re.sub(r"^[•\d\.\s]+", "", txt) if style in ("Normal", "No Spacing", "Title", "Subtitle") else txt
            if style == "List Bullet" and not txt.startswith("• "): txt2 = "• " + re.sub(r"^[•\d\.\s]+", "", txt)
            if style == "List Number" and not re.match(r"^\d+\.\s", txt): txt2 = "1. " + re.sub(r"^[•\d\.\s]+", "", txt)
            if txt2 != txt:
                self.text.delete(f"{ln}.0", f"{ln}.end"); self.text.insert(f"{ln}.0", txt2)
            if style not in ("Normal", "No Spacing"):
                self.text.tag_add(style, f"{ln}.0", f"{ln}.end")
            cust = self.custom_styles.get(style)
            if cust:
                if cust.get("bold"): self.text.tag_add("bold", f"{ln}.0", f"{ln}.end")
                if cust.get("italic"): self.text.tag_add("italic", f"{ln}.0", f"{ln}.end")
        self.refresh_nav(); self.update_status()

    def clear_format(self):
        a, b = self._sel()
        for t in self.text.tag_names(a):
            if t not in ("sel",): self.text.tag_remove(t, a, b)
        execute(self.text_to_doc(), "format.clear")
        self._log_macro("format.clear")

    def change_case(self):
        def apply(mode):
            try: t = self.text.get("sel.first", "sel.last")
            except Exception:
                messagebox.showinfo("Change case", "Select text first."); return
            if mode == "upper": out = t.upper()
            elif mode == "lower": out = t.lower()
            elif mode == "title": out = t.title()
            else: out = "".join(c.lower() if c.isupper() else c.upper() for c in t)
            self.text.delete("sel.first", "sel.last"); self.text.insert("insert", out)
            self.last_action = lambda: None
        dialogs.change_case_dialog(self.root, apply)

    def sort_paras(self):
        a, b = self._sel()
        la, lb = int(self.text.index(a).split(".")[0]), int(self.text.index(b).split(".")[0])
        lines = sorted(self.text.get(f"{la}.0", f"{lb}.end").split("\n"))
        self.text.delete(f"{la}.0", f"{lb}.end"); self.text.insert(f"{la}.0", "\n".join(lines))
        self.update_status("Sorted A→Z")

    def format_painter(self):
        try:
            self._painter = [t for t in self.text.tag_names("sel.first") if t != "sel"]
            self.update_status("Format copied — select text, Ctrl+Shift+V to apply")
        except Exception: messagebox.showinfo("Format painter", "Select some text first.")

    def paste_format(self):
        if not hasattr(self, "_painter"): return
        a, b = self._sel()
        for t in self._painter: self.text.tag_add(t, a, b)

    def indent(self, d):
        a, b = self._sel()
        la, lb = int(self.text.index(a).split(".")[0]), int(self.text.index(b).split(".")[0])
        for ln in range(la, lb + 1):
            old = next((t for t in self.text.tag_names(f"{ln}.0") if t.startswith("pgind-")), None)
            cur = 0
            if old:
                try: cur = int(int(self.text.tag_cget(old, "lmargin2")) / 18)
                except Exception: pass
                self.text.tag_remove(old, f"{ln}.0", f"{ln}.end")
            cur = max(0, cur + d)
            if cur:
                tg = f"pgind-{ln}.{cur}"
                self.text.tag_configure(tg, lmargin1=cur * 18, lmargin2=cur * 18)
                self.text.tag_add(tg, f"{ln}.0", f"{ln}.end")

    def line_spacing_menu(self):
        v = simpledialog.askstring("Line spacing", "1.0 / 1.15 / 1.5 / 2.0 (applies on export):",
                                   initialvalue="1.15", parent=self.root)
        if v: messagebox.showinfo("Line spacing", f"Set to {v} for the document (DOCX/PDF export).")

    def para_spacing(self):
        v = simpledialog.askstring("Paragraph spacing", "Space after in pt:", initialvalue="6", parent=self.root)
        if v and v.isdigit():
            for b in self.text_to_doc().blocks:
                if hasattr(b, "spacing_after"): b.spacing_after = int(v)
            messagebox.showinfo("Spacing", f"Space after set to {v} pt (applies on export).")

    def para_shading(self):
        c, hx = colorchooser.askcolor(title="Paragraph shading")
        if hx:
            a, b = self._sel(); tg = f"sh-{hx.strip('#').upper()}"
            self.text.tag_configure(tg, background=hx); self.text.tag_add(tg, a, b)

    def para_border(self):
        a, b = self._sel()
        la, lb = int(self.text.index(a).split(".")[0]), int(self.text.index(b).split(".")[0])
        for ln in range(la, lb + 1):
            self.text.tag_add("Intense Quote", f"{ln}.0", f"{ln}.end")

    def styles_pane(self):
        dialogs.style_manager_dialog(self.root, self.set_style, self.custom_styles)

    def new_style(self):
        d = dialogs.new_style_dialog(self.root)
        if d:
            self.custom_styles[d["name"]] = d; self._save_settings()
            if d["name"] not in STYLES: STYLES.append(d["name"])
            self.set_style(d["name"])

    # ---- find / replace / goto / select ----
    def find_dialog(self):
        w = simpledialog.askstring("Find", "Find what (prefix with re: for regex):", parent=self.root)
        if not w: return
        use_re = w.startswith("re:")
        pat = w[3:] if use_re else re.escape(w)
        self.text.tag_remove("found", "1.0", "end")
        self.text.tag_configure("found", background="yellow", foreground="black")
        body = self.text.get("1.0", "end-1c"); n = 0
        for m in re.finditer(pat, body):
            s = f"1.0+{m.start()}c"; e = f"1.0+{m.end()}c"
            self.text.tag_add("found", s, e); n += 1
            if n > 500: break
        if n: self.text.see("found.first")
        messagebox.showinfo("Find", f"Found {n} match(es).")

    def replace_dialog(self):
        win = tk.Toplevel(self.root); win.title("Find and Replace"); win.geometry("360x190")
        ttk.Label(win, text="Find what:").pack(); e1 = ttk.Entry(win, width=38); e1.pack(pady=2)
        ttk.Label(win, text="Replace with:").pack(); e2 = ttk.Entry(win, width=38); e2.pack(pady=2)
        def go():
            body = self.text.get("1.0", "end-1c"); w, r = e1.get(), e2.get()
            if not w: return
            n = len(re.findall(re.escape(w), body))
            self.text.delete("1.0", "end-1c"); self.text.insert("1.0", body.replace(w, r))
            messagebox.showinfo("Replace", f"Replaced {n}."); win.destroy(); self.update_status()
        ttk.Button(win, text="Replace All", command=go).pack(pady=8)

    def goto(self):
        dialogs.goto_dialog(self.root, self._goto_impl)

    def _goto_impl(self, what: str):
        w = what.strip().lower()
        m = re.match(r"page\s+(\d+)", w)
        if m:
            line = min(int(m.group(1)) * 45, int(self.text.index("end-1c").split(".")[0]))
            self.text.see(f"{line}.0"); self.text.mark_set("insert", f"{line}.0"); return
        m = re.match(r"line\s+(\d+)", w)
        if m:
            self.text.see(f"{m.group(1)}.0"); self.text.mark_set("insert", f"{m.group(1)}.0"); return
        idx = self.text.search(what, "1.0", nocase=True, stopindex="end")
        if idx: self.text.see(idx); self.text.mark_set("insert", idx)
        else: messagebox.showinfo("Go To", "Not found.")

    def select_menu(self):
        self.text.tag_add("sel", "1.0", "end")

    # ================= insert =================
    def insert_text(self, t):
        self.text.insert("insert", t)
        if self.track.get():
            self.text.tag_add("tracked", f"insert-{len(t)}c", "insert")
            self.doc.changes.append({"author": self.doc.author, "kind": "insert", "text": t[:40]})
        self.update_status()

    def insert_cover(self):
        self.text.insert("1.0", "DOCUMENT TITLE\nSubtitle goes here\nAuthor name — 2026\n— — — Page Break — — —\n")
        self.update_status("Cover page inserted")

    def insert_blank_page(self):
        self.insert_text("\n— — — Page Break — — —\n")

    def insert_page_break(self):
        line = self.text.index("insert").split(".")[0]
        self.text.insert("insert", "\n")
        self.text.insert("insert", "— — — Page Break — — —\n", ("pagebreak",))
        self.update_status("Page break inserted")

    def insert_table(self):
        t = dialogs.insert_table_dialog(self.root)
        if not t: return
        line = self.text.index("insert").split(".")[0]
        self.text.insert("insert", f"\n[Table {len(t.rows)}x{len(t.rows[0])} — double-click to edit]\n")
        self.tables[line] = t
        self.update_status("Table inserted")

    def edit_table(self):
        line = self.text.index("insert").split(".")[0]
        # nearest table marker at or above cursor
        for ln in [str(line)] + [str(i) for i in range(line - 1, max(0, line - 6), -1)]:
            if ln in self.tables:
                if dialogs.edit_table_dialog(self.root, self.tables[ln]):
                    self.update_status("Table updated")
                return
        messagebox.showinfo("Table", "Place the cursor on a [Table ...] line first.")

    def insert_image(self):
        from PIL import Image, ImageTk
        p = filedialog.askopenfilename(filetypes=[("Images", "*.png *.jpg *.jpeg *.gif *.bmp *.webp"), ("All", "*.*")])
        if not p: return
        try:
            img = Image.open(p); img.thumbnail((480, 480))
            ph = ImageTk.PhotoImage(img); self._images.append(ph)
            self.text.image_create("insert", image=ph)
            self.text.insert("insert", f" [Figure: {os.path.basename(p)}]\n")
        except Exception as ex: messagebox.showerror("Error", str(ex))

    def insert_shape(self):
        shapes = {"★ Star": "★", "→ Arrow": "→", "● Circle": "●", "■ Square": "■", "▲ Triangle": "▲", "✔ Check": "✔"}
        pick = simpledialog.askstring("Shape", "Pick: " + ", ".join(shapes), parent=self.root)
        if pick and pick in shapes: self.insert_text(f" {shapes[pick]} ")

    def insert_textbox(self):
        self.insert_text("\n")
        start = self.text.index("insert")
        self.insert_text("Type your text box content here.")
        self.text.tag_add("Intense Quote", start, "insert")
        self.insert_text("\n")

    def insert_pagenum(self):
        n = len(self.text.get("1.0", "end-1c").split("\n")) // 45 + 1
        self.insert_text(f" [Page {n}] ")

    def insert_datetime(self):
        self.insert_text(datetime.datetime.now().strftime("%B %d, %Y  %H:%M"))

    def insert_field(self):
        pick = simpledialog.askstring("Field", "TITLE / AUTHOR / DATE / WORDS / PAGES:", parent=self.root)
        if not pick: return
        k = pick.strip().upper()
        vals = {"TITLE": self.doc.title, "AUTHOR": self.doc.author,
                "DATE": datetime.date.today().isoformat(),
                "WORDS": str(self.text_to_doc().stats()["words"]),
                "PAGES": str(self.text_to_doc().stats()["pages"])}
        self.insert_text(vals.get(k, f"[{k}]"))

    def drop_cap(self):
        try:
            line = self.text.index("insert").split(".")[0]
            t = self.text.get(f"{line}.0", f"{line}.end")
            m = re.search(r"[A-Za-z]", t)
            if not m: return
            idx = f"{line}.{m.start()}"
            self.text.tag_configure("dropcap", font=(self.doc.theme.heading_font, int(36 * self.zoom), "bold"))
            self.text.tag_add("dropcap", idx, f"{idx}+1c")
        except Exception as ex: messagebox.showerror("Error", str(ex))

    def symbols(self):
        dialogs.symbol_dialog(self.root, self.insert_text)

    def insert_equation(self):
        eq = simpledialog.askstring("Equation", "Type the equation (unicode ok, e.g. E = mc²):", parent=self.root)
        if eq:
            self.caption_no["Equation"] += 1
            self.insert_text(f"  {eq}   (Equation {self.caption_no['Equation']})  ")

    # ================= draw =================
    def draw_mode(self, kind):
        self.ink.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.ink.lift()
        self._pen["width"] = 3 if kind == "pen" else 10
        self.update_status(f"Draw mode: {kind} — drag over the page. Esc hides.")

    def pen_color(self):
        c, hx = colorchooser.askcolor(title="Pen color")
        if hx: self._pen["color"] = hx

    def _ink_start(self, e):
        self._stroke_pts = [(e.x, e.y)]

    def _ink_draw(self, e):
        pts = getattr(self, "_stroke_pts", [])
        if not pts: return
        x0, y0 = pts[-1]
        self.ink.create_line(x0, y0, e.x, e.y, fill=self._pen["color"],
                             width=self._pen["width"], capstyle="round", smooth=True)
        pts.append((e.x, e.y))
        self._strokes.append((self._pen["color"], self._pen["width"], [(x0, y0), (e.x, e.y)]))

    def draw_clear(self):
        self.ink.delete("all"); self._strokes.clear(); self.ink.place_forget()

    def ink_to_picture(self):
        if not self._strokes:
            messagebox.showinfo("Ink", "Draw something first."); return
        from PIL import Image, ImageDraw
        w, h = max(100, self.ink.winfo_width()), max(100, self.ink.winfo_height())
        img = Image.new("RGBA", (w, h), (255, 255, 255, 0))
        d = ImageDraw.Draw(img)
        for color, width, pts in self._strokes:
            if len(pts) >= 2: d.line(pts, fill=color, width=width)
        from PIL import ImageTk
        ph = ImageTk.PhotoImage(img); self._images.append(ph)
        self.ink.place_forget()
        self.text.image_create("insert", image=ph)
        self.text.insert("insert", " [Ink drawing]\n")
        self._strokes.clear(); self.ink.delete("all")

    # ================= design =================
    def pick_theme(self):
        name = dialogs.theme_dialog(self.root, self.doc.theme.name)
        if name:
            self.doc.theme = THEMES[name]
            self._config_tags()
            self.load_doc(self.text_to_doc())
            self.update_status(f"Theme: {name}")

    def theme_fonts(self):
        h = simpledialog.askstring("Theme fonts", "Heading font:", initialvalue=self.doc.theme.heading_font, parent=self.root)
        b = simpledialog.askstring("Theme fonts", "Body font:", initialvalue=self.doc.theme.body_font, parent=self.root)
        if h: self.doc.theme.heading_font = h
        if b: self.doc.theme.body_font = b
        self._config_tags(); self.update_status("Theme fonts updated")

    def set_watermark(self):
        w = dialogs.watermark_dialog(self.root, self.doc.watermark)
        if w is not None:
            self.doc.watermark = w
            self.update_status(f"Watermark: {w or 'none'} (shows in DOCX/PDF/HTML export)")

    def page_color(self):
        c, hx = colorchooser.askcolor(title="Page color")
        if hx:
            self.doc.page_color = hx.strip("#")
            self.text.configure(bg=hx)

    def page_borders(self):
        self.text.configure(highlightthickness=2, highlightbackground=BRAND["blue"])
        messagebox.showinfo("Page borders", "A blue page border is shown; DOCX/PDF export uses document margins.")

    # ================= layout =================
    def page_setup(self):
        if dialogs.page_setup_dialog(self.root, self.doc.setup):
            self._apply_margins(); self._draw_ruler(); self.update_status()

    def margins_menu(self):
        pick = simpledialog.askstring("Margins", "Normal / Narrow / Wide (cm L,R or default):",
                                      initialvalue="Normal", parent=self.root)
        presets = {"normal": (2.54, 2.54), "narrow": (1.27, 1.27), "wide": (5.08, 5.08)}
        if pick and pick.lower() in presets:
            l, _ = presets[pick.lower()]
            self.doc.setup.margin_left = self.doc.setup.margin_right = l
            self._apply_margins(); self._draw_ruler(); self.update_status(f"Margins: {pick}")

    def set_orientation(self, o):
        self.doc.setup.orientation = o
        self.update_status(f"Orientation: {o} (applies on DOCX/PDF export)")

    def paper_size(self):
        pick = simpledialog.askstring("Paper size", "A4 / Letter / Legal:", initialvalue=self.doc.setup.size, parent=self.root)
        if pick and pick in ("A4", "Letter", "Legal"):
            self.doc.setup.size = pick
            self.update_status(f"Paper: {pick}")

    def columns_menu(self):
        v = simpledialog.askinteger("Columns", "1 - 3:", initialvalue=self.doc.setup.columns,
                                    minvalue=1, maxvalue=3, parent=self.root)
        if v:
            self.doc.setup.columns = v
            self.update_status(f"Columns: {v} (applies on DOCX export)")

    def breaks_menu(self):
        pick = simpledialog.askstring("Breaks", "page / column / section:", initialvalue="page", parent=self.root)
        if pick and pick.startswith("page"): self.insert_page_break()
        elif pick: self.insert_text(f"\n— — {pick} break — —\n")

    def toggle_line_numbers(self):
        self.doc.setup.line_numbers = not self.doc.setup.line_numbers
        self._sync_gutter(); self.update_status(f"Line numbers: {'on' if self.doc.setup.line_numbers else 'off'}")

    def toggle_hyphenation(self):
        self.doc.setup.hyphenation = not self.doc.setup.hyphenation
        self.update_status(f"Hyphenation: {'on' if self.doc.setup.hyphenation else 'off'} (applies at export)")

    def edit_header_footer(self):
        if dialogs.header_footer_dialog(self.root, self.doc.hf):
            self.update_status("Header/footer saved (shows in DOCX/PDF/HTML export)")

    # ================= references =================
    def insert_toc(self):
        line = self.text.index("insert").split(".")[0]
        toc = TOCField("Contents", 2)
        self.text.insert("insert", f"{toc.title}\n", ("toc", "bold"))
        for style, text, page in self.text_to_doc().build_toc(2):
            self.text.insert("insert", f"{text} ... {page}\n", ("toc",))
        self.tocs[line] = toc
        self.update_status("Table of contents inserted")

    def update_toc(self):
        self.load_doc(self.text_to_doc())
        self.update_status("TOC updated")

    def insert_footnote(self):
        t = simpledialog.askstring("Footnote", "Footnote text:", parent=self.root)
        if t is None: return
        nid = len(self.doc.footnotes) + len(self.doc.endnotes) + 1
        self.doc.footnotes.append(Footnote(nid, t or ""))
        self.text.insert("insert", f"[{nid}]", ("fnmark",))
        self.update_status(f"Footnote [{nid}] inserted")

    def insert_endnote(self):
        t = simpledialog.askstring("Endnote", "Endnote text:", parent=self.root)
        if t is None: return
        nid = len(self.doc.footnotes) + len(self.doc.endnotes) + 1
        self.doc.endnotes.append(Footnote(nid, t or ""))
        self.text.insert("insert", f"[{nid}]", ("fnmark",))

    def show_footnotes(self):
        notes = [f"[{f.nid}] {f.text}" for f in self.doc.footnotes + self.doc.endnotes]
        messagebox.showinfo("Footnotes & endnotes", "\n".join(notes) if notes else "No notes yet.")

    def citation_manager(self):
        dialogs.citation_manager_dialog(self.root, self.text_to_doc(), self.insert_text, self.update_status)
        # sync back citations added in the dialog
        self.doc.citations = self.text_to_doc().citations

    def insert_caption(self):
        label = simpledialog.askstring("Caption", "Label: Table / Figure / Equation", initialvalue="Table", parent=self.root)
        if not label: return
        label = label.capitalize()
        self.caption_no[label] = self.caption_no.get(label, 0) + 1
        self.insert_text(f"\n{label} {self.caption_no[label]}: description\n")

    def cross_ref(self):
        heads = self.text_to_doc().headings(9)
        if not heads:
            messagebox.showinfo("Cross-reference", "No headings to reference."); return
        pick = simpledialog.askstring("Cross-reference", "Heading:\n" + "\n".join(h[1][:50] for h in heads[:10]), parent=self.root)
        if pick: self.insert_text(f" (see “{pick}”) ")

    def mark_index(self):
        try: term = self.text.get("sel.first", "sel.last")
        except Exception:
            term = simpledialog.askstring("Mark index entry", "Term:", parent=self.root)
        if term:
            from engine import IndexEntry
            self.doc.index.append(IndexEntry(term.strip(), 1))
            self.update_status(f"Indexed: {term.strip()}")

    def insert_index(self):
        doc = self.text_to_doc()
        idx = doc.build_index()
        self.insert_text("\nIndex\n" + "\n".join(f"{t} ... {','.join(map(str, p))}" for t, p in idx) + "\n")

    # ================= mailings =================
    def _merge_state(self) -> dict:
        if not hasattr(self, "_mstate"):
            self._mstate = {"fields": [], "records": [], "i": 0}
        return self._mstate

    def mail_merge(self):
        st = self._merge_state()
        def preview(rec): self.update_status(f"Preview: {', '.join(f'{k}={v}' for k, v in list(rec.items())[:3])}")
        def finish(records):
            merged = self.text_to_doc().merge_all(records)
            self.load_doc(merged)
            self.update_status(f"Merged {len(records)} records")
        dialogs.mail_merge_dialog(self.root, self.text_to_doc(), self.insert_text, preview, finish)

    def insert_merge_field(self):
        st = self._merge_state()
        if st["fields"]:
            f = simpledialog.askstring("Merge field", f"Fields: {', '.join(st['fields'])}", parent=self.root)
            if f: self.insert_text(f"<<{f.strip()}>>")
        else:
            f = simpledialog.askstring("Merge field", "Field name:", parent=self.root)
            if f: self.insert_text(f"<<{f.strip()}>>")

    def merge_preview(self):
        st = self._merge_state()
        if not st["records"]:
            messagebox.showinfo("Preview", "Start a merge and load a CSV first."); return
        st["i"] = (st["i"] + 1) % len(st["records"])
        rec = st["records"][st["i"]]
        messagebox.showinfo("Preview", f"Record {st['i']+1}/{len(st['records'])}:\n" +
                            "\n".join(f"{k}: {v}" for k, v in rec.items()))

    def merge_finish(self):
        st = self._merge_state()
        if not st["records"]:
            messagebox.showinfo("Finish", "No recipients loaded."); return
        merged = self.text_to_doc().merge_all(st["records"])
        self.load_doc(merged)
        p = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=SAVE_DOCX)
        if p:
            formats.save_docx(merged, p)
            messagebox.showinfo("Merge", f"Merged {len(st['records'])} records →\n{p}")

    # ================= review =================
    def spell_dialog(self):
        words = re.findall(r"[A-Za-z']+", self.text.get("1.0", "end-1c"))
        freq: dict[str, int] = {}
        for w in words:
            if len(w) > 2: freq[w.lower()] = freq.get(w.lower(), 0) + 1
        vocab = set(freq) | self.user_dict
        suspects = []
        for w in words:
            wl = w.lower()
            if wl in self.user_dict or len(w) < 3: continue
            if re.search(r"(.)\1{3,}", w) or (freq[wl] == 1 and len(w) > 9 and not difflib.get_close_matches(wl, vocab - {wl}, 1, 0.9)):
                if w not in suspects: suspects.append(w)
            if len(suspects) >= 12: break
        if not suspects:
            messagebox.showinfo("Spelling (F7)", "Spelling check complete. No issues found. ✓"); return
        win = tk.Toplevel(self.root); win.title("Spelling"); win.geometry("380x300")
        ttk.Label(win, text="Suspect words (double-click to locate):").pack(pady=4)
        lb = tk.Listbox(win); lb.pack(fill="both", expand=True, padx=8)
        for w in suspects:
            sug = difflib.get_close_matches(w.lower(), vocab, 3, 0.8)
            lb.insert("end", f"{w}  →  {', '.join(sug) if sug else '(no suggestion)'}")
        def locate(_e=None):
            if lb.curselection():
                w = suspects[lb.curselection()[0]]
                idx = self.text.search(w, "1.0", stopindex="end")
                if idx: self.text.see(idx); self.text.mark_set("insert", idx)
        lb.bind("<Double-Button-1>", locate)
        brow = ttk.Frame(win); brow.pack(pady=4)
        ttk.Button(brow, text="Add all to dictionary",
                   command=lambda: (self.user_dict.update(s.lower() for s in suspects), self._save_settings(), win.destroy())).pack(side="left", padx=4)
        ttk.Button(brow, text="Close", command=win.destroy).pack(side="left", padx=4)

    def thesaurus(self):
        try: w = self.text.get("sel.first", "sel.last").strip().lower()
        except Exception: w = ""
        if not w:
            w = simpledialog.askstring("Thesaurus", "Word:", parent=self.root)
            if not w: return
            w = w.strip().lower()
        syns = THESAURUS.get(w, difflib.get_close_matches(
            w, set(re.findall(r"[a-z]+", self.text.get("1.0", "end-1c").lower())), 5, 0.8))
        messagebox.showinfo("Thesaurus", f"{w}: {', '.join(syns) if syns else 'no synonyms found'}")

    def wordcount_dialog(self):
        dialogs.word_count_dialog(self.root, self.text_to_doc().stats())

    def set_language(self):
        lang = simpledialog.askstring("Proofing language", "e.g. en-US, fr-FR, ar-SA:",
                                      initialvalue=self.settings.get("lang", "en-US"), parent=self.root)
        if lang:
            self.settings["lang"] = lang; self.doc.language = lang
            self.lang_btn.configure(text=lang); self._save_settings(); self.update_status()

    # ---- comments ----
    def toggle_comments(self):
        if self.cmt_frame.winfo_ismapped(): self.cmt_frame.pack_forget()
        else: self.cmt_frame.pack(side="right", fill="y")

    def refresh_comments(self):
        self.cmt_lb.delete(0, "end")
        for c in self.text_to_doc().comments:
            flag = "✓ " if c.resolved else ""
            self.cmt_lb.insert("end", f"{flag}[{c.cid}] {c.author}: {c.text[:60]}")

    def add_comment(self):
        try: sel = self.text.get("sel.first", "sel.last")
        except Exception: sel = ""
        t = simpledialog.askstring("New comment", f"Comment on: “{sel[:40]}”", parent=self.root)
        if t is None: return
        doc = self.text_to_doc()
        doc.comments.append(Comment(len(doc.comments) + 1, doc.author, sel[:80], t))
        if sel:
            try: self.text.tag_add("cmark", "sel.first", "sel.last")
            except Exception: pass
        self.doc.comments = doc.comments
        self.refresh_comments(); self.update_status("Comment added")

    def resolve_comment(self):
        sel = self.cmt_lb.curselection()
        doc = self.text_to_doc()
        if sel and sel[0] < len(doc.comments):
            doc.comments[sel[0]].resolved = True
            self.doc.comments = doc.comments
            self.refresh_comments()

    def delete_comment(self):
        sel = self.cmt_lb.curselection()
        doc = self.text_to_doc()
        if sel and sel[0] < len(doc.comments):
            doc.comments.pop(sel[0])
            self.doc.comments = doc.comments
            self.refresh_comments()

    # ---- tracking / compare / protect ----
    def toggle_track(self):
        self.track.set(not self.track.get())
        self.track_btn.configure(text=f"Track: {'ON' if self.track.get() else 'OFF'}")
        self.update_status()

    def accept_all(self):
        doc = self.text_to_doc()
        execute(doc, "review.acceptAll")
        self.text.tag_remove("tracked", "1.0", "end")
        self.doc.changes = doc.changes; self.doc.comments = doc.comments
        self.refresh_comments(); self.update_status("All changes accepted")

    def compare_docs(self):
        dialogs.compare_dialog(self.root, self._compare_impl)

    def _compare_impl(self, a, b):
        ta = formats.load(a).full_text().split("\n")
        tb = formats.load(b).full_text().split("\n")
        diff = list(difflib.unified_diff(ta, tb, lineterm=""))
        added = sum(1 for l in diff if l.startswith("+") and not l.startswith("+++"))
        removed = sum(1 for l in diff if l.startswith("-") and not l.startswith("---"))
        messagebox.showinfo("Compare", f"Added lines: {added}\nRemoved lines: {removed}\n(see details in message)")
        win = tk.Toplevel(self.root); win.title("Compare result"); win.geometry("600x400")
        t = tk.Text(win); t.pack(fill="both", expand=True)
        t.insert("1.0", "\n".join(diff[:2000]))

    def restrict_editing(self):
        if self.doc.protection:
            pw = simpledialog.askstring("Restrict editing", "Password to unlock:", show="*", parent=self.root)
            if pw == self.doc.protection:
                self.doc.protection = None; self.text.configure(state="normal")
                self.update_status("Protection removed")
            else: messagebox.showerror("Error", "Wrong password.")
        else:
            pw = simpledialog.askstring("Restrict editing", "Set a password:", show="*", parent=self.root)
            if pw:
                self.doc.protection = pw; self.text.configure(state="disabled")
                self.update_status("Document locked")

    def inspect_dialog(self):
        doc = self.text_to_doc()
        def clean(items):
            for it in items:
                if it == "Comments": doc.comments.clear()
                elif it == "Tracked changes": doc.changes.clear()
                elif it == "Hidden metadata": doc.subject = doc.keywords = ""
                elif it == "Headers/footers": doc.hf.header = doc.hf.footer = ""
                elif it == "Watermark": doc.watermark = ""
            self.load_doc(doc)
            messagebox.showinfo("Inspector", f"Removed: {', '.join(items)}")
        dialogs.inspector_dialog(self.root, doc, clean)

    def accessibility(self):
        messagebox.showinfo("Accessibility checker", dialogs.accessibility_report(self.root, self.text_to_doc()))

    # ================= view =================
    def set_view(self, v):
        self.view = v
        if v == "print":
            self.text.configure(padx=40, wrap="word"); self._apply_margins()
            self.nav_frame.pack_forget(); self.canvas_bg.pack_forget()
            self.nav_frame.pack(side="left", fill="y"); self.canvas_bg.pack(side="left", fill="both", expand=True)
            if self.cmt_frame.winfo_ismapped():
                self.cmt_frame.pack_forget(); self.cmt_frame.pack(side="right", fill="y")
        elif v == "web":
            self.text.configure(padx=60, wrap="word")
        elif v == "draft":
            self.text.configure(padx=8, wrap="word")
        elif v in ("read", "focus"):
            self.nav_frame.pack_forget(); self.cmt_frame.pack_forget()
            self.text.configure(padx=90 if v == "read" else 140, wrap="word")
            if v == "focus": self.update_status("Focus mode — Esc or View > Print to exit")
        self.update_status(f"View: {v}")

    def set_zoom(self, z):
        self.zoom = max(0.5, min(3.0, z))
        self._config_tags()
        for tg in self.text.tag_names():
            if tg.startswith("sz-"):
                try: self.text.tag_configure(tg, font=(self.doc.theme.body_font, int(int(tg[3:]) * self.zoom)))
                except Exception: pass
        try: self.zoom_scale.set(self.zoom * 100)
        except Exception: pass
        self._draw_ruler(); self.update_status()

    def zoom_preset(self, p):
        if p == "100%": self.set_zoom(1.0)
        elif p == "150%": self.set_zoom(1.5)
        elif p == "Page width": self.set_zoom(1.3)

    def toggle_nav(self):
        if self.nav_frame.winfo_ismapped(): self.nav_frame.pack_forget()
        else: self.nav_frame.pack(side="left", fill="y")

    def toggle_marks(self):
        self.marks = not self.marks
        if self.marks:
            self.text.tag_configure("pmark", foreground="#8b94a7")
            for ln in range(1, int(self.text.index("end-1c").split(".")[0]) + 1):
                self.text.insert(f"{ln}.end", "¶", ("pmark",))
        else:
            self._strip_marks()
        self.update_status(f"Formatting marks: {'on' if self.marks else 'off'}")

    def _strip_marks(self):
        for r in self.text.tag_ranges("pmark"):
            pass
        rng = self.text.tag_ranges("pmark")
        for i in range(0, len(rng), 2):
            try: self.text.delete(rng[i], rng[i + 1])
            except Exception: pass

    def new_window(self):
        win = tk.Toplevel(self.root); win.title(f"{BRAND['name']} — second window"); win.geometry("900x700")
        t = tk.Text(win, wrap="word", font=("Calibri", 11)); t.pack(fill="both", expand=True)
        t.insert("1.0", self.text.get("1.0", "end-1c"))

    # ---- macros ----
    def macro_record(self):
        if self.recording:
            name = simpledialog.askstring("Macro", "Save macro as:", parent=self.root) or "macro1"
            self.recording.name = name; self.recording.recording = False
            self.macros[name] = self.recording; self.recording = None
            self.update_status(f"Macro saved: {name}")
        else:
            self.recording = Macro("recording"); self.recording.recording = True
            self.update_status("Recording macro... click ● Rec to stop")

    def macro_dialog(self):
        dialogs.macro_dialog(self.root, self.macros,
                             lambda: setattr(self, "recording", None) or None,
                             self._macro_play)

    def _macro_play(self, name):
        if name in self.macros:
            self.load_doc(self.macros[name].play(self.text_to_doc()))
            self.update_status(f"Macro played: {name}")

    # ================= mac studio: gallery, prefs, versions, goals, extras =================
    def _tip(self, widget, text):
        tip = [None]

        def show(_e=None):
            x, y = widget.winfo_pointerx() + 12, widget.winfo_pointery() + 12
            t = tk.Toplevel(self.root); t.wm_overrideredirect(True); t.wm_geometry(f"+{x}+{y}")
            tk.Label(t, text=text, bg="#1C1C1E" if self.dark_mode else "#FFF8DC",
                     fg="white" if self.dark_mode else "black", padx=6, pady=3,
                     font=("Segoe UI", 9)).pack()
            tip[0] = t

        def hide(_e=None):
            if tip[0]:
                try: tip[0].destroy()
                except Exception: pass
                tip[0] = None

        widget.bind("<Enter>", show); widget.bind("<Leave>", hide)

    def show_gallery(self):
        win = tk.Toplevel(self.root); win.title(f"{BRAND['name']} — Document gallery")
        win.geometry("720x480"); win.transient(self.root)
        left = tk.Frame(win, bg=self.theme["panel2"], width=200)
        left.pack(side="left", fill="y"); left.pack_propagate(False)
        tk.Label(left, text=f"{BRAND['name']}\n{BRAND['tagline']}", bg=self.theme["panel2"],
                 fg=self.theme["fg"], font=("Segoe UI", 11, "bold"), pady=12).pack()
        tk.Label(left, text="Recent", bg=self.theme["panel2"], fg=self.theme["muted"]).pack()
        rlb = tk.Listbox(left, bg=self.theme["panel2"], fg=self.theme["fg"], relief="flat")
        rlb.pack(fill="both", expand=True, padx=6)
        for p in self.settings.get("recent", [])[:10]:
            rlb.insert("end", os.path.basename(p))
        right = tk.Frame(win, bg=self.theme["panel"]); right.pack(side="left", fill="both", expand=True)
        tk.Label(right, text="Create new", bg=self.theme["panel"], fg=self.theme["fg"],
                 font=("Segoe UI", 14, "bold")).pack(pady=10)
        grid = tk.Frame(right, bg=self.theme["panel"]); grid.pack()
        templates = [("Blank", "blank"), ("Cover", "cover"), ("Sample book", "sample"),
                     ("Letter", "letter"), ("Report", "report"), ("Resume", "resume")]
        def pick(kind):
            self.load_doc(WordDoc.template(kind)); self.current_path = None; win.destroy()

        for i, (label, kind) in enumerate(templates):
            b = tk.Button(grid, text=f"▦\n{label}", width=14, height=4, font=("Segoe UI", 10),
                          bg=self.theme["panel2"], fg=self.theme["fg"], relief="flat",
                          command=lambda k=kind: pick(k))
            b.grid(row=i // 3, column=i % 3, padx=8, pady=8)

        def open_recent(_e=None):
            if rlb.curselection():
                self.open_path(self.settings["recent"][rlb.curselection()[0]]); win.destroy()

        rlb.bind("<Double-Button-1>", open_recent)

    def preferences(self):
        win = tk.Toplevel(self.root); win.title("Preferences"); win.geometry("400x420")
        win.transient(self.root); win.grab_set()
        user = tk.StringVar(value=self.doc.author)
        lang = tk.StringVar(value=self.settings.get("lang", "en-US"))
        mins = tk.IntVar(value=self.settings.get("autosave_min", 5))
        auto = tk.BooleanVar(value=self.settings.get("autosave", True))
        gal = tk.BooleanVar(value=self.settings.get("gallery", True))
        restore = tk.BooleanVar(value=self.settings.get("restore", False))
        goal = tk.IntVar(value=self.goal)
        for label, var, w in [("Your name:", user, 30), ("Language:", lang, 12)]:
            ttk.Label(win, text=label).pack(pady=(8, 0)); ttk.Entry(win, textvariable=var, width=w).pack()
        ttk.Label(win, text="AutoRecover every (minutes):").pack(pady=(8, 0))
        ttk.Spinbox(win, from_=1, to=60, textvariable=mins, width=8).pack()
        ttk.Checkbutton(win, text="AutoRecover", variable=auto).pack()
        ttk.Checkbutton(win, text="Show gallery on launch", variable=gal).pack()
        ttk.Checkbutton(win, text="Restore last session", variable=restore).pack()
        ttk.Label(win, text="Writing goal (words, 0 = off):").pack(pady=(8, 0))
        ttk.Entry(win, textvariable=goal, width=10).pack()

        def ok():
            self.doc.author = user.get()
            self.settings.update(lang=lang.get(), autosave_min=mins.get(), autosave=auto.get(),
                                 gallery=gal.get(), restore=restore.get(), goal=goal.get(), user=user.get())
            self.goal = goal.get()
            self._save_settings(); win.destroy(); self.update_status("Preferences saved")

        ttk.Button(win, text="OK", command=ok).pack(pady=12)

    def _autosave_tick(self):
        try:
            if self.settings.get("autosave", True) and self._dirty:
                stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
                name = re.sub(r"[^\w\- ]+", "", self.doc.title or "document").strip()[:40] or "document"
                formats.save_docx(self.text_to_doc(), os.path.join(self.versions_dir, f"{name} {stamp}.docx"))
                self._dirty = False
        except Exception:
            pass
        self.root.after(max(1, int(self.settings.get("autosave_min", 5))) * 60000, self._autosave_tick)

    def versions_browser(self):
        win = tk.Toplevel(self.root); win.title("Versions / AutoRecover"); win.geometry("480x380")
        win.transient(self.root)
        lb = tk.Listbox(win); lb.pack(fill="both", expand=True, padx=8, pady=6)
        files = sorted((os.path.join(self.versions_dir, f) for f in os.listdir(self.versions_dir)
                                      if f.endswith(".docx")), reverse=True) if os.path.isdir(self.versions_dir) else []
        for f in files:
            lb.insert("end", f"{os.path.basename(f)}  ({os.path.getsize(f) // 1024} KB)")

        def restore():
            if lb.curselection():
                self.open_path(files[lb.curselection()[0]]); win.destroy()

        brow = ttk.Frame(win); brow.pack(pady=4)
        ttk.Button(brow, text="Restore", command=restore).pack(side="left", padx=4)
        ttk.Button(brow, text="Close", command=win.destroy).pack(side="left", padx=4)

    # ---- writer pro tools ----
    def set_goal(self):
        v = simpledialog.askinteger("Writing goal", "Target words (0 = off):",
                                    initialvalue=self.goal, minvalue=0, parent=self.root)
        if v is not None:
            self.goal = v; self.settings["goal"] = v; self._save_settings(); self.update_status()

    def editor_pane(self):
        extras.editor_dialog(self.root, self.text.get("1.0", "end-1c").replace("¶", ""))

    def autocorrect_mgr(self):
        extras.autocorrect_manager(self.root, AUTOCORRECT)
        self.settings["autocorrect"] = {k: v for k, v in AUTOCORRECT.items() if k not in _BUILTIN_AC}
        self._save_settings()

    def import_bibtex(self):
        doc = self.text_to_doc()
        if extras.import_bibtex_dialog(self.root, doc):
            self.doc.citations = doc.citations
            self.update_status(f"{len(doc.citations)} sources")

    def quick_parts(self):
        parts = self.settings.setdefault("autotext", {"Signature": "Best regards,\n" + self.doc.author,
                                                      "Address block": "Name\nStreet\nCity"})
        win = tk.Toplevel(self.root); win.title("Quick Parts"); win.geometry("360x300")
        win.transient(self.root)
        lb = tk.Listbox(win); lb.pack(fill="both", expand=True, padx=8, pady=6)
        for k in parts: lb.insert("end", k)

        def insert():
            if lb.curselection():
                self.insert_text(parts[lb.get(lb.curselection()[0])]); win.destroy()

        def save_sel():
            try: t = self.text.get("sel.first", "sel.last")
            except Exception:
                messagebox.showinfo("Quick Parts", "Select text first.", parent=win); return
            name = simpledialog.askstring("Quick Parts", "Save as:", parent=win)
            if name:
                parts[name] = t; self._save_settings(); win.destroy()

        brow = ttk.Frame(win); brow.pack(pady=4)
        ttk.Button(brow, text="Insert", command=insert).pack(side="left", padx=4)
        ttk.Button(brow, text="Save selection", command=save_sel).pack(side="left", padx=4)

    def insert_chart(self):
        spec = extras.chart_dialog(self.root)
        if not spec:
            return
        try:
            img = extras.build_chart_image(spec)
        except ImportError:
            messagebox.showerror("Chart", "Charts need matplotlib:\n\npip install matplotlib")
            return
        from PIL import ImageTk
        ph = ImageTk.PhotoImage(img); self._images.append(ph)
        self.text.image_create("insert", image=ph)
        self.text.insert("insert", f"\nFigure: {spec.get('title', 'Chart')}\n")

    def insert_wordart(self):
        spec = extras.wordart_dialog(self.root)
        if not spec:
            return
        from PIL import ImageTk
        ph = ImageTk.PhotoImage(extras.build_wordart_image(spec["text"], spec["style"]))
        self._images.append(ph)
        self.text.image_create("insert", image=ph)
        self.text.insert("insert", "\n")

    def insert_screenshot(self):
        try:
            self.root.iconify(); self.root.update()
            import time as _t; _t.sleep(0.4)
            img = extras.screenshot_grab()
            self.root.deiconify()
            img.thumbnail((640, 640))
            from PIL import ImageTk
            ph = ImageTk.PhotoImage(img); self._images.append(ph)
            self.text.image_create("insert", image=ph)
            self.text.insert("insert", " [Screenshot]\n")
        except Exception as ex:
            try: self.root.deiconify()
            except Exception: pass
            messagebox.showerror("Screenshot", str(ex))

    def toggle_typewriter(self):
        self.typewriter = not self.typewriter
        self.update_status(f"Typewriter mode: {'on' if self.typewriter else 'off'}")

    def toggle_fullscreen(self):
        try:
            cur = bool(self.root.attributes("-fullscreen"))
            self.root.attributes("-fullscreen", not cur)
        except Exception: pass

    def refresh_changes(self):
        try:
            self.chg_lb.delete(0, "end")
            for c in getattr(self.doc, "changes", []):
                if isinstance(c, dict):
                    self.chg_lb.insert("end", f"{c.get('author', '?')} {c.get('kind', '')}: {c.get('text', '')[:50]}")
        except Exception: pass

    # ================= nav / panes =================
    def refresh_nav(self):
        self.nav_lb.delete(0, "end"); self._nav_pos = []
        idx = "1.0"
        while True:
            nl = self.text.search("\n", idx, stopindex="end")
            if not nl: break
            for h in ("Title", "Heading 1", "Heading 2", "Heading 3", "Heading 4"):
                if h in self.text.tag_names(idx):
                    t = self.text.get(idx, nl).strip()[:60]
                    if t and "¶" not in t:
                        prefix = {"Title": "#", "Heading 1": "#", "Heading 2": "##",
                                  "Heading 3": "###", "Heading 4": "####"}[h]
                        self.nav_lb.insert("end", f"{prefix} {t}")
                        self._nav_pos.append(idx)
                    break
            idx = f"{nl}+1c"
            if self.text.compare(idx, ">=", "end-1c"): break
        self.pages_lb.delete(0, "end"); self._page_pos = []
        total = max(1, int(self.text.index("end-1c").split(".")[0]) // 45 + 1)
        for p in range(1, total + 1):
            self.pages_lb.insert("end", f"Page {p}")
            self._page_pos.append(f"{(p - 1) * 45 + 1}.0")
        self._sync_gutter()

    def _sync_gutter(self):
        n = int(self.text.index("end-1c").split(".")[0])
        self.gutter.configure(state="normal"); self.gutter.delete("1.0", "end")
        if self.doc.setup.line_numbers:
            if not self.gutter.winfo_ismapped():
                self.gutter.pack(side="left", fill="y")
            self.gutter.insert("1.0", "\n".join(str(i) for i in range(1, n + 1)))
        else:
            self.gutter.pack_forget()
        self.gutter.configure(state="disabled")

    def nav_goto(self, _e=None):
        try:
            i = self.nav_lb.curselection()[0]
            self.text.see(self._nav_pos[i]); self.text.mark_set("insert", self._nav_pos[i])
        except Exception: pass

    def page_goto(self, _e=None):
        try:
            i = self.pages_lb.curselection()[0]
            self.text.see(self._page_pos[i]); self.text.mark_set("insert", self._page_pos[i])
        except Exception: pass

    def nav_search(self):
        w = self.search_var.get()
        self.search_lb.delete(0, "end"); self._search_pos = []
        if not w: return
        idx, n = "1.0", 0
        while True:
            idx = self.text.search(w, idx, nocase=True, stopindex="end")
            if not idx or n > 200: break
            line = idx.split(".")[0]
            self.search_lb.insert("end", self.text.get(f"{line}.0", f"{line}.end")[:60])
            self._search_pos.append(idx); idx = f"{idx}+{len(w)}c"; n += 1

    def search_goto(self, _e=None):
        try:
            i = self.search_lb.curselection()[0]
            self.text.see(self._search_pos[i]); self.text.mark_set("insert", self._search_pos[i])
        except Exception: pass

    # ================= events =================
    def _bind_keys(self):
        r = self.root
        r.bind("<Control-s>", lambda e: (self.file_save(), "break")[1])
        r.bind("<Control-o>", lambda e: self.file_open())
        r.bind("<Control-n>", lambda e: (self.file_new(), "break")[1])
        r.bind("<Control-p>", lambda e: (self.print_dialog(), "break")[1])
        r.bind("<Control-f>", lambda e: self.find_dialog())
        r.bind("<Control-h>", lambda e: self.replace_dialog())
        r.bind("<Control-g>", lambda e: self.goto())
        r.bind("<F7>", lambda e: self.spell_dialog())
        r.bind("<F11>", lambda e: self.toggle_fullscreen())
        r.bind("<Control-comma>", lambda e: self.preferences())
        r.bind("<F4>", lambda e: self.repeat())
        r.bind("<Control-b>", lambda e: (self.toggle_tag("bold"), "break")[1])
        r.bind("<Control-i>", lambda e: (self.toggle_tag("italic"), "break")[1])
        r.bind("<Control-u>", lambda e: (self.toggle_tag("underline"), "break")[1])
        r.bind("<Control-e>", lambda e: (self.set_align("center"), "break")[1])
        r.bind("<Control-l>", lambda e: (self.set_align("left"), "break")[1])
        r.bind("<Control-r>", lambda e: (self.set_align("right"), "break")[1])
        r.bind("<Control-j>", lambda e: (self.set_align("justify"), "break")[1])
        r.bind("<Control-d>", lambda e: (self.apply_font(), "break")[1])
        r.bind("<Control-m>", lambda e: (self.indent(1), "break")[1])
        r.bind("<Control-Enter>", lambda e: (self.insert_page_break(), "break")[1])
        r.bind("<Escape>", lambda e: (self.hide_backstage(), self.ink.place_forget()))
        self.text.bind("<Control-Shift-V>", lambda e: self.paste_format())
        self.text.bind("<Double-Button-1>", self._dbl_click)

    def _dbl_click(self, _e=None):
        try:
            idx = self.text.index(f"@{self.text.winfo_pointerx() - self.text.winfo_rootx()},"
                                  f"@{self.text.winfo_pointery() - self.text.winfo_rooty()}")
            if "table-ph" in self.text.tag_names(idx):
                self.edit_table(); return "break"
        except Exception: pass
        return None

    def on_modified(self, _e=None):
        if self.text.edit_modified():
            self.text.edit_modified(False)
            self.refresh_nav(); self.refresh_comments(); self.refresh_changes(); self.update_status()

    def on_keyrelease(self, e=None):
        if self.smart.get() and e and e.char in (" ", "\r", "."):
            try:
                end = self.text.index("insert-1c")
                for k, v in AUTOCORRECT.items():
                    s = f"{end}-{len(k)}c"
                    if self.text.get(s, end) == k:
                        self.text.delete(s, end); self.text.insert(s, v)
                        break
            except Exception: pass
        if e and e.keysym == "Return":
            try:
                prev = int(self.text.index("insert").split(".")[0]) - 1
                t = self.text.get(f"{prev}.0", f"{prev}.end")
                if t.startswith("• "):
                    if t.strip() in ("•",):
                        self.text.delete(f"{prev}.0", f"{prev}.end")
                    else:
                        self.text.insert("insert", "• ")
                        self.text.tag_add("List Bullet", "insert linestart", "insert lineend")
                elif re.match(r"^1\.\s", t):
                    self.text.insert("insert", "1. ")
                    self.text.tag_add("List Number", "insert linestart", "insert lineend")
                m = re.match(r"^=lorem\((\d+)?\)$", t.strip(), re.I)
                if m:
                    self.text.delete(f"{prev}.0", f"{prev}.end")
                    self.text.insert(f"{prev}.0", extras.lorem_text(int(m.group(1) or 3)))
            except Exception: pass
        self._dirty = True
        if self.typewriter:
            try:
                self.text.see("insert")
                bbox = self.text.dlineinfo("insert")
                if bbox:
                    first, last = self.text.yview()
                    caret_abs = first + (bbox[1] / max(1, self.text.winfo_height())) * (last - first)
                    self.text.yview_moveto(max(0.0, caret_abs - (last - first) / 2))
            except Exception: pass
        self.update_status()

    def update_status(self, msg: str | None = None):
        txt = self.text.get("1.0", "end-1c").replace("¶", "")
        words = len(re.findall(r"\S+", txt))
        pages = max(1, round(words / 400 + 0.5))
        cur = max(1, int(self.text.index("insert").split(".")[0]) // 45 + 1)
        self.page_lbl.configure(text=f"Page {min(cur, pages)} of {pages}")
        self.word_lbl.configure(text=f"{msg + '  •  ' if msg else ''}{words} words")
        self.track_lbl.configure(text=f"Track: {'ON' if self.track.get() else 'off'}")
        self.zoom_lbl.configure(text=f"{int(self.zoom * 100)}%")
        if self.goal > 0:
            pct = min(100, int(words / self.goal * 100))
            self.goal_lbl.configure(text=f"Goal: {pct}% of {self.goal}  {'█' * (pct // 10)}{'░' * (10 - pct // 10)}")
        else:
            self.goal_lbl.configure(text="")

    def about(self):
        messagebox.showinfo(f"About {BRAND['name']}",
            f"{BRAND['name']} {BRAND['version']}\n{BRAND['tagline']}\n\n"
            "Full Word-style suite: backstage, ribbon, ruler, navigation, tables,\n"
            "headers & footers, TOC, footnotes, citations, mail merge, review,\n"
            "macros, and DOCX/PDF/HTML/MD export.\n\nCopyright (c) 2026 salim-slimani. MIT License.")


def main(path=None):
    root = tk.Tk()
    WordSoftApp(root, path=path)
    root.mainloop()


if __name__ == "__main__":
    import sys
    main(sys.argv[1] if len(sys.argv) > 1 else None)
