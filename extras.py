# -*- coding: utf-8 -*-
"""WordSoft extras — heavyweight showcase features.

Charts (matplotlib), WordArt (PIL), readability + Editor pane,
BibTeX import, lorem generator, screenshots, AutoCorrect manager.

Copyright (c) 2026 salim-slimani. MIT License.
"""
from __future__ import annotations
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from engine import Citation

LOREM = ("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor "
         "incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud "
         "exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure "
         "dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ")


def lorem_text(paras: int = 3) -> str:
    return "\n\n".join(LOREM * 2 for _ in range(max(1, paras)))


# ---------------- readability & Editor ----------------
def readability(text: str) -> dict:
    words = re.findall(r"[A-Za-z']+", text)
    sents = [s for s in re.split(r"[.!?]+", text) if s.strip()] or [text]
    syll = sum(max(1, len(re.findall(r"[aeiouy]+", w.lower())) - (1 if w.lower().endswith("e") else 0)) for w in words) or 1
    nw, ns = max(1, len(words)), max(1, len(sents))
    ease = 206.835 - 1.015 * (nw / ns) - 84.6 * (syll / nw)
    grade = 0.39 * (nw / ns) + 11.8 * (syll / nw) - 15.59
    return {"words": len(words), "sentences": len(sents), "syllables": syll,
            "flesch": round(ease, 1), "grade": round(max(0, grade), 1),
            "level": "Very easy" if ease >= 80 else "Easy" if ease >= 60 else
                     "Fair" if ease >= 50 else "Hard" if ease >= 30 else "Very hard"}


WEASEL = {"very", "really", "quite", "rather", "just", "actually", "basically", "literally",
          "stuff", "things", "a lot", "very unique"}


def editor_check(text: str) -> tuple[int, list[str]]:
    """Rule-based writing review. Returns (score 0-100, findings)."""
    sents = [s.strip() for s in re.split(r"(?<=[.!?])\s+", text) if s.strip()]
    findings = []
    long = sum(1 for s in sents if len(s.split()) > 25)
    if long: findings.append(f"{long} long sentence(s) — consider splitting over 25 words.")
    passive = sum(1 for s in sents if re.search(r"\b(am|is|are|was|were|be|been|being)\b\s+\w+ed\b", s, re.I))
    if passive: findings.append(f"{passive} possible passive construction(s).")
    adv = len(re.findall(r"\b\w+ly\b", text, re.I))
    if adv > max(3, len(sents) // 2): findings.append(f"{adv} adverbs — strong verbs beat adverbs.")
    weasel = sorted({w for w in WEASEL if re.search(r"\b" + re.escape(w) + r"\b", text, re.I)})
    if weasel: findings.append("Weasel words: " + ", ".join(weasel) + ".")
    words = [w.lower() for w in re.findall(r"[A-Za-z']{4,}", text)]
    from collections import Counter
    rep = [w for w, n in Counter(words).most_common(5) if n >= 5]
    if rep: findings.append("Very frequent words: " + ", ".join(rep) + ".")
    score = max(20, 100 - 8 * len(findings) - long - passive)
    if not findings:
        findings.append("Clean and readable. Nice writing! ✓")
    return score, findings


def editor_dialog(parent, text: str):
    r = readability(text)
    score, findings = editor_check(text)
    messagebox.showinfo("Editor",
        f"Writing score: {score}/100\nFlesch Reading Ease: {r['flesch']} ({r['level']})\n"
        f"Grade level: {r['grade']}  •  Sentences: {r['sentences']}\n\n" + "\n".join("• " + f for f in findings),
        parent=parent)


# ---------------- BibTeX ----------------
def parse_bibtex(path: str) -> list[Citation]:
    with open(path, encoding="utf-8", errors="replace") as f:
        src = f.read()
    out = []
    starts = [m for m in re.finditer(r"@\w+\s*\{\s*([^,]+),", src)]
    for i, m in enumerate(starts):
        key = m.group(1).strip()
        end = starts[i + 1].start() if i + 1 < len(starts) else len(src)
        body = src[m.end():end].rstrip()
        if body.endswith("}"):
            body = body[:-1]
        fields: dict[str, str] = {}
        for fm in re.finditer(r"(\w+)\s*=\s*[{|\"](.*?)[}|\"]\s*,?", body, re.S):
            fields[fm.group(1).lower()] = fm.group(2).strip()
        author = fields.get("author", "").replace(" and ", ", ")
        out.append(Citation(key=key, author=author, title=fields.get("title", ""),
                            year=fields.get("year", ""), publisher=fields.get("journal", fields.get("publisher", ""))))
    return out


def import_bibtex_dialog(parent, doc) -> int:
    p = filedialog.askopenfilename(filetypes=[("BibTeX", "*.bib")], parent=parent)
    if not p:
        return 0
    cites = parse_bibtex(p)
    doc.citations.extend(cites)
    messagebox.showinfo("BibTeX", f"Imported {len(cites)} source(s).", parent=parent)
    return len(cites)


# ---------------- charts ----------------
def chart_dialog(parent):
    win = tk.Toplevel(parent); win.title("Insert chart"); win.geometry("420x420")
    win.transient(parent); win.grab_set()
    kind = tk.StringVar(value="bar")
    ttk.Label(win, text="Type:").pack()
    ttk.Combobox(win, textvariable=kind, values=["bar", "line", "pie"], state="readonly", width=12).pack()
    ttk.Label(win, text="One label + value per line (Label, value):").pack(pady=(8, 0))
    txt = tk.Text(win, height=10); txt.pack(fill="both", expand=True, padx=10)
    txt.insert("1.0", "Q1, 40\nQ2, 65\nQ3, 52\nQ4, 80")
    ttk.Label(win, text="Title:").pack()
    title = tk.Entry(win, width=40); title.pack(); title.insert(0, "Sales 2026")
    out: dict = {}

    def ok():
        out.update(kind=kind.get(), title=title.get(), raw=txt.get("1.0", "end-1c"))
        win.destroy()

    ttk.Button(win, text="Insert chart", command=ok).pack(pady=8)
    win.wait_window()
    return out or None


def build_chart_image(spec: dict, width=640, height=400):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    labels, values = [], []
    for line in spec.get("raw", "").splitlines():
        if "," in line:
            lab, val = line.split(",", 1)
            try:
                labels.append(lab.strip() or "?"); values.append(float(val.strip()))
            except ValueError:
                continue
    if not values:
        labels, values = ["A", "B", "C"], [1, 2, 3]
    fig, ax = plt.subplots(figsize=(width / 100, height / 100), dpi=100)
    kind = spec.get("kind", "bar")
    if kind == "pie":
        ax.pie(values, labels=labels, autopct="%1.0f%%")
    elif kind == "line":
        ax.plot(labels, values, marker="o"); ax.grid(True, alpha=0.3)
    else:
        ax.bar(labels, values, color="#2563EB")
    ax.set_title(spec.get("title", "Chart"))
    fig.tight_layout()
    fig.canvas.draw()
    import numpy as np
    buf = np.asarray(fig.canvas.buffer_rgba())
    plt.close(fig)
    from PIL import Image
    return Image.fromarray(buf).convert("RGB")


# ---------------- WordArt ----------------
def wordart_dialog(parent):
    text = simpledialog.askstring("WordArt", "Text:", initialvalue="WordSoft", parent=parent)
    if not text:
        return None
    win = tk.Toplevel(parent); win.title("WordArt style"); win.geometry("300x200")
    win.transient(parent); win.grab_set()
    style = tk.StringVar(value="Blue fade")
    for s in ["Blue fade", "Sunset", "Forest", "Mono outline"]:
        ttk.Radiobutton(win, text=s, value=s, variable=style).pack(anchor="w", padx=20, pady=2)
    out: dict = {}

    def ok():
        out.update(text=text, style=style.get()); win.destroy()

    ttk.Button(win, text="Insert", command=ok).pack(pady=10)
    win.wait_window()
    return out or None


def build_wordart_image(text: str, style: str = "Blue fade", width=640, height=160):
    from PIL import Image, ImageDraw, ImageFont
    img = Image.new("RGBA", (width, height), (255, 255, 255, 0))
    d = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("arial.ttf", 84)
    except Exception:
        font = ImageFont.load_default()
    palettes = {"Blue fade": ((37, 99, 235), (147, 197, 253)),
                "Sunset": ((194, 65, 12), (253, 186, 116)),
                "Forest": ((4, 120, 87), (167, 243, 208)),
                "Mono outline": ((51, 65, 85), (148, 163, 184))}
    c1, c2 = palettes.get(style, palettes["Blue fade"])
    l, t, r, b = d.textbbox((0, 0), text, font=font)
    tw, th = r - l, b - t
    x, y = (width - tw) // 2, (height - th) // 2
    # gradient slab, then punch the text shape through a mask
    mask = Image.new("L", (width, height), 0)
    ImageDraw.Draw(mask).text((x, y), text, font=font, fill=255)
    from PIL import Image as _I
    solid = _I.new("RGBA", (width, height), (0, 0, 0, 0))
    sd = ImageDraw.Draw(solid)
    for i in range(th):
        f = i / max(1, th)
        col = tuple(int(c1[k] + (c2[k] - c1[k]) * f) for k in range(3)) + (255,)
        sd.line([x, y + i, x + tw, y + i], fill=col)
    solid.putalpha(mask)
    if style == "Mono outline":
        edge = _I.new("RGBA", (width, height), (0, 0, 0, 0))
        ImageDraw.Draw(edge).text((x, y), text, font=font, fill=(51, 65, 85, 255))
        return edge
    return solid


# ---------------- misc writer tools ----------------
def screenshot_grab():
    from PIL import ImageGrab
    return ImageGrab.grab()


def autocorrect_manager(parent, table: dict):
    win = tk.Toplevel(parent); win.title("AutoCorrect entries"); win.geometry("380x380")
    win.transient(parent); win.grab_set()
    lb = tk.Listbox(win); lb.pack(fill="both", expand=True, padx=8, pady=6)

    def reload():
        lb.delete(0, "end")
        for k, v in sorted(table.items()):
            lb.insert("end", f"{k}  →  {v}")

    bar = ttk.Frame(win); bar.pack(fill="x", padx=8, pady=4)

    def add():
        k = simpledialog.askstring("AutoCorrect", "Replace:", parent=win)
        v = simpledialog.askstring("AutoCorrect", "With:", parent=win) if k else None
        if k and v:
            table[k] = v; reload()

    def delete():
        if lb.curselection():
            table.pop(lb.get(lb.curselection()[0]).split("  →")[0].strip(), None); reload()

    ttk.Button(bar, text="Add", command=add).pack(side="left", padx=2)
    ttk.Button(bar, text="Delete", command=delete).pack(side="left", padx=2)
    reload()
