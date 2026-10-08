# -*- coding: utf-8 -*-
"""WordSoft CLI — convert, inspect and automate documents headlessly.

Usage:
  python cli.py convert in.docx out.pdf
  python cli.py text in.docx
  python cli.py stats in.docx
  python cli.py inspect in.docx
  python cli.py new --template sample --save out.docx

Copyright (c) 2026 salim-slimani. MIT License.
"""
import sys, json, argparse
from engine import WordDoc
import formats


def cmd_convert(a):
    doc = formats.load(a.src)
    formats.save(doc, a.dst)
    print(f"OK: {a.src} -> {a.dst}")

def cmd_text(a):
    print(formats.load(a.src).full_text())

def cmd_stats(a):
    d = formats.load(a.src).stats()
    print(json.dumps(d, ensure_ascii=False, indent=2))

def cmd_inspect(a):
    doc = formats.load(a.src)
    out = {"title": doc.title, "stats": doc.stats(),
           "blocks": [{"type": type(b).__name__,
                       "style": getattr(b, "style", ""),
                       "text": (b.text()[:120] if hasattr(b, "text") else str(b.rows[:2]))} for b in doc.blocks]}
    print(json.dumps(out, ensure_ascii=False, indent=2))

def cmd_new(a):
    doc = WordDoc.template(a.template)
    formats.save(doc, a.save)
    print(f"OK: new({a.template}) -> {a.save}")


def main(argv=None):
    ap = argparse.ArgumentParser(prog="wordsoft-cli", description="WordSoft CLI — headless document automation")
    sub = ap.add_subparsers(dest="cmd", required=True)
    c = sub.add_parser("convert", help="convert between docx/pdf/html/md/txt"); c.add_argument("src"); c.add_argument("dst"); c.set_defaults(f=cmd_convert)
    t = sub.add_parser("text", help="print plain text"); t.add_argument("src"); t.set_defaults(f=cmd_text)
    s = sub.add_parser("stats", help="print word/char/paragraph stats"); s.add_argument("src"); s.set_defaults(f=cmd_stats)
    ins = sub.add_parser("inspect", help="print document structure as JSON"); ins.add_argument("src"); ins.set_defaults(f=cmd_inspect)
    n = sub.add_parser("new", help="create a document from a template"); n.add_argument("--template", default="sample"); n.add_argument("--save", required=True); n.set_defaults(f=cmd_new)
    a = ap.parse_args(argv)
    a.f(a)

if __name__ == "__main__":
    main()
