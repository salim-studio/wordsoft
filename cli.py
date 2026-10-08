# -*- coding: utf-8 -*-
"""WordSoft CLI — convert, inspect, merge and automate documents headlessly.

Usage:
  python cli.py convert in.docx out.pdf
  python cli.py text in.docx
  python cli.py stats in.docx
  python cli.py inspect in.docx
  python cli.py toc in.docx
  python cli.py new --template sample --save out.docx
  python cli.py merge template.docx list.csv --save merged.docx
  python cli.py batch in.docx --cmd insert.toc --cmd "design.watermark text=DRAFT" --save out.docx
  python cli.py commands

Copyright (c) 2026 salim-slimani. MIT License.
"""
import sys, json, argparse
from engine import WordDoc
import formats
from commands import COMMANDS, list_commands, batch


def cmd_convert(a):
    doc = formats.load(a.src)
    formats.save(doc, a.dst)
    print(f"OK: {a.src} -> {a.dst}")

def cmd_text(a):
    print(formats.load(a.src).full_text())

def cmd_stats(a):
    print(json.dumps(formats.load(a.src).stats(), ensure_ascii=False, indent=2))

def cmd_inspect(a):
    doc = formats.load(a.src)
    out = {"title": doc.title, "stats": doc.stats(),
           "blocks": [{"type": type(b).__name__,
                       "style": getattr(b, "style", ""),
                       "text": (b.text()[:120] if hasattr(b, "text") else str(getattr(b, "rows", [[""]])[0][:2]))}
                      for b in doc.blocks]}
    print(json.dumps(out, ensure_ascii=False, indent=2))

def cmd_toc(a):
    doc = formats.load(a.src)
    for style, text, page in doc.build_toc():
        print(f"{style:12} p.{page:3}  {text}")

def cmd_new(a):
    doc = WordDoc.template(a.template)
    formats.save(doc, a.save)
    print(f"OK: new({a.template}) -> {a.save}")

def cmd_merge(a):
    doc = formats.load(a.template)
    fields, records = WordDoc.load_csv(a.csv)
    merged = doc.merge_all(records)
    formats.save(merged, a.save)
    print(f"OK: merged {len(records)} records ({', '.join(fields)}) -> {a.save}")

def _parse_cmd(spec: str) -> dict:
    parts = spec.split()
    step = {"cmd": parts[0]}
    for p in parts[1:]:
        if "=" in p:
            k, v = p.split("=", 1)
            step[k] = v
    return step

def cmd_batch(a):
    doc = formats.load(a.src)
    batch(doc, [_parse_cmd(c) for c in a.cmd])
    formats.save(doc, a.save)
    print(f"OK: batch({len(a.cmd)} commands) -> {a.save}")

def cmd_commands(a):
    for c in list_commands():
        print(f"{c['id']:22} [{c['tab']}]  {c['title']}")


def main(argv=None):
    ap = argparse.ArgumentParser(prog="wordsoft-cli", description="WordSoft CLI — headless document automation")
    sub = ap.add_subparsers(dest="cmd", required=True)
    c = sub.add_parser("convert", help="convert between docx/pdf/html/md/txt"); c.add_argument("src"); c.add_argument("dst"); c.set_defaults(f=cmd_convert)
    t = sub.add_parser("text", help="print plain text"); t.add_argument("src"); t.set_defaults(f=cmd_text)
    s = sub.add_parser("stats", help="print word/char/paragraph stats"); s.add_argument("src"); s.set_defaults(f=cmd_stats)
    ins = sub.add_parser("inspect", help="print document structure as JSON"); ins.add_argument("src"); ins.set_defaults(f=cmd_inspect)
    tc = sub.add_parser("toc", help="print table of contents"); tc.add_argument("src"); tc.set_defaults(f=cmd_toc)
    n = sub.add_parser("new", help="create a document from a template"); n.add_argument("--template", default="sample"); n.add_argument("--save", required=True); n.set_defaults(f=cmd_new)
    m = sub.add_parser("merge", help="mail merge a template with a CSV list"); m.add_argument("template"); m.add_argument("csv"); m.add_argument("--save", required=True); m.set_defaults(f=cmd_merge)
    b = sub.add_parser("batch", help="run command ids headlessly"); b.add_argument("src"); b.add_argument("--cmd", action="append", required=True); b.add_argument("--save", required=True); b.set_defaults(f=cmd_batch)
    lc = sub.add_parser("commands", help="list all command ids"); lc.set_defaults(f=cmd_commands)
    a = ap.parse_args(argv)
    a.f(a)

if __name__ == "__main__":
    main()
