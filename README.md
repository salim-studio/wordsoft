<p align="center">
  <img src="assets/logo.svg" width="120" alt="WordSoft logo">
</p>

<h1 align="center">WordSoft</h1>
<p align="center"><b>Write beautifully.</b> A full Word-style word processor in Python.</p>

<p align="center">
  <img src="https://img.shields.io/badge/python-3.10%2B-2563EB?style=flat-square" alt="Python">
  <img src="https://img.shields.io/badge/license-MIT-2563EB?style=flat-square" alt="MIT">
  <img src="https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-2563EB?style=flat-square" alt="Platform">
</p>

## Why WordSoft

- **Familiar.** Quick Access toolbar, File backstage, ribbon tabs, live style gallery, ruler, navigation pane, status bar with zoom slider — the Word workflow you already know.
- **Your files.** Opens and saves `.docx` (styles, tables, headers/footers, page setup, TOC fields, footnotes), exports real PDF (selectable text, headers, watermark), HTML, Markdown and TXT.
- **Fast.** Zero build step — `python app.py` and you are writing. Offline by default.
- **Everywhere.** One Python codebase for Windows, macOS and Linux, plus a browser version (Flask).
- **Automatable.** Every action is a command with an id (`python cli.py commands`): the same ids drive the ribbon, shortcuts, headless batch runs and recorded macros.
- **Private.** Spelling, thesaurus and everything else work offline.

## Quick start (Windows)

```bat
cd Wordsoft
pip install -r requirements.txt
python app.py
```

Or double-click `run.bat`. Other entry points:

```bat
python -m Wordsoft        :: same desktop app
python cli.py --help      :: headless CLI
python api/index.py       :: web version at http://127.0.0.1:5000
```

## Features

| Area | Highlights |
|---|---|
| **File backstage** | New (6 templates), Open + recent, Info/properties, Save As, Export, Print + preview |
| **Home** | Cut/copy/paste, format painter, fonts, grow/shrink, B/I/U/strike/super/sub, change case, colors, highlight, clear, bullets, numbering, indents, alignment, line spacing, shading, borders, sort A→Z, style gallery + manager + custom styles, find/replace/regex, go to |
| **Insert** | Cover page, blank page, breaks, tables (insert + cell editor), pictures, shapes, text boxes, headers/footers, live page numbers, date/time, fields, drop caps, symbols, equations, comments |
| **Draw** | Pen/marker/eraser, ink overlay, keep drawing as picture |
| **Design** | 6 themes, theme fonts, watermark, page color, page borders |
| **Layout** | Margins (presets + ruler drag + custom), orientation, paper size, columns, breaks, paragraph spacing, line numbers gutter, hyphenation |
| **References** | Auto TOC (insert + update, real Word field on export), footnotes/endnotes, citations (APA/MLA/Chicago/IEEE) + bibliography, captions, cross-references, index |
| **Mailings** | CSV merge: load list, insert `<<fields>>`, preview, finish to document; envelopes + labels |
| **Review** | F7 spelling with suggestions + custom dictionary, thesaurus, detailed word count, language, comments pane (add/resolve/delete), track changes, accept all, compare documents, restrict editing (password), document inspector, accessibility checker |
| **View** | Print/Web/Draft/Read/Focus, ruler, ¶ marks, nav pane, zoom presets + slider, new window, macros (record/play) |
| **Status bar** | Live page/words, language, tracking flag, view buttons, zoom slider |

Shortcuts: `Ctrl+N/O/S/P/F/H/G`, `Ctrl+B/I/U/E/L/R/J/D/M`, `Ctrl+Enter` page break, `F4` repeat, `F7` spelling, `Ctrl+Shift+V` paste format, `Esc` close overlay.

## CLI

```bat
python cli.py convert report.docx report.pdf
python cli.py text report.docx
python cli.py stats report.docx
python cli.py toc report.docx
python cli.py new --template sample --save out.docx
python cli.py merge template.docx list.csv --save merged.docx
python cli.py batch in.docx --cmd insert.toc --cmd "design.watermark text=DRAFT" --save out.docx
python cli.py commands
```

## Web version

```
cd Wordsoft
python api/index.py
```

Open http://127.0.0.1:5000 — formatting, headings, lists, tables, DOCX upload, DOCX/PDF/HTML export.

## Project structure

```
Wordsoft/
  assets/      # logo.svg + icon.png — the WordSoft visual identity
  engine.py    # document model: blocks, setup, theme, notes, comments, merge
  commands.py  # every action as command id + macros (ribbon/CLI/macro runner)
  formats.py   # open/save: docx (headers, setup, TOC, notes), txt, md, html, pdf
  dialogs.py   # table editor, styles, setup, merge, citations, inspector...
  app.py       # desktop suite: backstage, ribbon, ruler, panes, views
  cli.py       # headless automation
  api/index.py # web editor (Flask)
```

## License and credits

Copyright (c) 2026 salim-slimani. MIT License — see [LICENSE](LICENSE).
