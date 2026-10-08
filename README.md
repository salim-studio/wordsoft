<p align="center">
  <img src="assets/logo.svg" width="120" alt="WordSoft logo">
</p>

<h1 align="center">WordSoft</h1>
<p align="center"><b>Write beautifully.</b> A fast, friendly word processor in Python.</p>

<p align="center">
  <img src="https://img.shields.io/badge/python-3.10%2B-2563EB?style=flat-square" alt="Python">
  <img src="https://img.shields.io/badge/license-MIT-2563EB?style=flat-square" alt="MIT">
  <img src="https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-2563EB?style=flat-square" alt="Platform">
</p>

## Why WordSoft

- **Familiar.** Ribbon tabs (Home / Insert / Layout / Review / View), live style gallery, keyboard shortcuts you already know: `Ctrl+B`, `Ctrl+S`, `F7`.
- **Your files.** Opens and saves `.docx` that works in Word; exports real PDF (selectable text), HTML, Markdown and TXT.
- **Fast.** Zero build step — `python app.py` and you are writing. Instant startup, offline by default.
- **Everywhere.** One Python codebase for Windows, macOS and Linux, plus a browser version (Flask).
- **Automatable.** The CLI converts, inspects and generates documents headlessly — scripts and agents welcome.
- **Private.** Spelling and everything else work offline.

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

- **Writing:** undo/redo, list continuation on Enter, smart autocorrect, format painter, find & replace
- **Formatting:** fonts & sizes, bold/italic/underline/strike, text color & highlight, alignment, RTL support, indents
- **Styles:** Normal, Title, Heading 1/2/3, Quote, bullets, numbering + headings navigation pane
- **Tables:** insert n×m, saved into DOCX/PDF/HTML
- **Pictures:** insert with automatic thumbnailing
- **Review:** word count, F7 spelling check, track-changes flag, comments
- **Files:** DOCX round-trip with Word + PDF (real text) / HTML / MD / TXT export
- **Templates:** Blank, Open Studio Handbook (sample), Letter, Report, Resume
- **View:** zoom, dark/light mode, focus mode, navigation pane

## CLI

```bat
python cli.py convert report.docx report.pdf
python cli.py text report.docx
python cli.py stats report.docx
python cli.py inspect report.docx
python cli.py new --template sample --save out.docx
```

## Web version

```
cd Wordsoft
python api/index.py
```

Open http://127.0.0.1:5000 — a full editor: formatting, headings, lists, tables, DOCX upload, DOCX/PDF/HTML export.

## Project structure

```
Wordsoft/
  assets/     # logo.svg + icon.png — the WordSoft visual identity
  engine.py   # document model: paragraphs + runs + tables, samples & templates
  formats.py  # open/save: docx, txt, md, html, pdf
  app.py      # desktop UI in the Word style (Tkinter)
  cli.py      # headless automation (convert / text / stats / inspect / new)
  api/index.py# web editor (Flask)
```

## License and credits

Copyright (c) 2026 salim-slimani. MIT License — see [LICENSE](LICENSE).

Bundled fonts keep their own licenses. The sample documents are original text written for WordSoft.
