# -*- coding: utf-8 -*-
"""WordSoft Web — a full document editor in the browser (Flask) with a REST API.

Copyright (c) 2026 salim-slimani. MIT License.
"""
import os, sys, io
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from flask import Flask, request, send_file, jsonify, render_template_string
from engine import WordDoc
import formats

app = Flask(__name__)

PAGE = r"""<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>WordSoft Web — Write beautifully.</title>
<style>
:root{--blue:#2563EB}
body{background:#10141b;color:#e8eaf0;font-family:'Segoe UI',Tahoma,Arial;margin:0}
header{background:#1f2530;padding:10px 16px;display:flex;gap:8px;align-items:center;flex-wrap:wrap;border-bottom:1px solid #39435a}
header .logo{background:var(--blue);color:#fff;font-weight:bold;border-radius:8px;padding:4px 10px;font-size:18px}
header .tag{color:#8b94a7;font-style:italic}
button,select{background:#2a3242;color:#e8eaf0;border:1px solid #39435a;border-radius:6px;padding:6px 10px;cursor:pointer}
button:hover{background:var(--blue)}
main{display:flex;gap:0;min-height:calc(100vh - 110px)}
nav{width:220px;background:#1f2530;padding:10px;border-left:1px solid #39435a}
#ed{flex:1;background:#fff;color:#111;margin:14px;border-radius:8px;padding:34px;min-height:60vh;outline:none;line-height:1.9;max-width:900px}
#ed h1{font-size:28px}#ed h2{font-size:20px;color:var(--blue)}#ed blockquote{border-left:4px solid var(--blue);background:#f3f6ff;padding:8px 14px;margin:8px 0}
footer{background:#1f2530;padding:8px 16px;color:#8b94a7;font-size:13px;display:flex;gap:14px}
table{border-collapse:collapse;width:100%}td,th{border:1px solid #999;padding:6px}th{background:#e8eefc}
</style></head><body>
<header><span class="logo">W</span><b>WordSoft</b><span class="tag">Write beautifully.</span>
<button onclick="cmd('bold')">B</button><button onclick="cmd('italic')">I</button><button onclick="cmd('underline')">U</button>
<select onchange="cmd('formatBlock',this.value)"><option value="p">Normal</option><option value="h1">Heading 1</option><option value="h2">Heading 2</option><option value="blockquote">Quote</option></select>
<button onclick="cmd('insertUnorderedList')">Bullets</button><button onclick="cmd('insertOrderedList')">Numbering</button>
<button onclick="cmd('justifyCenter')">Center</button><button onclick="cmd('justifyLeft')">Left</button>
<button onclick="insTable()">Table</button>
<button onclick="newDoc('sample')">Sample</button><button onclick="newDoc('letter')">Letter</button>
<button onclick="dl('docx')">DOCX</button><button onclick="dl('pdf')">PDF</button><button onclick="dl('html')">HTML</button>
<label>Open DOCX <input type=file id=up accept=.docx hidden onchange="upload(this)"></label>
</header>
<main><div style="flex:1;display:flex;justify-content:center"><div id="ed" contenteditable="true"><h1>The Open Studio Handbook — WordSoft</h1><p>Welcome to <b>WordSoft</b> on the web — select text, format it from the toolbar, then export DOCX or PDF.</p><h2>Styles</h2><ul><li>Headings, lists and quotes</li><li>Real tables</li><li>Opens in Word</li></ul></div></div>
<nav><b>Navigation</b><div id="nav"></div><hr><b>Stats</b><div id="st"></div><hr><small>Copyright (c) 2026 salim-slimani</small></nav></main>
<footer><span id="words">Words: 0</span><span>WordSoft Web 2.0</span></footer>
<script>
function cmd(c,v){document.execCommand(c,false,v||null);document.getElementById('ed').focus();upd()}
function insTable(){let r=prompt('Rows:',3),c=prompt('Columns:',3);if(!r||!c)return;let h='<table>';for(let i=0;i<r;i++){h+='<tr>';for(let j=0;j<c;j++)h+=`<td>${i==0?'Header '+(j+1):'...'}</td>`;h+='</tr>'}document.execCommand('insertHTML',false,h+'</table>');upd()}
function upd(){let t=document.getElementById('ed').innerText;let w=(t.match(/\S+/g)||[]).length;document.getElementById('words').textContent='Words: '+w+' | Pages ~ '+Math.max(1,Math.round(w/400+.5));
let nav=document.getElementById('nav');nav.innerHTML='';document.querySelectorAll('#ed h1,#ed h2').forEach(h=>{let a=document.createElement('div');a.textContent='• '+h.textContent.slice(0,40);a.style.cursor='pointer';a.onclick=()=>h.scrollIntoView();nav.appendChild(a)})}
async function dl(f){let r=await fetch('/api/export/'+f,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({html:document.getElementById('ed').innerHTML})});let b=await r.blob();let a=document.createElement('a');a.href=URL.createObjectURL(b);a.download='wordsoft.'+f;a.click()}
async function newDoc(k){let r=await fetch('/api/template/'+k);let j=await r.json();document.getElementById('ed').innerHTML=j.html;upd()}
async function upload(inp){let fd=new FormData();fd.append('file',inp.files[0]);let r=await fetch('/api/upload',{method:'POST',body:fd});let j=await r.json();document.getElementById('ed').innerHTML=j.html;upd()}
document.getElementById('ed').addEventListener('input',upd);upd();setInterval(upd,2000);
</script></body></html>"""

def html_to_doc(h: str, title="Web document") -> WordDoc:
    import re, html as H
    doc = WordDoc(title=title)
    # Tables first
    tables = re.findall(r"<table.*?</table>", h, re.S | re.I)
    text = re.sub(r"<table.*?</table>", "\n[[TABLE]]\n", h, flags=re.S | re.I)
    parts = re.split(r"<h1>|<h2>|<blockquote>|<li>|<p>|</h1>|</h2>|</blockquote>|</li>|</p>|<br.*?>|\n", text)
    ti = 0
    for raw in parts:
        t = re.sub(r"<.*?>", "", raw).strip()
        t = H.unescape(t)
        if not t: continue
        if t == "[[TABLE]]" and ti < len(tables):
            cells = re.findall(r"<t[dh].*?>(.*?)</t[dh]>", tables[ti], re.S | re.I)
            import html as HH
            cells = [HH.unescape(re.sub(r"<.*?>", "", c).strip()) for c in cells]
            ncols = max(1, len(re.findall(r"<tr", tables[ti], re.I)) and len(cells) // max(1, len(re.findall(r"<tr", tables[ti], re.I))))
            rows = [cells[i:i + ncols] for i in range(0, len(cells), ncols)] or [cells]
            doc.blocks.append(__import__("engine").Table(rows=rows)); ti += 1
            continue
        if "<li" in raw:
            doc.blocks.append(__import__("engine").Paragraph([__import__("engine").Run(t)], style="List Bullet"))
        else:
            doc.blocks.append(__import__("engine").Paragraph([__import__("engine").Run(t)]))
    # First block becomes the title
    if doc.blocks and isinstance(doc.blocks[0], __import__("engine").Paragraph):
        doc.blocks[0].style = "Title"
    return doc or WordDoc.blank()

@app.route("/")
def index(): return render_template_string(PAGE)

@app.route("/api/template/<kind>")
def tpl(kind):
    import html as H
    doc = WordDoc.template(kind if kind in ("sample", "letter", "report", "resume", "blank") else "sample")
    h = ""
    for b in doc.blocks:
        if isinstance(b, __import__("engine").Table):
            h += "<table>" + "".join("<tr>" + "".join(f"<td>{H.escape(c)}</td>" for c in row) + "</tr>" for row in b.rows) + "</table>"
        else:
            t = H.escape(b.text())
            h += {"Title": f"<h1>{t}</h1>", "Heading 1": f"<h2>{t}</h2>"}.get(b.style, f"<p>{t}</p>")
    return jsonify({"html": h})

@app.route("/api/upload", methods=["POST"])
def upload():
    import html as H
    f = request.files["file"]
    p = "/tmp/_ws_up.docx"; f.save(p)
    doc = formats.load(p)
    h = ""
    for b in doc.blocks:
        if isinstance(b, __import__("engine").Table):
            h += "<table>" + "".join("<tr>" + "".join(f"<td>{H.escape(c)}</td>" for c in row) + "</tr>" for row in b.rows) + "</table>"
        else:
            h += f"<p>{H.escape(b.text())}</p>"
    return jsonify({"html": h})

@app.route("/api/export/<fmt>", methods=["POST"])
def export(fmt):
    doc = html_to_doc(request.json.get("html", ""), "WordSoft")
    import tempfile
    suf = {"docx": ".docx", "pdf": ".pdf", "html": ".html", "md": ".md", "txt": ".txt"}.get(fmt, ".docx")
    fd, path = tempfile.mkstemp(suffix=suf); os.close(fd)
    formats.save(doc, path)
    mime = {"docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "pdf": "application/pdf", "html": "text/html", "md": "text/markdown", "txt": "text/plain"}.get(fmt)
    return send_file(path, as_attachment=True, download_name=f"wordsoft{suf}", mimetype=mime)

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
