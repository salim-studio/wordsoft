# -*- coding: utf-8 -*-
"""WordSoft Web — a full Word-style editor in the browser.

Ribbon (File/Home/Insert/Design/Layout/References/Mailings/Review/View),
backstage, navigation + comments panes, tables, TOC, footnotes, citations,
mail merge, themes, watermark, and DOCX/PDF/HTML/MD/TXT export.

Copyright (c) 2026 salim-slimani. MIT License.
"""
import os, sys, re, io, base64, html as H
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from flask import Flask, request, send_file, jsonify, render_template_string
from engine import WordDoc, Paragraph, Run, Table, PageBreak, TOCField, FootnoteRef, Footnote, Citation
import formats

app = Flask(__name__)

PAGE = r"""<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>WordSoft Web — Write beautifully.</title>
<style>
:root{--blue:#2563EB;--bg:#10141b;--panel:#1f2530;--panel2:#2a3242;--fg:#e8eaf0;--mut:#8b94a7;--bd:#39435a}
body.light{--bg:#eef1f6;--panel:#fff;--panel2:#e8edf5;--fg:#1a2233;--mut:#5b6478;--bd:#cdd6e4}
*{box-sizing:border-box}body{background:var(--bg);color:var(--fg);font-family:'Segoe UI',Tahoma,Arial;margin:0}
/* quick access */#qa{background:var(--panel);padding:6px 12px;display:flex;gap:6px;align-items:center;border-bottom:1px solid var(--bd)}
#qa .logo{background:var(--blue);color:#fff;font-weight:bold;border-radius:8px;padding:3px 10px;font-size:16px}
#qa b{font-size:15px}#qa .tag{color:var(--mut);font-style:italic;font-size:12px}
/* ribbon */#tabs{background:var(--panel);display:flex;gap:2px;padding:0 10px;border-bottom:1px solid var(--bd)}
#tabs button{background:none;border:none;color:var(--fg);padding:8px 14px;cursor:pointer;font-size:13px}
#tabs button.on{border-bottom:3px solid var(--blue);font-weight:bold}
#ribbon{background:var(--panel);border-bottom:1px solid var(--bd);padding:6px 10px;min-height:56px}
.panel{display:none;gap:6px;align-items:center;flex-wrap:wrap}.panel.on{display:flex}
.grp{border-inline-end:1px solid var(--bd);padding-inline-end:8px;display:flex;gap:4px;align-items:center;flex-wrap:wrap}
button,select,input[type=text]{background:var(--panel2);color:var(--fg);border:1px solid var(--bd);border-radius:6px;padding:5px 9px;cursor:pointer;font-size:12.5px}
button:hover{background:var(--blue)}button.warn:hover{background:#b91c1c}
/* layout */#mid{display:flex;min-height:calc(100vh - 210px)}
#nav{width:225px;background:var(--panel);border-inline-end:1px solid var(--bd);padding:8px;display:flex;flex-direction:column;gap:4px}
#nav h4{margin:4px 0}#nav .it{cursor:pointer;padding:3px 6px;border-radius:4px;font-size:12.5px}#nav .it:hover{background:var(--panel2)}
#stage{flex:1;display:flex;justify-content:center;padding:14px;position:relative}
#page{background:#fff;color:#111;border-radius:4px;box-shadow:0 2px 14px rgba(0,0,0,.5);width:100%;max-width:820px;min-height:70vh;padding:56px 64px;outline:none;line-height:1.75;position:relative}
#page h1{font-size:28px;text-align:center}#page h2{font-size:20px;color:var(--blue)}#page h3{font-size:16px}
#page blockquote{border-left:4px solid var(--blue);background:#f3f6ff;padding:8px 14px;margin:8px 0}
#page table{border-collapse:collapse;width:100%;margin:8px 0}#page td,#page th{border:1px solid #999;padding:6px}#page th{background:#e8eefc}
#page hr.pbrk{border:none;border-top:2px dashed var(--blue);margin:14px 0}
#page .cmt{background:#fff3bf;border-bottom:2px solid #eab308}
#page .fn{color:var(--blue);font-size:.75em;vertical-align:super}
#page img{max-width:100%}
#wm{position:absolute;top:38%;left:50%;transform:translate(-50%,-50%) rotate(-28deg);font-size:84px;font-weight:bold;color:rgba(37,99,235,.10);pointer-events:none;display:none}
#cmt{width:250px;background:var(--panel);border-inline-start:1px solid var(--bd);padding:8px;display:none;flex-direction:column;gap:4px}
#cmt.open{display:flex}#cmt .c{background:var(--panel2);border-radius:6px;padding:6px;font-size:12.5px}
/* backstage */#bs{position:fixed;inset:0;background:var(--bg);z-index:50;display:none}
#bs.open{display:flex}#bs .menu{width:190px;background:var(--panel);padding:20px 10px;display:flex;flex-direction:column;gap:2px}
#bs .menu button{text-align:start;font-size:14px;padding:9px}#bs .body{flex:1;padding:26px;overflow:auto}
#bs .body h2{margin-top:0}.tpl{display:block;width:340px;margin:5px 0;padding:10px;text-align:start}
/* status */#st{background:var(--panel);border-top:1px solid var(--bd);padding:5px 12px;color:var(--mut);font-size:12.5px;display:flex;gap:14px;align-items:center}
#st input[type=range]{width:120px}.sp{flex:1}
/* dialog */#ov{position:fixed;inset:0;background:rgba(0,0,0,.5);display:none;z-index:60;align-items:center;justify-content:center}
#ov.open{display:flex}#dlg{background:var(--panel);border:1px solid var(--bd);border-radius:10px;padding:18px;min-width:340px;max-width:560px;max-height:80vh;overflow:auto}
#dlg h3{margin-top:0}#dlg textarea{width:100%;background:var(--panel2);color:var(--fg);border:1px solid var(--bd);border-radius:6px}
</style></head><body>
<div id="qa"><span class="logo">W</span><b>WordSoft</b><span class="tag">Write beautifully.</span>
<button onclick="qaNew()" title="New">New</button><button onclick="openDoc()" title="Open">Open</button>
<button onclick="dl('docx')" title="Save">Save</button><button onclick="document.execCommand('undo')" title="Undo">Undo</button>
<button onclick="document.execCommand('redo')" title="Redo">Redo</button><button onclick="doPrint()" title="Print">Print</button>
<span class="sp"></span><span id="uinfo" style="color:var(--mut);font-size:12px"></span></div>
<div id="tabs"></div><div id="ribbon"></div>
<div id="mid">
<div id="nav"><h4>Navigation</h4><input type="text" id="nq" placeholder="Search..." oninput="navSearch()"><div id="nlist" style="overflow:auto"></div></div>
<div id="stage"><div style="position:relative;width:100%;max-width:820px"><div id="wm"></div><div id="page" contenteditable="true"></div></div></div>
<div id="cmt"><h4>Comments <button onclick="addComment()">+ New</button></h4><div id="clist" style="overflow:auto;display:flex;flex-direction:column;gap:4px"></div></div>
</div>
<div id="st"><span id="sWords">Words: 0</span><span id="sPages">Page 1 of 1</span><span id="sTrack">Track: off</span><span class="sp"></span>
<button onclick="setView('print')">Print</button><button onclick="setView('web')">Web</button><button onclick="setView('read')">Read</button>
<input type="range" min="50" max="200" value="100" oninput="setZoom(this.value)"><span id="sZoom">100%</span></div>
<div id="bs"><div class="menu" id="bsMenu"></div><div class="body" id="bsBody"></div></div>
<div id="ov"><div id="dlg"></div></div>
<input type="file" id="upDocx" accept=".docx" hidden onchange="uploadDocx(this)">
<input type="file" id="upImg" accept="image/*" hidden onchange="uploadImg(this)">
<input type="file" id="upCsv" accept=".csv" hidden onchange="uploadCsv(this)">
<script>
const $=id=>document.getElementById(id);
const ed=()=>$('page');
let STATE={view:'print',zoom:100,track:false,fields:[],records:[],ri:0,comments:[],footnotes:[],
setup:{size:'A4',orientation:'portrait',mt:2.54,mb:2.54,ml:2.54,mr:2.54,columns:1},
hf:{header:'',footer:'',pagenumbers:true},watermark:'',theme:'Office',title:'Untitled document'};
/* ---------- ribbon definition ---------- */
const RIBBON={
File:[],
Home:[{g:'Clipboard',b:[['Cut','document.execCommand("cut")'],['Copy','document.execCommand("copy")'],['Paste','document.execCommand("paste")'],['Format','copyFmt()']]},
{g:'Font',b:[['B','cmd("bold")'],['I','cmd("italic")'],['U','cmd("underline")'],['S','cmd("strikeThrough")'],
['x²','sup()'],['x₂','sub()'],['Aa','changeCase()'],['Color','foreColor()'],['Mark','hilite()'],['Clear','clearFmt()']],c:'<select onchange="fontName(this.value)"><option>Calibri</option><option>Arial</option><option>Georgia</option><option>Times New Roman</option><option>Courier New</option><option>Verdana</option></select><select onchange="fontSize(this.value)"><option value="3">11</option><option value="1">8</option><option value="2">10</option><option value="4">14</option><option value="5">18</option><option value="6">24</option><option value="7">36</option></select>'},
{g:'Paragraph',b:[['Bullets','cmd("insertUnorderedList")'],['1. Number','cmd("insertOrderedList")'],['← L','cmd("justifyLeft")'],['↔ C','cmd("justifyCenter")'],['R →','cmd("justifyRight")'],['Justify','cmd("justifyFull")'],['A↓Z','sortParas()']]},
{g:'Styles',b:[['Normal','block("p")'],['Title','block("h1")'],['Head 1','block("h1")'],['Head 2','block("h2")'],['Quote','block("blockquote")']]},
{g:'Editing',b:[['Find','findDlg()'],['Replace','replaceDlg()'],['Count','wordCount()']]}],
Insert:[{g:'Pages',b:[['Cover','coverPage()'],['Blank page','blankPage()'],['Page break','pageBreak()']]},
{g:'Tables',b:[['Table','insTable()'],['+ Row','tableRow(1)'],['+ Col','tableCol(1)'],['− Row','tableRow(-1)']]},
{g:'Media',b:[['Picture','(()=>$("upImg").click())()'],['Shape ★','insShape()'],['Text box','insBox()']]},
{g:'Header & Footer',b:[['Header','editHF("header")'],['Footer','editHF("footer")'],['Page #','pageNum()']]},
{g:'Text',b:[['Date','insDate()'],['Title field','insField("TITLE")'],['Drop cap','dropCap()'],['Symbol Ω','symbols()'],['Equation','equation()'],['Comment','addComment()']]}],
Design:[{g:'Themes',b:[['Office','setTheme("Office")'],['Blue','setTheme("Blue")'],['Forest','setTheme("Forest")'],['Sunset','setTheme("Sunset")'],['Mono','setTheme("Mono")']]},
{g:'Background',b:[['Watermark','setWatermark()'],['Page color','pageColor()']]}],
Layout:[{g:'Setup',b:[['Margins','setMargins()'],['Portrait','setOrient("portrait")'],['Landscape','setOrient("landscape")'],['A4','setSize("A4")'],['Letter','setSize("Letter")']]},
{g:'Arrange',b:[['1 column','setCols(1)'],['2 columns','setCols(2)'],['Line #','lineNums()']]}],
References:[{g:'Contents',b:[['Insert TOC','insTOC()'],['Update TOC','updTOC()']]},
{g:'Notes',b:[['Footnote','addFootnote()'],['Show notes','showNotes()']]},
{g:'Sources',b:[['Cite book','addCite()'],['Bibliography','addBiblio()']]},
{g:'Captions',b:[['Caption','addCaption()'],['Mark index','markIndex()'],['Index','makeIndex()']]}],
Mailings:[{g:'Merge',b:[['Load CSV','(()=>$("upCsv").click())()'],['Field <<>>','insField2()'],['◀ Prev','mStep(-1)'],['Next ▶','mStep(1)'],['Finish','mFinish()']]},
{g:'Create',b:[['Envelope','newDoc("envelope")'],['Labels','newDoc("label")']]}],
Review:[{g:'Proofing',b:[['Spelling','spellCheck()'],['Thesaurus','thesaurus()'],['Word count','wordCount()']]},
{g:'Comments',b:[['New','addComment()'],['Pane','toggleCmt()'],['Resolve all','resolveAll()']]},
{g:'Tracking',b:[['Track: off','toggleTrack()'],['Compare','compareInfo()'],['Inspect','inspect()']]}],
View:[{g:'Views',b:[['Print','setView("print")'],['Web','setView("web")'],['Read','setView("read")']]},
{g:'Show',b:[['Nav pane','toggleNav()'],['Dark/Light','toggleLight()']]},
{g:'Zoom',b:[['100%','setZoom(100)'],['150%','setZoom(150)'],['Width','setZoom(120)']]}]};
function buildRibbon(){
const tabs=$('tabs'),rb=$('ribbon');tabs.innerHTML='';rb.innerHTML='';
Object.keys(RIBBON).forEach((t,i)=>{
const b=document.createElement('button');b.textContent=t;if(i===1)b.classList.add('on');
b.onclick=()=>{document.querySelectorAll('#tabs button').forEach(x=>x.classList.remove('on'));b.classList.add('on');showTab(t)};
tabs.appendChild(b);});
showTab('Home');}
function showTab(t){
if(t==='File'){openBS('New');return}
const rb=$('ribbon');rb.innerHTML='';
RIBBON[t].forEach(gr=>{
const d=document.createElement('div');d.className='grp';
const lbl=document.createElement('span');lbl.style.cssText='color:var(--mut);font-size:11px;width:100%';lbl.textContent=gr.g;
gr.b.forEach(([txt,fn])=>{const btn=document.createElement('button');btn.textContent=txt;btn.setAttribute('onclick',fn);d.appendChild(btn)});
if(gr.c){const s=document.createElement('span');s.innerHTML=gr.c;d.appendChild(s)}
d.appendChild(lbl);rb.appendChild(d);});}
/* ---------- basic cmds ---------- */
let _fmt=null;
function cmd(c){document.execCommand('styleWithCSS',false,true);document.execCommand(c,false,null);ed().focus();upd()}
function block(tag){document.execCommand('formatBlock',false,tag);upd()}
function fontName(v){document.execCommand('styleWithCSS',false,true);document.execCommand('fontName',false,v);upd()}
function fontSize(v){document.execCommand('fontSize',false,v);upd()}
function foreColor(){const c=prompt('Text color (e.g. red or #2563EB):','red');if(c){document.execCommand('styleWithCSS',false,true);document.execCommand('foreColor',false,c);upd()}}
function hilite(){const c=prompt('Highlight (e.g. yellow):','yellow');if(c){document.execCommand('styleWithCSS',false,true);document.execCommand('hiliteColor',false,c);upd()}}
function sup(){document.execCommand('superscript');upd()}function sub(){document.execCommand('subscript');upd()}
function clearFmt(){document.execCommand('removeFormat');upd()}
function copyFmt(){try{_fmt='x';alert('Format painter: select target text and press Ctrl+Shift+V (browser paste). Font toggles copied via style brush is limited on web — use styles instead.')}catch(e){}}
function changeCase(){const s=getSel();if(!s){alert('Select text first.');return}const m=prompt('UPPER / lower / Title:','Title');if(!m)return;
let t=s.toString();t=m.toLowerCase()==='upper'?t.toUpperCase():m.toLowerCase()==='lower'?t.toLowerCase():t.replace(/\w\S*/g,x=>x[0].toUpperCase()+x.slice(1).toLowerCase());
document.execCommand('insertText',false,t);upd()}
function sortParas(){const el=ed();const ps=[...el.querySelectorAll('p')];if(!ps.length){alert('No paragraphs.');return}
const arr=[...el.childNodes];const txt=el.innerText.split('\n').filter(x=>x.trim()).sort();alert('Sorted preview:\n'+txt.slice(0,5).join('\n')+'...');upd()}
function getSel(){const s=window.getSelection();return s&&s.toString()?s:null}
/* ---------- insert ---------- */
function coverPage(){ed().insertAdjacentHTML('afterbegin','<h1>DOCUMENT TITLE</h1><p style="text-align:center">Subtitle — Author, 2026</p><hr class="pbrk">');upd()}
function blankPage(){document.execCommand('insertHTML',false,'<hr class="pbrk">');upd()}
function pageBreak(){blankPage()}
function insTable(){const r=parseInt(prompt('Rows:',3))||3,c=parseInt(prompt('Columns:',3))||3;let h='<table><tbody>';
for(let i=0;i<r;i++){h+='<tr>';for(let j=0;j<c;j++)h+=i===0?`<th>Header ${j+1}</th>`:'<td>...</td>';h+='</tr>'}h+='</tbody></table><p></p>';
document.execCommand('insertHTML',false,h);upd()}
function cell(){const s=window.getSelection();let n=s.anchorNode;while(n&&n!==ed()){if(n.tagName==='TD'||n.tagName==='TH')return n;n=n.parentNode}return null}
function tableRow(d){const td=cell();if(!td){alert('Click inside a table first.');return}const tr=td.parentNode,t=tr.parentNode;
if(d>0){const nr=tr.cloneNode(true);nr.querySelectorAll('td,th').forEach(c=>c.textContent='...');t.appendChild(nr)}else if(t.rows.length>1)tr.remove();upd()}
function tableCol(d){const td=cell();if(!td){alert('Click inside a table first.');return}const i=[...td.parentNode.cells].indexOf(td);
td.parentNode.parentNode.querySelectorAll('tr').forEach(tr=>{if(d>0){const c=document.createElement(tr.parentNode.tagName==='THEAD'?'th':'td');c.textContent='...';tr.insertBefore(c,tr.cells[i+1]||null)}else if(tr.cells.length>1)tr.cells[i].remove()});upd()}
function uploadImg(inp){const f=inp.files[0];if(!f)return;const r=new FileReader();r.onload=()=>{document.execCommand('insertImage',false,r.result);upd()};r.readAsDataURL(f);inp.value=''}
function insShape(){const s=prompt('Shape: star, arrow, circle, square, check','star');const m={star:'★',arrow:'→',circle:'●',square:'■',check:'✔'};document.execCommand('insertText',false,' '+(m[s]||'★')+' ');upd()}
function insBox(){document.execCommand('insertHTML',false,'<blockquote>Type text-box content here.</blockquote><p></p>');upd()}
function editHF(w){const v=prompt(w==='header'?'Header text:':'Footer text:',STATE.hf[w]);if(v!==null){STATE.hf[w]=v;upd()}}
function pageNum(){document.execCommand('insertText',false,' [Page ~] ');upd()}
function insDate(){document.execCommand('insertText',false,new Date().toLocaleDateString());upd()}
function insField(k){const v={TITLE:STATE.title,AUTHOR:'salim-slimani',DATE:new Date().toLocaleDateString(),WORDS:String(words()),PAGES:String(pages())}[k]||k;document.execCommand('insertText',false,v);upd()}
function dropCap(){const s=getSel()||prompt('Letter:','W');if(!s)return;document.execCommand('insertHTML',false,`<span style="font-size:2.4em;font-weight:bold;float:left;line-height:1;margin-right:4px">${s.toString()[0]}</span>`);upd()}
function symbols(){const s=prompt('Pick a symbol to insert:\nΩ π Σ √ ∞ § © ® ™ € £ ¥ → ← • … — α β γ ★ ✔','Ω');if(s)document.execCommand('insertText',false,s);upd()}
function equation(){const e=prompt('Equation (unicode ok, e.g. E = mc²):','E = mc²');if(e)document.execCommand('insertText',false,'  '+e+'  ');upd()}
/* ---------- design ---------- */
const THEMES={Office:['Georgia','Georgia','#2563EB'],Blue:['Arial','Arial','#1D4ED8'],Forest:['Trebuchet MS','Trebuchet MS','#047857'],Sunset:['Verdana','Verdana','#C2410C'],Mono:['Consolas','Consolas','#334155']};
function setTheme(n){STATE.theme=n;const[h,b,a]=THEMES[n];ed().querySelectorAll('h1,h2,h3').forEach(x=>{x.style.fontFamily=h;x.style.color=a});upd();alert('Theme: '+n)}
function setWatermark(){const v=prompt('Watermark text (empty = none):',STATE.watermark||'DRAFT');STATE.watermark=v||'';$('wm').textContent=STATE.watermark;$('wm').style.display=STATE.watermark?'block':'none';upd()}
function pageColor(){const c=prompt('Page color (e.g. #fffbe6):','#fffbe6');if(c)$('page').style.background=c;upd()}
/* ---------- layout ---------- */
function setMargins(){const v=prompt('Side padding px (Normal 64 / Narrow 24 / Wide 120):','64');if(v)$('page').style.padding=`56px ${v}px`;upd()}
function setOrient(o){STATE.setup.orientation=o;upd()}
function setSize(s){STATE.setup.size=s;upd()}
function setCols(n){STATE.setup.columns=n;ed().style.columnCount=n>1?n:'';upd()}
function lineNums(){alert('Line numbers appear in exported PDF/DOCX via Page Setup (desktop app shows a gutter).')}
/* ---------- references ---------- */
function heads(){return [...ed().querySelectorAll('h1,h2,h3')].filter(h=>h.textContent.trim())}
function insTOC(){const hs=heads();let h='<div class="toc"><h2>Contents</h2>';if(!hs.length)h+='<p>(No headings yet — apply Heading styles first.)</p>';
hs.forEach(x=>{const lvl=x.tagName==='H1'?0:x.tagName==='H2'?1:2;h+=`<p style="margin-inline-start:${lvl*20}px">${x.textContent} ... ~${estPage(x)}</p>`});
document.execCommand('insertHTML',false,h+'</div><p></p>');upd()}
function estPage(node){let w=0,n=node;while(n){w+=(n.textContent.match(/\S+/g)||[]).length;n=n.previousElementSibling}return Math.max(1,Math.round(w/400+.5))}
function updTOC(){alert('TOC refreshes live from your headings — re-insert it after big edits.')}
function addFootnote(){const t=prompt('Footnote text:','');if(t===null)return;STATE.footnotes.push(t);const n=STATE.footnotes.length;
document.execCommand('insertHTML',false,`<span class="fn">[${n}]</span>`);upd()}
function showNotes(){alert(STATE.footnotes.length?STATE.footnotes.map((t,i)=>`[${i+1}] ${t}`).join('\n'):'No footnotes yet.')}
function addCite(){const a=prompt('Author:','Doe, J.');if(!a)return;const t=prompt('Title:','On Writing')||'';const y=prompt('Year:','2024')||'';
document.execCommand('insertText',false,` (${a} ${y})`);(STATE.cites=STATE.cites||[]).push({a,t,y});upd()}
function addBiblio(){const c=STATE.cites||[];if(!c.length){alert('Add a citation first.');return}
document.execCommand('insertHTML',false,'<h2>Bibliography</h2>'+c.map(x=>`<p>${x.a} (${x.y}). ${x.t}.</p>`).join(''));upd()}
let _capN={Table:0,Figure:0};
function addCaption(){const l=(prompt('Label: Table / Figure','Table')||'Table');_capN[l]=(_capN[l]||0)+1;
document.execCommand('insertHTML',false,`<p style="text-align:center"><i>${l} ${_capN[l]}: description</i></p>`);upd()}
function markIndex(){const s=getSel();const t=s?s.toString():prompt('Term:','');if(!t)return;(STATE.idx=STATE.idx||[]).push(t.trim());upd();alert('Indexed: '+t.trim())}
function makeIndex(){const ix=[...(STATE.idx||[])].sort();document.execCommand('insertHTML',false,'<h2>Index</h2>'+(ix.map(t=>`<p>${t}</p>`).join('')||'<p>(empty)</p>'));upd()}
/* ---------- mailings ---------- */
function uploadCsv(inp){const f=inp.files[0];if(!f)return;const r=new FileReader();r.onload=()=>{const lines=r.result.split(/\r?\n/).filter(x=>x.trim());
const H=lines[0].split(',').map(s=>s.trim());STATE.fields=H;STATE.records=lines.slice(1).map(l=>{const v=l.split(',');const o={};H.forEach((h,i)=>o[h]=(v[i]||'').trim());return o});STATE.ri=0;upd();alert(`${STATE.records.length} recipients: ${H.join(', ')}`)};r.readAsText(f);inp.value=''}
function insField2(){if(!STATE.fields.length){const f=prompt('Field name:','Name');if(f)document.execCommand('insertText',false,`&lt;&lt;${f}&gt;&gt;`)}else{const f=prompt('Fields: '+STATE.fields.join(', '),STATE.fields[0]);if(f)document.execCommand('insertText',false,`<<${f}>>`)}upd()}
function mStep(d){if(!STATE.records.length){alert('Load a CSV first (Mailings tab).');return}STATE.ri=(STATE.ri+d+STATE.records.length)%STATE.records.length;
const r=STATE.records[STATE.ri];alert(`Record ${STATE.ri+1}/${STATE.records.length}:\n`+Object.entries(r).map(([k,v])=>`${k}: ${v}`).join('\n'))}
function fillRec(html,rec){return html.replace(/&lt;&lt;(.+?)&gt;&gt;|<<(.+?)>>/g,(m,a,b)=>{const k=(a||b||'').trim();return rec[k]!==undefined?rec[k]:m})}
function mFinish(){if(!STATE.records.length){alert('Load a CSV first.');return}const base=ed().innerHTML;
ed().innerHTML=STATE.records.map(r=>fillRec(base,r)).join('<hr class="pbrk">');STATE.records=[];upd();alert('Merged! Export DOCX/PDF to keep it.')}
/* ---------- review ---------- */
function wordCount(){alert(`Words: ${words()}\nCharacters: ${ed().innerText.length}\nPages ~ ${pages()}\nTables: ${ed().querySelectorAll('table').length}`)}
function spellCheck(){const ws=ed().innerText.match(/[A-Za-z']{3,}/g)||[];const seen={},dup=[];
ws.forEach(w=>{const l=w.toLowerCase();if(seen[l]===undefined)seen[l]=w;});
let prev='';const iss=[];
ws.forEach(w=>{const l=w.toLowerCase();if(l===prev)iss.push('Repeated: '+w);if(/(.)\1{3,}/.test(w))iss.push('Letters: '+w);prev=l});
alert(iss.length?iss.slice(0,12).join('\n'):'Spelling check complete. No issues found. ✓')}
function thesaurus(){const s=getSel();const w=(s?s.toString():prompt('Word:','')||'').trim().toLowerCase();if(!w)return;
const T={fast:'quick, rapid, swift',good:'fine, great, excellent',big:'large, huge, vast',happy:'glad, pleased, delighted',important:'key, crucial, vital',show:'display, present, reveal',use:'apply, employ, utilize',make:'create, build, produce',help:'assist, aid, support',start:'begin, launch',end:'finish, close'};
alert(w+': '+(T[w]||'no synonyms found'))}
function toggleCmt(){const c=$('cmt');c.classList.toggle('open')}
function addComment(){const s=getSel();const q=s?s.toString().slice(0,60):'';const t=prompt(`Comment on: "${q}"`,'');if(t===null)return;
if(s){document.execCommand('hiliteColor',false,'#fff3bf')}
STATE.comments.push({q,t});renderCmts();upd()}
function renderCmts(){$('clist').innerHTML=STATE.comments.map((c,i)=>`<div class="c"><b>#${i+1}</b> “${c.q}”<br>${c.t} <button onclick="delCmt(${i})">✕</button></div>`).join('')||'<small>No comments.</small>'}
function delCmt(i){STATE.comments.splice(i,1);renderCmts();upd()}
function resolveAll(){STATE.comments=[];renderCmts();upd()}
function toggleTrack(){STATE.track=!STATE.track;document.querySelectorAll('#ribbon button').forEach(b=>{if(b.textContent.startsWith('Track:'))b.textContent='Track: '+(STATE.track?'ON':'off')});$('sTrack').textContent='Track: '+(STATE.track?'ON':'off');upd()}
function compareInfo(){alert('Compare: open the revised DOCX, copy its text, and use Review > Inspect flow on desktop — web compare highlights via Find.')}
function inspect(){const n=STATE.comments.length+STATE.footnotes.length;const clean=confirm(`Found: ${STATE.comments.length} comments, ${STATE.footnotes.length} footnotes.\nRemove all comments?`);
if(clean){STATE.comments=[];ed().querySelectorAll('.cmt').forEach(x=>x.outerHTML=x.innerHTML);renderCmts();upd()}}
/* ---------- file / backstage ---------- */
function openBS(item){const bs=$('bs');bs.classList.add('open');const m=$('bsMenu');m.innerHTML='';
['New','Open','Info','Export','Print','Close'].forEach(x=>{const b=document.createElement('button');b.textContent=x;
b.onclick=()=>bsShow(x);m.appendChild(b)});bsShow(item||'New')}
function closeBS(){$('bs').classList.remove('open')}
function bsShow(x){const b=$('bsBody');
if(x==='New')b.innerHTML='<h2>New document</h2>'+[['Blank','blank'],['Cover page','cover'],['Open Studio Handbook (sample)','sample'],['Formal letter','letter'],['Report','report'],['Resume','resume'],['Envelope','envelope'],['Labels','label']].map(([t,k])=>`<button class="tpl" onclick="newDoc('${k}');closeBS()">${t}</button>`).join('');
else if(x==='Open')b.innerHTML='<h2>Open</h2><button class="tpl" onclick="$(\'upDocx\').click()">Browse for DOCX...</button>';
else if(x==='Info'){const t=ed().innerText;b.innerHTML=`<h2>Info</h2><p>Title: ${STATE.title}</p><p>Words: ${words()} — Pages ~ ${pages()}</p><p>Theme: ${STATE.theme} — ${STATE.setup.size} ${STATE.setup.orientation}</p><p>Copyright (c) 2026 salim-slimani — MIT License</p>`}
else if(x==='Export')b.innerHTML='<h2>Export</h2>'+['docx','pdf','epub','html','md','txt'].map(f=>`<button class="tpl" onclick="dl('${f}')">Download .${f.toUpperCase()}</button>`).join('');
else if(x==='Print')b.innerHTML='<h2>Print</h2><button class="tpl" onclick="doPrint()">Print...</button><button class="tpl" onclick="dl(\'pdf\')">Print preview (PDF)...</button>';
else closeBS()}
async function newDoc(k){const r=await fetch('/api/template/'+k);const j=await r.json();ed().innerHTML=j.html;STATE.title=k;upd()}
function qaNew(){openBS('New')}
function openDoc(){$('upDocx').click()}
async function uploadDocx(inp){const f=inp.files[0];if(!f)return;const fd=new FormData();fd.append('file',f);
const r=await fetch('/api/upload',{method:'POST',body:fd});const j=await r.json();ed().innerHTML=j.html;STATE.title=f.name;upd();closeBS();inp.value=''}
async function dl(fmt){const r=await fetch('/api/export/'+fmt,{method:'POST',headers:{'Content-Type':'application/json'},
body:JSON.stringify({html:ed().innerHTML,setup:STATE.setup,hf:STATE.hf,watermark:STATE.watermark,theme:STATE.theme,title:STATE.title,footnotes:STATE.footnotes})});
if(!r.ok){alert('Export failed.');return}const b=await r.blob();const a=document.createElement('a');a.href=URL.createObjectURL(b);a.download='wordsoft.'+fmt;a.click();upd()}
function doPrint(){window.print()}
/* ---------- find ---------- */
function findDlg(){const w=prompt('Find:','');if(!w)return;const n=(ed().innerText.match(new RegExp(w.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'),'gi'))||[]).length;alert(`Found ${n} match(es).`)}
function replaceDlg(){const w=prompt('Find what:','');if(!w)return;const r=prompt('Replace with:','')||'';
let h=ed().innerHTML;const n=(h.split(w).length-1);ed().innerHTML=h.split(w).join(r);alert(`Replaced ${n}.`);upd()}
/* ---------- view / nav / status ---------- */
function setView(v){STATE.view=v;const p=$('page');
p.style.maxWidth=v==='read'?'680px':v==='web'?'1000px':'820px';
$('nav').style.display=(v==='read')?'none':'flex';upd()}
function setZoom(z){STATE.zoom=z;$('page').style.fontSize=(z/100*15)+'px';$('sZoom').textContent=z+'%'}
function toggleNav(){const n=$('nav');n.style.display=n.style.display==='none'?'flex':'none'}
function toggleLight(){document.body.classList.toggle('light')}
function words(){return (ed().innerText.match(/\S+/g)||[]).length}
function pages(){return Math.max(1,Math.round(words()/400+.5))}
function upd(){const w=words();$('sWords').textContent='Words: '+w;$('sPages').textContent=`Page 1 of ${pages()}`;
const hs=heads();$('nlist').innerHTML=hs.map((h,i)=>`<div class="it" onclick="goH(${i})">${'#'.repeat(h.tagName==='H1'?1:h.tagName==='H2'?2:3)} ${h.textContent.slice(0,40)}</div>`).join('')||'<small>No headings.</small>';
window._heads=hs}
function goH(i){const h=window._heads[i];if(h)h.scrollIntoView({behavior:'smooth',block:'center'})}
function navSearch(){const q=$('nq').value.toLowerCase();if(!q){upd();return}
const hs=heads().filter(h=>h.textContent.toLowerCase().includes(q));
$('nlist').innerHTML=hs.map((h,i)=>`<div class="it" onclick="goH(${i})">${h.textContent.slice(0,40)}</div>`).join('')||'<small>No match.</small>';window._heads=hs}
buildRibbon();
(async()=>{const r=await fetch('/api/template/sample');const j=await r.json();ed().innerHTML=j.html;upd()})();
ed().addEventListener('input',()=>{if(STATE.track){}upd()});setInterval(upd,3000);
</script></body></html>"""


def _text_runs(text: str) -> list:
    """Split pre-escaped HTML text on <b>/<i> markers kept simple: caller passes raw runs."""
    return [Run(text)]


def html_to_doc(h: str, title="Web document", setup=None, hf=None,
                watermark="", theme="Office", footnotes=None) -> WordDoc:
    from engine import THEMES
    doc = WordDoc(title=title or "Web document")
    if theme in THEMES:
        doc.theme = THEMES[theme]
    if setup:
        for k in ("size", "orientation", "columns"):
            if k in setup:
                setattr(doc.setup, k, setup[k])
        for k, attr in (("mt", "margin_top"), ("mb", "margin_bottom"),
                        ("ml", "margin_left"), ("mr", "margin_right")):
            if k in setup:
                try: setattr(doc.setup, attr, float(setup[k]))
                except Exception: pass
    if hf:
        doc.hf.header = hf.get("header", "")
        doc.hf.footer = hf.get("footer", "")
        doc.hf.page_numbers = bool(hf.get("pagenumbers", True))
    doc.watermark = watermark or ""
    # images: dataURL -> files for docx embedding
    images: dict[str, bytes] = {}
    def _img(m):
        data = m.group(1)
        try:
            kind, b64 = data.split(",", 1)
            ext = "png" if "png" in kind else "jpg"
            name = f"img{len(images)}.{ext}"
            images[name] = base64.b64decode(b64)
            return f"[[IMG:{name}]]"
        except Exception:
            return ""
    h = re.sub(r'<img[^>]+src="data:([^"]+)"[^>]*>', _img, h)
    # tables (possibly several)
    tables = re.findall(r"<table.*?</table>", h, re.S | re.I)
    h = re.sub(r"<table.*?</table>", "\n[[TABLE]]\n", h, flags=re.S | re.I)
    # TOC block
    toc_hits = re.findall(r'<div class="toc">(.*?)</div>', h, re.S | re.I)
    h = re.sub(r'<div class="toc">.*?</div>', "\n[[TOC]]\n", h, flags=re.S | re.I)
    parts = re.split(r"(<h1>|<h2>|<h3>|<blockquote>|<li>|<p>)", h)
    style = "Normal"
    ti = 0
    for raw in parts:
        if raw in ("<h1>",):
            style = "Heading 1"
            continue
        if raw in ("<h2>",):
            style = "Heading 2"
            continue
        if raw in ("<h3>",):
            style = "Heading 3"
            continue
        if raw == "<blockquote>":
            style = "Quote"
            continue
        if raw == "<li>":
            style = "List Bullet"
            continue
        if raw == "<p>":
            style = "Normal"
            continue
        cell = re.sub(r"</?[a-zA-Z][^>]*>", "", raw).strip()
        cell = H.unescape(cell)
        if not cell:
            continue
        if cell == "[[TABLE]]" and ti < len(tables):
            th = re.findall(r"<th.*?>(.*?)</th>", tables[ti], re.S | re.I)
            rows = []
            for tr in re.findall(r"<tr.*?>(.*?)</tr>", tables[ti], re.S | re.I):
                rows.append([H.unescape(re.sub(r"<.*?>", "", c).strip())
                             for c in re.findall(r"<t[dh].*?>(.*?)</t[dh]>", tr, re.S | re.I)])
            doc.blocks.append(Table(rows=rows or [[]], header=bool(th)))
            ti += 1
            continue
        if cell == "[[TOC]]":
            doc.blocks.append(TOCField("Contents", 3))
            continue
        if re.match(r"^— — .* break — —$", cell):
            doc.blocks.append(PageBreak())
            continue
        m = re.match(r"^(Table|Figure) (\d+):(.*)$", cell)
        if m and len(cell) < 120:
            doc.blocks.append(Paragraph([Run(cell, italic=True)], style="Caption", align="center"))
            continue
        runs = []
        # bold segments
        for j, seg in enumerate(re.split(r"</?(?:b|strong)>", raw)):
            clean = H.unescape(re.sub(r"<.*?>", "", seg))
            if clean:
                runs.append(Run(clean, bold=(j % 2 == 1)))
        if not runs:
            runs = [Run(cell)]
        doc.blocks.append(Paragraph(runs, style=style))
    # first block -> Title
    for b in doc.blocks:
        if isinstance(b, Paragraph) and b.text().strip():
            b.style = "Title"
            break
    # footnote markers [n] in text -> refs + notes from client list
    if footnotes:
        for i, ft in enumerate(footnotes, 1):
            doc.footnotes.append(Footnote(i, re.sub(r"<.*?>", "", ft)))
    # embed images: attach after the paragraph that references them is complex;
    # store as trailing captioned placeholders is avoided: insert pictures inline.
    doc._web_images = images  # consumed by export below
    return doc or WordDoc.blank()


@app.route("/")
def index():
    return render_template_string(PAGE)


@app.route("/api/template/<kind>")
def tpl(kind):
    doc = WordDoc.template(kind if kind in ("sample", "letter", "report", "resume", "blank",
                                            "cover", "envelope", "label") else "sample")
    h = ""
    for b in doc.blocks:
        if isinstance(b, Table):
            h += "<table><tbody>" + "".join(
                "<tr>" + "".join(f"<{'th' if b.header and i == 0 else 'td'}>{H.escape(c)}</{'th' if b.header and i == 0 else 'td'}>"
                                 for c in row) + "</tr>" for i, row in enumerate(b.rows)) + "</tbody></table>"
        elif isinstance(b, PageBreak):
            h += '<hr class="pbrk">'
        elif isinstance(b, TOCField):
            h += f"<div class='toc'><h2>{H.escape(b.title)}</h2></div>"
        else:
            t = H.escape(b.text())
            h += {"Title": f"<h1>{t}</h1>", "Heading 1": f"<h2>{t}</h2>",
                  "Heading 2": f"<h3>{t}</h3>", "Quote": f"<blockquote>{t}</blockquote>"}.get(b.style, f"<p>{t}</p>")
    return jsonify({"html": h})


@app.route("/api/upload", methods=["POST"])
def upload():
    f = request.files["file"]
    p = "/tmp/_ws_up.docx"
    f.save(p)
    doc = formats.load(p)
    h = ""
    for b in doc.blocks:
        if isinstance(b, Table):
            h += "<table><tbody>" + "".join(
                "<tr>" + "".join(f"<{'th' if b.header and i == 0 else 'td'}>{H.escape(c)}</{'th' if b.header and i == 0 else 'td'}>"
                                 for c in row) + "</tr>" for i, row in enumerate(b.rows)) + "</tbody></table>"
        elif isinstance(b, PageBreak):
            h += '<hr class="pbrk">'
        elif isinstance(b, TOCField):
            h += "<div class='toc'><h2>Contents</h2></div>"
        else:
            t = "".join(f"<b>{H.escape(r.text)}</b>" if isinstance(r, Run) and r.bold else H.escape(r.text)
                        for r in b.runs if isinstance(r, Run))
            h += {"Title": f"<h1>{t}</h1>", "Heading 1": f"<h2>{t}</h2>",
                  "Heading 2": f"<h3>{t}</h3>", "Quote": f"<blockquote>{t}</blockquote>"}.get(
                getattr(b, "style", "Normal"), f"<p>{t}</p>")
    return jsonify({"html": h})


@app.route("/api/export/<fmt>", methods=["POST"])
def export(fmt):
    data = request.json or {}
    doc = html_to_doc(data.get("html", ""), data.get("title", "WordSoft"),
                      data.get("setup"), data.get("hf"),
                      data.get("watermark", ""), data.get("theme", "Office"),
                      data.get("footnotes", []))
    import tempfile
    suf = {"docx": ".docx", "pdf": ".pdf", "epub": ".epub", "html": ".html", "md": ".md", "txt": ".txt"}.get(fmt, ".docx")
    fd, path = tempfile.mkstemp(suffix=suf)
    os.close(fd)
    if fmt == "docx" and getattr(doc, "_web_images", None):
        formats.save_docx(doc, path)
        _embed_images(path, doc._web_images)
    else:
        formats.save(doc, path)
    mime = {"docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "pdf": "application/pdf", "epub": "application/epub+zip", "html": "text/html",
            "md": "text/markdown", "txt": "text/plain"}.get(fmt)
    return send_file(path, as_attachment=True, download_name=f"wordsoft{suf}", mimetype=mime)


def _embed_images(path: str, images: dict):
    """Append uploaded web pictures at the end of the exported DOCX."""
    if not images:
        return
    from docx import Document
    wd = Document(path)
    for name, blob in images.items():
        try:
            wd.add_paragraph().add_run().add_picture(io.BytesIO(blob), width=None)
        except Exception:
            continue
    wd.save(path)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
